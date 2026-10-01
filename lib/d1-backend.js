// Cloudflare D1(DB)로 Code.gs를 실행하는 서버
//
// 방식: 요청마다 DB의 시트 데이터를 메모리로 읽어 '스프레드시트처럼 보이는 객체'를 만들고,
//       Code.gs(gas-core.js)의 doPost를 그대로 실행한 뒤, 바뀐 줄만 DB에 다시 쓴다.
//       → 권한·검증 로직은 Apps Script와 100% 같은 코드. 저장소만 스프레드시트 → DB로 바뀐다.
//
// DB 구조 (시트 구조를 그대로 옮김)
//   rows(sheet, k, pos, data)  — sheet: 시트 이름, k: 첫 칸 값(id·email·date, 제목 줄은 '#header'),
//                                pos: 줄 순서, data: 그 줄의 값 배열(JSON)
//   kv(k, v, exp)              — CacheService(로그인 실패 횟수, 요청 ID) · PropertiesService(TOKEN_SECRET) 대용
import { createHash, createHmac, randomUUID } from 'node:crypto';
import { Buffer } from 'node:buffer';
import { createGas } from './gas-core.js';

export const SCHEMA = [
  'CREATE TABLE IF NOT EXISTS rows (sheet TEXT NOT NULL, k TEXT NOT NULL, pos REAL NOT NULL, data TEXT NOT NULL, PRIMARY KEY (sheet, k))',
  'CREATE TABLE IF NOT EXISTS kv (k TEXT PRIMARY KEY, v TEXT NOT NULL, exp INTEGER)',
];

const HEADER = '#header';
const rowKey = (row, i) => (i === 0 ? HEADER : String(row[0]));

/** 스프레드시트 시트 흉내 (Code.gs가 쓰는 메서드만) */
class MemSheet {
  constructor(rows) { this.rows = rows; }
  getMaxRows() { return 1000; }
  getLastRow() { return this.rows.length; }
  setFrozenRows() {}
  deleteRow(r) { this.rows.splice(r - 1, 1); }
  getRange(r, c, nr = 1, nc = 1) {
    const sh = this;
    return {
      getValue: () => (sh.rows[r - 1] || [])[c - 1] ?? '',
      getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => (sh.rows[r - 1 + i] || [])[c - 1 + j] ?? '')),
      setValues(v) {
        v.forEach((row, i) => {
          while (sh.rows.length < r + i) sh.rows.push([]);
          const cur = sh.rows[r + i - 1].slice();
          row.forEach((val, j) => (cur[c - 1 + j] = val));
          sh.rows[r + i - 1] = cur;
        });
        return this;
      },
      setNumberFormat() { return this; },
      setFontWeight() { return this; },
    };
  }
}

/** Asia/Seoul 날짜 (Code.gs의 Utilities.formatDate(..., 'Asia/Seoul', 'yyyy-MM-dd') 용) */
const kstDate = (d) => new Date(d.getTime() + 9 * 3600e3).toISOString().slice(0, 10);
const bytes = (b) => Array.from(b);

function services(sheets, kv, kvChanges) {
  const book = {
    getSheetByName: (n) => (sheets[n] ? sheets[n].sheet : null),
    insertSheet: (n) => { sheets[n] = { sheet: new MemSheet([]), before: new Map(), maxPos: 0 }; return sheets[n].sheet; },
  };
  const kvGet = (k) => {
    const e = kv.get(k);
    return e && (!e.exp || e.exp > Date.now()) ? e.v : null;
  };
  const kvSet = (k, v, ttlSec) => {
    const e = { v: String(v), exp: ttlSec ? Date.now() + ttlSec * 1000 : null };
    kv.set(k, e); kvChanges.set(k, e);
  };
  const kvDel = (k) => { kv.delete(k); kvChanges.set(k, null); };
  const html = { setTitle() { return this; }, setXFrameOptionsMode() { return this; }, addMetaTag() { return this; } };
  return {
    SpreadsheetApp: { getActiveSpreadsheet: () => book, openById: () => book },
    Utilities: {
      getUuid: () => randomUUID(),
      DigestAlgorithm: { SHA_256: 'sha256' },
      computeDigest: (a, s) => bytes(createHash('sha256').update(s).digest()),
      computeHmacSha256Signature: (v, k) => bytes(createHmac('sha256', k).update(v).digest()),
      base64Encode: (b) => Buffer.from(b).toString('base64'),
      base64EncodeWebSafe: (b) => Buffer.from(typeof b === 'string' ? Buffer.from(b, 'utf8') : b).toString('base64').replace(/\+/g, '-').replace(/\//g, '_'),
      base64DecodeWebSafe: (s) => bytes(Buffer.from(s.replace(/-/g, '+').replace(/_/g, '/'), 'base64')),
      newBlob: (b) => ({ getDataAsString: () => Buffer.from(b).toString('utf8') }),
      formatDate: (d) => kstDate(d),
    },
    CacheService: { getScriptCache: () => ({ get: (k) => kvGet('c:' + k), put: (k, v, ttl) => kvSet('c:' + k, v, ttl || 600), remove: (k) => kvDel('c:' + k) }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => kvGet('p:' + k), setProperty: (k, v) => kvSet('p:' + k, v, 0) }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) }, // 요청 하나 안에서는 순차 실행
    ScriptApp: { getService: () => ({ getUrl: () => '' }) }, // 화면이 자기 주소를 안내 메시지에 넣는다
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: (t) => ({ t, setMimeType() { return this; } }) },
    HtmlService: { XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' }, createHtmlOutput: () => html, createHtmlOutputFromFile: () => html },
    Session: { getActiveUser: () => ({ getEmail: () => '' }) },
  };
}

/** DB → 메모리 */
async function load(db) {
  const [rowsRes, kvRes] = await db.batch([
    db.prepare('SELECT sheet, k, pos, data FROM rows ORDER BY sheet, pos'),
    db.prepare('SELECT k, v, exp FROM kv'),
  ]);
  const sheets = {};
  for (const r of rowsRes.results) {
    const s = (sheets[r.sheet] = sheets[r.sheet] || { rows: [], before: new Map(), maxPos: 0 });
    s.rows.push(JSON.parse(r.data));
    s.before.set(r.k, { pos: r.pos, data: r.data });
    s.maxPos = Math.max(s.maxPos, r.pos);
  }
  for (const name of Object.keys(sheets)) sheets[name].sheet = new MemSheet(sheets[name].rows);
  const kv = new Map(kvRes.results.map((r) => [r.k, { v: r.v, exp: r.exp }]));
  return { sheets, kv };
}

/** 메모리 → DB (바뀐 줄만) */
function changes(db, sheets, kvChanges) {
  const stmts = [];
  for (const [name, s] of Object.entries(sheets)) {
    const seen = new Set();
    let next = s.maxPos;
    s.sheet.rows.forEach((row, i) => {
      if (!row.length) return;
      const k = rowKey(row, i);
      if (i > 0 && k === '') return; // 빈 줄은 저장하지 않음 (Code.gs도 무시)
      if (seen.has(k)) return; // 같은 키가 두 줄이면 첫 줄만 (Code.gs도 첫 줄만 찾는다)
      seen.add(k);
      const data = JSON.stringify(row.map((v) => (v == null ? '' : v)));
      const old = s.before.get(k);
      if (old && old.data === data) return;
      const pos = old ? old.pos : (i === 0 ? 0 : ++next);
      stmts.push(db.prepare('INSERT INTO rows (sheet, k, pos, data) VALUES (?1, ?2, ?3, ?4) ON CONFLICT(sheet, k) DO UPDATE SET data = excluded.data').bind(name, k, pos, data));
    });
    for (const k of s.before.keys()) {
      if (!seen.has(k)) stmts.push(db.prepare('DELETE FROM rows WHERE sheet = ?1 AND k = ?2').bind(name, k));
    }
  }
  for (const [k, e] of kvChanges) {
    stmts.push(e
      ? db.prepare('INSERT INTO kv (k, v, exp) VALUES (?1, ?2, ?3) ON CONFLICT(k) DO UPDATE SET v = excluded.v, exp = excluded.exp').bind(k, e.v, e.exp)
      : db.prepare('DELETE FROM kv WHERE k = ?1').bind(k));
  }
  return stmts;
}

/**
 * 요청 하나 처리: body = '{"fn": "...", "args": [...], "rid": "..."}' → 응답 JSON 문자열
 * (Code.gs의 doPost를 그대로 실행하므로 허용 함수 목록·요청 ID 중복 방지·오류 형식이 Apps Script와 같다)
 */
let schemaReady = false; // 서버(isolate)마다 처음 한 번만 표 확인
export async function handleD1(db, body) {
  if (!schemaReady) { await db.batch(SCHEMA.map((sql) => db.prepare(sql))); schemaReady = true; }
  const { sheets, kv } = await load(db);
  const kvChanges = new Map();
  const gas = createGas(services(sheets, kv, kvChanges));
  const text = gas.doPost({ postData: { contents: body } }).t;
  const stmts = changes(db, sheets, kvChanges);
  if (stmts.length) await db.batch(stmts); // 한 번에 (전부 되거나 전부 안 되거나)
  // 만료된 캐시 정리 (가끔)
  if (Math.random() < 0.02) await db.prepare('DELETE FROM kv WHERE exp IS NOT NULL AND exp < ?1').bind(Date.now()).run();
  return text;
}

/** 스프레드시트에서 내보낸 데이터 → DB에 넣을 SQL 문장들 (이전·최종 동기화용) */
export function importStatements(exported) {
  const out = ['DELETE FROM rows', "DELETE FROM kv WHERE k LIKE 'p:%'"];
  const q = (v) => "'" + String(v).replace(/'/g, "''") + "'";
  for (const [name, rows] of Object.entries(exported.sheets)) {
    const seen = new Set();
    rows.forEach((row, i) => {
      if (!row.length || (i > 0 && String(row[0]) === '')) return;
      const k = rowKey(row, i);
      if (seen.has(k)) return;
      seen.add(k);
      out.push(`INSERT INTO rows (sheet, k, pos, data) VALUES (${q(name)}, ${q(k)}, ${i}, ${q(JSON.stringify(row))})`);
    });
  }
  for (const [k, v] of Object.entries(exported.props || {})) out.push(`INSERT OR REPLACE INTO kv (k, v, exp) VALUES (${q('p:' + k)}, ${q(v)}, NULL)`);
  return out;
}

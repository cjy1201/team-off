// 자동 생성 파일 — tools/gen-gas-core.mjs 가 Code.gs 로 만든다 (build.sh). 직접 고치지 마세요.
// Cloudflare(D1) 서버가 Apps Script와 똑같은 로직을 쓰도록 Code.gs 전체를 함수 하나로 감싼 것.
/* eslint-disable */
export function createGas({ SpreadsheetApp, Utilities, CacheService, PropertiesService, LockService, ScriptApp, ContentService, HtmlService, Session, TeamLog }) {
/**
 * 슈퍼쏠 전담반 일정관리 — Google Apps Script 백엔드
 *
 * - 웹 앱은 '배포한 사람(관리자)' 권한으로 실행된다. 스프레드시트는 팀원에게 공유하지 않는다.
 * - 사용자 식별은 앱 자체 로그인(이메일 + 비밀번호)으로 한다.
 * - 권한은 3단계. 화면에서 숨기는 게 아니라 서버가 볼 수 있는 일정만 내려준다.
 *     PM(관리자)  : 전체 일정 + 팀원/공휴일 관리
 *     PL(파트 리더): 전체 직원 일정 조회 (수정은 본인 일정만)
 *     파트원       : 같은 파트 직원 일정 조회 (수정은 본인 일정만)
 *   + '전체 일정 조회' 옵션(viewAll): 파트원도 전체 일정을 볼 수만 있다 (근태 보고 담당 등)
 * 중요 일정(전담반 공지성 일정)은 PM만 등록하고 전 직원이 본다.
 *
 * 데이터는 아래 스프레드시트의 '일정앱_팀원' / '일정앱_일정' / '일정앱_공휴일' 시트에 저장된다.
 * 시트가 없으면 처음 접속할 때 자동으로 만든다.
 * 같은 이름의 시트가 이미 있는데 이 앱의 형식이 아니면 건드리지 않고 오류를 낸다.
 */

// 비워두면 이 Apps Script가 붙어 있는 스프레드시트(확장 프로그램 → Apps Script로 연 시트)를 쓴다.
// 다른 스프레드시트를 쓰려면 그 주소의 /d/ 와 /edit 사이 ID를 넣는다.
const SHEET_ID = '';
const TZ = 'Asia/Seoul';
const TOKEN_DAYS = 30;
const REMEMBER_DAYS = 180; // '이 기기에서 로그인 유지' (홈 화면 추가용)
const MAX_LOGIN_FAILS = 5;

const SHEETS = {
  // annual: 연도별 휴가 부여 일수 JSON (예: {"2026":{"일반휴가":15,"대체휴가":1}}) — PM이 입력, 본인과 PM만 봄 (참고용)
  users: { name: '일정앱_팀원', headers: ['email', 'name', 'role', 'createdAt', 'passwordHash', 'salt', 'mustChange', 'part', 'defaultSubs', 'viewAll', 'annual'] },
  leaves: {
    name: '일정앱_일정',
    headers: ['id', 'email', 'type', 'startDate', 'endDate', 'substitute', 'memo', 'createdAt', 'updatedAt', 'updatedBy'],
  },
  holidays: { name: '일정앱_공휴일', headers: ['date', 'name'] },
  events: {
    name: '일정앱_중요일정',
    headers: ['id', 'title', 'startDate', 'endDate', 'time', 'memo', 'createdAt', 'updatedAt', 'updatedBy'],
  },
  // 파트별 업무 일정 (휴가 일정과 완전히 별도)
  tasks: {
    name: '일정앱_업무일정',
    headers: ['id', 'part', 'title', 'category', 'startDate', 'endDate', 'time', 'assignees', 'status', 'visibility', 'memo',
      'createdBy', 'createdAt', 'updatedAt', 'updatedBy'],
  },
  // 변경 이력 (PM '변경 이력' 탭). Cloudflare(DB) 서버는 매 요청마다 읽지 않도록 따로 보관한다 (historyStore_)
  history: { name: '일정앱_이력', headers: ['id', 'at', 'actor', 'kind', 'action', 'owner', 'summary'] },
  // 새 소식 — 사람마다 한 줄 (마지막으로 읽은 시각 + 최근 소식 목록 JSON)
  news: { name: '일정앱_소식', headers: ['email', 'seenAt', 'items'] },
};

const LEAVE_TYPES = ['연차', '오전반차', '오후반차', '병가', '연수/교육', '기타'];
const HALF_TYPES = ['오전반차', '오후반차'];
const WEEK = ['일', '월', '화', '수', '목', '금', '토'];
const PARTS = ['기획', 'QA', '개발', '퍼블', 'NP', '디자인'];
const ROLES = ['pm', 'pl', 'member'];
const MAX_SUBSTITUTES = 3;
const TASK_CATEGORIES = ['개발', '배포', '테스트', '회의/리뷰', '마감', '기타'];
const TASK_STATUSES = ['예정', '진행 중', '완료', '지연'];
const MAX_ASSIGNEES = 10;
const ANNUAL_KINDS = ['일반휴가', '특별휴가', '대체휴가']; // PM이 입력하는 부여 일수 구분 (사용은 연차·반차에서 합쳐서 차감)
const HISTORY_DAYS = 90; // 변경 이력·새 소식 보관 기간
const NEWS_MAX = 20; // 사람마다 보관하는 새 소식 개수
const ROLE_LABELS = { pm: 'PM', pl: 'PL', member: '파트원' };

// 관리자 화면 '공휴일' 탭에서 추가/삭제 가능
const DEFAULT_HOLIDAYS = [
  ['2026-01-01', '신정'],
  ['2026-02-16', '설날 연휴'], ['2026-02-17', '설날'], ['2026-02-18', '설날 연휴'],
  ['2026-03-02', '삼일절 대체공휴일'],
  ['2026-05-05', '어린이날'],
  ['2026-05-25', '부처님오신날 대체공휴일'],
  ['2026-06-03', '전국동시지방선거'],
  ['2026-08-17', '광복절 대체공휴일'],
  ['2026-09-24', '추석 연휴'], ['2026-09-25', '추석'], ['2026-09-26', '추석 연휴'],
  ['2026-10-05', '개천절 대체공휴일'],
  ['2026-10-09', '한글날'],
  ['2026-12-25', '성탄절'],
];

/* ───────────── Web App ───────────── */

function doGet(e) {
  // ?export=… : DB 전환용 데이터 내보내기 (Migration.gs가 있을 때만 동작, 전환 후 Migration.gs 삭제)
  if (e && e.parameter && e.parameter.export && typeof migrationExport_ === 'function') return migrationExport_(e);
  // ?ping=1 : 연결 확인용. 브라우저 주소창에 입력해서 {"ok":true,...} 가 보이면 fetch 방식(구글 응답 전달 경로)이 통하는 망
  if (e && e.parameter && e.parameter.ping === '1') {
    return ContentService.createTextOutput(JSON.stringify({ ok: true, pong: new Date().toISOString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // ?bridge=1 : GitHub Pages 화면이 숨겨서 띄우는 통신용 페이지 (아래 BRIDGE_HTML)
  if (e && e.parameter && e.parameter.bridge === '1') {
    return HtmlService.createHtmlOutput(BRIDGE_HTML)
      .setTitle('bridge')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // 다른 사이트에서 iframe으로 띄울 수 있게
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('슈퍼쏠 전담반 일정')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/*
 * 통신용 페이지. 외부 화면(GitHub Pages)이 보이지 않는 iframe으로 띄우고 postMessage로 요청을 보내면
 * google.script.run으로 API 함수를 실행해 결과를 돌려준다.
 * doPost(fetch) 방식은 서버는 빨리 끝나도 구글의 응답 전달(macros/echo)이 자주 404로 실패해서 이 방식을 기본으로 쓴다.
 * 주의: Apps Script가 HTML을 내보낼 때 JS 주석과 문자열 속 슬래시 두 개를 잘라내므로 아래 코드에는 주석·URL을 넣지 않는다.
 * 뜰 때 ping으로 실제 통신을 시험해서 되면 teamoff-ready, 안 되면 teamoff-broken 을 알린다 (회사망 등에서 google.script.run이 막히는 경우).
 * 허용하는 화면 주소: *.github.io, *.netlify.app, *.web.app, *.firebaseapp.com, *.vercel.app, localhost (그 외는 무시)
 */
const BRIDGE_HTML = `<!DOCTYPE html><html><body><script>
(function () {
  var ALLOWED = [/\\.github\\.io$/, /\\.netlify\\.app$/, /\\.web\\.app$/, /\\.firebaseapp\\.com$/, /\\.vercel\\.app$/, /^localhost$/, /^127\\.0\\.0\\.1$/];
  function allowed(origin) {
    try { var h = new URL(origin).hostname; return ALLOWED.some(function (r) { return r.test(h); }); } catch (e) { return false; }
  }
  window.addEventListener('message', function (ev) {
    var d = ev.data;
    if (!d || d.type !== 'teamoff-call' || !allowed(ev.origin)) return;
    function reply(x) { x.type = 'teamoff-reply'; x.id = d.id; ev.source.postMessage(x, ev.origin); }
    try {
      var runner = google.script.run
        .withSuccessHandler(function (r) { reply({ ok: true, data: r }); })
        .withFailureHandler(function (e) { reply({ ok: false, error: (e && e.message) || String(e) }); });
      if (typeof runner[d.fn] !== 'function') return reply({ ok: false, error: '알 수 없는 요청입니다.' });
      runner[d.fn].apply(runner, d.args || []);
    } catch (e) { reply({ ok: false, error: (e && e.message) || String(e) }); }
  });
  google.script.run
    .withSuccessHandler(function () { window.top.postMessage({ type: 'teamoff-ready' }, '*'); })
    .withFailureHandler(function (e) { window.top.postMessage({ type: 'teamoff-broken', error: String((e && e.message) || e) }, '*'); })
    .ping();
})();
</script></body></html>`;

/**
 * GitHub Pages 등 외부에 올린 화면에서 부르는 API 입구.
 * 요청: POST 본문 {"fn": "함수 이름", "args": [...]}  → 응답: {"ok": true, "data": ...} 또는 {"ok": false, "error": "..."}
 * google.script.run으로 부를 수 있던 함수와 같은 것만 허용한다.
 */
const API_FUNCTIONS = {
  getBootstrap, setupAdmin, login, changePassword, resetPassword,
  checkLeave, saveLeave, deleteLeave,
  saveUser, bulkAddUsers, deleteUser, setMyDefaultSubs,
  saveEvent, deleteEvent, saveHoliday, deleteHoliday,
  getTasks, saveTask, deleteTask, checkTaskAbsence,
  getBackupInfo, getHistory, markNewsRead, setAnnualDays, bulkAddLeaves,
};

function doPost(e) {
  const json = (text) => ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
  let req = {};
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { req = {}; }

  // 저장 요청 중복 방지: 구글 응답 전달(echo)이 응답을 잃어버리면 전달 서버가 같은 rid로 다시 보낸다.
  // 이미 처리한 rid면 다시 실행하지 않고 저장해 둔 결과를 돌려준다 (10분 보관).
  const rid = /^[A-Za-z0-9_-]{8,64}$/.test(String(req.rid || '')) ? 'rid:' + req.rid : '';
  const cache = CacheService.getScriptCache();
  if (rid) {
    const done = cache.get(rid);
    if (done) return json(done);
  }

  let out;
  try {
    const fn = Object.prototype.hasOwnProperty.call(API_FUNCTIONS, req.fn) ? API_FUNCTIONS[req.fn] : null;
    if (!fn) throw new Error('알 수 없는 요청입니다.');
    out = { ok: true, data: fn.apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    out = { ok: false, error: (err && err.message) || String(err) };
  }
  const text = JSON.stringify(out);
  if (rid) {
    // 캐시 한 칸은 100KB 제한 → 결과가 크면 '처리 완료'만 기록 (화면은 다시 불러오면 됨)
    cache.put(rid, text.length < 90000 ? text : JSON.stringify({ ok: false, error: 'SAVED:처리는 완료되었습니다. 화면을 새로 불러옵니다.' }), 600);
  }
  return json(text);
}

/** PM: 마지막 백업 시각 (DB → 스프레드시트 백업이 데이터를 가져간 시각. Cloudflare /api-backup 이 기록) */
function getBackupInfo(token) {
  requireAdmin_(token);
  return { lastBackupAt: PropertiesService.getScriptProperties().getProperty('LAST_BACKUP_AT') || '' };
}

/** 통신용 페이지가 뜰 때 실제로 서버와 통신되는지 시험 */
function ping() {
  return 'pong';
}

/* ───────────── 인증 API ───────────── */

/** 첫 화면. 토큰이 없거나 만료면 needLogin, 팀원이 한 명도 없으면 needsSetup */
function getBootstrap(token) {
  if (readAll_('users').length === 0) return { needsSetup: true };
  const ctx = auth_(token, true);
  if (!ctx) return { needLogin: true };
  return { state: state_(ctx) };
}

/** 최초 관리자 등록 — 팀원이 한 명도 없을 때만 가능 */
function setupAdmin(input) {
  return withLock_(() => {
    if (readAll_('users').length > 0) throw new Error('이미 관리자가 등록되어 있습니다. 로그인하세요.');
    const email = cleanEmail_(input.email);
    const name = cleanName_(input.name);
    checkPassword_(input.password);
    const salt = Utilities.getUuid();
    const user = {
      email, name, role: 'pm', part: PARTS.indexOf(input.part) >= 0 ? input.part : '', createdAt: new Date().toISOString(),
      passwordHash: hash_(input.password, salt), salt, mustChange: '',
    };
    upsert_('users', user);
    logChange_({ email }, 'user', '관리자 등록', email, `${name} · PM`);
    return loginResult_(user);
  });
}

function login(email, password, remember) {
  email = String(email || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  const failKey = 'fail:' + email;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_LOGIN_FAILS) throw new Error('로그인 시도가 너무 많습니다. 10분 후 다시 시도하세요.');

  const user = readAll_('users').find((u) => u.email === email);
  if (!user || !user.passwordHash || hash_(String(password || ''), user.salt) !== user.passwordHash) {
    cache.put(failKey, String(fails + 1), 600);
    throw new Error('이메일 또는 비밀번호가 올바르지 않습니다.');
  }
  cache.remove(failKey);
  logChange_({ email }, 'login', '로그인', email, remember ? '로그인 유지' : '');
  return loginResult_(user, remember);
}

function changePassword(token, current, next, remember) {
  return withLock_(() => {
    const ctx = auth_(token);
    const user = ctx.user;
    if (hash_(String(current || ''), user.salt) !== user.passwordHash) throw new Error('현재 비밀번호가 올바르지 않습니다.');
    checkPassword_(next);
    if (next === current) throw new Error('새 비밀번호가 현재 비밀번호와 같습니다.');
    const salt = Utilities.getUuid();
    Object.assign(user, { passwordHash: hash_(next, salt), salt, mustChange: '' });
    upsert_('users', user);
    logChange_(ctx, 'user', '비밀번호 변경', ctx.email, '');
    return loginResult_(user, remember);
  });
}

/** 관리자: 팀원 비밀번호를 임시 비밀번호로 초기화 */
function resetPassword(token, email) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    const user = ctx.users.find((u) => u.email === String(email).toLowerCase());
    if (!user) throw new Error('팀원을 찾을 수 없습니다.');
    if (user.email === ctx.email) throw new Error('본인 비밀번호는 [비밀번호 변경]에서 바꾸세요.');
    const temp = tempPassword_();
    const salt = Utilities.getUuid();
    Object.assign(user, { passwordHash: hash_(temp, salt), salt, mustChange: 'Y' });
    upsert_('users', user);
    logChange_(ctx, 'user', '비밀번호 초기화', user.email, user.name);
    return { state: state_(auth_(token)), tempPassword: temp, email: user.email, appUrl: appUrl_() };
  });
}

/* ───────────── 일정 API ───────────── */

/** 등록 폼에서 실시간 경고 확인용 (저장하지 않음) */
function checkLeave(token, input) {
  const ctx = auth_(token);
  const draft = normalizeLeave_(ctx, input, null);
  const leaves = readAll_('leaves');
  const dup = duplicateOf_(draft, leaves);
  return {
    duplicate: dup ? `이미 ${rangeText_(dup)} ${dup.type} 일정이 있어 등록할 수 없습니다.` : '',
    warnings: draft.substitutes.length ? conflictsOf_(draft, leaves, ctx.users) : [],
  };
}

function saveLeave(token, input) {
  return withLock_(() => {
    const ctx = auth_(token);
    const isAdmin = isPM_(ctx);
    const leaves = readAll_('leaves');
    const now = new Date().toISOString();

    let existing = null;
    if (input.id) {
      existing = leaves.find((l) => l.id === input.id);
      if (!existing) throw new Error('일정을 찾을 수 없습니다. 새로고침 후 다시 시도하세요.');
      if (!isAdmin && existing.email !== ctx.email) throw new Error('본인 일정만 수정할 수 있습니다.');
    }

    const l = normalizeLeave_(ctx, input, existing);
    if (!l.substitutes.length) throw new Error('대직자를 1명 이상 지정해야 합니다.');
    if (l.substitutes.length > MAX_SUBSTITUTES) throw new Error(`대직자는 최대 ${MAX_SUBSTITUTES}명까지 지정할 수 있습니다.`);
    if (l.substitutes.indexOf(l.email) >= 0) throw new Error('본인을 대직자로 지정할 수 없습니다.');
    if (l.substitutes.some((e) => !ctx.users.some((u) => u.email === e))) throw new Error('대직자가 팀원 목록에 없습니다.');
    const dup = duplicateOf_(l, leaves);
    if (dup) throw new Error(`이미 등록된 일정과 겹칩니다: ${rangeText_(dup)} ${dup.type}`);

    const memo = String(input.memo || '').slice(0, 200);
    upsert_('leaves', {
      id: existing ? existing.id : Utilities.getUuid(),
      email: l.email,
      type: l.type,
      startDate: l.startDate,
      endDate: l.endDate,
      substitute: l.substitutes.join(','),
      memo,
      createdAt: existing ? existing.createdAt : now,
      updatedAt: now,
      updatedBy: ctx.email,
    });
    const desc = (x) => `${x.type} ${rangeText_(x)} · 대직 ${x.substitutes.map((e) => nameOf_(ctx, e)).join(', ')}`;
    const changed = existing && desc(existing) !== desc(l) ? `${desc(existing)} → ${desc(l)}` : desc(l) + (existing && existing.memo !== memo ? ' (메모 수정)' : '');
    logChange_(ctx, 'leave', existing ? '수정' : '등록', l.email, changed);
    leaveNews_(ctx, existing, l);
    if (input.saveDefault) {
      const owner = ctx.users.find((u) => u.email === l.email);
      owner.defaultSubs = l.substitutes.join(',');
      upsert_('users', owner);
    }
    return { state: state_(ctx) };
  });
}

function deleteLeave(token, id) {
  return withLock_(() => {
    const ctx = auth_(token);
    const leave = readAll_('leaves').find((l) => l.id === id);
    if (!leave) throw new Error('이미 삭제된 일정입니다.');
    if (!isPM_(ctx) && leave.email !== ctx.email) throw new Error('본인 일정만 삭제할 수 있습니다.');
    deleteRow_('leaves', id);
    logChange_(ctx, 'leave', '삭제', leave.email,
      `${leave.type} ${rangeText_(leave)} · 대직 ${leave.substitutes.map((e) => nameOf_(ctx, e)).join(', ')}${leave.memo ? ' · 메모: ' + leave.memo : ''}`);
    leaveNews_(ctx, leave, null);
    return { state: state_(ctx) };
  });
}

/* ───────────── 관리자 API ───────────── */

/** 신규 팀원이면 임시 비밀번호를 발급해 돌려준다 */
function saveUser(token, input) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    const email = cleanEmail_(input.email);
    const name = cleanName_(input.name);
    const role = ROLES.indexOf(input.role) >= 0 ? input.role : 'member';
    const part = PARTS.indexOf(input.part) >= 0 ? input.part : '';
    if (role !== 'pm' && !part) throw new Error('PL과 파트원은 파트를 지정해야 합니다.');

    const existing = ctx.users.find((u) => u.email === email);
    if (input.isNew && existing) throw new Error('이미 등록된 이메일입니다.');
    if (existing && existing.role === 'pm' && role !== 'pm' && pmCount_(ctx.users) <= 1) {
      throw new Error('PM(관리자)은 최소 1명 이상 있어야 합니다.');
    }

    const defaultSubs = checkDefaultSubs_(input.defaultSubs, email, ctx.users).join(',');
    const viewAll = role === 'member' && input.viewAll ? 'Y' : ''; // PM·PL은 원래 전체를 봄

    let temp = '';
    const desc = (u) => `${u.name} · ${ROLE_LABELS[u.role] || u.role}${u.part ? ' · ' + u.part : ''}${u.viewAll === 'Y' ? ' · 전체 일정 조회' : ''}`;
    if (existing) {
      const before = desc(existing);
      Object.assign(existing, { name, role, part, defaultSubs, viewAll });
      upsert_('users', existing);
      if (before !== desc(existing)) logChange_(ctx, 'user', '정보 수정', email, `${before} → ${desc(existing)}`);
    } else {
      temp = tempPassword_();
      const salt = Utilities.getUuid();
      upsert_('users', {
        email, name, role, part, defaultSubs, viewAll, createdAt: new Date().toISOString(),
        passwordHash: hash_(temp, salt), salt, mustChange: 'Y',
      });
      logChange_(ctx, 'user', '추가', email, desc({ name, role, part, viewAll }));
    }
    return { state: state_(auth_(token)), tempPassword: temp, email, appUrl: appUrl_() };
  });
}

/** 본인 기본 대직자 저장 (누구나) */
function setMyDefaultSubs(token, subs) {
  return withLock_(() => {
    const ctx = auth_(token);
    ctx.user.defaultSubs = checkDefaultSubs_(subs, ctx.email, ctx.users).join(',');
    upsert_('users', ctx.user);
    return { state: state_(ctx) };
  });
}

function checkDefaultSubs_(v, owner, users) {
  const subs = parseSubs_(v);
  if (subs.length > MAX_SUBSTITUTES) throw new Error(`기본 대직자는 최대 ${MAX_SUBSTITUTES}명까지 지정할 수 있습니다.`);
  if (subs.indexOf(owner) >= 0) throw new Error('본인을 대직자로 지정할 수 없습니다.');
  if (subs.some((e) => !users.some((u) => u.email === e))) throw new Error('대직자가 팀원 목록에 없습니다.');
  return subs;
}

/**
 * 일괄 등록. rows: [{ email, name, part, role }] (role: pm/pl/member 또는 PM/PL/파트원, 비우면 파트원)
 * 행마다 검사해서 통과한 사람만 등록하고, 실패한 행은 사유와 함께 돌려준다.
 */
function bulkAddUsers(token, rows) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    if (!Array.isArray(rows) || !rows.length) throw new Error('등록할 직원이 없습니다.');
    if (rows.length > 200) throw new Error('한 번에 200명까지 등록할 수 있습니다.');

    const taken = {};
    ctx.users.forEach((u) => (taken[u.email] = true));
    const created = [];
    const errors = [];
    const newRows = [];
    const now = new Date().toISOString();

    rows.forEach((r, i) => {
      try {
        const email = cleanEmail_(r.email);
        const name = cleanName_(r.name);
        const role = roleOf_(r.role);
        const part = PARTS.indexOf(String(r.part || '').trim()) >= 0 ? String(r.part).trim() : '';
        if (role !== 'pm' && !part) throw new Error(`파트를 확인하세요 (${PARTS.join('/')})`);
        if (taken[email]) throw new Error('이미 등록된 이메일입니다.');
        taken[email] = true;

        const temp = tempPassword_();
        const salt = Utilities.getUuid();
        const user = { email, name, role, createdAt: now, passwordHash: hash_(temp, salt), salt, mustChange: 'Y', part };
        newRows.push(SHEETS.users.headers.map((h) => String(user[h] == null ? '' : user[h])));
        created.push({ email, name, part, role, tempPassword: temp });
      } catch (e) {
        errors.push({ line: i + 1, email: String(r.email || ''), message: e.message });
      }
    });

    if (newRows.length) {
      const sh = sheet_('users');
      sh.getRange(sh.getLastRow() + 1, 1, newRows.length, newRows[0].length).setNumberFormat('@').setValues(newRows);
      created.forEach((u) => logChange_(ctx, 'user', '추가 (일괄)', u.email, `${u.name} · ${ROLE_LABELS[u.role]}${u.part ? ' · ' + u.part : ''}`));
    }
    return { state: state_(auth_(token)), created, errors, appUrl: appUrl_() };
  });
}

/**
 * 연차 부여 일수 — 연도별로 휴가 종류(일반휴가·특별휴가·대체휴가)마다 따로 입력
 * '{"2026":{"일반휴가":15,"대체휴가":1}}' → { '2026': { 일반휴가: 15, 대체휴가: 1 } }
 * (처음 버전은 '{"2026":15}' 숫자 하나 → 일반휴가로 읽는다. 이상한 값은 무시)
 */
function parseAnnual_(v) {
  try {
    const o = JSON.parse(v || '{}');
    const out = {};
    Object.keys(o).forEach((y) => {
      if (!/^\d{4}$/.test(y)) return;
      const src = typeof o[y] === 'number' ? { 일반휴가: o[y] } : o[y] && typeof o[y] === 'object' ? o[y] : {};
      const k = {};
      ANNUAL_KINDS.forEach((kind) => { if (typeof src[kind] === 'number') k[kind] = src[kind]; });
      if (Object.keys(k).length) out[y] = k;
    });
    return out;
  } catch (e) {
    return {};
  }
}

/**
 * PM: 휴가 부여 일수 입력. days = { email: { 일반휴가: 15, 특별휴가: '', 대체휴가: 1 } } (숫자 하나만 주면 일반휴가)
 * 빈칸('')이면 그 종류를 지운다. 0~60일, 0.5일 단위.
 * 본인과 PM만 볼 수 있고, 남은 휴가는 화면에서 '세 종류 합계 − 사용(연차·반차, 예정 포함)'으로 계산한다 (참고용)
 */
function setAnnualDays(token, year, days) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    year = String(year || '');
    if (!/^20\d{2}$/.test(year)) throw new Error('연도를 확인하세요.');
    if (!days || typeof days !== 'object') throw new Error('입력한 값이 없습니다.');
    const desc = (k) => (k && Object.keys(k).length ? ANNUAL_KINDS.filter((x) => x in k).map((x) => x.slice(0, 2) + ' ' + k[x]).join('·') : '-');
    const changes = [];
    Object.keys(days).forEach((email) => {
      const u = ctx.users.find((x) => x.email === String(email).toLowerCase());
      if (!u) return;
      const entry = days[email] && typeof days[email] === 'object' ? days[email] : { 일반휴가: days[email] };
      const map = parseAnnual_(u.annual);
      const before = desc(map[year]);
      const k = Object.assign({}, map[year] || {});
      ANNUAL_KINDS.forEach((kind) => {
        if (!(kind in entry)) return;
        const raw = String(entry[kind] == null ? '' : entry[kind]).trim();
        if (raw === '') { delete k[kind]; return; }
        const n = Number(raw);
        if (!isFinite(n) || n < 0 || n > 60 || Math.round(n * 2) !== n * 2) throw new Error(`${u.name}: ${kind} 일수는 0~60, 0.5일 단위로 입력하세요.`);
        k[kind] = n;
      });
      if (Object.keys(k).length) map[year] = k;
      else delete map[year];
      const after = desc(map[year]);
      if (before === after) return;
      u.annual = Object.keys(map).length ? JSON.stringify(map) : '';
      upsert_('users', u);
      changes.push(`${u.name} ${before} → ${after}`);
    });
    if (changes.length) logChange_(ctx, 'user', `${year}년 휴가 일수`, '', changes.join(', '));
    return { state: state_(ctx), changed: changes.length };
  });
}

/**
 * PM: 휴가 일괄 등록 (지난 휴가 옮겨 넣기 등). rows: [{ email, type, startDate, endDate, substitutes: [email], memo, createdAt: 'yyyy-MM-dd' }]
 * 화면이 그룹웨어 표를 읽어 이메일로 바꿔 보낸다. 줄마다 검사해서 통과한 줄만 등록하고 실패한 줄은 사유와 함께 돌려준다.
 * 이미 지난 일정은 대직자를 비워도 된다 (예전 기록은 대직자를 모르는 경우가 많음). 앞으로의 일정은 평소처럼 대직자 필수.
 */
function bulkAddLeaves(token, rows) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    if (!Array.isArray(rows) || !rows.length) throw new Error('등록할 휴가가 없습니다.');
    if (rows.length > 500) throw new Error('한 번에 500건까지 등록할 수 있습니다.');
    const leaves = readAll_('leaves');
    const today = today_();
    const now = new Date().toISOString();
    const added = [];
    const errors = [];
    rows.forEach((row, i) => {
      try {
        const l = normalizeLeave_(ctx, { email: row.email, type: row.type, startDate: row.startDate, endDate: row.endDate, substitutes: row.substitutes }, null);
        if (l.substitutes.length > MAX_SUBSTITUTES) throw new Error(`대직자는 최대 ${MAX_SUBSTITUTES}명까지 지정할 수 있습니다.`);
        if (l.substitutes.indexOf(l.email) >= 0) throw new Error('본인을 대직자로 지정할 수 없습니다.');
        if (l.substitutes.some((e) => !ctx.users.some((u) => u.email === e))) throw new Error('대직자가 팀원 목록에 없습니다.');
        if (!l.substitutes.length && l.endDate >= today) throw new Error('오늘 이후 일정은 대직자를 지정해야 합니다.');
        const dup = duplicateOf_(l, leaves.concat(added));
        if (dup) throw new Error(`이미 등록된 일정과 겹칩니다: ${rangeText_(dup)} ${dup.type}`);
        const created = isDate_(String(row.createdAt || '')) ? new Date(row.createdAt + 'T09:00:00+09:00').toISOString() : now;
        added.push({
          id: Utilities.getUuid(), email: l.email, type: l.type, startDate: l.startDate, endDate: l.endDate,
          substitute: l.substitutes.join(','), substitutes: l.substitutes, memo: String(row.memo || '').slice(0, 200),
          createdAt: created, updatedAt: now, updatedBy: ctx.email,
        });
      } catch (e) {
        errors.push({ line: i + 1, message: e.message });
      }
    });
    if (added.length) {
      const sh = sheet_('leaves');
      const values = added.map((o) => SHEETS.leaves.headers.map((h) => String(o[h] == null ? '' : o[h])));
      sh.getRange(sh.getLastRow() + 1, 1, values.length, values[0].length).setNumberFormat('@').setValues(values);
      const byName = {};
      added.forEach((l) => { const n = nameOf_(ctx, l.email); byName[n] = (byName[n] || 0) + 1; });
      logChange_(ctx, 'leave', '일괄 등록', '', `${added.length}건 · ` + Object.keys(byName).map((n) => `${n} ${byName[n]}건`).join(', '));
      added.forEach((l) => leaveNews_(ctx, null, l)); // 지난 일정은 알리지 않음
    }
    return { state: state_(ctx), created: added.length, errors };
  });
}

function roleOf_(v) {
  const s = String(v || '').trim().toLowerCase();
  if (s === 'pm' || s === '관리자') return 'pm';
  if (s === 'pl' || s === '리더') return 'pl';
  return 'member';
}

function deleteUser(token, email) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    email = String(email || '').toLowerCase();
    if (email === ctx.email) throw new Error('본인 계정은 삭제할 수 없습니다.');
    const target = ctx.users.find((u) => u.email === email);
    if (!target) throw new Error('이미 삭제된 팀원입니다.');
    if (target.role === 'pm' && pmCount_(ctx.users) <= 1) throw new Error('PM(관리자)은 최소 1명 이상 있어야 합니다.');
    deleteRow_('users', email);
    deleteRow_('news', email);
    logChange_(ctx, 'user', '삭제', email, `${target.name} · ${ROLE_LABELS[target.role] || target.role}${target.part ? ' · ' + target.part : ''}`);
    return { state: state_(auth_(token)) };
  });
}

/* ───────────── 업무 일정 (파트별, 휴가와 별도) ─────────────
 * 등록: 누구나 (파트원·PL은 자기 파트만, PM은 모든 파트)
 * 수정·삭제: 작성자, 그 파트의 PL, PM
 * 조회: '전체 공개'는 모두, '파트 내부'는 그 파트 사람과 PM, '나만 보기'는 등록한 사람만 (PM도 못 봄)
 */

function taskVisible_(ctx, t) {
  if (t.visibility === 'private') return t.createdBy === ctx.email;
  return t.visibility !== 'part' || isPM_(ctx) || ctx.user.part === t.part;
}

function taskEditable_(ctx, t) {
  if (t.visibility === 'private') return t.createdBy === ctx.email;
  return isPM_(ctx) || t.createdBy === ctx.email || (ctx.user.role === 'pl' && ctx.user.part === t.part);
}

/**
 * 담당자 부재 안내 — 휴가 공개 규칙(visibleLeaves_)을 그대로 따른다.
 * 볼 수 없는 사람의 휴가는 안내하지 않고, 휴가 종류는 알려주지 않는다 (겹치는 날짜와 오전/오후만).
 */
function taskAbsences_(ctx, task, visibleLeaves) {
  const nameOf = (email) => (ctx.users.find((u) => u.email === email) || { name: email }).name;
  const out = [];
  (task.assignees || []).forEach((email) => {
    visibleLeaves.forEach((l) => {
      if (l.email !== email || l.startDate > task.endDate || task.startDate > l.endDate) return;
      const s = l.startDate > task.startDate ? l.startDate : task.startDate;
      const e = l.endDate < task.endDate ? l.endDate : task.endDate;
      const half = l.type === '오전반차' ? ' 오전' : l.type === '오후반차' ? ' 오후' : '';
      out.push(`${nameOf(email)}님 ${rangeText_({ startDate: s, endDate: e })}${half} 부재`);
    });
  });
  return out;
}

function tasksState_(ctx) {
  const visibleLeaves = visibleLeaves_(ctx, readAll_('leaves'));
  return {
    tasks: readAll_('tasks').filter((t) => taskVisible_(ctx, t)).map((t) => ({
      id: t.id, part: t.part, title: t.title, category: t.category, startDate: t.startDate, endDate: t.endDate,
      time: t.time, assignees: t.assignees, status: t.status || '예정', visibility: t.visibility || 'all', memo: t.memo,
      createdBy: t.createdBy, updatedAt: t.updatedAt, canEdit: taskEditable_(ctx, t),
      absences: taskAbsences_(ctx, t, visibleLeaves),
    })),
  };
}

function getTasks(token) {
  return tasksState_(auth_(token));
}

/** 업무 등록 화면에서 담당자 부재 미리 확인 (저장하지 않음) */
function checkTaskAbsence(token, input) {
  const ctx = auth_(token);
  const startDate = String(input.startDate || '');
  const endDate = String(input.endDate || startDate);
  if (!isDate_(startDate) || !isDate_(endDate) || endDate < startDate) return { absences: [] };
  const task = { startDate, endDate, assignees: parseSubs_(input.assignees) };
  return { absences: taskAbsences_(ctx, task, visibleLeaves_(ctx, readAll_('leaves'))) };
}

function saveTask(token, input) {
  return withLock_(() => {
    const ctx = auth_(token);
    const tasks = readAll_('tasks');
    const existing = input.id ? tasks.find((t) => t.id === input.id) : null;
    if (input.id && !existing) throw new Error('업무 일정을 찾을 수 없습니다. 새로고침 후 다시 시도하세요.');
    if (existing && !taskEditable_(ctx, existing)) throw new Error('작성자, 해당 파트 PL, PM만 수정할 수 있습니다.');

    // 파트원·PL은 자기 파트 일정만 (수정 시 파트는 그대로), PM은 모든 파트
    let part = String(input.part || '');
    if (!isPM_(ctx)) part = existing ? existing.part : ctx.user.part;
    if (PARTS.indexOf(part) < 0) throw new Error(isPM_(ctx) ? '파트를 선택하세요.' : '파트가 지정되지 않은 계정입니다. PM에게 파트 지정을 요청하세요.');

    const title = String(input.title || '').trim().slice(0, 60);
    const category = String(input.category || '');
    const startDate = String(input.startDate || '');
    const endDate = String(input.endDate || startDate);
    if (!title) throw new Error('제목을 입력하세요.');
    if (TASK_CATEGORIES.indexOf(category) < 0) throw new Error('분류를 선택하세요.');
    if (!isDate_(startDate) || !isDate_(endDate)) throw new Error('날짜 형식이 올바르지 않습니다.');
    if (endDate < startDate) throw new Error('종료일이 시작일보다 빠릅니다.');
    const visibility = ['all', 'part', 'private'].indexOf(input.visibility) >= 0 ? input.visibility : 'all';
    if (visibility === 'private' && (existing ? existing.createdBy : ctx.email) !== ctx.email) throw new Error("'나만 보기'는 등록한 사람만 설정할 수 있습니다.");
    // 나만 보기: 담당자는 등록한 사람으로 고정
    const assignees = visibility === 'private' ? [existing ? existing.createdBy : ctx.email] : parseSubs_(input.assignees);
    if (assignees.length > MAX_ASSIGNEES) throw new Error(`담당자는 최대 ${MAX_ASSIGNEES}명까지 지정할 수 있습니다.`);
    if (assignees.some((e) => !ctx.users.some((u) => u.email === e))) throw new Error('담당자가 팀원 목록에 없습니다.');
    const status = TASK_STATUSES.indexOf(input.status) >= 0 ? input.status : existing ? existing.status || '예정' : '예정';

    const now = new Date().toISOString();
    const saved = {
      id: existing ? existing.id : Utilities.getUuid(),
      part, title, category, startDate, endDate,
      time: String(input.time || '').trim().slice(0, 30),
      assignees: assignees.join(','),
      status,
      visibility,
      memo: String(input.memo || '').trim().slice(0, 300),
      createdBy: existing ? existing.createdBy : ctx.email,
      createdAt: existing ? existing.createdAt : now,
      updatedAt: now,
      updatedBy: ctx.email,
    };
    upsert_('tasks', saved);
    saved.assignees = assignees;
    logChange_(ctx, 'task', existing ? '수정' : '등록', '', taskDesc_(ctx, existing, saved));
    taskNews_(ctx, existing, saved);
    return tasksState_(ctx);
  });
}

function deleteTask(token, id) {
  return withLock_(() => {
    const ctx = auth_(token);
    const t = readAll_('tasks').find((x) => x.id === id);
    if (!t) throw new Error('이미 삭제된 업무 일정입니다.');
    if (!taskEditable_(ctx, t)) throw new Error('작성자, 해당 파트 PL, PM만 삭제할 수 있습니다.');
    deleteRow_('tasks', id);
    logChange_(ctx, 'task', '삭제', '', taskDesc_(ctx, null, t));
    taskNews_(ctx, t, null);
    return tasksState_(ctx);
  });
}

/** 중요 일정 등록/수정 (PM) */
function saveEvent(token, input) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    const title = String(input.title || '').trim().slice(0, 50);
    const startDate = String(input.startDate || '');
    const endDate = String(input.endDate || startDate);
    if (!title) throw new Error('제목을 입력하세요.');
    if (!isDate_(startDate) || !isDate_(endDate)) throw new Error('날짜 형식이 올바르지 않습니다.');
    if (endDate < startDate) throw new Error('종료일이 시작일보다 빠릅니다.');

    const existing = input.id ? readAll_('events').find((e) => e.id === input.id) : null;
    if (input.id && !existing) throw new Error('중요 일정을 찾을 수 없습니다. 새로고침 후 다시 시도하세요.');
    const now = new Date().toISOString();
    const ev = {
      id: existing ? existing.id : Utilities.getUuid(),
      title, startDate, endDate,
      time: String(input.time || '').trim().slice(0, 30),
      memo: String(input.memo || '').trim().slice(0, 300),
      createdAt: existing ? existing.createdAt : now,
      updatedAt: now,
      updatedBy: ctx.email,
    };
    upsert_('events', ev);
    const desc = (e) => `${e.title} ${rangeText_(e)}${e.time ? ' ' + e.time : ''}`;
    logChange_(ctx, 'event', existing ? '수정' : '등록', '', existing && desc(existing) !== desc(ev) ? `${desc(existing)} → ${desc(ev)}` : desc(ev));
    const everyone = ctx.users.map((u) => u.email);
    if (ev.endDate >= today_()) {
      if (!existing) pushNews_(ctx, everyone, `새 중요 일정: ${desc(ev)}`, 'day:' + ev.startDate);
      else if (desc(existing) !== desc(ev)) pushNews_(ctx, everyone, `중요 일정 변경: ${desc(ev)}`, 'day:' + ev.startDate);
    }
    return { state: state_(ctx) };
  });
}

function deleteEvent(token, id) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    const ev = readAll_('events').find((e) => e.id === String(id));
    deleteRow_('events', String(id));
    if (ev) {
      logChange_(ctx, 'event', '삭제', '', `${ev.title} ${rangeText_(ev)}`);
      if (ev.endDate >= today_()) pushNews_(ctx, ctx.users.map((u) => u.email), `중요 일정 취소: ${ev.title} ${rangeText_(ev)}`, 'tab:events');
    }
    return { state: state_(ctx) };
  });
}

function saveHoliday(token, input) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    const date = String(input.date || '');
    const name = String(input.name || '').trim().slice(0, 30);
    if (!isDate_(date)) throw new Error('날짜 형식이 올바르지 않습니다.');
    if (!name) throw new Error('공휴일 이름을 입력하세요.');
    upsert_('holidays', { date, name });
    logChange_(ctx, 'holiday', '등록', '', `${date} ${name}`);
    return { state: state_(ctx) };
  });
}

function deleteHoliday(token, date) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    const h = readAll_('holidays').find((x) => x.date === String(date));
    deleteRow_('holidays', String(date));
    if (h) logChange_(ctx, 'holiday', '삭제', '', `${h.date} ${h.name}`);
    return { state: state_(ctx) };
  });
}

/* ───────────── 변경 이력 · 새 소식 ─────────────
 * 변경 이력: 누가 언제 무엇을 바꿨는지 한 줄씩 (PM만 조회, 최근 HISTORY_DAYS일)
 * 새 소식  : 그중 '나와 관련된 것'만 사람마다 따로 쌓는다 (최근 NEWS_MAX개). 내가 한 일은 나에게 오지 않는다.
 * 대직 소식에는 이름·날짜만 넣고 휴가 종류는 넣지 않는다.
 */

/** 이력 저장소. Cloudflare(DB) 서버는 이력을 매 요청 읽지 않도록 따로 보관하는 TeamLog를 넣어준다. Apps Script는 시트 그대로 */
function historyStore_() {
  if (typeof TeamLog !== 'undefined' && TeamLog) return TeamLog;
  return {
    append: (row) => {
      const sh = sheet_('history');
      sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
    },
    read: () => {
      const sh = sheet_('history');
      const last = sh.getLastRow();
      return last < 2 ? [] : sh.getRange(2, 1, last - 1, SHEETS.history.headers.length).getValues();
    },
  };
}

function logChange_(ctx, kind, action, owner, summary) {
  const h = { id: Utilities.getUuid(), at: new Date().toISOString(), actor: ctx.email, kind, action, owner: owner || '', summary: String(summary || '').slice(0, 300) };
  historyStore_().append(SHEETS.history.headers.map((k) => String(h[k])));
}

/** PM: 최근 변경 이력 (새것부터, 최대 1000건) */
function getHistory(token) {
  requireAdmin_(token);
  const since = new Date(Date.now() - HISTORY_DAYS * 864e5).toISOString();
  const H = SHEETS.history.headers;
  const items = historyStore_().read()
    .map((r) => { const o = {}; H.forEach((h, i) => (o[h] = norm_(r[i]))); return o; })
    .filter((o) => o.id && o.at >= since)
    .reverse()
    .sort((a, b) => (a.at < b.at ? 1 : a.at > b.at ? -1 : 0))
    .slice(0, 1000);
  return { items, days: HISTORY_DAYS };
}

function today_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

function nameOf_(ctx, email) {
  return (ctx.users.find((u) => u.email === email) || { name: email }).name;
}

function parseNews_(v) {
  try {
    const a = JSON.parse(v || '[]');
    return Array.isArray(a) ? a : [];
  } catch (e) {
    return [];
  }
}

/** emails에게 소식 하나 (나 자신·없는 사람은 제외) */
function pushNews_(ctx, emails, text, go) {
  const targets = emails.filter((e, i, a) => e && e !== ctx.email && a.indexOf(e) === i && ctx.users.some((u) => u.email === e));
  if (!targets.length) return;
  const rows = readAll_('news');
  const now = new Date().toISOString();
  const since = new Date(Date.now() - HISTORY_DAYS * 864e5).toISOString();
  targets.forEach((email) => {
    const row = rows.find((n) => n.email === email) || { email, seenAt: '' };
    const items = [{ at: now, text: String(text).slice(0, 200), go: go || '' }].concat(parseNews_(row.items))
      .filter((x) => x.at >= since).slice(0, NEWS_MAX);
    upsert_('news', { email, seenAt: row.seenAt || '', items: JSON.stringify(items) });
  });
}

function newsOf_(ctx) {
  const row = readAll_('news').find((n) => n.email === ctx.email);
  if (!row) return { items: [], unread: 0, seenAt: '' };
  const items = parseNews_(row.items);
  return { items, unread: items.filter((x) => x.at > (row.seenAt || '')).length, seenAt: row.seenAt || '' };
}

/** 새 소식 창을 열면 읽음 처리 */
function markNewsRead(token) {
  return withLock_(() => {
    const ctx = auth_(token);
    const row = readAll_('news').find((n) => n.email === ctx.email);
    if (!row) return { seenAt: '' };
    row.seenAt = new Date().toISOString();
    upsert_('news', row);
    return { seenAt: row.seenAt };
  });
}

/** 휴가 등록·수정(old, now)·삭제(old, null) → 대직자와 (남이 바꿨으면) 본인에게 소식. 지난 일정은 알리지 않음 */
function leaveNews_(ctx, old, now) {
  const today = today_();
  const cur = now || old;
  const who = nameOf_(ctx, cur.email);
  const before = old ? old.substitutes : [];
  const after = now ? now.substitutes : [];
  const datesChanged = old && now && (old.startDate !== now.startDate || old.endDate !== now.endDate);
  if (now && now.endDate >= today) {
    pushNews_(ctx, after.filter((e) => before.indexOf(e) < 0), `${who}님이 나를 대직자로 지정했어요 · ${rangeText_(now)}`, 'day:' + now.startDate);
    if (datesChanged) pushNews_(ctx, after.filter((e) => before.indexOf(e) >= 0), `${who}님 일정 날짜가 바뀌었어요 (대직) · ${rangeText_(old)} → ${rangeText_(now)}`, 'day:' + now.startDate);
  }
  if (old && old.endDate >= today) {
    const dropped = before.filter((e) => after.indexOf(e) < 0);
    pushNews_(ctx, dropped, now ? `${who}님 일정의 대직자에서 빠졌어요 · ${rangeText_(old)}` : `${who}님이 일정을 취소했어요 (대직 해제) · ${rangeText_(old)}`, now ? 'day:' + old.startDate : '');
  }
  // PM이 다른 사람의 일정을 바꾼 경우 본인에게
  if (cur.email !== ctx.email && cur.endDate >= today) {
    const me = nameOf_(ctx, ctx.email);
    const text = !old ? `${me}님이 내 일정을 등록했어요` : !now ? `${me}님이 내 일정을 삭제했어요` : `${me}님이 내 일정을 수정했어요`;
    pushNews_(ctx, [cur.email], `${text} · ${cur.type} ${rangeText_(cur)}`, now ? 'day:' + now.startDate : '');
  }
}

/** 업무 이력 문구 ('나만 보기' 업무는 PM도 볼 수 없으므로 제목을 남기지 않음) */
function taskDesc_(ctx, old, t) {
  if (t.visibility === 'private' || (old && old.visibility === 'private')) return `[${t.part}] (나만 보기 업무)`;
  const desc = (x) => `[${x.part}] ${x.title} ${rangeText_(x)} · ${x.status || '예정'} · 담당 ${(x.assignees || []).map((e) => nameOf_(ctx, e)).join(', ') || '없음'}`;
  return old && desc(old) !== desc(t) ? `${desc(old)} → ${desc(t)}` : desc(t);
}

/** 업무 등록·수정·삭제 → 담당자에게 소식 (그 업무를 볼 수 있는 사람만, 지난 업무는 알리지 않음) */
function taskNews_(ctx, old, now) {
  const today = today_();
  const canSee = (t) => (e) => {
    const user = ctx.users.find((u) => u.email === e);
    return user && taskVisible_({ email: e, user, users: ctx.users }, t);
  };
  const before = old ? old.assignees || [] : [];
  const after = now ? now.assignees || [] : [];
  if (now && now.endDate >= today) {
    const see = canSee(now);
    pushNews_(ctx, after.filter((e) => before.indexOf(e) < 0 && see(e)), `업무 배정: ${now.title} · ${rangeText_(now)}`, 'tab:work');
    if (old && (old.startDate !== now.startDate || old.endDate !== now.endDate)) {
      pushNews_(ctx, after.filter((e) => before.indexOf(e) >= 0 && see(e)), `업무 일정 변경: ${now.title} · ${rangeText_(old)} → ${rangeText_(now)}`, 'tab:work');
    }
  }
  if (old && old.endDate >= today) {
    const see = canSee(old);
    pushNews_(ctx, before.filter((e) => after.indexOf(e) < 0 && see(e)), now ? `업무 담당에서 빠졌어요: ${old.title}` : `업무 삭제: ${old.title} · ${rangeText_(old)}`, 'tab:work');
  }
}

/* ───────────── 권한 / 상태 ───────────── */

/** 토큰 검증. 실패하면 'AUTH:'로 시작하는 오류 (quiet면 null) */
function auth_(token, quiet) {
  const fail = () => {
    if (quiet) return null;
    throw new Error('AUTH:로그인이 만료되었습니다. 다시 로그인하세요.');
  };
  if (!token || String(token).indexOf('.') < 0) return fail();
  const parts = String(token).split('.');
  let payload;
  try {
    payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
  } catch (e) {
    return fail();
  }
  if (sign_(payload) !== parts[1]) return fail();
  const [email, exp, pwTag] = payload.split('|');
  if (Number(exp) < Date.now()) return fail();

  const users = readAll_('users');
  const user = users.find((u) => u.email === email);
  if (!user || user.passwordHash.slice(0, 12) !== pwTag) return fail(); // 비밀번호 변경·초기화 시 기존 토큰 무효
  return { email, user, users };
}

function isPM_(ctx) {
  return ctx.user.role === 'pm';
}

function requireAdmin_(token) {
  const ctx = auth_(token);
  if (!isPM_(ctx)) throw new Error('PM(관리자)만 사용할 수 있는 기능입니다.');
  return ctx;
}

function loginResult_(user, remember) {
  const days = remember ? REMEMBER_DAYS : TOKEN_DAYS;
  const payload = [user.email, Date.now() + days * 864e5, user.passwordHash.slice(0, 12)].join('|');
  const token = Utilities.base64EncodeWebSafe(payload) + '.' + sign_(payload);
  return { token, state: state_({ email: user.email, user, users: readAll_('users') }) };
}

/** PM·PL·전체 조회 옵션: 전체 / 파트원: 본인 + 같은 파트 + PM·PL(리더)의 휴가 */
function visibleLeaves_(ctx, all) {
  const role = ctx.user.role;
  if (role === 'pm' || role === 'pl' || ctx.user.viewAll === 'Y') return all;
  return all.filter((l) => l.email === ctx.email || inMyPart_(ctx, l) || isLeader_(ctx, l.email));
}
function inMyPart_(ctx, l) {
  return !!ctx.user.part && (ctx.users.find((u) => u.email === l.email) || {}).part === ctx.user.part;
}
function isLeader_(ctx, email) {
  return ['pm', 'pl'].indexOf((ctx.users.find((u) => u.email === email) || {}).role) >= 0;
}
/** 파트원이 다른 파트 리더의 휴가를 볼 때는 날짜·종류·대직자만 (메모·대직 겹침 내용은 숨김) */
function briefOnly_(ctx, l) {
  return ctx.user.role === 'member' && ctx.user.viewAll !== 'Y' && l.email !== ctx.email && !inMyPart_(ctx, l);
}

function state_(ctx) {
  const pm = isPM_(ctx);
  const all = readAll_('leaves');
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const nameOf = (email) => (ctx.users.find((u) => u.email === email) || { name: email }).name;
  // 삭제된 직원은 기본 대직자에서 자동으로 빠진다
  const defaults = (u) => parseSubs_(u.defaultSubs).filter((e) => ctx.users.some((x) => x.email === e));
  return {
    me: {
      email: ctx.email, name: ctx.user.name, role: ctx.user.role, part: ctx.user.part,
      mustChange: ctx.user.mustChange === 'Y', defaultSubs: defaults(ctx.user), viewAll: ctx.user.viewAll === 'Y',
      annual: parseAnnual_(ctx.user.annual), // 내 연차 부여 일수 (연도별)
    },
    // 대직자 선택용 명단 (이름·파트). 권한 정보는 PM에게만
    users: ctx.users.map((u) => (pm
      ? { email: u.email, name: u.name, role: u.role, part: u.part, defaultSubs: defaults(u), viewAll: u.viewAll === 'Y', annual: parseAnnual_(u.annual) }
      : { email: u.email, name: u.name, part: u.part })),
    leaves: visibleLeaves_(ctx, all).map((l) => ({
      id: l.id, email: l.email, type: l.type, startDate: l.startDate, endDate: l.endDate,
      substitutes: l.substitutes, memo: briefOnly_(ctx, l) ? '' : l.memo, createdAt: l.createdAt, updatedAt: l.updatedAt,
      conflicts: briefOnly_(ctx, l) ? [] : conflictsOf_(l, all, ctx.users),
    })),
    holidays: readAll_('holidays'),
    // 중요 일정은 모든 직원에게
    events: readAll_('events').map((e) => ({
      id: e.id, title: e.title, startDate: e.startDate, endDate: e.endDate, time: e.time, memo: e.memo, createdAt: e.createdAt,
    })),
    // 앞으로 있을 일정에서 나를 대직자로 지정한 사람 — 이름만 (날짜·종류는 노출하지 않음)
    designatedBy: all
      .filter((l) => l.substitutes.indexOf(ctx.email) >= 0 && l.endDate >= today)
      .map((l) => nameOf(l.email))
      .filter((n, i, a) => a.indexOf(n) === i)
      .sort(),
    news: newsOf_(ctx),
  };
}

function pmCount_(users) {
  return users.filter((u) => u.role === 'pm').length;
}

/* ───────────── 일정 검증 / 겹침 ───────────── */

function normalizeLeave_(ctx, input, existing) {
  const email = isPM_(ctx) && input.email ? String(input.email).toLowerCase() : existing ? existing.email : ctx.email;
  const type = String(input.type || '');
  const startDate = String(input.startDate || '');
  let endDate = String(input.endDate || startDate);
  if (!ctx.users.some((u) => u.email === email)) throw new Error('등록되지 않은 팀원입니다.');
  if (LEAVE_TYPES.indexOf(type) < 0) throw new Error('일정 유형을 선택하세요.');
  if (!isDate_(startDate) || !isDate_(endDate)) throw new Error('날짜 형식이 올바르지 않습니다.');
  if (HALF_TYPES.indexOf(type) >= 0) endDate = startDate;
  if (endDate < startDate) throw new Error('종료일이 시작일보다 빠릅니다.');
  return {
    id: existing ? existing.id : String(input.id || ''),
    email, type, startDate, endDate,
    substitutes: parseSubs_(input.substitutes || input.substitute),
  };
}

/** 대직자 목록 — 배열 또는 '가@x.com,나@x.com' 문자열 → 중복 없는 소문자 배열 */
function parseSubs_(v) {
  const list = Array.isArray(v) ? v : String(v || '').split(',');
  return list
    .map((e) => String(e || '').trim().toLowerCase())
    .filter((e, i, a) => e && a.indexOf(e) === i);
}

function slots_(type) {
  if (type === '오전반차') return 1;
  if (type === '오후반차') return 2;
  return 3; // 종일
}

function overlaps_(a, b) {
  return a.startDate <= b.endDate && b.startDate <= a.endDate && (slots_(a.type) & slots_(b.type)) !== 0;
}

function duplicateOf_(l, leaves) {
  return leaves.find((x) => x.id !== l.id && x.email === l.email && overlaps_(x, l)) || null;
}

/**
 * 대직 충돌 경고 문구. 상대방 휴가의 종류·전체 기간은 노출하지 않고 겹치는 날짜만 알려준다.
 * ① 내 대직자(여러 명 중 누구라도)가 같은 시간에 부재
 * ② 나를 대직자로 지정한 사람과 같은 시간에 부재
 */
function conflictsOf_(l, leaves, users) {
  const nameOf = (email) => {
    const u = users.find((x) => x.email === email);
    return u ? u.name : email;
  };
  const out = [];
  leaves.forEach((m) => {
    if (m.id === l.id || !overlaps_(l, m)) return;
    const s = l.startDate > m.startDate ? l.startDate : m.startDate;
    const e = l.endDate < m.endDate ? l.endDate : m.endDate;
    const when = rangeText_({ startDate: s, endDate: e }) + (m.type === '오전반차' ? ' 오전' : m.type === '오후반차' ? ' 오후' : '');
    // 대직자 중 한 명이라도 같은 시간에 부재면 경고
    if (l.substitutes.indexOf(m.email) >= 0) out.push(`대직자 ${nameOf(m.email)}님이 ${when} 부재`);
    else if (m.substitutes.indexOf(l.email) >= 0) out.push(`${nameOf(m.email)}님이 ${nameOf(l.email)}님을 대직자로 지정했는데 ${when} 둘 다 부재`);
  });
  return out;
}

function rangeText_(l) {
  const f = (s) => {
    const [y, mo, d] = s.split('-').map(Number);
    return `${mo}/${d}(${WEEK[new Date(y, mo - 1, d).getDay()]})`;
  };
  return l.startDate === l.endDate ? f(l.startDate) : `${f(l.startDate)} ~ ${f(l.endDate)}`;
}

/* ───────────── 비밀번호 / 토큰 ───────────── */

function hash_(password, salt) {
  let h = salt + '|' + password;
  for (let i = 0; i < 300; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt));
  }
  return h;
}

function sign_(payload) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, secret_()));
}

function secret_() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('TOKEN_SECRET');
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('TOKEN_SECRET', s);
  }
  return s;
}

/** 안내 메시지에 넣을 웹 앱 주소 (/exec) */
function appUrl_() {
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (e) {
    return '';
  }
}

function tempPassword_() {
  const chars = 'abcdefghjkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < 8; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

function checkPassword_(pw) {
  if (String(pw || '').length < 8) throw new Error('비밀번호는 8자 이상이어야 합니다.');
}

function cleanEmail_(v) {
  const email = String(v || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('이메일 형식이 올바르지 않습니다.');
  return email;
}

function cleanName_(v) {
  const name = String(v || '').trim().slice(0, 30);
  if (!name) throw new Error('이름을 입력하세요.');
  return name;
}

/* ───────────── 시트 입출력 ───────────── */

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function sheet_(key) {
  const def = SHEETS[key];
  const book = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('연결된 스프레드시트가 없습니다. 스프레드시트에서 확장 프로그램 → Apps Script로 만들거나 Code.gs의 SHEET_ID를 입력하세요.');
  let sh = book.getSheetByName(def.name);
  if (!sh) {
    sh = book.insertSheet(def.name);
    sh.getRange(1, 1, sh.getMaxRows(), def.headers.length).setNumberFormat('@'); // 날짜 자동변환 방지
    sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (key === 'holidays') {
      sh.getRange(2, 1, DEFAULT_HOLIDAYS.length, 2).setValues(DEFAULT_HOLIDAYS);
    }
  } else {
    // 다른 용도의 시트를 덮어쓰지 않도록 첫 칸 헤더로 이 앱의 시트인지 확인
    if (norm_(sh.getRange(1, 1).getValue()) !== def.headers[0]) {
      throw new Error(`스프레드시트에 '${def.name}' 시트가 이미 다른 형식으로 있습니다. 시트 이름을 바꾸거나 삭제한 뒤 다시 접속하세요.`);
    }
    if (!sh.getRange(1, def.headers.length).getValue()) {
      sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]); // 컬럼 추가된 버전으로 업그레이드
    }
  }
  return sh;
}

function readAll_(key) {
  const def = SHEETS[key];
  const sh = sheet_(key);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh
    .getRange(2, 1, last - 1, def.headers.length)
    .getValues()
    .filter((r) => r[0] !== '')
    .map((r) => {
      const o = {};
      def.headers.forEach((h, i) => (o[h] = norm_(r[i])));
      if (o.email) o.email = o.email.toLowerCase();
      if (key === 'leaves') o.substitutes = parseSubs_(o.substitute);
      if (key === 'tasks') o.assignees = parseSubs_(o.assignees);
      if (key === 'users') o.role = o.role === 'admin' ? 'pm' : ROLES.indexOf(o.role) >= 0 ? o.role : 'member';
      return o;
    });
}

function norm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v).trim();
}

function findRow_(sh, key) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const col = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < col.length; i++) {
    if (norm_(col[i][0]).toLowerCase() === String(key).toLowerCase()) return i + 2;
  }
  return -1;
}

function upsert_(key, obj) {
  const def = SHEETS[key];
  const sh = sheet_(key);
  const row = def.headers.map((h) => (obj[h] == null ? '' : String(obj[h])));
  const r = findRow_(sh, row[0]);
  if (r > 0) {
    sh.getRange(r, 1, 1, row.length).setValues([row]);
  } else {
    const target = sh.getRange(sh.getLastRow() + 1, 1, 1, row.length);
    target.setNumberFormat('@').setValues([row]);
  }
}

function deleteRow_(key, id) {
  const sh = sheet_(key);
  const r = findRow_(sh, id);
  if (r > 0) sh.deleteRow(r);
}

function isDate_(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(s);
}

  return { doPost, API_FUNCTIONS, SHEETS };
}

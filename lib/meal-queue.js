// 식권 승인 대기열 — 집 맥북 승인 프로그램(meal-agent)이 대신 승인한다
//
// 2026-11 PAYCO가 로그인 기록이 있는 기기에서만 승인 화면을 열어 주도록 바뀌어서, 이 서버가 PAYCO에 직접 승인할 수 없다.
// 그래서 직원이 [승인 요청]을 누르면 요청을 대기열(meal_jobs)에 넣고, 집 맥북이 운영 시간(평일 17~22시)에 5분마다 가져가서
// 승인권자가 로그인해 둔 PAYCO 창에서 승인한 뒤 결과를 보고한다.
//   직원 화면  → /api  fn=checkMeal(상태) · approveMeal(요청) · mealStatus(결과 확인)
//   집 맥북    → /api-meal-agent  action=poll(가져가기, 상태 알림) · report(결과 보고)  (MEAL_AGENT_KEY 필요)
//
// 표 (백업 대상인 rows · kv와 분리 — 스프레드시트 백업에 섞이지 않게)
//   meal_jobs(id, email, name, status, message, approved, failed, created, taken, done)
//     status: queued(대기) → working(맥북이 처리 중) → done · failed  /  72시간 지나면 expired
//   meal_agent(id=1, state, message, seen)  — 맥북이 마지막으로 확인한 시각과 PAYCO 로그인 상태
import { randomUUID } from 'node:crypto';
import { logHistory } from './d1-backend.js';

const SCHEMA = [
  `CREATE TABLE IF NOT EXISTS meal_jobs (id TEXT PRIMARY KEY, email TEXT NOT NULL, name TEXT NOT NULL, status TEXT NOT NULL,
    message TEXT NOT NULL DEFAULT '', approved TEXT NOT NULL DEFAULT '[]', failed TEXT NOT NULL DEFAULT '[]',
    created INTEGER NOT NULL, taken INTEGER, done INTEGER)`,
  'CREATE INDEX IF NOT EXISTS meal_jobs_email ON meal_jobs (email, created)',
  'CREATE TABLE IF NOT EXISTS meal_agent (id INTEGER PRIMARY KEY CHECK (id = 1), state TEXT NOT NULL, message TEXT NOT NULL, seen INTEGER NOT NULL)',
];

// 운영 시간 — 맥북 프로그램(meal-agent/agent.mjs의 HOURS)과 같게 유지할 것. 한국 시간 기준
export const HOURS = { days: [1, 2, 3, 4, 5], start: 17, end: 22, everyMin: 5 }; // 월~금, 17:00 ≤ 시각 < 22:00, 5분마다 확인
export function inHours(now = Date.now()) {
  const k = new Date(now + 9 * 3600e3);
  const h = k.getUTCHours();
  return HOURS.days.includes(k.getUTCDay()) && h >= HOURS.start && h < HOURS.end;
}

export const EXPIRE_MS = 72 * 3600e3; // 운영 시간 밖 요청은 다음 운영 시간에 처리 — 금요일 밤 요청이 월요일 17시까지 남도록
const RETRY_MS = 3 * 60e3; // 처리 중에 맥북이 멈추면 이 시간 뒤 다시 내준다 (승인 전 목록을 다시 읽으므로 두 번 승인되지 않음)
const OFFLINE_MS = (HOURS.everyMin * 2 + 2) * 60e3; // 운영 시간에 이보다 오래 소식이 없으면 꺼진 것으로 본다
const KEEP_MS = 7 * 864e5; // 끝난 요청은 7일 보관

let schemaReady = false;
async function ready(db) {
  if (!schemaReady) { await db.batch(SCHEMA.map((sql) => db.prepare(sql))); schemaReady = true; }
}

const list = (s) => { try { const a = JSON.parse(s || '[]'); return Array.isArray(a) ? a.map(String) : []; } catch (e) { return []; } };

/** 직원 화면에 보여줄 요청 상태 */
function view(row, now = Date.now()) {
  if (!row) return null;
  const expired = (row.status === 'queued' || row.status === 'working') && row.created < now - EXPIRE_MS;
  return {
    id: row.id, status: expired ? 'expired' : row.status,
    message: expired ? '승인 PC가 3일 동안 처리하지 못해 요청이 만료됐어요. 승인권자에게 직접 요청해 주세요.' : row.message,
    approved: list(row.approved), failed: list(row.failed),
    created: new Date(row.created).toISOString(), done: row.done ? new Date(row.done).toISOString() : '',
  };
}

/**
 * 승인 PC 상태: ok(정상) · login(PAYCO 로그인 풀림) · offline(운영 시간인데 소식 없음) · closed(운영 시간 아님) · none(아직 연결된 적 없음)
 * hours: 화면 안내용 운영 시간
 */
async function agentState(db, now = Date.now()) {
  const a = await db.prepare('SELECT state, message, seen FROM meal_agent WHERE id = 1').first();
  const hours = HOURS;
  // 운영 시간 밖에는 맥북이 서버에 연락하지 않으므로, 연결 기록이 없어도 '운영 시간 아님'으로 안내한다
  if (!inHours(now)) return { state: 'closed', seen: a ? new Date(a.seen).toISOString() : '', hours };
  if (!a) return { state: 'none', seen: '', hours };
  const seen = new Date(a.seen).toISOString();
  // 운영 시간이 막 시작됐으면 첫 확인 전까지는 어제 기록이 남아 있으니 꺼진 것으로 보지 않는다
  const k = new Date(now + 9 * 3600e3);
  const sinceOpen = ((k.getUTCHours() - HOURS.start) * 60 + k.getUTCMinutes()) * 60e3;
  if (a.seen < now - OFFLINE_MS && sinceOpen > OFFLINE_MS) return { state: 'offline', seen, hours };
  return { state: a.state === 'login' ? 'login' : 'ok', seen, hours };
}

const latest = (db, email) => db.prepare('SELECT * FROM meal_jobs WHERE email = ?1 ORDER BY created DESC LIMIT 1').bind(email).first();

/** 직원: 식권 버튼 → 승인 PC 상태 + 내 최근 요청 */
export async function mealOverview(env, me) {
  await ready(env.DB);
  return { mode: 'agent', agent: await agentState(env.DB), job: view(await latest(env.DB, me.email)) };
}

/** 직원: [승인 요청] → 대기열에 넣기 (이미 대기 · 처리 중인 요청이 있으면 그걸 돌려줌) */
export async function requestMeal(env, me) {
  await ready(env.DB);
  const now = Date.now();
  const cur = view(await latest(env.DB, me.email), now);
  if (cur && (cur.status === 'queued' || cur.status === 'working')) return { agent: await agentState(env.DB, now), job: cur, existing: true };
  const id = randomUUID();
  await env.DB.prepare("INSERT INTO meal_jobs (id, email, name, status, created) VALUES (?1, ?2, ?3, 'queued', ?4)").bind(id, me.email, me.name, now).run();
  // 가끔 오래된 요청 정리
  if (Math.random() < 0.05) await env.DB.prepare('DELETE FROM meal_jobs WHERE created < ?1').bind(now - KEEP_MS).run();
  return { agent: await agentState(env.DB, now), job: view(await latest(env.DB, me.email), now), existing: false };
}

/** 직원: 요청 결과 확인 (본인 요청만) */
export async function mealStatus(env, me, id) {
  await ready(env.DB);
  const row = await env.DB.prepare('SELECT * FROM meal_jobs WHERE id = ?1 AND email = ?2').bind(String(id || ''), me.email).first();
  if (!row) throw new Error('요청을 찾을 수 없습니다.');
  return { agent: await agentState(env.DB), job: view(row) };
}

/**
 * 맥북: 30초마다 확인. 상태(ok · login)를 기록하고, PAYCO에 로그인돼 있으면 가장 오래된 대기 요청 하나를 내준다.
 * peek: 설정 확인(--check)용 — 키 확인만 하고 아무것도 바꾸지 않는다.
 */
export async function agentPoll(env, body) {
  await ready(env.DB);
  if (body.peek) return { job: null };
  const now = Date.now();
  const state = body.state === 'login' ? 'login' : 'ok';
  await env.DB.prepare('INSERT INTO meal_agent (id, state, message, seen) VALUES (1, ?1, ?2, ?3) ON CONFLICT(id) DO UPDATE SET state = excluded.state, message = excluded.message, seen = excluded.seen')
    .bind(state, String(body.message || '').slice(0, 200), now).run();
  if (state === 'login') return { job: null }; // 로그인이 풀린 동안 요청은 대기로 둔다
  const job = await env.DB.prepare(
    `UPDATE meal_jobs SET status = 'working', taken = ?1
     WHERE id = (SELECT id FROM meal_jobs WHERE created > ?2 AND (status = 'queued' OR (status = 'working' AND taken < ?3)) ORDER BY created LIMIT 1)
     RETURNING id, name`,
  ).bind(now, now - EXPIRE_MS, now - RETRY_MS).first();
  return { job: job ? { id: job.id, name: job.name } : null };
}

/**
 * 맥북: 결과 보고. status = done · failed · blocked(처리 중 PAYCO 로그인이 풀림 → 다시 대기로)
 * 처리 중(working)인 요청에만 반영 — 같은 보고가 두 번 와도 한 번만 기록된다.
 */
export async function agentReport(env, body) {
  await ready(env.DB);
  const id = String(body.id || '');
  const status = ['done', 'failed', 'blocked'].includes(body.status) ? body.status : 'failed';
  const message = String(body.message || '').slice(0, 300);
  const approved = list(JSON.stringify(body.approved || [])).slice(0, 20);
  const failed = list(JSON.stringify(body.failed || [])).slice(0, 20);
  const now = Date.now();
  if (status === 'blocked') {
    await env.DB.prepare("UPDATE meal_jobs SET status = 'queued', taken = NULL, message = ?2 WHERE id = ?1 AND status = 'working'").bind(id, message).run();
    return { saved: true };
  }
  const row = await env.DB.prepare(
    "UPDATE meal_jobs SET status = ?2, message = ?3, approved = ?4, failed = ?5, done = ?6 WHERE id = ?1 AND status = 'working' RETURNING email",
  ).bind(id, status, message, JSON.stringify(approved), JSON.stringify(failed), now).first();
  if (row && (approved.length || failed.length)) {
    // 승인은 이미 끝났으므로 이력 기록이 실패해도 보고는 성공으로 둔다
    const summary = [...approved, ...failed.map((x) => `${x} (승인 안 됨)`)].join(', ');
    await logHistory(env.DB, { actor: row.email, kind: 'meal', action: failed.length ? '승인 실패' : '승인', owner: row.email, summary })
      .catch((e) => console.log(`[meal] 이력 기록 실패: ${e && e.stack}`));
  }
  return { saved: !!row };
}

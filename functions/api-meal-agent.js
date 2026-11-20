// 집 맥북 식권 승인 프로그램(meal-agent) 전용 입구 (/api-meal-agent)
// 비밀 키(MEAL_AGENT_KEY, Cloudflare 비밀 변수로만 설정 — 저장소에 없음)를 x-agent-key 헤더로 보낸 요청만 처리하고, 그 외는 404.
// 이 키로는 대기열 확인(poll)과 결과 보고(report)만 할 수 있다. (lib/meal-queue.js)
import { createHash, timingSafeEqual } from 'node:crypto';
import { agentPoll, agentReport } from '../lib/meal-queue.js';

const HEADERS = { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' };
const json = (obj, status = 200) => new Response(JSON.stringify(obj), { status, headers: HEADERS });
const digest = (s) => createHash('sha256').update(String(s)).digest();

export async function onRequest({ request, env }) {
  const key = request.headers.get('x-agent-key');
  if (!env.MEAL_AGENT_KEY || !env.DB || !key || !timingSafeEqual(digest(key), digest(env.MEAL_AGENT_KEY))) return new Response('Not Found', { status: 404 });
  if (request.method !== 'POST') return json({ ok: false, error: 'POST만 허용됩니다.' }, 405);
  let body;
  try { body = JSON.parse(await request.text()); } catch (e) { return json({ ok: false, error: '요청 형식이 올바르지 않습니다.' }, 400); }
  try {
    if (body.action === 'poll') return json({ ok: true, data: await agentPoll(env, body) });
    if (body.action === 'report') return json({ ok: true, data: await agentReport(env, body) });
    return json({ ok: false, error: '알 수 없는 요청입니다.' }, 400);
  } catch (e) {
    console.log(`[api-meal-agent] ${body.action} 오류: ${e && e.stack}`);
    return json({ ok: false, error: '서버 오류: ' + ((e && e.message) || e) }, 500);
  }
}

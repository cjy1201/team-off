// Cloudflare Pages Function: 화면(/api) → Google Apps Script(doPost) 전달
// 회사망처럼 브라우저가 구글과 직접 통신할 수 없는 환경에서도, Cloudflare 서버가 대신 요청해서 동작하게 한다.
// 요청/응답 형식은 Apps Script doPost와 같다: {"fn": "...", "args": [...]} → {"ok": true, "data": ...}
// (Netlify용 netlify/functions/api.mjs 와 같은 동작)
//
// 환경 변수(wrangler.toml [vars] 또는 Cloudflare 대시보드)로 동작을 바꾼다:
//   BACKEND = 'gas' (기본) → Apps Script로 전달 / 'd1' → Cloudflare DB(D1)에서 직접 처리 (Code.gs 로직 그대로)
//   MAINTENANCE = 'readonly' → 저장 요청만 잠시 막음 (DB 전환 순간에 사용)
import { handleD1 } from '../lib/d1-backend.js';
import { approveMyMeal, checkMyMeal } from '../lib/meal.js';

const DEPLOY_ID = 'AKfycbyJlnAadwmv1SzNsrGRL01ornXEGcSefstb79pJKqld6RsVVSzzbnIzzzEiIUU3mQzt';

// 조회만 하는 요청은 구글 응답 전달(echo)이 실패하면 다시 보내도 안전하다
const READ_ONLY = new Set(['getBootstrap', 'checkLeave', 'getTasks', 'checkTaskAbsence', 'checkMeal']);

const HEADERS = { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' };
const json = (obj, status = 200) => new Response(JSON.stringify(obj), { status, headers: HEADERS });

/**
 * 식권 (Code.gs 밖의 기능): 로그인한 직원 본인 이름으로 들어온 PAYCO 신청만 조회(checkMeal)·승인(approveMeal)한다.
 * 토큰 검증은 getBootstrap을 그대로 써서 Code.gs와 같은 규칙(만료·비밀번호 변경 시 무효)을 따른다.
 */
const MEAL = { checkMeal: checkMyMeal, approveMeal: approveMyMeal };
async function meal(env, fn, body) {
  try {
    if (!env || env.BACKEND !== 'd1' || !env.DB) throw new Error('지금은 사용할 수 없는 기능입니다.');
    const token = (JSON.parse(body).args || [])[0];
    const boot = JSON.parse(await handleD1(env.DB, JSON.stringify({ fn: 'getBootstrap', args: [token] })));
    const me = boot.ok && boot.data && boot.data.state && boot.data.state.me;
    if (!me) throw new Error('AUTH:로그인이 만료되었습니다. 다시 로그인하세요.');
    return { ok: true, data: await MEAL[fn](env, me.name) };
  } catch (e) {
    console.log(`[api] ${fn} 오류: ${e && e.stack}`);
    return { ok: false, error: (e && e.message) || String(e) };
  }
}

async function forward(url, body) {
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'content-type': 'text/plain;charset=utf-8' },
    body,
    redirect: 'follow', // Apps Script는 302로 응답 주소(macros/echo)를 알려준다
    signal: AbortSignal.timeout(20000),
  });
  const text = await res.text();
  if (!res.ok || !text.trim().startsWith('{')) throw new Error(`upstream ${res.status}`);
  return text;
}

/** GET /api?q=… 로 온 요청 내용 복원 (회사망이 POST를 막을 때 화면이 GET으로 보낸다) */
function fromB64Url(q) {
  const b64 = q.replace(/-/g, '+').replace(/_/g, '/');
  const bin = atob(b64 + '='.repeat((4 - (b64.length % 4)) % 4));
  return new TextDecoder().decode(Uint8Array.from(bin, (c) => c.charCodeAt(0)));
}

async function readBody(request) {
  if (request.method === 'POST') return request.text();
  if (request.method === 'GET') {
    const q = new URL(request.url).searchParams.get('q');
    if (q) return fromB64Url(q);
  }
  return null;
}

export async function onRequest({ request, env }) {
  const url = (env && env.GAS_URL) || `https://script.google.com/macros/s/${DEPLOY_ID}/exec`;
  let body, fn = '';
  try {
    body = await readBody(request);
    if (body === null) return json({ ok: false, error: 'POST만 허용됩니다.' }, 405);
    fn = String(JSON.parse(body).fn || '');
  } catch (e) { return json({ ok: false, error: '요청 형식이 올바르지 않습니다.' }, 400); }

  if (env && env.MAINTENANCE === 'readonly' && !READ_ONLY.has(fn)) {
    return json({ ok: false, error: '잠시 점검 중입니다. 1~2분 뒤 다시 시도해 주세요.' });
  }

  if (Object.prototype.hasOwnProperty.call(MEAL, fn)) return json(await meal(env, fn, body));

  // DB 서버
  if (env && env.BACKEND === 'd1' && env.DB) {
    try {
      return new Response(await handleD1(env.DB, body), { headers: HEADERS });
    } catch (e) {
      console.log(`[api:d1] ${fn} 오류: ${e && e.stack}`);
      return json({ ok: false, error: '서버 오류가 발생했습니다. 잠시 후 다시 시도해 주세요.' }, 500);
    }
  }

  // Apps Script로 전달
  // 조회, 그리고 요청 ID(rid)가 붙은 저장 요청은 다시 보내도 서버가 한 번만 처리하므로 3번까지 시도
  let rid = '';
  try { rid = String(JSON.parse(body).rid || ''); } catch (e) {}
  const tries = READ_ONLY.has(fn) || rid ? 3 : 1;
  for (let i = 1; i <= tries; i++) {
    try {
      return new Response(await forward(url, body), { headers: HEADERS });
    } catch (e) {
      console.log(`[api] ${fn} 시도 ${i}/${tries} 실패: ${e.message}`);
      if (i < tries) await new Promise((r) => setTimeout(r, 500 * i));
    }
  }
  // 502 → 화면에서 '응답 유실'로 처리 (조회는 다시 시도, 저장은 반영 여부 확인 안내)
  return json({ ok: false, error: '서버 응답을 받지 못했습니다.' }, 502);
}

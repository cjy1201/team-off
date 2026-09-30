// Netlify Function: 화면(/api) → Google Apps Script(doPost) 전달
// 회사망처럼 브라우저가 구글과 직접 통신할 수 없는 환경에서도, Netlify 서버가 대신 요청해서 동작하게 한다.
// 요청/응답 형식은 Apps Script doPost와 같다: {"fn": "...", "args": [...]} → {"ok": true, "data": ...}

const DEPLOY_ID = 'AKfycbyJlnAadwmv1SzNsrGRL01ornXEGcSefstb79pJKqld6RsVVSzzbnIzzzEiIUU3mQzt';
const GAS_URL = process.env.GAS_URL || `https://script.google.com/macros/s/${DEPLOY_ID}/exec`;

// 조회만 하는 요청은 구글 응답 전달(echo)이 실패하면 다시 보내도 안전하다
const READ_ONLY = new Set(['getBootstrap', 'checkLeave']);

const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), { status, headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' } });

async function forward(body) {
  const res = await fetch(GAS_URL, {
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

export default async (req) => {
  let body, fn = '';
  try {
    body = await readBody(req);
    if (body === null) return json({ ok: false, error: 'POST만 허용됩니다.' }, 405);
    fn = String(JSON.parse(body).fn || '');
  } catch (e) { return json({ ok: false, error: '요청 형식이 올바르지 않습니다.' }, 400); }

  const tries = READ_ONLY.has(fn) ? 3 : 1;
  for (let i = 1; i <= tries; i++) {
    try {
      const text = await forward(body);
      return new Response(text, { headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' } });
    } catch (e) {
      console.log(`[api] ${fn} 시도 ${i}/${tries} 실패: ${e.message}`);
      if (i < tries) await new Promise((r) => setTimeout(r, 500 * i));
    }
  }
  // 502 → 화면에서 '응답 유실'로 처리 (조회는 다시 시도, 저장은 반영 여부 확인 안내)
  return json({ ok: false, error: '서버 응답을 받지 못했습니다.' }, 502);
};

export const config = { path: '/api' };

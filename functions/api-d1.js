// 전환 전 검증용 DB 서버 입구 (/api-d1)
// 비밀 키(D1_TEST_KEY, Cloudflare 비밀 변수로만 설정 — 저장소에 없음)를 x-test-key 헤더로 보낸 요청만 처리하고, 그 외는 404.
// 직원들이 쓰는 /api 에는 영향이 없다. 전환이 끝나면 D1_TEST_KEY를 지워서 닫는다.
import { handleD1 } from '../lib/d1-backend.js';

const HEADERS = { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' };

export async function onRequest({ request, env }) {
  if (!env.D1_TEST_KEY || !env.DB || request.headers.get('x-test-key') !== env.D1_TEST_KEY) return new Response('Not Found', { status: 404 });
  if (request.method !== 'POST') return new Response('POST only', { status: 405 });
  const t0 = Date.now();
  try {
    const text = await handleD1(env.DB, await request.text());
    return new Response(text, { headers: { ...HEADERS, 'x-server-ms': String(Date.now() - t0) } });
  } catch (e) {
    console.log(`[api-d1] 오류: ${e && e.stack}`);
    return new Response(JSON.stringify({ ok: false, error: '서버 오류: ' + (e && e.message) }), { status: 500, headers: HEADERS });
  }
}

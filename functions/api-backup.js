// DB → 스프레드시트 백업용 데이터 내보내기 (/api-backup)
// Apps Script(Backup.gs)가 1시간마다 비밀 키(BACKUP_KEY, Cloudflare 비밀 변수 — 저장소에 없음)를 x-backup-key 헤더로 보내 가져간다.
// 키가 없거나 틀리면 404. 읽기만 하고, 가져간 시각을 LAST_BACKUP_AT 으로 기록한다 (PM 팀원 관리 화면에 표시).
const HEADERS = { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' };

export async function onRequest({ request, env }) {
  if (!env.BACKUP_KEY || !env.DB || request.headers.get('x-backup-key') !== env.BACKUP_KEY) return new Response('Not Found', { status: 404 });
  const { results } = await env.DB.prepare('SELECT sheet, data FROM rows ORDER BY sheet, pos').all();
  const sheets = {};
  for (const r of results) (sheets[r.sheet] = sheets[r.sheet] || []).push(JSON.parse(r.data));
  const exportedAt = new Date().toISOString();
  await env.DB.prepare("INSERT INTO kv (k, v, exp) VALUES ('p:LAST_BACKUP_AT', ?1, NULL) ON CONFLICT(k) DO UPDATE SET v = excluded.v").bind(exportedAt).run();
  return new Response(JSON.stringify({ ok: true, exportedAt, sheets }), { headers: HEADERS });
}

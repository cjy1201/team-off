// 식권(PAYCO 비즈플러스) 본인 신청 승인
//
// 직원이 PAYCO에서 식권을 신청하면 승인권자가 승인해야 발급된다.
// 이 사이트에 로그인한 직원이 '식권 승인'을 누르면, 승인권자 계정으로 PAYCO 승인 목록을 읽어
// 그 직원 이름으로 들어온 대기 신청만 승인한다. (다른 사람 신청은 건드리지 않는다)
//
// PAYCO 승인 화면은 휴대폰 번호만으로 로그인되므로, 번호는 저장소(공개)에 두지 않고
// Cloudflare 비밀 변수로만 넣는다:
//   PAYCO_TEL_NO       승인권자 휴대폰 번호 (숫자만)
//   PAYCO_COMPANY_CODE 회사 코드 (승인 화면 주소의 companyCode 값)
const BASE = 'https://bizplus.payco.com/apply/nonLogin';
const UA = 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15';

const norm = (s) => String(s || '').replace(/\s+/g, '');

async function payco(path, body) {
  const res = await fetch(BASE + path, {
    method: body ? 'POST' : 'GET',
    headers: {
      'user-agent': UA,
      'x-requested-with': 'XMLHttpRequest',
      ...(body ? { 'content-type': 'application/json; charset=utf-8' } : {}),
    },
    body: body ? JSON.stringify(body) : undefined,
    signal: AbortSignal.timeout(15000),
  });
  if (!res.ok) throw new Error(`PAYCO 서버 응답 오류 (HTTP ${res.status})`);
  return res.text();
}

/** 승인권자 로그인 → 승인 화면 주소에 붙는 번호 */
async function login(env) {
  const text = await payco(
    `/ajax/bizCouponApplyApprove/login.json?companyCode=${encodeURIComponent(env.PAYCO_COMPANY_CODE)}`,
    { telNo: env.PAYCO_TEL_NO },
  );
  let res = null;
  try { res = JSON.parse(text); } catch (e) {}
  const seq = res && res.isSuccess && res.result && res.result.bizRoleAssignedEmployeeSeq;
  if (!seq) throw new Error('PAYCO 승인권자 로그인에 실패했습니다. 관리자에게 알려주세요.');
  return `companyCode=${encodeURIComponent(env.PAYCO_COMPANY_CODE)}&bizRoleAssignedEmployeeSeq=${seq}`;
}

/**
 * 승인 목록 HTML → [{ name, empNo, seq, pending, title, date, time }]
 * 대기 신청은 승인 버튼에 applyApprove(신청번호) 호출이 붙어 있고, 처리된 신청은 '승인 완료' 표시만 있다.
 */
export function parseList(html) {
  const items = [];
  for (const [, , block] of html.matchAll(/<li\b([^>]*)>([\s\S]*?)<\/li>/g)) {
    const text = (re) => ((block.match(re) || [])[1] || '').replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim();
    const staff = text(/class="txt_staff"[^>]*>([\s\S]*?)<\/span>/); // '최준영 11011009'
    if (!staff) continue;
    const m = staff.match(/^(.*?)\s*(\d+)?$/);
    const seq = (block.match(/applyApprove\(\s*['"]?(\d+)/) || [])[1] || '';
    items.push({
      name: m[1], empNo: m[2] || '', seq, pending: !!seq,
      status: text(/<button[^>]*>([\s\S]*?)<\/button>/), // '승인 완료' · '반려 완료' (대기건은 '승인')
      title: text(/class="txt_tit"[^>]*>([\s\S]*?)<\/strong>/),
      date: text(/class="txt_date"[^>]*>([\s\S]*?)<\/span>/),
      time: text(/<dd>([\s\S]*?)<\/dd>/),
    });
  }
  return items;
}

const label = (i) => `${i.date} ${i.time} ${i.title}`.trim();

/** 승인 목록에서 name(로그인한 직원 이름)으로 들어온 대기 신청 찾기 */
async function findMine(env, name) {
  if (!env.PAYCO_TEL_NO || !env.PAYCO_COMPANY_CODE) throw new Error('식권 승인 기능이 아직 설정되지 않았습니다. 관리자에게 알려주세요.');
  const me = norm(name);
  if (!me) throw new Error('이름 정보가 없습니다.');

  const q = await login(env);
  const list = parseList(await payco(`/bizCouponApplyApprove/list.nhn?${q}`));
  const mine = list.filter((i) => i.pending && norm(i.name) === me);
  // 같은 이름이 다른 사번으로 둘 이상 있으면 누구 것인지 알 수 없으므로 승인하지 않는다
  const empNos = new Set(mine.map((i) => i.empNo));
  if (empNos.size > 1) throw new Error(`'${name}' 이름의 신청이 서로 다른 사번(${[...empNos].join(', ')})으로 들어와 있어 자동 승인할 수 없습니다. 승인권자에게 직접 요청해 주세요.`);

  let message = '';
  if (!mine.length) {
    const done = list.find((i) => !i.pending && norm(i.name) === me);
    message = done
      ? `대기 중인 신청이 없습니다. 가장 최근 신청(${label(done)})은 '${done.status || '처리 완료'}' 상태입니다.`
      : '대기 중인 식권 신청이 없습니다. PAYCO에서 먼저 신청해 주세요.';
  }
  return { q, mine, message };
}

/** 조회만: 내 대기 신청 → { pending: [설명…], message } */
export async function checkMyMeal(env, name) {
  const { mine, message } = await findMine(env, name);
  return { pending: mine.map(label), message };
}

/** 내 대기 신청만 승인 → { approved: [설명…], failed: [설명…], message } */
export async function approveMyMeal(env, name) {
  const { q, mine, message } = await findMine(env, name);
  if (!mine.length) return { approved: [], failed: [], message };

  for (const i of mine) {
    await payco(`/bizCouponApplyApprove/applyApprove.nhn?${q}`, { bizCouponApplySeq: Number(i.seq) });
  }

  // 승인 응답은 팝업용 HTML이라 믿지 않고, 목록을 다시 읽어 실제로 처리됐는지 확인한다
  const after = parseList(await payco(`/bizCouponApplyApprove/list.nhn?${q}`));
  const stillPending = new Set(after.filter((i) => i.pending).map((i) => i.seq));
  const approved = mine.filter((i) => !stillPending.has(i.seq)).map(label);
  const failed = mine.filter((i) => stillPending.has(i.seq)).map(label);
  return {
    approved, failed,
    message: failed.length
      ? `${approved.length}건 승인, ${failed.length}건은 승인되지 않았습니다. 승인권자에게 직접 요청해 주세요.`
      : `식권 ${approved.length}건을 승인했습니다.`,
  };
}

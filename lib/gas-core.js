// 자동 생성 파일 — tools/gen-gas-core.mjs 가 Code.gs 로 만든다 (build.sh). 직접 고치지 마세요.
// Cloudflare(D1) 서버가 Apps Script와 똑같은 로직을 쓰도록 Code.gs 전체를 함수 하나로 감싼 것.
/* eslint-disable */
export function createGas({ SpreadsheetApp, Utilities, CacheService, PropertiesService, LockService, ScriptApp, ContentService, HtmlService, Session, TeamLog }) {
/**
 * 슈퍼SOL 전담반 일정관리 — Google Apps Script 백엔드
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
  // annual: 연도별 휴가 부여 JSON (일반휴가 일수 + 특별·대체휴가 부여 내역, parseAnnual_ 참고) — PM이 입력, 본인과 PM만 봄 (참고용)
  // ds: 'Y'면 DS직원 — 상단 식권(PAYCO) 승인 버튼 사용 (PM이 팀원 관리에서 체크. 실제 승인은 Cloudflare functions/api.js + lib/meal.js)
  //     (이 칸은 처음에 'meal'이라는 이름이었다. 칸 순서로 읽으므로 이름만 바뀌고 값은 그대로)
  users: { name: '일정앱_팀원', headers: ['email', 'name', 'role', 'createdAt', 'passwordHash', 'salt', 'mustChange', 'part', 'defaultSubs', 'viewAll', 'annual', 'ds', 'navOrder'] },
  leaves: {
    name: '일정앱_일정',
    headers: ['id', 'email', 'type', 'startDate', 'endDate', 'substitute', 'memo', 'createdAt', 'updatedAt', 'updatedBy'],
  },
  holidays: { name: '일정앱_공휴일', headers: ['date', 'name'] },
  events: {
    name: '일정앱_중요일정',
    headers: ['id', 'title', 'startDate', 'endDate', 'time', 'memo', 'createdAt', 'updatedAt', 'updatedBy', 'assignees'], // assignees: 담당자 이메일(쉼표)
  },
  // 파트별 업무 일정 (휴가 일정과 완전히 별도)
  tasks: {
    name: '일정앱_업무일정',
    headers: ['id', 'part', 'title', 'category', 'startDate', 'endDate', 'time', 'assignees', 'status', 'visibility', 'memo',
      'createdBy', 'createdAt', 'updatedAt', 'updatedBy', 'series'], // series: 반복 등록으로 함께 만든 일정끼리 같은 값
  },
  // 변경 이력 (PM '변경 이력' 탭). Cloudflare(DB) 서버는 매 요청마다 읽지 않도록 따로 보관한다 (historyStore_)
  history: { name: '일정앱_이력', headers: ['id', 'at', 'actor', 'kind', 'action', 'owner', 'summary'] },
  // 새 소식 — 사람마다 한 줄 (마지막으로 읽은 시각 + 최근 소식 목록 JSON)
  news: { name: '일정앱_소식', headers: ['email', 'seenAt', 'items'] },
  // 요청함: targets·questions는 JSON (requestsState_ 참고). status: 'open' | 'closed', shared: 'Y'면 대상자끼리 응답 공개
  requests: { name: '일정앱_요청함', headers: ['id', 'title', 'desc', 'targets', 'questions', 'deadline', 'status', 'shared', 'createdBy', 'createdAt', 'updatedAt'] },
  // 응답: 사람마다 요청당 한 줄. answers는 {질문id: 값} JSON, taskId는 '시간대 고르기'로 자동 등록된 업무 일정
  responses: { name: '일정앱_응답', headers: ['id', 'requestId', 'email', 'answers', 'createdAt', 'updatedAt', 'taskId'] },
  // 사다리 추첨: 회차마다 한 줄 (후보 · 제외 · 당첨 · 설정은 JSON)
  early: { name: '일정앱_조기퇴근', headers: ['id', 'email', 'date', 'memo', 'createdAt', 'createdBy', 'updatedAt'] }, // 1시간 조기퇴근권 사용 기록 (PM만)
  draws: { name: '일정앱_추첨', headers: ['id', 'kind', 'targetId', 'title', 'date', 'round', 'pool', 'candidates', 'excluded', 'quota', 'picked', 'fixed', 'note', 'by', 'at'] },
};

const LEAVE_TYPES = ['연차', '오전반차', '오후반차', '반반차', '병가', '연수/교육', '기타'];
const CROWD_PERCENT = 20; // 총원(테스터 제외) 중 이 비율 이상이 같은 날 부재하면 휴가 등록 창에 안내 (사유 필수는 아님)
const HALF_TYPES = ['오전반차', '오후반차', '반반차']; // 하루짜리 (반반차는 오후에만, 0.25일 — 일수 계산은 화면에서)
const WEEK = ['일', '월', '화', '수', '목', '금', '토'];
const PARTS = ['기획', 'QA', '개발', '퍼블', 'NP', '디자인'];
const TASK_PARTS = ['전체'].concat(PARTS, ['PM']); // 업무 일정의 파트 ('전체' · 'PM' 업무는 PM만 등록. '전체'는 모든 파트 공통 업무)
const ROLES = ['pm', 'pl', 'member'];
const MAX_SUBSTITUTES = 3;
const TASK_CATEGORIES = ['개발', '배포', '테스트', '회의/리뷰', '마감', '인프라', '기타'];
const TASK_STATUSES = ['예정', '진행 중', '완료', '지연'];
const MEMO_MAX = 2000; // 업무 · 중요 일정 메모 글자 수 (정기배포 업무 내역 등 길게)
const MAX_ASSIGNEES = 10;
const MAX_EVENT_ASSIGNEES = 30; // 중요 일정 담당자
const ANNUAL_KINDS = ['일반휴가', '특별휴가', '대체휴가']; // PM이 입력하는 부여 일수 구분 (사용은 연차·반차에서 합쳐서 차감)
const GRANT_KINDS = ['특별휴가', '대체휴가']; // 생길 때마다 한 줄씩 쌓는 부여 내역 (날짜·일수·사유)
const HISTORY_DAYS = 90; // 변경 이력·새 소식 보관 기간
const NEWS_MAX = 20; // 사람마다 보관하는 새 소식 개수
const REQ_QTYPES = ['text', 'long', 'number', 'link', 'choice', 'multi', 'date', 'slot']; // 요청함 질문 유형
const REQ_MAX_QUESTIONS = 20;
const REQ_MAX_SLOTS = 200;
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
    .setTitle('슈퍼SOL 전담반 일정')
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
  getTasks, saveTask, deleteTask, checkTaskAbsence, extendTaskSeries,
  getRequests, saveRequest, setRequestStatus, deleteRequest, remindRequest, submitResponse, cancelResponse,
  runDraw, replaceDraw, getDraw,
  getEarly, saveEarly, deleteEarly,
  getBackupInfo, getHistory, markNewsRead, setAnnualDays, setAnnualGrants, bulkAddLeaves, setDsStaff, setViewAll, setMyNavOrder, bulkAddHolidays,
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
  const warnings = draft.substitutes.length ? conflictsOf_(draft, leaves, ctx.users) : [];
  const crowded = crowdedDays_(draft, leaves, ctx.users, readAll_('holidays'));
  return {
    duplicate: dup ? `이미 ${rangeText_(dup)} ${dup.type} 일정이 있어 등록할 수 없습니다.` : '',
    warnings,
    crowded,
    reasonRequired: reasonRequired_(draft, warnings),
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
      if (!isAdmin && existing.endDate < today_()) throw new Error('지난 일정은 PM만 수정할 수 있습니다.');
    }

    const l = normalizeLeave_(ctx, input, existing);
    // 대직자 필수. 단 PM이 지난 일정을 고칠 때는 비워도 된다 (일괄 등록한 예전 기록은 대직자가 없을 수 있음)
    if (!l.substitutes.length && !(isAdmin && l.endDate < today_())) throw new Error('대직자를 1명 이상 지정해야 합니다.');
    if (l.substitutes.length > MAX_SUBSTITUTES) throw new Error(`대직자는 최대 ${MAX_SUBSTITUTES}명까지 지정할 수 있습니다.`);
    if (l.substitutes.indexOf(l.email) >= 0) throw new Error('본인을 대직자로 지정할 수 없습니다.');
    if (l.substitutes.some((e) => !ctx.users.some((u) => u.email === e))) throw new Error('대직자가 팀원 목록에 없습니다.');
    const dup = duplicateOf_(l, leaves);
    if (dup) throw new Error(`이미 등록된 일정과 겹칩니다: ${rangeText_(dup)} ${dup.type}`);

    const memo = String(input.memo || '').slice(0, 200);
    if (!memo.trim() && reasonRequired_(l, l.substitutes.length ? conflictsOf_(l, leaves, ctx.users) : [])) {
      throw new Error('대직자와 일정이 겹쳐서 사유를 꼭 적어야 합니다.');
    }
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
    if (!isPM_(ctx) && leave.endDate < today_()) throw new Error('지난 일정은 PM만 삭제할 수 있습니다.');
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
    const desc = (u) => `${u.name} · ${ROLE_LABELS[u.role] || u.role}${u.part ? ' · ' + u.part : ''}${u.viewAll === 'Y' ? ' · 전체 일정 조회' : ''}${u.ds === 'Y' ? ' · DS직원' : ''}`;
    const ds = input.ds === undefined ? (existing ? existing.ds || '' : '') : input.ds ? 'Y' : '';
    if (existing) {
      const before = desc(existing);
      Object.assign(existing, { name, role, part, defaultSubs, viewAll, ds });
      upsert_('users', existing);
      if (before !== desc(existing)) logChange_(ctx, 'user', '정보 수정', email, `${before} → ${desc(existing)}`);
    } else {
      temp = tempPassword_();
      const salt = Utilities.getUuid();
      upsert_('users', {
        email, name, role, part, defaultSubs, viewAll, ds, createdAt: new Date().toISOString(),
        passwordHash: hash_(temp, salt), salt, mustChange: 'Y',
      });
      logChange_(ctx, 'user', '추가', email, desc({ name, role, part, viewAll, ds }));
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

/** 내 메뉴 순서 (탭 키 목록). 권한은 그대로 — 보이는 탭만 이 순서로 정렬된다. 빈 목록이면 기본 순서 */
const NAV_TABS = ['cal', 'events', 'work', 'mine', 'requests', 'all', 'members', 'early', 'holidays', 'history'];
function setMyNavOrder(token, order) {
  return withLock_(() => {
    const ctx = auth_(token);
    const keys = (Array.isArray(order) ? order : []).map(String).filter((k, i, a) => NAV_TABS.indexOf(k) >= 0 && a.indexOf(k) === i);
    ctx.user.navOrder = keys.join(',');
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
 * 휴가 부여 일수 — 연도별. 일반휴가는 숫자 하나, 특별휴가·대체휴가는 생길 때마다 한 줄씩 쌓는 '부여 내역'
 * 저장: '{"2026":{"일반휴가":15,"내역":[{"kind":"대체휴가","days":1,"date":"2026-10-01","note":"9/20 주말 근무"}]}}'
 * 읽으면 내역의 종류별 합계를 특별휴가·대체휴가 숫자로 함께 돌려준다 (화면은 합계로 남은 휴가 계산)
 * 예전 형식도 읽는다: 숫자 하나({"2026":15}) → 일반휴가, {"특별휴가":2,"사유":{...}} → 내역 한 줄
 */
function parseAnnual_(v) {
  try {
    const o = JSON.parse(v || '{}');
    const out = {};
    Object.keys(o).forEach((y) => {
      if (!/^\d{4}$/.test(y)) return;
      const src = typeof o[y] === 'number' ? { 일반휴가: o[y] } : o[y] && typeof o[y] === 'object' ? o[y] : {};
      const k = {};
      if (typeof src.일반휴가 === 'number') k.일반휴가 = src.일반휴가;
      let list = [];
      if (Array.isArray(src.내역)) {
        list = src.내역.filter((g) => g && GRANT_KINDS.indexOf(g.kind) >= 0 && typeof g.days === 'number')
          .map((g) => ({ kind: g.kind, days: g.days, date: isDate_(String(g.date || '')) ? g.date : '', note: String(g.note || '') }));
      } else {
        GRANT_KINDS.forEach((kind) => {
          if (typeof src[kind] === 'number') list.push({ kind, days: src[kind], date: '', note: (src.사유 && typeof src.사유[kind] === 'string' && src.사유[kind]) || '' });
        });
      }
      if (list.length) {
        k.내역 = list;
        GRANT_KINDS.forEach((kind) => {
          const mine = list.filter((g) => g.kind === kind);
          if (mine.length) k[kind] = mine.reduce((sum, g) => sum + g.days, 0);
        });
      }
      if (Object.keys(k).length) out[y] = k;
    });
    return out;
  } catch (e) {
    return {};
  }
}

/** parseAnnual_ 결과 → 저장 문자열 (합계는 저장하지 않고 내역만) */
function storeAnnual_(map) {
  const o = {};
  Object.keys(map).forEach((y) => {
    const k = map[y];
    const x = {};
    if (typeof k.일반휴가 === 'number') x.일반휴가 = k.일반휴가;
    if (k.내역 && k.내역.length) x.내역 = k.내역;
    if (Object.keys(x).length) o[y] = x;
  });
  return Object.keys(o).length ? JSON.stringify(o) : '';
}

/** 내역(날짜·사유)을 뺀 합계만 (전체 일정 조회 권한자에게 보낼 때 — 내역은 본인과 PM만) */
function annualNoNotes_(v) {
  const map = parseAnnual_(v);
  Object.keys(map).forEach((y) => delete map[y].내역);
  return map;
}

/** 이력용 문구: '일반 15 · 대체 10/1 1일(9/20 주말 근무)' */
function annualDesc_(k) {
  if (!k) return '-';
  const md = (d) => (d ? Number(d.slice(5, 7)) + '/' + Number(d.slice(8)) + ' ' : '');
  const parts = [];
  if (typeof k.일반휴가 === 'number') parts.push('일반 ' + k.일반휴가);
  (k.내역 || []).forEach((g) => parts.push(`${g.kind.slice(0, 2)} ${md(g.date)}${g.days}일${g.note ? `(${g.note})` : ''}`));
  return parts.join(' · ') || '-';
}

function checkDays_(n, max, label) {
  if (!isFinite(n) || n < 0 || n > max || Math.round(n * 4) !== n * 4) throw new Error(`${label} 일수는 0~${max}, 0.25일 단위로 입력하세요.`); // 반반차(0.25일)까지
  return n;
}

/**
 * PM: 일반휴가 부여 일수 입력 (여러 명 한 번에). days = { email: 15 } 또는 { email: { 일반휴가: 15 } }. 빈칸('')이면 지운다. 0~60일, 0.25일 단위.
 * 특별휴가·대체휴가는 setAnnualGrants 로 내역을 따로 관리한다.
 * 본인과 PM만 볼 수 있고, 남은 휴가는 화면에서 '일반 + 특별 + 대체 합계 − 사용(연차·반차, 예정 포함)'으로 계산한다 (참고용)
 */
function setAnnualDays(token, year, days) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    year = String(year || '');
    if (!/^20\d{2}$/.test(year)) throw new Error('연도를 확인하세요.');
    if (!days || typeof days !== 'object') throw new Error('입력한 값이 없습니다.');
    const changes = [];
    Object.keys(days).forEach((email) => {
      const u = ctx.users.find((x) => x.email === String(email).toLowerCase());
      if (!u) return;
      const raw0 = days[email] && typeof days[email] === 'object' ? days[email].일반휴가 : days[email];
      if (raw0 === undefined) return;
      const map = parseAnnual_(u.annual);
      const before = annualDesc_(map[year]);
      const k = map[year] || {};
      const raw = String(raw0 == null ? '' : raw0).trim();
      if (raw === '') delete k.일반휴가;
      else k.일반휴가 = checkDays_(Number(raw), 60, `${u.name}: 일반휴가`);
      map[year] = k;
      const after = annualDesc_(k);
      if (before === after) return;
      u.annual = storeAnnual_(map);
      upsert_('users', u);
      changes.push(`${u.name} ${before} → ${after}`);
    });
    if (changes.length) logChange_(ctx, 'user', `${year}년 휴가 일수`, '', changes.join(', '));
    return { state: state_(ctx), changed: changes.length };
  });
}

/**
 * PM: 한 사람의 특별휴가·대체휴가 부여 내역을 통째로 저장.
 * grants = [{ kind: '특별휴가'|'대체휴가', days: 1, date: 'yyyy-MM-dd'(생긴 날, 선택), note: '사유'(50자, 선택) }] — 빈 배열이면 모두 지움
 */
function setAnnualGrants(token, year, email, grants) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    year = String(year || '');
    if (!/^20\d{2}$/.test(year)) throw new Error('연도를 확인하세요.');
    const u = ctx.users.find((x) => x.email === String(email || '').toLowerCase());
    if (!u) throw new Error('팀원을 찾을 수 없습니다.');
    if (!Array.isArray(grants)) throw new Error('입력한 값이 없습니다.');
    if (grants.length > 50) throw new Error('한 해에 50줄까지 입력할 수 있습니다.');
    const list = grants.map((g, i) => {
      const kind = String((g && g.kind) || '');
      if (GRANT_KINDS.indexOf(kind) < 0) throw new Error(`${i + 1}번째 줄: 종류를 고르세요.`);
      const days = checkDays_(Number(g.days), 30, `${i + 1}번째 줄:`);
      if (!days) throw new Error(`${i + 1}번째 줄: 일수를 입력하세요.`);
      const date = String(g.date || '');
      if (date && !isDate_(date)) throw new Error(`${i + 1}번째 줄: 날짜 형식이 올바르지 않습니다.`);
      return { kind, days, date, note: String(g.note || '').trim().slice(0, 50) };
    }).sort((a, b) => (a.date || '9999').localeCompare(b.date || '9999'));
    const map = parseAnnual_(u.annual);
    const before = annualDesc_(map[year]);
    const k = map[year] || {};
    if (list.length) k.내역 = list;
    else delete k.내역;
    map[year] = k;
    u.annual = storeAnnual_(map);
    const after = annualDesc_(parseAnnual_(u.annual)[year]);
    if (before !== after) {
      upsert_('users', u);
      logChange_(ctx, 'user', `${year}년 휴가 부여 내역`, u.email, `${before} → ${after}`);
    }
    return { state: state_(ctx), changed: before !== after ? 1 : 0 };
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

/** PM: 팀원 관리 목록의 'DS직원' 체크 (DS직원에게 상단 식권 승인 버튼) */
function setDsStaff(token, email, on) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    const u = ctx.users.find((x) => x.email === String(email || '').toLowerCase());
    if (!u) throw new Error('팀원을 찾을 수 없습니다.');
    const ds = on ? 'Y' : '';
    if ((u.ds || '') !== ds) {
      u.ds = ds;
      upsert_('users', u);
      logChange_(ctx, 'user', on ? 'DS직원 지정' : 'DS직원 해제', u.email, u.name);
    }
    return { state: state_(auth_(token)) };
  });
}

/** 팀원관리 목록에서 '전체 조회' 켜고 끄기 (PM만, 파트원에게만 의미 있음 — PM·PL은 원래 전체를 봄) */
function setViewAll(token, email, on) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    const u = ctx.users.find((x) => x.email === String(email || '').toLowerCase());
    if (!u) throw new Error('팀원을 찾을 수 없습니다.');
    if (u.role !== 'member') throw new Error('PM·PL은 원래 전체 일정을 볼 수 있습니다.');
    const v = on ? 'Y' : '';
    if ((u.viewAll || '') !== v) {
      u.viewAll = v;
      upsert_('users', u);
      logChange_(ctx, 'user', on ? '전체 조회 지정' : '전체 조회 해제', u.email, u.name);
    }
    return { state: state_(auth_(token)) };
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
 * 조회: '전체 공개'는 모두, '파트 내부'는 그 파트 사람과 PM, 'DS직원만'은 DS직원과 PM, '담당자만'은 담당자와 등록한 사람,
 *       '나만 보기'는 등록한 사람만 (PM도 못 봄)
 */

function taskVisible_(ctx, t) {
  if (t.visibility === 'private') return t.createdBy === ctx.email;
  if (t.visibility === 'ds') return isPM_(ctx) || ctx.user.ds === 'Y' || t.createdBy === ctx.email; // DS직원만 (+PM)
  if (t.visibility === 'assignees') return t.createdBy === ctx.email || (t.assignees || []).indexOf(ctx.email) >= 0; // 담당자만 (등록한 사람 포함, PM도 담당이 아니면 못 봄)
  return t.visibility !== 'part' || isPM_(ctx) || t.part === '전체' || ctx.user.part === t.part; // '전체' 파트 업무는 파트 내부여도 모두
}

function taskEditable_(ctx, t) {
  if (t.visibility === 'private') return t.createdBy === ctx.email;
  if (!taskVisible_(ctx, t)) return false;
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
      const half = l.type === '오전반차' ? ' 오전' : l.type === '오후반차' || l.type === '반반차' ? ' 오후' : '';
      out.push(`${nameOf(email)}님 ${rangeText_({ startDate: s, endDate: e })}${half} 부재`);
    });
  });
  return out;
}

function tasksState_(ctx) {
  const visibleLeaves = visibleLeaves_(ctx, readAll_('leaves'));
  const draws = latestDraws_();
  return {
    tasks: readAll_('tasks').filter((t) => taskVisible_(ctx, t)).map((t) => ({
      drawId: draws['task:' + t.id] || '',
      id: t.id, part: t.part, title: t.title, category: t.category, startDate: t.startDate, endDate: t.endDate,
      time: t.time, assignees: t.assignees, status: t.status || '예정', visibility: t.visibility || 'all', memo: t.memo,
      createdBy: t.createdBy, updatedAt: t.updatedAt, canEdit: taskEditable_(ctx, t), series: t.series || '',
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
    if (TASK_PARTS.indexOf(part) < 0) throw new Error(isPM_(ctx) ? '파트를 선택하세요.' : '파트가 지정되지 않은 계정입니다. PM에게 파트 지정을 요청하세요.');

    const title = String(input.title || '').trim().slice(0, 60);
    const category = String(input.category || '');
    const startDate = String(input.startDate || '');
    const endDate = String(input.endDate || startDate);
    if (!title) throw new Error('제목을 입력하세요.');
    if (TASK_CATEGORIES.indexOf(category) < 0) throw new Error('분류를 선택하세요.');
    if (!isDate_(startDate) || !isDate_(endDate)) throw new Error('날짜 형식이 올바르지 않습니다.');
    if (endDate < startDate) throw new Error('종료일이 시작일보다 빠릅니다.');
    const visibility = ['all', 'part', 'private', 'ds', 'assignees'].indexOf(input.visibility) >= 0 ? input.visibility : 'all';
    if (visibility === 'ds' && !isPM_(ctx) && ctx.user.ds !== 'Y') throw new Error("'DS직원만'은 DS직원과 PM만 설정할 수 있습니다.");
    if (visibility === 'private' && (existing ? existing.createdBy : ctx.email) !== ctx.email) throw new Error("'나만 보기'는 등록한 사람만 설정할 수 있습니다.");
    // 나만 보기: 담당자는 등록한 사람으로 고정
    const assignees = visibility === 'private' ? [existing ? existing.createdBy : ctx.email] : parseSubs_(input.assignees);
    if (assignees.length > MAX_ASSIGNEES) throw new Error(`담당자는 최대 ${MAX_ASSIGNEES}명까지 지정할 수 있습니다.`);
    if (assignees.some((e) => !ctx.users.some((u) => u.email === e))) throw new Error('담당자가 팀원 목록에 없습니다.');
    const status = TASK_STATUSES.indexOf(input.status) >= 0 ? input.status : existing ? existing.status || '예정' : '예정';

    const now = new Date().toISOString();
    // 반복 등록 (새 일정만): 날짜마다 한 건씩 만들고 같은 series로 묶는다
    if (!existing && input.repeat && input.repeat.freq) {
      const holidays = {};
      readAll_('holidays').forEach((h) => (holidays[h.date] = true));
      const dates = repeatDates_(startDate, endDate, input.repeat, holidays);
      const series = Utilities.getUuid();
      const rows = dates.map(([s1, e1]) => ({
        id: Utilities.getUuid(), part, title, category, startDate: s1, endDate: e1,
        time: String(input.time || '').trim().slice(0, 30), assignees: assignees.join(','), status, visibility,
        memo: String(input.memo || '').trim().slice(0, MEMO_MAX), createdBy: ctx.email, createdAt: now, updatedAt: now, updatedBy: ctx.email, series,
      }));
      const sh = sheet_('tasks');
      const values = rows.map((o) => SHEETS.tasks.headers.map((h) => String(o[h] == null ? '' : o[h])));
      sh.getRange(sh.getLastRow() + 1, 1, values.length, values[0].length).setNumberFormat('@').setValues(values);
      const first = Object.assign({}, rows[0], { assignees });
      const how = repeatLabel_(input.repeat);
      logChange_(ctx, 'task', '반복 등록', '', `${taskDesc_(ctx, null, first)} · ${how} · ${rows.length}회 (~${rangeText_({ startDate: rows[rows.length - 1].startDate, endDate: rows[rows.length - 1].startDate })})`);
      // 담당자에게는 한 번만 알림 (볼 수 있는 사람만)
      const see = (e) => { const u = ctx.users.find((x) => x.email === e); return u && taskVisible_({ email: e, user: u, users: ctx.users }, first); };
      if (visibility !== 'private') pushNews_(ctx, assignees.filter(see), `반복 업무 배정: ${title} · ${how} · ${rows.length}회 (${rangeText_(first)}부터)`, 'tab:work');
      return tasksState_(ctx);
    }
    const saved = {
      id: existing ? existing.id : Utilities.getUuid(),
      part, title, category, startDate, endDate,
      time: String(input.time || '').trim().slice(0, 30),
      assignees: assignees.join(','),
      status,
      visibility,
      memo: String(input.memo || '').trim().slice(0, MEMO_MAX),
      createdBy: existing ? existing.createdBy : ctx.email,
      createdAt: existing ? existing.createdAt : now,
      updatedAt: now,
      updatedBy: ctx.email,
      series: existing ? existing.series || '' : '',
    };
    upsert_('tasks', saved);
    // 반복 일정 수정 범위: 'after'(이 날짜 이후) · 'all'(전체). 날짜·상태는 이 일정에만, 나머지 내용은 같은 반복에 함께 반영
    const scope = input.scope === 'after' || input.scope === 'all' ? input.scope : '';
    const others = existing && existing.series && scope
      ? tasks.filter((x) => x.series === existing.series && x.id !== existing.id && (scope === 'all' || x.startDate >= existing.startDate) && taskEditable_(ctx, x))
      : [];
    others.forEach((x) => upsert_('tasks', Object.assign({}, x, {
      part, title, category, time: saved.time, assignees: saved.assignees, visibility, memo: saved.memo, updatedAt: now, updatedBy: ctx.email,
    })));
    saved.assignees = assignees;
    if (others.length) logChange_(ctx, 'task', '반복 수정', '', `${taskDesc_(ctx, existing, saved)} · ${scope === 'all' ? '전체' : '이후'} ${others.length + 1}건`);
    else logChange_(ctx, 'task', existing ? '수정' : '등록', '', taskDesc_(ctx, existing, saved));
    taskNews_(ctx, existing, saved);
    return tasksState_(ctx);
  });
}

/** scope: 'after'면 같은 반복의 이 날짜 이후 일정까지 모두 (수정 권한이 있는 것만) */
function deleteTask(token, id, scope) {
  return withLock_(() => {
    const ctx = auth_(token);
    const all = readAll_('tasks');
    const t = all.find((x) => x.id === id);
    if (!t) throw new Error('이미 삭제된 업무 일정입니다.');
    if (!taskEditable_(ctx, t)) throw new Error('작성자, 해당 파트 PL, PM만 삭제할 수 있습니다.');
    if (scope === 'after' && t.series) {
      const targets = all.filter((x) => x.series === t.series && x.startDate >= t.startDate && taskEditable_(ctx, x));
      targets.forEach((x) => deleteRow_('tasks', x.id));
      logChange_(ctx, 'task', '반복 삭제', '', `${taskDesc_(ctx, null, t)} 부터 ${targets.length}건`);
      const see = (e) => { const u = ctx.users.find((x) => x.email === e); return u && taskVisible_({ email: e, user: u, users: ctx.users }, t); };
      if (t.endDate >= today_() && t.visibility !== 'private') pushNews_(ctx, (t.assignees || []).filter(see), `반복 업무 삭제: ${t.title} (${rangeText_(t)}부터 ${targets.length}건)`, 'tab:work');
      return tasksState_(ctx);
    }
    deleteRow_('tasks', id);
    logChange_(ctx, 'task', '삭제', '', taskDesc_(ctx, null, t));
    taskNews_(ctx, t, null);
    return tasksState_(ctx);
  });
}

/**
 * 반복 연장: 같은 반복의 마지막 회차 다음부터 repeat.until까지 이어서 만든다 (같은 series).
 * 내용(파트·제목·분류·시간·담당자·공개 범위·메모·기간)은 마지막 회차를 따르고 상태는 '예정'.
 * 규칙 기준일: 매월이면 가장 많이 쓰인 날짜(휴일로 옮겨진 회차 제외), 그 밖에는 첫 회차.
 */
function extendTaskSeries(token, id, repeat) {
  return withLock_(() => {
    const ctx = auth_(token);
    const tasks = readAll_('tasks');
    const t = tasks.find((x) => x.id === id);
    if (!t) throw new Error('업무 일정을 찾을 수 없습니다. 새로고침 후 다시 시도하세요.');
    if (!t.series) throw new Error('반복 일정이 아닙니다.');
    if (!taskEditable_(ctx, t)) throw new Error('작성자, 해당 파트 PL, PM만 수정할 수 있습니다.');
    const rep = repeat || {};
    const list = tasks.filter((x) => x.series === t.series).sort((a, b) => (a.startDate < b.startDate ? -1 : a.startDate > b.startDate ? 1 : 0));
    const last = list[list.length - 1];
    let base = list[0];
    if (rep.freq === 'monthly') {
      const cnt = {};
      list.forEach((x) => (cnt[x.startDate.slice(8)] = (cnt[x.startDate.slice(8)] || 0) + 1));
      const day = Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a] || (a < b ? -1 : 1))[0];
      base = list.find((x) => x.startDate.slice(8) === day);
    }
    const DAY = 864e5;
    const toT = (ds) => Date.UTC(+ds.slice(0, 4), +ds.slice(5, 7) - 1, +ds.slice(8, 10));
    const span = toT(last.endDate) - toT(last.startDate);
    const holidays = {};
    readAll_('holidays').forEach((h) => (holidays[h.date] = true));
    const have = {};
    list.forEach((x) => (have[x.startDate] = true));
    const dates = repeatDates_(base.startDate, new Date(toT(base.startDate) + span).toISOString().slice(0, 10), rep, holidays, last.startDate)
      .filter(([s1]) => !have[s1]);
    if (!dates.length) throw new Error('반복 조건에 맞는 날짜가 없습니다.');
    const now = new Date().toISOString();
    const rows = dates.map(([s1, e1]) => Object.assign({}, last, {
      id: Utilities.getUuid(), startDate: s1, endDate: e1, status: '예정', createdAt: now, updatedAt: now, updatedBy: ctx.email,
    }));
    const sh = sheet_('tasks');
    const values = rows.map((o) => SHEETS.tasks.headers.map((h) => String(o[h] == null ? '' : o[h])));
    sh.getRange(sh.getLastRow() + 1, 1, values.length, values[0].length).setNumberFormat('@').setValues(values);
    const first = Object.assign({}, rows[0], { assignees: parseSubs_(last.assignees) });
    const end = rows[rows.length - 1];
    const how = repeatLabel_(rep);
    logChange_(ctx, 'task', '반복 연장', '', `${taskDesc_(ctx, null, first)} · ${how} · ${rows.length}회 추가 (~${rangeText_({ startDate: end.startDate, endDate: end.startDate })})`);
    const see = (e) => { const u = ctx.users.find((x) => x.email === e); return u && taskVisible_({ email: e, user: u, users: ctx.users }, first); };
    if (first.visibility !== 'private') pushNews_(ctx, first.assignees.filter(see), `반복 업무 연장: ${first.title} · ${how} · ${rows.length}회 (~${rangeText_({ startDate: end.startDate, endDate: end.startDate })})`, 'tab:work');
    return tasksState_(ctx);
  });
}

/**
 * 반복 날짜 만들기. repeat = { freq: 'weekday'(매일·평일) | 'weekly' | 'biweekly' | 'monthly', days: [0~6](주간), until: 'yyyy-MM-dd',
 *   offDay: 쉬는 날에 걸리면 'skip'(건너뛰기) | 'next'(다음 영업일) | 'prev'(이전 영업일) | 'keep'(그대로) }  (예전 화면: skipHolidays → skip)
 * 쉬는 날 = 공휴일, 그리고 매월 반복이면 주말도 (매주·격주는 고른 요일을 그대로 존중)
 * 첫 일정의 기간(시작~종료 일수)을 그대로 유지한다. 최대 1년, 100회까지.
 * after: 반복 연장용 — 기준 일정(startDate)의 규칙은 그대로, 이 날짜 다음 회차부터만 만든다 (1년·100회는 after부터 셈)
 */
function repeatDates_(startDate, endDate, rep, holidays, after) {
  const DAY = 864e5;
  const toT = (ds) => Date.UTC(+ds.slice(0, 4), +ds.slice(5, 7) - 1, +ds.slice(8, 10));
  const toS = (t) => new Date(t).toISOString().slice(0, 10);
  const freq = String(rep.freq || '');
  if (['weekday', 'weekly', 'biweekly', 'monthly'].indexOf(freq) < 0) throw new Error('반복 방식을 확인하세요.');
  const until = String(rep.until || '');
  if (!isDate_(until) || until < startDate || (after && until <= after)) throw new Error('반복 종료일을 확인하세요.');
  const s0 = toT(startDate), u = toT(until), span = toT(endDate) - s0;
  if (u - (after ? toT(after) : s0) > 366 * DAY) throw new Error('반복은 최대 1년까지 등록할 수 있습니다.');
  const days = (Array.isArray(rep.days) ? rep.days : []).map(Number).filter((d) => d >= 0 && d <= 6);
  const weekdays = days.length ? days : [new Date(s0).getUTCDay()];
  const week0 = s0 - new Date(s0).getUTCDay() * DAY; // 첫 주 일요일
  const mode = ['skip', 'next', 'prev', 'keep'].indexOf(rep.offDay) >= 0 ? rep.offDay : rep.skipHolidays ? 'skip' : 'keep';
  const weekend = (tt) => { const w = new Date(tt).getUTCDay(); return w === 0 || w === 6; };
  const isOff = (tt) => !!holidays[toS(tt)] || (freq === 'monthly' && weekend(tt));
  const business = (tt) => !weekend(tt) && !holidays[toS(tt)];
  const seen = {};
  const out = [];
  for (let t = s0; t <= u; t += DAY) {
    const d = new Date(t), dow = d.getUTCDay();
    let hit = false;
    if (freq === 'weekday') hit = dow >= 1 && dow <= 5;
    else if (freq === 'weekly') hit = weekdays.indexOf(dow) >= 0;
    else if (freq === 'biweekly') hit = weekdays.indexOf(dow) >= 0 && Math.floor((t - week0) / (7 * DAY)) % 2 === 0;
    else hit = d.getUTCDate() === new Date(s0).getUTCDate();
    if (!hit) continue;
    let at = t;
    if (isOff(t)) {
      if (mode === 'skip') continue;
      if (mode === 'next') while (!business(at)) at += DAY;
      if (mode === 'prev') while (!business(at)) at -= DAY;
    }
    if (seen[at]) continue; // 옮긴 날짜가 다른 회차와 겹치면 하나만
    seen[at] = true;
    if (after && toS(at) <= after) continue;
    out.push([toS(at), toS(at + span)]);
    if (out.length > 100) throw new Error('반복은 한 번에 100회까지 등록할 수 있습니다. 종료일을 앞당겨 주세요.');
  }
  if (!out.length) throw new Error('반복 조건에 맞는 날짜가 없습니다.');
  return out;
}

function repeatLabel_(rep) {
  const W = ['일', '월', '화', '수', '목', '금', '토'];
  const days = (Array.isArray(rep.days) ? rep.days : []).map(Number).sort().map((d) => W[d]).join('·');
  return rep.freq === 'weekday' ? '매일(평일)' : rep.freq === 'monthly' ? '매월' : `${rep.freq === 'biweekly' ? '격주' : '매주'}${days ? ' ' + days : ''}`;
}

/* ───────────── 요청함 (도서구매 신청 · 면담 일정 취합 등) ───────────── */
/*
 * PM이 질문을 만들어 대상자에게 요청하고, 대상자가 응답한다.
 * - 응답은 본인과 PM만 본다. '응답 공개'를 켜면 대상자끼리도 응답을 본다 (예: 도서 중복 구매 방지).
 * - 질문이 바뀌면 기존 응답과 맞지 않으므로, 응답이 하나라도 있으면 질문은 고칠 수 없다 (제목·설명·마감일·대상·공개는 가능).
 * - '시간대 고르기'(slot): 시간대마다 정원이 있고 선착순. 고르면 만든 사람과 응답자를 담당자로 하는 업무 일정('담당자만' 공개)이 자동으로 생긴다.
 * - 이력에는 요청 등록·수정·마감·삭제만 남기고 응답 내용은 남기지 않는다.
 */
function reqParse_(v, fallback) {
  try { const x = JSON.parse(v || ''); return x == null ? fallback : x; } catch (e) { return fallback; }
}

function reqOpen_(r) {
  return r.status !== 'closed' && today_() <= r.deadline;
}

/** 대상자 이메일 — 지금 팀원 기준 (새로 들어온 사람도 포함). 전체·파트는 테스터를 빼고, 만든 사람은 '전체'일 때만 뺀다 (직접 지정·파트로 고르면 본인도 응답) */
function reqTargets_(ctx, r) {
  const t = reqParse_(r.targets, { mode: 'all' });
  let list;
  if (t.mode === 'people') list = parseSubs_(t.emails).filter((e) => ctx.users.some((u) => u.email === e));
  else list = ctx.users.filter((u) => !isTester_(u) && (t.mode !== 'parts' || (t.parts || []).indexOf(u.part || (u.role === 'pm' ? 'PM' : '')) >= 0)).map((u) => u.email);
  return t.mode === 'all' ? list.filter((e) => e !== r.createdBy) : list;
}

function reqResponses_(id) {
  return readAll_('responses').filter((x) => x.requestId === id);
}

function reqSlotQ_(questions) {
  return questions.find((q) => q.type === 'slot') || null;
}

/** 요청함 목록 (내가 대상이거나 PM). PM에게는 응답 전체와 미응답자, 대상자에게는 내 응답 · 시간대 남은 자리 (공개면 응답 목록) */
function requestsState_(ctx) {
  const pm = isPM_(ctx);
  const all = readAll_('responses');
  const nameOf = (e) => nameOf_(ctx, e);
  const partOf = (e) => (ctx.users.find((u) => u.email === e) || {}).part || '';
  const requests = readAll_('requests')
    .map((r) => {
      const targets = reqTargets_(ctx, r);
      const isTarget = targets.indexOf(ctx.email) >= 0;
      if (!pm && !isTarget) return null;
      const questions = reqParse_(r.questions, []);
      const resps = all.filter((x) => x.requestId === r.id);
      const mine = resps.find((x) => x.email === ctx.email);
      const sq = reqSlotQ_(questions);
      const taken = {};
      if (sq) resps.forEach((x) => { const v = reqParse_(x.answers, {})[sq.id]; if (v) taken[v] = (taken[v] || 0) + 1; });
      const shared = r.shared === 'Y';
      const showAll = pm || shared;
      return {
        id: r.id, title: r.title, desc: r.desc, deadline: r.deadline, status: r.status === 'closed' ? 'closed' : 'open', open: reqOpen_(r),
        shared, targets: reqParse_(r.targets, { mode: 'all' }), questions, createdBy: r.createdBy, createdByName: nameOf(r.createdBy),
        createdAt: r.createdAt, updatedAt: r.updatedAt, isTarget, canManage: pm,
        targetCount: targets.length, responseCount: resps.filter((x) => targets.indexOf(x.email) >= 0).length,
        mine: mine ? { answers: reqParse_(mine.answers, {}), updatedAt: mine.updatedAt } : null,
        slotTaken: taken,
        responses: showAll ? resps.map((x) => ({ email: x.email, name: nameOf(x.email), part: partOf(x.email), answers: reqParse_(x.answers, {}), updatedAt: x.updatedAt })) : [],
        pending: pm ? targets.filter((e) => !resps.some((x) => x.email === e)).map((e) => ({ email: e, name: nameOf(e), part: partOf(e) })) : [],
      };
    })
    .filter(Boolean)
    .sort((a, b) => (a.open !== b.open ? (a.open ? -1 : 1) : a.open ? (a.deadline < b.deadline ? -1 : a.deadline > b.deadline ? 1 : 0) : a.deadline < b.deadline ? 1 : -1));
  return { requests, pending: requests.filter((r) => r.open && r.isTarget && !r.mine).length };
}

function getRequests(token) {
  const ctx = auth_(token);
  // 이 기능 전에 만든 요청은 PM이 요청함을 열 때 마감 일정을 한 번 채운다
  if (isPM_(ctx)) {
    const have = readAll_('tasks').map((t) => t.id);
    const missing = readAll_('requests').filter((r) => reqOpen_(r) && have.indexOf(r.id + '-due') < 0);
    if (missing.length) withLock_(() => missing.forEach((r) => reqDueTask_(ctx, r)));
  }
  return requestsState_(ctx);
}

/** 응답할 요청 수 (상단 메뉴 배지용 — 화면 첫 상태에 같이 보냄) */
function reqPendingCount_(ctx) {
  const resps = readAll_('responses');
  return readAll_('requests').filter((r) => reqOpen_(r) && reqTargets_(ctx, r).indexOf(ctx.email) >= 0
    && !resps.some((x) => x.requestId === r.id && x.email === ctx.email)).length;
}

function reqQuestions_(input) {
  const list = Array.isArray(input) ? input : [];
  if (!list.length) throw new Error('질문을 1개 이상 넣으세요.');
  if (list.length > REQ_MAX_QUESTIONS) throw new Error(`질문은 최대 ${REQ_MAX_QUESTIONS}개까지 넣을 수 있습니다.`);
  const ids = {};
  const prev = [];
  const out = list.map((q, i) => {
    const type = String((q && q.type) || '');
    if (REQ_QTYPES.indexOf(type) < 0) throw new Error(`${i + 1}번 질문의 유형을 확인하세요.`);
    const label = String(q.label || '').trim().slice(0, 100);
    if (!label) throw new Error(`${i + 1}번 질문 내용을 입력하세요.`);
    let id = /^[A-Za-z0-9_-]{1,20}$/.test(String(q.id || '')) ? String(q.id) : 'q' + (i + 1);
    while (ids[id]) id += 'x';
    ids[id] = true;
    const o = { id, type, label, required: !!q.required };
    if (type === 'choice' || type === 'multi') {
      o.options = (Array.isArray(q.options) ? q.options : []).map((x) => String(x || '').trim().slice(0, 50)).filter((x, k, a) => x && a.indexOf(x) === k).slice(0, 30);
      if (o.options.length < 2) throw new Error(`'${label}' 질문에 보기를 2개 이상 넣으세요.`);
    }
    // 조건부 질문: 앞에 있는 '하나 고르기' 질문의 답이 eq일 때만 보이고 검사한다
    if (q.showIf && q.showIf.q) {
      const src = prev.find((x) => x.id === String(q.showIf.q));
      if (!src || src.type !== 'choice' || src.options.indexOf(String(q.showIf.eq)) < 0) throw new Error(`'${label}' 질문의 보이기 조건을 확인하세요. (앞에 있는 '하나 고르기' 질문만 고를 수 있어요)`);
      o.showIf = { q: src.id, eq: String(q.showIf.eq) };
    }
    if (type === 'slot') {
      const seen = {};
      o.slots = (Array.isArray(q.slots) ? q.slots : []).map((s) => {
        const date = String((s && s.date) || ''), start = String(s.start || ''), end = String(s.end || '');
        const cap = Math.max(1, Math.min(50, Math.floor(Number(s.cap) || 1)));
        if (!isDate_(date) || !/^\d{2}:\d{2}$/.test(start) || !/^\d{2}:\d{2}$/.test(end) || end <= start) throw new Error(`'${label}' 시간대를 확인하세요.`);
        return { id: `${date} ${start}`, date, start, end, cap };
      }).filter((s) => !seen[s.id] && (seen[s.id] = true)).sort((a, b) => (a.id < b.id ? -1 : 1));
      if (!o.slots.length) throw new Error(`'${label}' 질문에 시간대를 1개 이상 만드세요.`);
      if (o.slots.length > REQ_MAX_SLOTS) throw new Error(`시간대는 최대 ${REQ_MAX_SLOTS}개까지 만들 수 있습니다.`);
    }
    prev.push(o);
    return o;
  });
  if (out.filter((q) => q.type === 'slot').length > 1) throw new Error("'시간대 고르기' 질문은 하나만 넣을 수 있습니다.");
  return out;
}

/** 요청 마감일을 업무 일정('마감')으로 자동 등록 · 갱신. id를 '<요청id>-due'로 고정해 따로 연결 컬럼 없이 찾는다.
 *  전체 대상이면 모두에게 보이게('전체 공개'), 파트·직접 지정이면 대상자와 만든 사람만('담당자만'). 마감하면 '완료'로 */
function reqDueTask_(ctx, r) {
  const id = r.id + '-due';
  const old = readAll_('tasks').find((t) => t.id === id);
  const all = reqParse_(r.targets, { mode: 'all' }).mode === 'all';
  const assignees = all ? [r.createdBy] : [r.createdBy].concat(reqTargets_(ctx, r).filter((e) => e !== r.createdBy));
  const now = new Date().toISOString();
  upsert_('tasks', {
    id, part: 'PM', title: `[요청함] ${r.title}`.slice(0, 60), category: '마감', startDate: r.deadline, endDate: r.deadline, time: '',
    assignees: assignees.join(','), status: r.status === 'closed' ? '완료' : old && old.status !== '완료' ? old.status : '예정',
    visibility: all ? 'all' : 'assignees', memo: '요청함 마감일 자동 등록 — 응답은 요청함 메뉴에서 해 주세요.',
    createdBy: r.createdBy, createdAt: old ? old.createdAt : now, updatedAt: now, updatedBy: ctx.email, series: '',
  });
}

/** PM: 요청 등록·수정. 새로 열면 대상자에게 소식 */
function saveRequest(token, input) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    const existing = input.id ? readAll_('requests').find((r) => r.id === input.id) : null;
    if (input.id && !existing) throw new Error('요청을 찾을 수 없습니다. 새로고침 후 다시 시도하세요.');
    const title = String(input.title || '').trim().slice(0, 60);
    if (!title) throw new Error('제목을 입력하세요.');
    const deadline = String(input.deadline || '');
    if (!isDate_(deadline)) throw new Error('마감일을 선택하세요.');
    if (!existing && deadline < today_()) throw new Error('마감일이 이미 지났습니다.');
    const t = input.targets || {};
    const mode = ['all', 'parts', 'people'].indexOf(t.mode) >= 0 ? t.mode : 'all';
    const targets = { mode };
    if (mode === 'parts') {
      targets.parts = (Array.isArray(t.parts) ? t.parts : []).filter((p, i, a) => PARTS.concat(['PM']).indexOf(p) >= 0 && a.indexOf(p) === i);
      if (!targets.parts.length) throw new Error('대상 파트를 고르세요.');
    }
    if (mode === 'people') {
      targets.emails = parseSubs_(t.emails).filter((e) => ctx.users.some((u) => u.email === e));
      if (!targets.emails.length) throw new Error('대상자를 1명 이상 고르세요.');
    }
    const hasResp = existing ? reqResponses_(existing.id).length > 0 : false;
    const questions = hasResp ? reqParse_(existing.questions, []) : reqQuestions_(input.questions);
    if (hasResp && input.questions && JSON.stringify(reqQuestions_(input.questions)) !== JSON.stringify(questions)) {
      throw new Error('응답이 있어서 질문은 바꿀 수 없습니다. 새 요청을 만들어 주세요.');
    }
    const now = new Date().toISOString();
    const r = {
      id: existing ? existing.id : Utilities.getUuid(), title, desc: String(input.desc || '').trim().slice(0, 500),
      targets: JSON.stringify(targets), questions: JSON.stringify(questions), deadline,
      status: existing ? existing.status || 'open' : 'open', shared: input.shared ? 'Y' : '',
      createdBy: existing ? existing.createdBy : ctx.email, createdAt: existing ? existing.createdAt : now, updatedAt: now,
    };
    upsert_('requests', r);
    reqDueTask_(ctx, r);
    logChange_(ctx, 'request', existing ? '수정' : '등록', '', `${title} · 마감 ${rangeText_({ startDate: deadline, endDate: deadline })}`);
    if (!existing) pushNews_(ctx, reqTargets_(ctx, r), `새 요청: ${title} (~${rangeText_({ startDate: deadline, endDate: deadline })})`, 'tab:requests');
    return requestsState_(ctx);
  });
}

/** PM: 마감('closed') · 다시 열기('open') */
function setRequestStatus(token, id, status) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    const r = readAll_('requests').find((x) => x.id === id);
    if (!r) throw new Error('요청을 찾을 수 없습니다.');
    if (status === 'open' && r.deadline < today_()) throw new Error('마감일이 지났습니다. 마감일을 먼저 바꿔 주세요.');
    r.status = status === 'closed' ? 'closed' : 'open';
    r.updatedAt = new Date().toISOString();
    upsert_('requests', r);
    reqDueTask_(ctx, r);
    logChange_(ctx, 'request', r.status === 'closed' ? '마감' : '다시 열기', '', r.title);
    return requestsState_(ctx);
  });
}

/** PM: 요청 삭제 — 응답과 자동 등록된 면담 업무 일정도 함께 */
function deleteRequest(token, id) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    const r = readAll_('requests').find((x) => x.id === id);
    if (!r) throw new Error('이미 삭제된 요청입니다.');
    const resps = reqResponses_(id);
    resps.forEach((x) => { if (x.taskId) deleteRow_('tasks', x.taskId); deleteRow_('responses', x.id); });
    deleteRow_('tasks', id + '-due');
    deleteRow_('requests', id);
    logChange_(ctx, 'request', '삭제', '', `${r.title} · 응답 ${resps.length}건`);
    return requestsState_(ctx);
  });
}

/** PM: 아직 응답하지 않은 대상자에게 다시 알림 */
function remindRequest(token, id) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    const r = readAll_('requests').find((x) => x.id === id);
    if (!r) throw new Error('요청을 찾을 수 없습니다.');
    if (!reqOpen_(r)) throw new Error('마감된 요청입니다.');
    const resps = reqResponses_(id);
    const pending = reqTargets_(ctx, r).filter((e) => !resps.some((x) => x.email === e));
    pushNews_(ctx, pending, `응답 부탁드려요: ${r.title} (~${rangeText_({ startDate: r.deadline, endDate: r.deadline })})`, 'tab:requests');
    return Object.assign(requestsState_(ctx), { reminded: pending.length });
  });
}

/** 대상자: 응답 저장 (마감 전까지 고칠 수 있음). 시간대는 선착순 정원, 고르면 면담 업무 일정 자동 등록 */
function submitResponse(token, requestId, answers) {
  return withLock_(() => {
    const ctx = auth_(token);
    const r = readAll_('requests').find((x) => x.id === requestId);
    if (!r) throw new Error('요청을 찾을 수 없습니다. 새로고침 후 다시 시도하세요.');
    if (reqTargets_(ctx, r).indexOf(ctx.email) < 0) throw new Error('이 요청의 대상자가 아닙니다.');
    if (!reqOpen_(r)) throw new Error('마감된 요청입니다.');
    const questions = reqParse_(r.questions, []);
    const resps = reqResponses_(requestId);
    const mine = resps.find((x) => x.email === ctx.email);
    const input = answers && typeof answers === 'object' ? answers : {};
    const out = {};
    questions.forEach((q) => {
      if (q.showIf && out[q.showIf.q] !== q.showIf.eq) return; // 조건에 안 맞는 질문은 건너뛰고 값도 버림
      let v = input[q.id];
      if (q.type === 'multi') {
        v = (Array.isArray(v) ? v : []).map(String).filter((x, k, a) => q.options.indexOf(x) >= 0 && a.indexOf(x) === k);
        if (q.required && !v.length) throw new Error(`'${q.label}' 항목을 고르세요.`);
        if (v.length) out[q.id] = v;
        return;
      }
      v = v == null ? '' : String(v).trim();
      if (!v) { if (q.required) throw new Error(`'${q.label}' 항목을 ${q.type === 'choice' || q.type === 'slot' ? '고르' : '입력하'}세요.`); return; }
      if (q.type === 'text') v = v.slice(0, 200);
      if (q.type === 'long') v = v.slice(0, 1000);
      if (q.type === 'link' && !/^https?:\/\/\S+$/i.test(v)) throw new Error(`'${q.label}'에는 http로 시작하는 주소를 넣으세요.`);
      if (q.type === 'link') v = v.slice(0, 500);
      if (q.type === 'number') { if (!isFinite(Number(v)) || v === '') throw new Error(`'${q.label}'에는 숫자를 넣으세요.`); v = String(Number(v)); }
      if (q.type === 'date' && !isDate_(v)) throw new Error(`'${q.label}' 날짜를 확인하세요.`);
      if (q.type === 'choice' && q.options.indexOf(v) < 0) throw new Error(`'${q.label}' 보기를 다시 고르세요.`);
      if (q.type === 'slot') {
        const s = q.slots.find((x) => x.id === v);
        if (!s) throw new Error(`'${q.label}' 시간대를 다시 고르세요.`);
        const before = mine ? reqParse_(mine.answers, {})[q.id] : '';
        const used = resps.filter((x) => x.email !== ctx.email && reqParse_(x.answers, {})[q.id] === v).length;
        if (v !== before && used >= s.cap) throw new Error(`${rangeText_({ startDate: s.date, endDate: s.date })} ${s.start} 시간대는 방금 마감됐어요. 다른 시간대를 골라 주세요.`);
      }
      out[q.id] = v;
    });
    const now = new Date().toISOString();
    const row = { id: mine ? mine.id : Utilities.getUuid(), requestId, email: ctx.email, answers: JSON.stringify(out), createdAt: mine ? mine.createdAt : now, updatedAt: now, taskId: mine ? mine.taskId || '' : '' };
    // 면담 업무 일정 자동 등록 · 변경 · 삭제
    const sq = reqSlotQ_(questions);
    const slot = sq && out[sq.id] ? sq.slots.find((x) => x.id === out[sq.id]) : null;
    const oldTask = row.taskId ? readAll_('tasks').find((t) => t.id === row.taskId) : null;
    if (slot) {
      const task = {
        id: oldTask ? oldTask.id : Utilities.getUuid(), part: 'PM', title: `${r.title} · ${ctx.user.name}`.slice(0, 60), category: '회의/리뷰',
        startDate: slot.date, endDate: slot.date, time: `${slot.start}~${slot.end}`, assignees: [r.createdBy, ctx.email].join(','),
        status: oldTask ? oldTask.status : '예정', visibility: 'assignees', memo: '요청함 응답으로 자동 등록',
        createdBy: r.createdBy, createdAt: oldTask ? oldTask.createdAt : now, updatedAt: now, updatedBy: ctx.email, series: '',
      };
      upsert_('tasks', task);
      row.taskId = task.id;
      const was = oldTask ? `${oldTask.startDate} ${String(oldTask.time).slice(0, 5)}` : '';
      if (was !== slot.id) pushNews_(ctx, [r.createdBy], `${oldTask ? '면담 시간 변경' : '면담 신청'}: ${ctx.user.name} · ${rangeText_({ startDate: slot.date, endDate: slot.date })} ${slot.start} (${r.title})`, 'tab:requests');
    } else if (oldTask) {
      deleteRow_('tasks', oldTask.id);
      row.taskId = '';
    }
    upsert_('responses', row);
    return requestsState_(ctx);
  });
}

/** 대상자: 내 응답 취소 (마감 전) — 면담 업무 일정도 지움 */
function cancelResponse(token, requestId) {
  return withLock_(() => {
    const ctx = auth_(token);
    const r = readAll_('requests').find((x) => x.id === requestId);
    if (!r) throw new Error('요청을 찾을 수 없습니다.');
    if (!reqOpen_(r)) throw new Error('마감된 요청입니다.');
    const mine = reqResponses_(requestId).find((x) => x.email === ctx.email);
    if (!mine) throw new Error('취소할 응답이 없습니다.');
    if (mine.taskId) {
      deleteRow_('tasks', mine.taskId);
      pushNews_(ctx, [r.createdBy], `면담 신청 취소: ${ctx.user.name} (${r.title})`, 'tab:requests');
    }
    deleteRow_('responses', mine.id);
    return requestsState_(ctx);
  });
}

/* ───────────── 사다리 추첨 (회식 참석자 등 — 일정의 담당자를 랜덤으로) ───────────── */
/*
 * PM이 업무 일정 · 중요 일정에서 후보와 인원(또는 파트별 인원)을 정해 추첨하면, 뽑힌 사람이 그 일정의 담당자로 들어간다.
 * - 뽑기는 서버에서 한다. 화면의 사다리는 정해진 결과를 보여주는 연출이다.
 * - 회차마다 기록이 남고(재추첨 · 1명 교체 포함) 이력에도 남는다. 결과는 후보 전원과 PM이 본다.
 * - 다시 추첨하면 지난 회차 당첨자는 담당자에서 빠지고 새로 뽑는다. 원래 있던 담당자(고정)는 그대로.
 */
function drawTarget_(kind, id) {
  const key = kind === 'event' ? 'events' : kind === 'task' ? 'tasks' : '';
  if (!key) throw new Error('추첨할 일정 종류를 확인하세요.');
  const row = readAll_(key).find((x) => x.id === id);
  if (!row) throw new Error('일정을 찾을 수 없습니다. 새로고침 후 다시 시도하세요.');
  return { key, row, max: kind === 'event' ? MAX_EVENT_ASSIGNEES : MAX_ASSIGNEES, assignees: parseSubs_(row.assignees) };
}

function drawsFor_(kind, id) {
  return readAll_('draws').filter((d) => d.kind === kind && d.targetId === id).sort((a, b) => Number(a.round) - Number(b.round));
}

function drawPartOf_(u) {
  return u.part || (u.role === 'pm' ? 'PM' : '');
}

/** 후보 명단: 전체 · 파트 · DS직원 · 직접 지정 (직접 지정 말고는 테스터 제외) */
function drawPool_(ctx, pool) {
  const mode = ['all', 'parts', 'ds', 'people'].indexOf(pool.mode) >= 0 ? pool.mode : 'all';
  if (mode === 'people') return parseSubs_(pool.emails).filter((e) => ctx.users.some((u) => u.email === e));
  return ctx.users.filter((u) => !isTester_(u) && (mode === 'all' || (mode === 'ds' ? u.ds === 'Y' : (pool.parts || []).indexOf(drawPartOf_(u)) >= 0))).map((u) => u.email);
}

function drawPick_(list, n) {
  const a = list.slice();
  for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); const t = a[i]; a[i] = a[j]; a[j] = t; }
  return a.slice(0, n);
}

/** 최근 n번의 추첨(다른 일정)에서 뽑힌 사람 — 각 일정의 마지막 회차 기준 */
function drawRecentWinners_(kind, id, n) {
  if (!n) return [];
  const last = {};
  readAll_('draws').forEach((d) => {
    if (d.kind === kind && d.targetId === id) return;
    const k = d.kind + ':' + d.targetId;
    if (!last[k] || Number(d.round) > Number(last[k].round)) last[k] = d;
  });
  const out = [];
  Object.keys(last).map((k) => last[k]).sort((a, b) => (a.at < b.at ? 1 : -1)).slice(0, n)
    .forEach((d) => reqParse_(d.picked, []).forEach((e) => { if (out.indexOf(e) < 0) out.push(e); }));
  return out;
}

function drawSaveTarget_(t, assignees, ctx) {
  t.row.assignees = assignees.join(',');
  t.row.updatedAt = new Date().toISOString();
  t.row.updatedBy = ctx.email;
  upsert_(t.key, t.row);
}

function drawView_(ctx, d) {
  const person = (e) => { const u = ctx.users.find((x) => x.email === e) || {}; return { email: e, name: u.name || e, part: u.part || (u.role === 'pm' ? 'PM' : '') }; };
  const rounds = drawsFor_(d.kind, d.targetId);
  return {
    id: d.id, kind: d.kind, targetId: d.targetId, title: d.title, date: d.date, round: Number(d.round),
    pool: reqParse_(d.pool, {}), quota: reqParse_(d.quota, {}), note: d.note, by: nameOf_(ctx, d.by), at: d.at,
    candidates: reqParse_(d.candidates, []).map(person), picked: reqParse_(d.picked, []).map(person), fixed: reqParse_(d.fixed, []).map(person),
    excluded: reqParse_(d.excluded, []).map((x) => Object.assign(person(x.email), { why: x.why })),
    latest: !rounds.length || rounds[rounds.length - 1].id === d.id, canManage: isPM_(ctx),
    rounds: rounds.map((r) => ({ id: r.id, round: Number(r.round), at: r.at, note: r.note, picked: reqParse_(r.picked, []).map((e) => person(e).name) })),
  };
}

function drawCanSee_(ctx, d) {
  if (isPM_(ctx)) return true;
  return ['candidates', 'fixed'].some((k) => reqParse_(d[k], []).indexOf(ctx.email) >= 0) || reqParse_(d.excluded, []).some((x) => x.email === ctx.email);
}

/** PM: 추첨 — input { kind, id, pool: {mode, parts, emails}, count, quota: {파트: 인원}, skipLeave, recent } */
function runDraw(token, input) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    const kind = String(input.kind || '');
    const t = drawTarget_(kind, String(input.id || ''));
    const prev = drawsFor_(kind, t.row.id).pop();
    const prevPicked = prev ? reqParse_(prev.picked, []) : [];
    const fixed = t.assignees.filter((e) => prevPicked.indexOf(e) < 0);
    const pool = input.pool || {};
    const candidates = drawPool_(ctx, pool).filter((e) => fixed.indexOf(e) < 0);
    if (!candidates.length) throw new Error('후보가 없습니다. 후보를 다시 고르세요.');
    const excluded = [];
    if (input.skipLeave !== false) {
      readAll_('leaves').forEach((l) => {
        if (candidates.indexOf(l.email) >= 0 && l.startDate <= t.row.endDate && t.row.startDate <= l.endDate && !excluded.some((x) => x.email === l.email)) excluded.push({ email: l.email, why: `휴가 (${l.type})` });
      });
    }
    const recent = Math.max(0, Math.min(5, Math.floor(Number(input.recent) || 0)));
    drawRecentWinners_(kind, t.row.id, recent).forEach((e) => {
      if (candidates.indexOf(e) >= 0 && !excluded.some((x) => x.email === e)) excluded.push({ email: e, why: '최근 당첨' });
    });
    const eligible = candidates.filter((e) => !excluded.some((x) => x.email === e));
    const quota = {};
    let picked = [];
    const q = input.quota && typeof input.quota === 'object' ? input.quota : null;
    if (q && Object.keys(q).some((k) => Number(q[k]) > 0)) {
      Object.keys(q).forEach((part) => {
        const n = Math.max(0, Math.min(50, Math.floor(Number(q[part]) || 0)));
        if (!n) return;
        const inPart = eligible.filter((e) => drawPartOf_(ctx.users.find((u) => u.email === e) || {}) === part);
        if (inPart.length < n) throw new Error(`${part} 후보가 ${inPart.length}명이라 ${n}명을 뽑을 수 없습니다.`);
        quota[part] = n;
        picked = picked.concat(drawPick_(inPart, n));
      });
    } else {
      const n = Math.floor(Number(input.count) || 0);
      if (n < 1) throw new Error('뽑을 인원을 1명 이상 넣으세요.');
      if (eligible.length < n) throw new Error(`추첨할 수 있는 후보가 ${eligible.length}명이라 ${n}명을 뽑을 수 없습니다.`);
      picked = drawPick_(eligible, n);
    }
    if (fixed.length + picked.length > t.max) throw new Error(`담당자는 최대 ${t.max}명이에요. (지금 담당자 ${fixed.length}명 + 뽑을 인원 ${picked.length}명)`);
    drawSaveTarget_(t, fixed.concat(picked), ctx);
    const d = {
      id: Utilities.getUuid(), kind, targetId: t.row.id, title: t.row.title, date: t.row.startDate, round: String(prev ? Number(prev.round) + 1 : 1),
      pool: JSON.stringify({ mode: pool.mode || 'all', parts: pool.parts || [], skipLeave: input.skipLeave !== false, recent, count: picked.length }),
      candidates: JSON.stringify(candidates), excluded: JSON.stringify(excluded), quota: JSON.stringify(quota), picked: JSON.stringify(picked), fixed: JSON.stringify(fixed),
      note: prev ? '다시 추첨' : '', by: ctx.email, at: new Date().toISOString(),
    };
    upsert_('draws', d);
    const when = rangeText_({ startDate: t.row.startDate, endDate: t.row.startDate });
    logChange_(ctx, kind, '추첨', '', `${t.row.title} (${when}) · ${d.round}회차 · 후보 ${candidates.length}명 중 ${picked.length}명: ${picked.map((e) => nameOf_(ctx, e)).join(', ')}`);
    pushNews_(ctx, picked, `${t.row.title}(${when}) 담당자로 뽑혔어요 (사다리 추첨)`, 'draw:' + d.id);
    pushNews_(ctx, candidates.filter((e) => picked.indexOf(e) < 0), `${t.row.title}(${when}) 추첨 결과가 나왔어요`, 'draw:' + d.id);
    return { draw: drawView_(ctx, d) };
  });
}

/** PM: 당첨자 1명만 빼고 같은 조건(파트별이면 같은 파트)에서 1명 다시 뽑기 */
function replaceDraw(token, drawId, email) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    const d = readAll_('draws').find((x) => x.id === drawId);
    if (!d) throw new Error('추첨 기록을 찾을 수 없습니다.');
    const rounds = drawsFor_(d.kind, d.targetId);
    if (rounds[rounds.length - 1].id !== d.id) throw new Error('마지막 회차에서만 바꿀 수 있습니다. 새로고침 후 다시 시도하세요.');
    const t = drawTarget_(d.kind, d.targetId);
    const picked = reqParse_(d.picked, []);
    if (picked.indexOf(email) < 0) throw new Error('당첨자가 아닙니다.');
    const quota = reqParse_(d.quota, {});
    const excluded = reqParse_(d.excluded, []).concat([{ email, why: '교체로 빠짐' }]);
    const partOf = (e) => drawPartOf_(ctx.users.find((u) => u.email === e) || {});
    const pool = reqParse_(d.candidates, []).filter((e) => picked.indexOf(e) < 0 && !excluded.some((x) => x.email === e) && t.assignees.indexOf(e) < 0
      && (!Object.keys(quota).length || partOf(e) === partOf(email)));
    if (!pool.length) throw new Error('대신 뽑을 후보가 없습니다.');
    const pick = drawPick_(pool, 1)[0];
    const next = picked.map((e) => (e === email ? pick : e));
    drawSaveTarget_(t, t.assignees.filter((e) => e !== email).concat(t.assignees.indexOf(pick) < 0 ? [pick] : []), ctx);
    const nd = Object.assign({}, d, {
      id: Utilities.getUuid(), round: String(Number(d.round) + 1), excluded: JSON.stringify(excluded), picked: JSON.stringify(next),
      note: `교체: ${nameOf_(ctx, email)} → ${nameOf_(ctx, pick)}`, by: ctx.email, at: new Date().toISOString(),
    });
    upsert_('draws', nd);
    const when = rangeText_({ startDate: d.date, endDate: d.date });
    logChange_(ctx, d.kind, '추첨 교체', '', `${d.title} (${when}) · ${nd.round}회차 · ${nd.note}`);
    pushNews_(ctx, [pick], `${d.title}(${when}) 담당자로 뽑혔어요 (사다리 추첨 · 교체)`, 'draw:' + nd.id);
    return { draw: drawView_(ctx, nd) };
  });
}

/** 추첨 결과 보기 (후보 · 고정 담당자 · PM) */
function getDraw(token, drawId) {
  const ctx = auth_(token);
  const d = readAll_('draws').find((x) => x.id === drawId);
  if (!d) throw new Error('추첨 기록을 찾을 수 없습니다.');
  if (!drawCanSee_(ctx, d)) throw new Error('이 추첨의 후보가 아니라서 볼 수 없습니다.');
  return { draw: drawView_(ctx, d) };
}

/** 일정마다 마지막 추첨 id (목록에서 '추첨 결과' 버튼용) */
function latestDraws_() {
  const last = {};
  readAll_('draws').forEach((d) => {
    const k = d.kind + ':' + d.targetId;
    if (!last[k] || Number(d.round) > Number(last[k].round)) last[k] = d;
  });
  const out = {};
  Object.keys(last).forEach((k) => (out[k] = last[k].id));
  return out;
}

/* ───────────── 1시간 조기퇴근권 (PM만 기록 · 조회, 한 사람 한 달에 1번) ───────────── */
function earlyState_(ctx) {
  return {
    items: readAll_('early').map((x) => ({ id: x.id, email: x.email, name: nameOf_(ctx, x.email), date: x.date, memo: x.memo, createdAt: x.createdAt, by: nameOf_(ctx, x.createdBy) }))
      .sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0)),
  };
}

function getEarly(token) {
  return earlyState_(requireAdmin_(token));
}

/** PM: 사용 기록 등록 · 수정 — 같은 사람이 같은 달에 이미 썼으면 거부 */
function saveEarly(token, input) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    const all = readAll_('early');
    const ex = input.id ? all.find((x) => x.id === input.id) : null;
    if (input.id && !ex) throw new Error('기록을 찾을 수 없습니다. 새로고침 후 다시 시도하세요.');
    const email = String(input.email || '').toLowerCase();
    if (!ctx.users.some((u) => u.email === email)) throw new Error('직원을 고르세요.');
    const date = String(input.date || '');
    if (!isDate_(date)) throw new Error('날짜를 고르세요.');
    const dup = all.find((x) => x.email === email && x.date.slice(0, 7) === date.slice(0, 7) && (!ex || x.id !== ex.id));
    if (dup) throw new Error(`${nameOf_(ctx, email)}님은 ${Number(date.slice(5, 7))}월에 이미 썼어요 (${rangeText_({ startDate: dup.date, endDate: dup.date })}). 한 달에 한 번만 쓸 수 있어요.`);
    const now = new Date().toISOString();
    const row = { id: ex ? ex.id : Utilities.getUuid(), email, date, memo: String(input.memo || '').trim().slice(0, 100),
      createdAt: ex ? ex.createdAt : now, createdBy: ex ? ex.createdBy : ctx.email, updatedAt: now };
    upsert_('early', row);
    logChange_(ctx, 'early', ex ? '수정' : '등록', email, `조기퇴근권 ${rangeText_({ startDate: date, endDate: date })}${row.memo ? ' · ' + row.memo : ''}`);
    return earlyState_(ctx);
  });
}

function deleteEarly(token, id) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    const ex = readAll_('early').find((x) => x.id === id);
    if (!ex) throw new Error('이미 삭제된 기록입니다.');
    deleteRow_('early', id);
    logChange_(ctx, 'early', '삭제', ex.email, `조기퇴근권 ${rangeText_({ startDate: ex.date, endDate: ex.date })}`);
    return earlyState_(ctx);
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
    // 담당자 (선택, 최대 MAX_EVENT_ASSIGNEES명). 화면이 담당자를 보내지 않으면(예전 화면) 기존 값 유지
    const assignees = input.assignees === undefined && existing ? parseSubs_(existing.assignees) : parseSubs_(input.assignees);
    if (assignees.length > MAX_EVENT_ASSIGNEES) throw new Error(`담당자는 최대 ${MAX_EVENT_ASSIGNEES}명까지 지정할 수 있습니다.`);
    if (assignees.some((e) => !ctx.users.some((u) => u.email === e))) throw new Error('담당자가 팀원 목록에 없습니다.');
    const now = new Date().toISOString();
    const ev = {
      id: existing ? existing.id : Utilities.getUuid(),
      title, startDate, endDate,
      time: String(input.time || '').trim().slice(0, 30),
      memo: String(input.memo || '').trim().slice(0, MEMO_MAX),
      createdAt: existing ? existing.createdAt : now,
      updatedAt: now,
      updatedBy: ctx.email,
      assignees: assignees.join(','),
    };
    upsert_('events', ev);
    const desc = (e) => `${e.title} ${rangeText_(e)}${e.time ? ' ' + e.time : ''}`;
    const who = (list) => (list.length ? ' · 담당 ' + list.map((e) => nameOf_(ctx, e)).join(', ') : '');
    const before = existing ? parseSubs_(existing.assignees) : [];
    const summary = (e, list) => desc(e) + who(list);
    logChange_(ctx, 'event', existing ? '수정' : '등록', '', existing && summary(existing, before) !== summary(ev, assignees) ? `${summary(existing, before)} → ${summary(ev, assignees)}` : summary(ev, assignees));
    const everyone = ctx.users.map((u) => u.email);
    if (ev.endDate >= today_()) {
      const added = assignees.filter((e) => before.indexOf(e) < 0);
      // 새로 담당자가 된 사람에게는 '담당' 소식, 나머지는 평소처럼 (같은 사람에게 두 번 가지 않게)
      pushNews_(ctx, added, `중요 일정 담당으로 지정됐어요: ${desc(ev)}`, 'day:' + ev.startDate);
      const rest = everyone.filter((e) => added.indexOf(e) < 0);
      if (!existing) pushNews_(ctx, rest, `새 중요 일정: ${desc(ev)}`, 'day:' + ev.startDate);
      else if (desc(existing) !== desc(ev)) pushNews_(ctx, rest, `중요 일정 변경: ${desc(ev)}`, 'day:' + ev.startDate);
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
    const ex = readAll_('holidays').find((h) => h.date === date);
    upsert_('holidays', { date, name });
    if (!ex) logChange_(ctx, 'holiday', '등록', '', `${date} ${name}`);
    else if (ex.name !== name) logChange_(ctx, 'holiday', '이름 수정', '', `${date} ${ex.name} → ${name}`);
    return { state: state_(ctx) };
  });
}

/** PM: 공휴일 일괄 추가 — rows = [{ date, name }]. 이미 있는 날짜는 건너뜀 (이름도 그대로) */
function bulkAddHolidays(token, rows) {
  return withLock_(() => {
    const ctx = requireAdmin_(token);
    if (!Array.isArray(rows) || !rows.length) throw new Error('등록할 공휴일이 없습니다.');
    if (rows.length > 100) throw new Error('한 번에 100건까지 등록할 수 있습니다.');
    const have = {};
    readAll_('holidays').forEach((h) => (have[h.date] = true));
    const added = [], errors = [];
    let skipped = 0;
    rows.forEach((r, i) => {
      const date = String((r && r.date) || '');
      const name = String((r && r.name) || '').trim().slice(0, 30);
      const real = isDate_(date) && new Date(Date.UTC(+date.slice(0, 4), +date.slice(5, 7) - 1, +date.slice(8, 10))).toISOString().slice(0, 10) === date; // 2027-02-30 같은 없는 날짜 거부
      if (!real || !name) return errors.push({ line: i + 1, message: '날짜 또는 이름이 올바르지 않습니다.' });
      if (have[date]) return skipped++;
      have[date] = true;
      added.push({ date, name });
    });
    added.forEach((h) => upsert_('holidays', h));
    if (added.length) logChange_(ctx, 'holiday', '일괄 등록', '', `${added.length}건 (${added[0].date} ~ ${added[added.length - 1].date})`);
    return { state: state_(ctx), created: added.length, skipped, errors };
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
  if (t.visibility === 'assignees' || (old && old.visibility === 'assignees')) return `[${t.part}] (담당자만 공개 업무)`;
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
  const extUsage = ctx.user.role === 'member' && ctx.user.viewAll === 'Y';
  return {
    me: {
      email: ctx.email, name: ctx.user.name, role: ctx.user.role, part: ctx.user.part,
      mustChange: ctx.user.mustChange === 'Y', defaultSubs: defaults(ctx.user), viewAll: ctx.user.viewAll === 'Y',
      annual: parseAnnual_(ctx.user.annual), // 내 연차 부여 일수 (연도별)
      ds: ctx.user.ds === 'Y', // DS직원 (상단 식권 승인 버튼)
      navOrder: String(ctx.user.navOrder || '').split(',').filter(Boolean), // 내 메뉴 순서
    },
    // 대직자 선택용 명단 (이름·파트). 권한 정보는 PM에게만
    users: ctx.users.map((u) => (pm
      ? { email: u.email, name: u.name, role: u.role, part: u.part, defaultSubs: defaults(u), viewAll: u.viewAll === 'Y', annual: parseAnnual_(u.annual), ds: u.ds === 'Y' }
      : extUsage && u.ds !== 'Y'
        ? { email: u.email, name: u.name, part: u.part, annual: annualNoNotes_(u.annual), usage: true } // '전체 일정 조회' 권한자: DS직원이 아닌 사람의 휴가 사용 현황용
        : { email: u.email, name: u.name, part: u.part })),
    leaves: visibleLeaves_(ctx, all).map((l) => ({
      id: l.id, email: l.email, type: l.type, startDate: l.startDate, endDate: l.endDate,
      substitutes: l.substitutes, memo: briefOnly_(ctx, l) ? '' : l.memo, createdAt: l.createdAt, updatedAt: l.updatedAt,
      conflicts: briefOnly_(ctx, l) ? [] : conflictsOf_(l, all, ctx.users),
      conflictDays: briefOnly_(ctx, l) ? [] : conflictDays_(l, all),
    })),
    holidays: readAll_('holidays'),
    // 중요 일정은 모든 직원에게
    events: ((draws) => readAll_('events').map((e) => ({
      drawId: draws['event:' + e.id] || '',
      id: e.id, title: e.title, startDate: e.startDate, endDate: e.endDate, time: e.time, memo: e.memo, createdAt: e.createdAt,
      assignees: parseSubs_(e.assignees).filter((x) => ctx.users.some((u) => u.email === x)), // 담당자 (삭제된 직원은 빼고)
    })))(latestDraws_()),
    // 앞으로 있을 일정에서 나를 대직자로 지정한 사람 — 이름만 (날짜·종류는 노출하지 않음)
    designatedBy: all
      .filter((l) => l.substitutes.indexOf(ctx.email) >= 0 && l.endDate >= today)
      .map((l) => nameOf(l.email))
      .filter((n, i, a) => a.indexOf(n) === i)
      .sort(),
    news: newsOf_(ctx),
    reqPending: reqPendingCount_(ctx), // 응답할 요청 수 (요청함 메뉴 배지)
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
  if (type === '오후반차' || type === '반반차') return 2; // 반반차는 오후에만
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
    const when = rangeText_({ startDate: s, endDate: e }) + (m.type === '오전반차' ? ' 오전' : m.type === '오후반차' || m.type === '반반차' ? ' 오후' : '');
    // 대직자 중 한 명이라도 같은 시간에 부재면 경고
    if (l.substitutes.indexOf(m.email) >= 0) out.push(`대직자 ${nameOf(m.email)}님이 ${when} 부재`);
    else if (m.substitutes.indexOf(l.email) >= 0) out.push(`${nameOf(m.email)}님이 ${nameOf(l.email)}님을 대직자로 지정했는데 ${when} 둘 다 부재`);
  });
  return out;
}

/** 테스터 계정 (이름에 '테스터'): 총원 · 부재 인원에서 뺀다 */
function isTester_(u) {
  return String((u && u.name) || '').indexOf('테스터') >= 0;
}

/**
 * 총원(테스터 제외)의 CROWD_PERCENT% 이상이 같은 시간에 부재하게 되는 날 (주말·공휴일 제외) — 안내용.
 * 이 일정과 시간(오전·오후)이 겹치는 일정만 센다. 누가 부재인지는 알려주지 않고 인원수만.
 */
function crowdedDays_(l, leaves, users, holidays) {
  const people = users.filter((u) => !isTester_(u));
  const testers = {};
  users.filter(isTester_).forEach((u) => (testers[u.email] = true));
  if (!people.length || testers[l.email] || l.endDate < today_()) return [];
  const min = Math.ceil((people.length * CROWD_PERCENT) / 100);
  const off = {};
  (holidays || []).forEach((h) => (off[h.date] = true));
  const out = [];
  const toT = (ds) => Date.UTC(+ds.slice(0, 4), +ds.slice(5, 7) - 1, +ds.slice(8, 10));
  for (let t = toT(l.startDate); t <= toT(l.endDate); t += 864e5) {
    const ds = new Date(t).toISOString().slice(0, 10);
    const w = new Date(t).getUTCDay();
    if (w === 0 || w === 6 || off[ds]) continue;
    const day = { startDate: ds, endDate: ds, type: l.type };
    const away = {};
    leaves.forEach((m) => { if (m.id !== l.id && m.email !== l.email && !testers[m.email] && overlaps_(day, m)) away[m.email] = true; });
    const n = Object.keys(away).length + 1;
    if (n >= min) out.push(`${rangeText_(day)} 부재 ${n}명 (총원 ${people.length}명의 ${Math.round((n * 100) / people.length)}%)`);
  }
  return out;
}

/** 오늘 이후 일정이 대직자와 겹치면 사유(메모) 필수 */
function reasonRequired_(l, warnings) {
  return l.endDate >= today_() && warnings.length > 0;
}

/** 대직 겹침이 있는 날짜만 (캘린더에서 겹치는 날에만 ⚠ 표시) — conflictsOf_와 같은 조건 */
function conflictDays_(l, leaves) {
  const days = {};
  const toT = (ds) => Date.UTC(+ds.slice(0, 4), +ds.slice(5, 7) - 1, +ds.slice(8, 10));
  leaves.forEach((m) => {
    if (m.id === l.id || !overlaps_(l, m)) return;
    if (l.substitutes.indexOf(m.email) < 0 && m.substitutes.indexOf(l.email) < 0) return;
    const s = l.startDate > m.startDate ? l.startDate : m.startDate;
    const e = l.endDate < m.endDate ? l.endDate : m.endDate;
    for (let t = toT(s); t <= toT(e); t += 864e5) days[new Date(t).toISOString().slice(0, 10)] = true;
  });
  return Object.keys(days).sort();
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

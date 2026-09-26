/**
 * K-Color 컬러진단표준시스템 (CASS) — Apps Script 게이트웨이
 * 색다른컬러연구소 · DARUN COLOR INSTITUTE
 *
 * ★ 이번 수정 (2026-08)
 *   - consultants 에 '국적' 열을 C(이름) 오른쪽 = D 에 추가.
 *   - 코드 형식 변경: 역할코드-YYNNN-국적   예) CS-26001-KR
 *       YY = 발급 연도 뒤 2자리, NNN = 그 해·그 역할의 일련번호(3자리), 국적 = KR 등
 *   - 국적 추가로 D 뒤 열이 한 칸씩 밀림 → 아래 열 번호를 모두 새로 맞춤.
 *
 * ★ 추가 (2026-08) — records 에 '선호 태그' 열을 '컨설턴트 메모'·'동의여부' 뒤, '결과ID' 앞에 넣었다.
 *
 * ★ 확장 (2026-09) — records 열을 대폭 늘렸다 (12유형 개편 반영).
 *     추가: 진단방식(드레이핑/선호), 측색 원본 4부위(턱좌·턱우·볼좌·볼우 L*a*b*),
 *           턱볼 Δb*·색차(Δab), 스타일그룹(선호 진단 시), 그리고 '시즌' 열을 '타임' 으로 바꿨다.
 *     readRecord_ 는 열 이름으로 찾도록 바꿔 열 순서가 바뀌어도 안전하다.
 *     ※ 기존 records 는 열 구성이 달라 어긋나므로, records 를 'records(구)' 로 이름 바꾸고
 *       새 [records] 시트를 만든 뒤 setup() 을 실행하는 것을 권장.
 *
 * ※ 적용 순서
 *   1) (권장) 기존 'records' 탭 이름을 'records(구)' 로 바꾼다. (옛 기록 보관용)
 *   2) 이 파일을 통째로 붙여넣고 setup() 을 한 번 실행한다 → 새 [records] 가 새 머리글로 만들어진다.
 *   3) 배포 > 배포 관리 > (기존 배포) 수정 > 새 버전 으로 재배포한다. (새 배포 아님)
 */

// ─────────────────────────────────────────────────────────────
// 설정
// ─────────────────────────────────────────────────────────────
var SHEET_CONSULTANTS = 'consultants';
var SHEET_RECORDS     = 'records';
var SHEET_RULES       = 'rules';

// A     B     C     D     E     F      G      H       I        J        K        L        M
var HEAD_CONSULTANTS = [
  '코드', '역할', '이름', '국적', '소속', '발급일', '만료일', '활성여부', '최근접속', '누적 진단수',
  '가입개월', '현재세션', '세션시각'
];

var HEAD_RECORDS = [
  '타임스탬프', '코드', '역할', '진단방식',
  '이름', '이메일', '생년월일', '성별', '연락처', '인종', '직업', '진단일',
  '피부명도', '모발색', '눈동자색',
  '피부색 밝기', '부위별 밝기차', '홍조 강도', '피부 균일도',
  '턱 L*', '턱 a*', '턱 b*', '볼 L*', '볼 a*', '볼 b*',
  '턱좌 L*', '턱좌 a*', '턱좌 b*', '턱우 L*', '턱우 a*', '턱우 b*',
  '볼좌 L*', '볼좌 a*', '볼좌 b*', '볼우 L*', '볼우 a*', '볼우 b*',
  '턱볼 Δb*', '턱볼 색차(Δab)',
  '타임', '유형코드', '유형', '스타일그룹',
  '베스트 오브 베스트', '베스트 컬러(BOC)',
  '컨설턴트 메모', '동의여부', '선호 태그', '결과ID'
];

var HEAD_RULES = ['항목', '값', '설명', '수정자', '수정일'];

// 열 번호 (1-based) — consultants 머리글과 반드시 일치
var COL_CODE     = 1;   // A 코드
var COL_ROLE     = 2;   // B 역할
var COL_NAME     = 3;   // C 이름
var COL_NATION   = 4;   // D 국적          (수동)
var COL_ORG      = 5;   // E 소속
var COL_ISSUED   = 6;   // F 발급일
var COL_EXPIRY   = 7;   // G 만료일
var COL_ACTIVE   = 8;   // H 활성여부
var COL_LASTSEEN = 9;   // I 최근접속
var COL_COUNT    = 10;  // J 누적 진단수
var COL_MONTHS   = 11;  // K 가입개월      (수동 전용 — 앱이 건드리지 않음)
var COL_SESSION  = 12;  // L 현재세션      (동시접속 차단)
var COL_SESSTIME = 13;  // M 세션시각

var SESSION_TIMEOUT_MS = 2 * 60 * 1000;   // 2분

// ─────────────────────────────────────────────────────────────
// 최초 설정 — 한 번만 실행
// ─────────────────────────────────────────────────────────────
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_CONSULTANTS, HEAD_CONSULTANTS);
  ensureSheet_(ss, SHEET_RECORDS,     HEAD_RECORDS);
  ensureSheet_(ss, SHEET_RULES,       HEAD_RULES);

  var cs = ss.getSheetByName(SHEET_CONSULTANTS);
  if (cs.getLastRow() < 2) {
    var today = new Date();
    var next  = new Date(today.getFullYear() + 1, today.getMonth(), today.getDate());
    // 코드, 역할, 이름, 국적, 소속, 발급일, 만료일, 활성, 최근접속, 누적, 가입개월, 현재세션, 세션시각
    cs.appendRow(['AD-26001-KR', 'Admin',      '유잰',  'KR', '색다른컬러연구소', today, next, 'Y', '', 0, 12, '', '']);
    cs.appendRow(['ED-26001-KR', 'Educator',   '(예시)', 'KR', '',                today, next, 'Y', '', 0, 12, '', '']);
    cs.appendRow(['CS-26001-KR', 'Consultant', '(예시)', 'KR', '',                today, next, 'Y', '', 0, 12, '', '']);
  }
  try {
    SpreadsheetApp.getUi().alert('설정 완료 — consultants / records / rules 머리글이 정리되었습니다.');
  } catch (e) { Logger.log('설정 완료'); }
}

function ensureSheet_(ss, name, head) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  var needHead = sh.getLastRow() === 0 || sh.getRange(1, 1).getValue() !== head[0]
              || sh.getLastColumn() < head.length;
  if (needHead) sh.getRange(1, 1, 1, head.length).setValues([head]);
  sh.getRange(1, 1, 1, head.length)
    .setFontWeight('bold').setBackground('#3F3B37').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  return sh;
}

// ─────────────────────────────────────────────────────────────
// 웹앱 진입점
// ─────────────────────────────────────────────────────────────
function doPost(e) {
  var req = {};
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'BAD_JSON' }); }
  try {
    if (req.action === 'verify') return json_(verifyCode_(req.code));
    if (req.action === 'save')   return json_(saveRecord_(req));
    if (req.action === 'read')   return json_(readRecord_(req.id));
    if (req.action === 'claim')  return json_(claimSession_(req.code, req.session));
    if (req.action === 'beat')   return json_(beatSession_(req.code, req.session));
    return json_({ ok: false, error: 'UNKNOWN_ACTION' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  var id = e && e.parameter && e.parameter.id;
  if (!id) return json_({ ok: true, service: 'K-Color CASS gateway' });
  return json_(readRecord_(id));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────────
// 1) 코드 검증
// ─────────────────────────────────────────────────────────────
var ROLE_SCOPE = {
  'ADMIN':      ['diag', 'edu', 'adm'],
  'EDUCATOR':   ['edu'],
  'CONSULTANT': ['diag'],
  'MASTER':     ['diag', 'edu']
};
var ROLE_LABEL  = { 'ADMIN': 'Admin', 'EDUCATOR': 'Educator', 'CONSULTANT': 'Consultant', 'MASTER': 'Master' };
var ROLE_PREFIX = { 'ADMIN': 'AD',    'EDUCATOR': 'ED',       'CONSULTANT': 'CS',         'MASTER': 'MS'     };
var PREFIX_ROLE = { 'AD': 'ADMIN',    'ED': 'EDUCATOR',       'CS': 'CONSULTANT',         'MS': 'MASTER'     };

function normCode_(s) {
  return String(s === null || s === undefined ? '' : s).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function findCodeRow_(sh, code) {
  var key = normCode_(code);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (normCode_(rows[i][0]) === key) return i + 1;
  }
  return -1;
}

function verifyCode_(code) {
  var key = normCode_(code);
  if (!key) return { ok: false, error: 'NO_CODE' };

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONSULTANTS);
  var rows = sh.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    if (normCode_(rows[i][0]) !== key) continue;

    var sheetCode = String(rows[i][COL_CODE   - 1]).trim();
    var role      = String(rows[i][COL_ROLE   - 1]).trim().toUpperCase();
    var nation    = String(rows[i][COL_NATION - 1] || '').trim();
    var org       = rows[i][COL_ORG    - 1];
    var expiry    = rows[i][COL_EXPIRY - 1];
    var active    = String(rows[i][COL_ACTIVE - 1]).trim().toUpperCase();

    // 역할 칸이 비면 코드 접두사(AD/ED/CS/MS)로 보완
    if (!ROLE_SCOPE[role]) {
      var byPrefix = PREFIX_ROLE[key.substring(0, 2)];
      if (byPrefix) role = byPrefix;
    }

    if (active !== 'Y') return { ok: false, error: 'INACTIVE' };
    if (expiry instanceof Date && expiry < new Date()) return { ok: false, error: 'EXPIRED' };
    if (!ROLE_SCOPE[role]) return { ok: false, error: 'BAD_ROLE' };

    sh.getRange(i + 1, COL_LASTSEEN).setValue(new Date());   // I 최근접속
    return {
      ok: true,
      code: sheetCode,
      role: ROLE_LABEL[role],
      name: rows[i][COL_NAME - 1],
      nation: nation,
      org:  org,
      modes: ROLE_SCOPE[role],
      expiry: expiry instanceof Date
        ? Utilities.formatDate(expiry, 'Asia/Seoul', 'yyyy-MM-dd') : ''
    };
  }
  return { ok: false, error: 'NOT_FOUND' };
}

// ─────────────────────────────────────────────────────────────
// 4) 동시접속 차단 — L(현재세션)·M(세션시각)만 쓴다. 다른 열은 건드리지 않음.
// ─────────────────────────────────────────────────────────────
function claimSession_(code, session) {
  var v = verifyCode_(code);
  if (!v.ok) return { ok: false, error: v.error };
  if (!session) return { ok: false, error: 'NO_SESSION' };

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONSULTANTS);
  var row = findCodeRow_(sh, code);
  if (row < 0) return { ok: false, error: 'NOT_FOUND' };

  var lock = LockService.getScriptLock();
  try { lock.waitLock(3000); } catch (e) {}
  sh.getRange(row, COL_SESSION).setValue(session);
  sh.getRange(row, COL_SESSTIME).setValue(new Date());
  try { lock.releaseLock(); } catch (e) {}

  return { ok: true, session: session };
}

function beatSession_(code, session) {
  if (!session) return { ok: false, error: 'NO_SESSION' };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONSULTANTS);
  var row = findCodeRow_(sh, code);
  if (row < 0) return { ok: false, error: 'NOT_FOUND' };

  var cur = String(sh.getRange(row, COL_SESSION).getValue() || '');
  if (cur && cur !== session) {
    var t = sh.getRange(row, COL_SESSTIME).getValue();
    var age = (t instanceof Date) ? (new Date() - t) : Infinity;
    if (age < SESSION_TIMEOUT_MS) return { ok: false, error: 'TAKEOVER' };
  }
  sh.getRange(row, COL_SESSION).setValue(session);
  sh.getRange(row, COL_SESSTIME).setValue(new Date());
  return { ok: true };
}

// ─────────────────────────────────────────────────────────────
// 2) 기록 저장
// ─────────────────────────────────────────────────────────────
function saveRecord_(r) {
  var v = verifyCode_(r.code);
  if (!v.ok) return { ok: false, error: 'AUTH_' + v.error };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ensureSheet_(ss, SHEET_RECORDS, HEAD_RECORDS);
  var id = 'KC' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMddHHmmss')
                + Math.floor(Math.random() * 900 + 100);

  var lab = r.lab_raw || {};
  var jaw = lab.jaw || {}, chk = lab.cheek || {}, st = lab.sites || {};
  var jl = st.jl || {}, jr = st.jr || {}, cl = st.cl || {}, cr = st.cr || {};

  sh.appendRow([
    new Date(), v.code, v.role, r.flow_ko || '',
    r.name || '', r.email || '', r.birth || '', r.sex || '', r.tel || '',
    r.race || '', r.job || '', r.date || '',
    r.skin || '', r.hair || '', r.eye || '',
    num_(r.skin_lightness), num_(r.area_diff), num_(r.redness), num_(r.evenness),
    num_(jaw.L), num_(jaw.a), num_(jaw.b), num_(chk.L), num_(chk.a), num_(chk.b),
    num_(jl.L), num_(jl.a), num_(jl.b), num_(jr.L), num_(jr.a), num_(jr.b),
    num_(cl.L), num_(cl.a), num_(cl.b), num_(cr.L), num_(cr.a), num_(cr.b),
    num_(lab.db), num_(lab.dAB),
    r.time || r.season || '', r.type || '', r.type_ko || '', r.style_group || '',
    r.bob || '', (r.best || []).join(', '),
    r.memo || '', r.agree ? 'Y' : 'N', r.pref || '', id
  ]);

  // J 누적 진단수 +1 (다른 열은 건드리지 않음)
  var cs = ss.getSheetByName(SHEET_CONSULTANTS);
  var rows = cs.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (normCode_(rows[i][0]) === normCode_(v.code)) {
      cs.getRange(i + 1, COL_COUNT).setValue((Number(rows[i][COL_COUNT - 1]) || 0) + 1);
      break;
    }
  }
  return { ok: true, id: id };
}

function num_(x) {
  return (x === null || x === undefined || x === '') ? '' : Number(x);
}

// ─────────────────────────────────────────────────────────────
// 3) 결과 조회 (공유 링크) — 열 이름으로 찾으므로 열 순서가 바뀌어도 안전
// ─────────────────────────────────────────────────────────────
function readRecord_(id) {
  if (!id) return { ok: false, error: 'NO_ID' };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECORDS);
  var rows = sh.getDataRange().getValues();
  if (rows.length < 2) return { ok: false, error: 'NOT_FOUND' };
  var H = rows[0];
  var col = function (n) { return H.indexOf(n); };
  var idc = col('결과ID');

  for (var i = rows.length - 1; i >= 1; i--) {
    if (idc < 0 || String(rows[i][idc]) !== String(id)) continue;
    var g = function (n) { var c = col(n); return c >= 0 ? rows[i][c] : ''; };
    return {
      ok: true,
      name:    g('이름'),
      date:    g('진단일'),
      season:  g('타임'),
      type:    g('유형코드'),
      type_ko: g('유형'),
      bob:     g('베스트 오브 베스트'),
      best:    String(g('베스트 컬러(BOC)')).split(',')
                 .map(function (s) { return s.trim(); })
                 .filter(function (s) { return s; })
    };
  }
  return { ok: false, error: 'NOT_FOUND' };
}

// ─────────────────────────────────────────────────────────────
// 관리 도구 — 코드 발급
//   아래 4줄만 바꿔서 실행하면 consultants 맨 아래에 새 코드가 추가된다.
//   코드 형식: 역할코드-YYNNN-국적   예) CS-26001-KR
// ─────────────────────────────────────────────────────────────
function issueCode() {
  var role   = 'Consultant';        // Admin / Educator / Consultant / Master
  var name   = '(이름)';
  var nation = 'KR';                // 국적 코드 (2자리 권장)
  var org    = '색다른컬러연구소';   // 취득기관
  var months = 3;                   // 가입개월 (숫자)

  var key = String(role).trim().toUpperCase();
  if (!ROLE_SCOPE[key]) throw new Error('역할은 Admin / Educator / Consultant / Master 중 하나여야 합니다: ' + role);
  months = Number(months) || 0;
  nation = String(nation).trim().toUpperCase();

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONSULTANTS);
  var prefix = ROLE_PREFIX[key];
  var yy = ('0' + (new Date().getFullYear() % 100)).slice(-2);   // 26

  // 같은 역할·같은 연도(YY) 중 일련번호 최댓값 + 1
  var rows = sh.getDataRange().getValues(), max = 0;
  for (var i = 1; i < rows.length; i++) {
    var parts = String(rows[i][0]).toUpperCase().split('-');   // [CS, 26001, KR]
    if (parts.length < 2) continue;
    if (parts[0].replace(/[^A-Z]/g, '') !== prefix) continue;
    var mid = parts[1].replace(/[^0-9]/g, '');                 // 26001
    if (mid.length < 3 || mid.substring(0, 2) !== yy) continue;
    var seq = Number(mid.substring(2));
    if (seq > max) max = seq;
  }
  var nnn = ('00' + (max + 1)).slice(-3);
  var code = prefix + '-' + yy + nnn + '-' + nation;          // CS-26001-KR

  var today  = new Date();
  var expiry = new Date(today.getFullYear(), today.getMonth() + months, today.getDate());

  // 코드, 역할, 이름, 국적, 소속, 발급일, 만료일, 활성, 최근접속, 누적, 가입개월, 현재세션, 세션시각
  sh.appendRow([code, ROLE_LABEL[key], name, nation, org, today, expiry, 'Y', '', 0, months, '', '']);
  Logger.log('발급: ' + code + ' / ' + ROLE_LABEL[key] + ' / 국적 ' + nation +
             ' / 가입 ' + months + '개월 / 만료 ' +
             Utilities.formatDate(expiry, 'Asia/Seoul', 'yyyy-MM-dd'));
}

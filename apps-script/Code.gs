/**
 * K-Color 컬러진단표준시스템 (CASS) — Apps Script 게이트웨이
 * 색다른컬러연구소 · DARUN COLOR INSTITUTE
 *
 * 하는 일
 *   1) 코드 검증 : 앱이 보낸 코드를 consultants 탭에서 찾아 권한을 돌려준다
 *   2) 기록 저장 : 진단 결과 1건을 records 탭에 1행으로 추가한다
 *   3) 결과 조회 : 공유 링크용 읽기 전용 조회
 *
 * 처음 한 번 할 일
 *   - 스프레드시트를 만들고, 확장 프로그램 > Apps Script 에 이 파일을 붙여넣는다
 *   - setup() 을 한 번 실행한다 (탭·헤더가 자동으로 만들어진다)
 *   - 배포 > 새 배포 > 웹 앱
 *       실행 계정 : 나
 *       액세스   : 모든 사용자
 *     → 나오는 /exec URL 을 index.html 의 API_URL 에 넣는다
 */

// ─────────────────────────────────────────────────────────────
// 설정
// ─────────────────────────────────────────────────────────────
var SHEET_CONSULTANTS = 'consultants';
var SHEET_RECORDS     = 'records';
var SHEET_RULES       = 'rules';

var HEAD_CONSULTANTS = [
  '코드', '역할', '이름', '소속', '발급일', '만료일', '활성여부', '최근접속', '누적 진단수'
];

var HEAD_RECORDS = [
  '타임스탬프', '코드', '역할',
  '이름', '이메일', '생년월일', '성별', '연락처', '인종', '직업', '진단일',
  '피부명도', '모발색', '눈동자색',
  '피부색 밝기', '부위별 밝기차', '홍조 강도', '피부 균일도',
  '턱 L*', '턱 a*', '턱 b*', '볼 L*', '볼 a*', '볼 b*',
  '시즌', '유형코드', '유형',
  '베스트 오브 베스트', '베스트 컬러(BOC)',
  '컨설턴트 메모', '동의여부', '결과ID'
];

var HEAD_RULES = ['항목', '값', '설명', '수정자', '수정일'];

// ─────────────────────────────────────────────────────────────
// 최초 설정 — 한 번만 실행
// ─────────────────────────────────────────────────────────────
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_CONSULTANTS, HEAD_CONSULTANTS);
  ensureSheet_(ss, SHEET_RECORDS,     HEAD_RECORDS);
  ensureSheet_(ss, SHEET_RULES,       HEAD_RULES);

  // 예시 코드 3개 — 실제 운영 전에 지우고 새로 발급하세요
  var cs = ss.getSheetByName(SHEET_CONSULTANTS);
  if (cs.getLastRow() < 2) {
    var today = new Date();
    var next  = new Date(today.getFullYear() + 1, today.getMonth(), today.getDate());
    cs.appendRow(['AD-0001', 'Admin',      '유잰',  '색다른컬러연구소', today, next, 'Y', '', 0]);
    cs.appendRow(['ED-0001', 'Educator',   '(예시)', '',                today, next, 'Y', '', 0]);
    cs.appendRow(['CS-0001', 'Consultant', '(예시)', '',                today, next, 'Y', '', 0]);
  }
  try {
    SpreadsheetApp.getUi().alert('설정 완료 — consultants / records / rules 탭이 준비되었습니다.');
  } catch (e) { Logger.log('설정 완료'); }
}

function ensureSheet_(ss, name, head) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0 || sh.getRange(1, 1).getValue() !== head[0]) {
    sh.getRange(1, 1, 1, head.length).setValues([head]);
  }
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
/**
 * 역할 값 — consultants 탭 '역할' 열에 아래 세 글자 중 하나를 정확히 적습니다.
 * 대소문자는 가리지 않습니다 (admin / Admin / ADMIN 모두 됨).
 *   Admin      관리자     — 진단 · 교육 · 관리자
 *   Educator   에듀케이터 — 교육
 *   Consultant 컨설턴트   — 진단
 */
var ROLE_SCOPE = {
  'ADMIN':      ['diag', 'edu', 'adm'],
  'EDUCATOR':   ['edu'],
  'CONSULTANT': ['diag']
};
var ROLE_LABEL  = { 'ADMIN': 'Admin', 'EDUCATOR': 'Educator', 'CONSULTANT': 'Consultant' };
var ROLE_PREFIX = { 'ADMIN': 'AD',    'EDUCATOR': 'ED',       'CONSULTANT': 'CS' };
var PREFIX_ROLE = { 'AD': 'ADMIN',    'ED': 'EDUCATOR',       'CS': 'CONSULTANT' };

/**
 * 코드 정규화 — 대문자로 올리고 영문·숫자만 남긴다.
 *   'ad 0002' · 'Ad-0002' · 'AD0002'           → 'AD0002'
 *   'ad-2608kr-001' · 'AD 2608 KR 001'         → 'AD2608KR001'
 * 시트에 적힌 코드와 앱에서 입력한 코드를 이 형태로 맞춰 비교한다.
 */
function normCode_(s) {
  return String(s === null || s === undefined ? '' : s).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function verifyCode_(code) {
  var key = normCode_(code);
  if (!key) return { ok: false, error: 'NO_CODE' };

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONSULTANTS);
  var rows = sh.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    if (normCode_(rows[i][0]) !== key) continue;

    var sheetCode = String(rows[i][0]).trim();          // 시트에 적힌 원래 표기
    var role   = String(rows[i][1]).trim().toUpperCase();
    var expiry = rows[i][5];
    var active = String(rows[i][6]).trim().toUpperCase();

    // 역할 칸이 비어 있으면 코드 접두사(AD/ED/CS)로 보완한다
    if (!ROLE_SCOPE[role]) {
      var byPrefix = PREFIX_ROLE[key.substring(0, 2)];
      if (byPrefix) role = byPrefix;
    }

    if (active !== 'Y') return { ok: false, error: 'INACTIVE' };
    if (expiry instanceof Date && expiry < new Date()) return { ok: false, error: 'EXPIRED' };
    if (!ROLE_SCOPE[role]) return { ok: false, error: 'BAD_ROLE' };

    sh.getRange(i + 1, 8).setValue(new Date());   // 최근접속
    return {
      ok: true,
      code: sheetCode,
      role: ROLE_LABEL[role],
      name: rows[i][2],
      org:  rows[i][3],
      modes: ROLE_SCOPE[role],
      expiry: expiry instanceof Date
        ? Utilities.formatDate(expiry, 'Asia/Seoul', 'yyyy-MM-dd') : ''
    };
  }
  return { ok: false, error: 'NOT_FOUND' };
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
  var jaw = lab.jaw || {}, chk = lab.cheek || {};

  sh.appendRow([
    new Date(), v.code, v.role,
    r.name || '', r.email || '', r.birth || '', r.sex || '', r.tel || '',
    r.race || '', r.job || '', r.date || '',
    r.skin || '', r.hair || '', r.eye || '',
    num_(r.skin_lightness), num_(r.area_diff), num_(r.redness), num_(r.evenness),
    num_(jaw.L), num_(jaw.a), num_(jaw.b), num_(chk.L), num_(chk.a), num_(chk.b),
    r.season || '', r.type || '', r.type_ko || '',
    r.bob || '', (r.best || []).join(', '),
    r.memo || '', r.agree ? 'Y' : 'N', id
  ]);

  // 누적 진단수 +1
  var cs = ss.getSheetByName(SHEET_CONSULTANTS);
  var rows = cs.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (normCode_(rows[i][0]) === normCode_(v.code)) {
      cs.getRange(i + 1, 9).setValue((Number(rows[i][8]) || 0) + 1);
      break;
    }
  }
  return { ok: true, id: id };
}

function num_(x) {
  return (x === null || x === undefined || x === '') ? '' : Number(x);
}

// ─────────────────────────────────────────────────────────────
// 3) 결과 조회 (공유 링크)
// ─────────────────────────────────────────────────────────────
function readRecord_(id) {
  if (!id) return { ok: false, error: 'NO_ID' };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECORDS);
  var rows = sh.getDataRange().getValues();
  var last = HEAD_RECORDS.length - 1;

  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][last]) !== String(id)) continue;
    return {
      ok: true,
      name:    rows[i][3],
      date:    rows[i][10],
      season:  rows[i][24],
      type:    rows[i][25],
      type_ko: rows[i][26],
      bob:     rows[i][27],
      best:    String(rows[i][28]).split(',')
                 .map(function (s) { return s.trim(); })
                 .filter(function (s) { return s; })
    };
  }
  return { ok: false, error: 'NOT_FOUND' };
}

// ─────────────────────────────────────────────────────────────
// 관리 도구 — 코드 발급
//   Apps Script 편집기에서 함수 목록에 issueCode 를 고르고 실행하면
//   consultants 탭 맨 아래에 새 코드가 추가된다.
//   기본값을 바꾸려면 아래 세 줄을 고쳐서 실행하세요.
// ─────────────────────────────────────────────────────────────
function issueCode() {
  var role   = 'Consultant';        // Admin / Educator / Consultant
  var name   = '(이름)';
  var org    = '';
  var months = 12;                  // 유효 개월 수

  var key = String(role).trim().toUpperCase();
  if (!ROLE_SCOPE[key]) throw new Error('역할은 Admin / Educator / Consultant 중 하나여야 합니다: ' + role);

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONSULTANTS);
  var prefix = ROLE_PREFIX[key];

  // 같은 접두사 중 맨 뒤 숫자가 가장 큰 값 + 1
  var rows = sh.getDataRange().getValues(), max = 0;
  for (var i = 1; i < rows.length; i++) {
    var n = normCode_(rows[i][0]);
    if (n.substring(0, 2) !== prefix) continue;
    var m = n.match(/(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  }
  var code = prefix + '-' + ('000' + (max + 1)).slice(-4);

  var today = new Date();
  var expiry = new Date(today.getFullYear(), today.getMonth() + months, today.getDate());

  sh.appendRow([code, ROLE_LABEL[key], name, org, today, expiry, 'Y', '', 0]);
  Logger.log('발급: ' + code + ' / ' + ROLE_LABEL[key] + ' / 만료 ' +
             Utilities.formatDate(expiry, 'Asia/Seoul', 'yyyy-MM-dd'));
}

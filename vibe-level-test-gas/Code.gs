/**
 * 바이브 레벨 테스트 - 구글 앱스크립트 서버 코드
 *
 * 이 스크립트는 구글 시트에 붙여서(확장 프로그램 > Apps Script) 사용합니다.
 * - doGet: 웹 앱 주소로 들어오면 index.html 화면을 보여 줍니다.
 * - saveResponse: 참여자 응답을 '응답' 시트에 한 줄씩 저장합니다.
 * - listResponses: 교사 화면에 회차별 응답을 돌려줍니다.
 */

var SHEET_NAME = '응답';
var HEADER = ['제출 시각', '회차', '이름', '총점', '레벨',
  '개념 이해', '도구 경험', '프롬프트 감각', '수업 적용',
  'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'Q11', 'Q12'];

/* 문항별 영역(index.html의 Q 배열 순서와 같아야 합니다) */
var Q_CATS = [0, 0, 0, 1, 1, 1, 2, 2, 2, 3, 3, 3];

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var boot = {
    params: params,
    url: ScriptApp.getService().getUrl()
  };
  var t = HtmlService.createTemplateFromFile('index');
  // </script> 로 스크립트가 끊기지 않도록 '<'를 이스케이프합니다.
  t.boot = JSON.stringify(boot).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle('바이브 레벨 테스트')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* 응답 저장 */
function saveResponse(room, rec) {
  room = cleanRoom_(room);
  if (!rec || !Array.isArray(rec.a) || rec.a.length !== Q_CATS.length) {
    throw new Error('응답 형식이 올바르지 않아요.');
  }
  // 점수는 서버에서 다시 계산합니다(문항별 점수만 믿고 나머지는 재계산).
  var a = rec.a.map(function (v) {
    v = Number(v);
    if (!(v >= 0 && v <= 3) || Math.floor(v) !== v) throw new Error('응답 값이 올바르지 않아요.');
    return v;
  });
  var c = [0, 0, 0, 0], s = 0;
  a.forEach(function (v, i) { c[Q_CATS[i]] += v; s += v; });
  var name = String(rec.n || '').trim().slice(0, 12);

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getSheet_();
    sh.appendRow([new Date(), safeText_(room), safeText_(name), s, levelOf_(s)]
      .concat(c).concat(a));
  } finally {
    lock.releaseLock();
  }
  return true;
}

/* 회차별 응답 목록 */
function listResponses(room) {
  room = cleanRoom_(room);
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, HEADER.length).getValues();
  var out = [];
  rows.forEach(function (r) {
    if (String(r[1]) !== room) return;
    var a = r.slice(9, 21).map(Number);
    var c = r.slice(5, 9).map(Number);
    var t = (r[0] instanceof Date) ? r[0].getTime() : 0; // Date는 그대로 넘길 수 없어서 숫자로 바꿉니다.
    out.push({ n: String(r[2] || ''), t: t, s: Number(r[3]), l: Number(r[4]), c: c, a: a });
  });
  return out;
}

/* ---------- 도우미 ---------- */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
  }
  return sh;
}

function levelOf_(s) { return s <= 8 ? 1 : s <= 16 ? 2 : s <= 24 ? 3 : s <= 31 ? 4 : 5; }

function cleanRoom_(r) {
  r = String(r || 'main').replace(/[^A-Za-z0-9가-힣_-]/g, '').slice(0, 20);
  return r || 'main';
}

/* 글자를 항상 '텍스트'로 저장합니다. 앞에 '를 붙이면 =수식, 5/3(날짜), 007(숫자) 같은 자동 변환을 막을 수 있어요.
   시트 화면과 읽어 온 값에는 '가 보이지 않습니다. */
function safeText_(v) {
  v = String(v);
  return v ? "'" + v : v;
}

/* 편집기에서 한 번 실행하면 시트 권한을 승인하고 '응답' 시트를 만들어 둡니다. */
function setup() {
  getSheet_();
}

/**
 * 생활 약속 기록 — Google 시트 저장 서버 (Apps Script)
 *
 * 설치 방법은 저장소의 SETUP-GOOGLE-SHEET.md 참고.
 * 이 스크립트는 스프레드시트에 연결된(컨테이너 바인드) Apps Script로 붙여넣고,
 * 웹 앱으로 배포해서 사용합니다.
 */

// ▼▼ 배포 전에 반드시 나만 아는 값으로 바꾸세요 (앱 설정의 '비밀 코드'와 동일하게) ▼▼
var TOKEN = '여기에-비밀코드-입력';

// ▼▼ 자녀 조회 페이지용 코드 (읽기 전용). TOKEN과 다른 값으로 정하고, 자녀에게는 이 코드만 알려주세요 ▼▼
var VIEW_TOKEN = '여기에-조회코드-입력';

var SHEET_NAME = '기록';
var HEADERS = ['자녀', '날짜', '기기사용(분)', '국영수 확인', '국어 문항', '수학 문항', '영어 문항',
               '문항 합계(구버전)', '암기 단어', '기상 시각', '아침 준비', '공기계 제출',
               '올림피아드', '특이사항', '수정시각'];

function doGet(e) {
  var p = (e && e.parameter) || {};
  // 조회(GET)는 비밀 코드 또는 조회 코드 둘 다 허용, 쓰기(POST)는 비밀 코드만
  var authorized = p.token && (p.token === TOKEN || p.token === VIEW_TOKEN);
  if (!authorized) return json_({ ok: false, error: 'unauthorized' });
  var child = p.child || '';
  var sh = getSheet_();
  var data = sh.getDataRange().getValues();
  var records = [];
  for (var i = 1; i < data.length; i++) {
    if (child && String(data[i][0]) !== child) continue;
    records.push(rowToRecord_(data[i]));
  }
  return json_({ ok: true, records: records });
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad request' });
  }
  if (body.token !== TOKEN) return json_({ ok: false, error: 'unauthorized' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getSheet_();
    if (body.action === 'save' && body.record) {
      upsert_(sh, String(body.child || ''), body.record);
    } else if (body.action === 'saveAll' && body.records) {
      for (var i = 0; i < body.records.length; i++) {
        upsert_(sh, String(body.child || ''), body.records[i]);
      }
    } else if (body.action === 'delete' && body.date) {
      var row = findRow_(sh, String(body.child || ''), String(body.date));
      if (row > 0) sh.deleteRow(row);
    } else {
      return json_({ ok: false, error: 'unknown action' });
    }
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    // 날짜·기상 시각 열이 자동 변환되지 않도록 텍스트 서식 고정
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange('J:J').setNumberFormat('@');
  }
  return sh;
}

function upsert_(sh, child, r) {
  var rowVals = [
    child,
    String(r.date || ''),
    numCell_(r.screenTime),
    boolLabel_(r.studyShown, '보여줌', '안 보여줌'),
    numCell_(r.koreanCount),
    numCell_(r.mathCount),
    numCell_(r.englishCount),
    numCell_(r.problemCount),
    numCell_(r.wordCount),
    r.wakeTime ? String(r.wakeTime) : '',
    boolLabel_(r.morningReady, '이행', '미이행'),
    boolLabel_(r.deviceSubmit, '제출', '미제출'),
    boolLabel_(r.olympiadPrep, '함', '안 함'),
    r.memo ? String(r.memo) : '',
    r.updatedAt ? String(r.updatedAt) : ''
  ];
  var row = findRow_(sh, child, String(r.date || ''));
  if (row > 0) {
    sh.getRange(row, 1, 1, HEADERS.length).setValues([rowVals]);
  } else {
    sh.appendRow(rowVals);
  }
}

function rowToRecord_(row) {
  return {
    date: dateStr_(row[1]),
    screenTime: cellNum_(row[2]),
    studyShown: labelBool_(row[3], '보여줌', '안 보여줌'),
    koreanCount: cellNum_(row[4]),
    mathCount: cellNum_(row[5]),
    englishCount: cellNum_(row[6]),
    problemCount: cellNum_(row[7]),
    wordCount: cellNum_(row[8]),
    wakeTime: row[9] ? timeStr_(row[9]) : null,
    morningReady: labelBool_(row[10], '이행', '미이행'),
    deviceSubmit: labelBool_(row[11], '제출', '미제출'),
    olympiadPrep: labelBool_(row[12], '함', '안 함'),
    memo: row[13] ? String(row[13]) : null,
    updatedAt: row[14] ? String(row[14]) : null
  };
}

function findRow_(sh, child, date) {
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === child && dateStr_(data[i][1]) === date) return i + 1;
  }
  return -1;
}

// 시트가 날짜/시각 문자열을 Date로 자동 변환한 경우까지 문자열로 복원
function dateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}
function timeStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  return String(v);
}
function boolLabel_(v, yes, no) {
  return v === true ? yes : (v === false ? no : '');
}
function labelBool_(v, yes, no) {
  v = String(v);
  return v === yes ? true : (v === no ? false : null);
}
function numCell_(v) {
  return typeof v === 'number' ? v : '';
}
function cellNum_(v) {
  return (v === '' || v === null || v === undefined) ? null : Number(v);
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

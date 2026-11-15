/* =========================================================
 *  새 앱(클라우드플레어) 연결용 — 기존 Apps Script 프로젝트에 파일로 추가
 *  스크립트 속성에 CRM2_URL(예: https://crm.drawyourlife.co.kr), CRM2_KEY 를 넣어주세요.
 * ========================================================= */

function crm2Conf_() {
  var p = PropertiesService.getScriptProperties();
  var url = String(p.getProperty('CRM2_URL') || '').replace(/\/+$/, '');
  var key = String(p.getProperty('CRM2_KEY') || '');
  if (!url || !key) throw new Error('스크립트 속성에 CRM2_URL, CRM2_KEY 를 먼저 넣어주세요.');
  return { url: url, key: key };
}

function crm2Post_(path, payload) {
  var c = crm2Conf_();
  var res = UrlFetchApp.fetch(c.url + path, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-crm-key': c.key }, payload: JSON.stringify(payload || {})
  });
  var code = res.getResponseCode();
  if (code >= 300) throw new Error(path + ' 실패(' + code + '): ' + res.getContentText().slice(0, 300));
  return JSON.parse(res.getContentText());
}

/* 5분마다: 새 네이버 예약 메일을 새 앱으로 전달 */
function crm2_메일전달() {
  var labelName = 'CRM2전달';
  var label = GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);
  var threads = GmailApp.search('subject:"네이버 예약" newer_than:3d -label:' + labelName, 0, 30);
  if (!threads.length) return;
  var mails = [];
  threads.forEach(function (th) {
    th.getMessages().forEach(function (m) {
      mails.push({ id: m.getId(), subject: m.getSubject(), date: m.getDate().getTime(),
                   from: m.getFrom(), plain: m.getPlainBody(), html: m.getBody() });
    });
  });
  var r = crm2Post_('/ingest', { mails: mails });
  threads.forEach(function (th) { th.addLabel(label); });
  Logger.log(JSON.stringify(r.stat));
}

/* 시계 작업들 */
function crm2_자동정산() { crm2Post_('/cron/settle'); }
function crm2_사전안내() { crm2Post_('/cron/advance'); }
function crm2_일일정리() { crm2Post_('/cron/daily'); }

/* 한 번만 실행: 위 작업들의 자동 실행 시간 등록 */
function crm2_자동실행_설치() {
  var fns = ['crm2_메일전달', 'crm2_자동정산', 'crm2_사전안내', 'crm2_일일정리'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (fns.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('crm2_메일전달').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('crm2_자동정산').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('crm2_사전안내').timeBased().atHour(14).nearMinute(0).everyDays(1).create();
  ScriptApp.newTrigger('crm2_일일정리').timeBased().atHour(1).nearMinute(15).everyDays(1).create();
  Logger.log('새 앱 자동실행 등록 완료');
}

/* 한 번만 실행: 지금 스프레드시트 데이터를 새 앱으로 통째로 복사 (새 앱 쪽 같은 탭은 덮어씀) */
function crm2_데이터_보내기() {
  var names = ['설정', '네이버예약_데이터', '고객DB', '발송로그', '시술메뉴', '매출내역', '동의서양식',
               '동의서기록', '간단일정', '수입지출', '선불포인트내역', '휴지통', '접속기록'];
  var book = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SHEET_ID);
  var report = {};
  names.forEach(function (n) {
    var sh = book.getSheetByName(n);
    if (!sh || sh.getLastRow() < 1) return;
    var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues().map(function (row) {
      return row.map(function (x) { return (x instanceof Date) ? { $d: x.getTime() } : x; });
    });
    var r = crm2Post_('/admin/import', { sheets: (function () { var o = {}; o[n] = v; return o; })() });
    report[n] = r.imported[n];
  });
  /* 앱 비밀번호·직원 PIN 은 설정 탭에 있어 같이 옮겨집니다. 접속 토큰만 안 옮겨서 한 번 다시 로그인하면 됩니다 */
  Logger.log(JSON.stringify(report));
  return report;
}

/* 자동 생성 파일 — 직접 고치지 말고 tools/gen.mjs 로 다시 만드세요 */
import { KDate, makeServices } from '../src/shim.js';

export function makeGas(book, io) {
  const Date = KDate;
  const { SpreadsheetApp, Utilities, UrlFetchApp, GmailApp, Logger, PropertiesService, CacheService,
          ScriptApp, CalendarApp, DriveApp, MailApp, HtmlService, Session, LockService } = makeServices(book, io);
  const globalThis = {};

/* ===== 원본 Code.gs ===== */
/* 편집기에서 바로 실행할 수 있는 단축 함수 (함수 목록 맨 위에 나옵니다) */
function A_지금_메일확인() { var s = processBookingMails(); Logger.log(JSON.stringify(s)); return s; }
function A_요청사항_다시읽기() { return rescanRequests(); }
function A_자동실행_시간_다시맞추기() { installTriggers(); Logger.log('자동실행 시간을 다시 맞췄습니다.'); return 'ok'; }
function A_시술메뉴_기본값으로_다시만들기() {
  var sh = ss().getSheetByName(SH.MENU);
  if (sh) ss().deleteSheet(sh);
  var s2 = buildMenuSheet();
  Logger.log('시술메뉴 ' + (s2.getLastRow() - 1) + '개 만들었습니다.');
  return s2.getLastRow() - 1;
}
function A_전체초기화_후_재수집() { return resetAndReimport(); }
function A_메일진단() {
  var out = [];
  var q1 = 'subject:"네이버 예약" newer_than:7d';
  var t1 = GmailApp.search(q1, 0, 30);
  out.push('[1] ' + q1 + '  ->  스레드 ' + t1.length + '개');
  t1.forEach(function (t) { t.getMessages().forEach(function (msg) {
    out.push('   ' + fmt(msg.getDate(), 'MM-dd HH:mm') + ' | ' + msg.getSubject());
  }); });
  var q2 = 'from:navercorp.com newer_than:7d';
  var t2 = GmailApp.search(q2, 0, 30);
  out.push('[2] ' + q2 + '  ->  스레드 ' + t2.length + '개');
  t2.forEach(function (t) { t.getMessages().forEach(function (msg) {
    out.push('   ' + fmt(msg.getDate(), 'MM-dd HH:mm') + ' | ' + msg.getFrom() + ' | ' + msg.getSubject());
  }); });
  var q3 = '취소 newer_than:7d';
  var t3 = GmailApp.search(q3, 0, 30);
  out.push('[3] ' + q3 + '  ->  스레드 ' + t3.length + '개');
  t3.forEach(function (t) { t.getMessages().forEach(function (msg) {
    out.push('   ' + fmt(msg.getDate(), 'MM-dd HH:mm') + ' | ' + msg.getFrom() + ' | ' + msg.getSubject());
  }); });
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/*************************************************************************
 *  뷰티샵 CRM  v1.0
 *  당신을 그리다(퍼스널컬러) / 닥터 뷰티랩(반영구)
 *  구글 스프레드시트 + Apps Script
 *
 *  ※ 이 코드는 수정할 필요가 없습니다.
 *    모든 설정은 스프레드시트의 [설정] 탭에서 바꿉니다.
 *************************************************************************/

/* ===================== 기본 상수 ===================== */

var SH = {
  HOME:    '홈',
  DATA:    '네이버예약_데이터',
  CUST:    '고객DB',
  PROFILE: '고객프로필',
  LOG:     '발송로그',
  CONFIG:  '설정',
  SALES:   '매출통계_',
  CAL:     '예약캘린더_',
  MENU:    '시술메뉴',
  SLOG:    '매출내역',
  FORM:    '동의서양식',
  CONSENT: '동의서기록',
  BLOCK:   '간단일정',
  CASH:    '수입지출',
  PPLOG:   '선불포인트내역',
  TRASH:   '휴지통',
  AUTH:    '접속기록'
};
var AUTH_HEAD = ['시각', '이름', '권한', '결과'];

var GRADES = ['VIP', '단골', '신규', '재방문', '휴면', '블랙리스트'];
var BLOCK_HEAD = ['ID', '샵', '날짜', '시작', '종료', '제목', '메모', '등록시각'];
var CASH_HEAD = ['ID', '거래일', '구분', '계정', '내용', '금액', '결제수단', '샵', '메모', '등록시각'];
var PP_HEAD = ['등록시각', '고객ID', '고객명', '종류', '구분', '금액', '잔액', '메모', '담당자'];
var TRASH_HEAD = ['삭제시각', '종류', '식별자', '요약', '데이터'];

/* 네이버예약_데이터 열 번호 */
var D = {
  RESNO:1, SHOP:2, STATUS:3, NAME:4, PHONE:5, SAFE:6, DATE:7, TIME:8,
  PRODUCT:9, OPTION:10, PEOPLE:11, REQUEST:12, AMOUNT:13, PAID:14,
  STAFF:15, SOURCE:16, MAILAT:17, PROCAT:18, CUSTID:19, CALID:20,
  MAILID:21, MEMO:22
};
var D_HEAD = ['예약번호','샵','상태','예약자명','연락처','안심번호','예약일','예약시간',
  '예약상품','선택메뉴(옵션)','인원','요청사항','예약금액','실결제금액','담당자',
  '등록경로','메일수신시각','처리시각','고객ID','구글캘린더ID','메일ID','비고'];

/* 고객DB 열 번호 */
var C = {
  ID:1, NAME:2, PHONE:3, SAFE:4, SHOP:5, FIRST:6, LAST:7, VISITS:8,
  PRODUCT:9, TOTAL:10, USED:11, LEFT:12, MEMO:13, SOURCE:14, FOLDER:15, NEXT:16,
  GENDER:17, BIRTH:18, STAFF:19, PREPAID:20, REL:21, NEEDLE:22, PIGMENT:23, TONE:24,
  GRADE:25, TAG:26, POINT:27, LANG:28, CWAY:29, CID:30, CAUTION:31, ROUTE:32
};
var C_HEAD = ['고객ID','이름','연락처','안심번호(050)','주이용샵','최초방문일','최근방문일',
  '총방문횟수','회차상품명','총회차','사용회차','잔여회차','고객메모(특이사항)',
  '등록경로','사진폴더','다음예약일',
  '성별','생년월일','담당자','잔여 선불액','연결 고객','니들','색소','퍼스널컬러 톤',
  '등급','태그','잔여 포인트','응대 언어','연락 방법','연락 아이디','주의사항(상담지)','방문경로(상담지)'];

/* 발송로그 열 */
var L_HEAD = ['발송시각','예약번호','샵','고객명','연락처','유형','채널','결과','메시지','응답'];

var STATUS = { OK:'확정', CANCEL:'취소', CHANGE:'변경', DONE:'이용완료', CHECK:'확인필요' };

/* ===================== 메뉴 ===================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📅 샵 관리')
      .addItem('🚀 처음 한 번만 - 초기세팅', 'setup')
      .addSeparator()
      .addItem('📩 지금 새 예약메일 확인', 'checkMailNow')
      .addItem('✍️ 수동예약 반영 (캘린더에 직접 쓴 예약)', 'importManualBookings')
      .addItem('🔗 고객 매칭 다시하기', 'rematchCustomers')
      .addSeparator()
      .addItem('🗓 다음달 캘린더 만들기', 'createNextMonthCalendars')
      .addItem('🗓 이번달 캘린더 다시 그리기', 'refreshThisMonthCalendars')
      .addSeparator()
      .addItem('✅ 지난 예약 이용완료 처리', 'completePastBookings')
      .addItem('💰 매출통계 새로고침', 'refreshAllSales')
      .addItem('🏠 홈 화면 새로고침', 'refreshHome')
      .addSeparator()
      .addItem('💬 선택한 예약에 문자/알림톡 보내기', 'sendToSelected')
      .addItem('💬 사전 안내 전체 발송 (수동)', 'sendReminderAdvanceManual')
      .addItem('💬 내일 예약자 전체 발송 (수동)', 'sendReminderTomorrowManual')
      .addItem('💬 오늘 예약자 전체 발송 (수동)', 'sendReminderTodayManual')
      .addItem('💬 발송 테스트 (내 번호로 1건)', 'sendTestMessage')
      .addSeparator()
      .addItem('👤 고객프로필 새로고침', 'refreshProfile')
      .addItem('📁 선택한 고객 사진폴더 만들기', 'createFolderForSelectedCustomer')
      .addSeparator()
      .addItem('🔧 자동실행 시간 다시 맞추기', 'installTriggers')
      .addItem('❓ 상태 점검(진단)', 'runDiagnostics')
      .addItem('🧪 최근 예약메일 읽기 테스트', 'testParseLatestMail')
    .addToUi();
}

/* ===================== 설정 읽기 ===================== */

var _cfgCache = null;

function cfg(key, dflt) {
  if (!_cfgCache) {
    _cfgCache = {};
    var sh = ss().getSheetByName(SH.CONFIG);
    if (sh) {
      var v = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 2).getValues();
      for (var i = 0; i < v.length; i++) {
        var k = String(v[i][0]).trim();
        if (k && k.indexOf('■') !== 0) _cfgCache[k] = v[i][1];
      }
    }
  }
  var val = _cfgCache[key];
  if (val === undefined || val === null || String(val).trim() === '') {
    return (dflt === undefined) ? '' : dflt;
  }
  return val;
}

function cfgStr(key, dflt) { return String(cfg(key, dflt)).trim(); }
function cfgOn(key) { return String(cfg(key, '')).trim() === '사용'; }

function setCfg(key, value) {
  var sh = ss().getSheetByName(SH.CONFIG);
  if (!sh) return;
  var v = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      if (_cfgCache) _cfgCache[key] = value;
      return;
    }
  }
  sh.appendRow([key, value, '']);
  if (_cfgCache) _cfgCache[key] = value;
}

var SHEET_ID = '1yT1GBi--CNHA3i4rVmmahmMo_6jX3bHq_3Y0wy5TkEQ';
function ss() {
  try { var a = SpreadsheetApp.getActiveSpreadsheet(); if (a) return a; } catch (e) {}
  return SpreadsheetApp.openById(SHEET_ID);
}
function tz() { return 'Asia/Seoul'; }

/* 샵 목록: [{name, short, keywords[], calKey, salesSheet, calPrefix}] */
function shops() {
  var out = [];
  for (var i = 1; i <= 2; i++) {
    var name = cfgStr('샵' + i + ' 이름', '');
    if (!name) continue;
    out.push({
      idx: i,
      name: name,
      short: cfgStr('샵' + i + ' 약칭', name),
      keywords: cfgStr('샵' + i + ' 식별키워드', name).split(',')
                  .map(function (s) { return s.trim(); })
                  .filter(function (s) { return s; }),
      calIdKey: '샵' + i + ' 구글캘린더ID',
      sendKey: '샵' + i + ' 발송 사용'
    });
  }
  return out;
}

function shopByName(name) {
  var list = shops();
  for (var i = 0; i < list.length; i++) if (list[i].name === name) return list[i];
  return null;
}

/* ===================== 유틸 ===================== */

function fmt(d, pattern) {
  if (!d) return '';
  if (!(d instanceof Date)) d = new Date(d);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, tz(), pattern);
}

function ymd(d) { return fmt(d, 'yyyy-MM-dd'); }
function hm(d)  { return fmt(d, 'HH:mm'); }

function toDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  var s = String(v).trim().replace(/\./g, '-').replace(/-+$/, '').replace(/\//g, '-');
  var m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

/* '오후 1:00' / '13:00' → 'HH:mm' */
function normTime(s) {
  if (!s) return '';
  if (s instanceof Date) return hm(s);
  var t = String(s).trim();
  var pm = /오후|PM/i.test(t);
  var am = /오전|AM/i.test(t);
  var m = t.match(/(\d{1,2})\s*[:시]\s*(\d{1,2})?/);
  if (!m) return '';
  var h = Number(m[1]);
  var mi = Number(m[2] || 0);
  if (pm && h < 12) h += 12;
  if (am && h === 12) h = 0;
  return ('0' + h).slice(-2) + ':' + ('0' + mi).slice(-2);
}

function timeStr(v) { if (v instanceof Date) return hm(v); return String(v || '').trim(); }

function onlyDigits(s) { return String(s || '').replace(/[^0-9]/g, ''); }

function last4(phone) {
  var d = onlyDigits(phone);
  return d.length >= 4 ? d.slice(-4) : '';
}

function normPhone(s) {
  var d = onlyDigits(s);
  if (!d) return '';
  if (d.length === 11) return d.slice(0, 3) + '-' + d.slice(3, 7) + '-' + d.slice(7);
  if (d.length === 10) return d.slice(0, 3) + '-' + d.slice(3, 6) + '-' + d.slice(6);
  return String(s).trim();
}

function isSafePhone(s) { return /^050/.test(onlyDigits(s)); }

function nowStr() { return fmt(new Date(), 'yyyy-MM-dd HH:mm:ss'); }

/* '2026-08-29 04:35:12' 같은 문자열을 비교 가능한 숫자로 */
function parseStamp(v) {
  if (v instanceof Date) return v.getTime();
  var m = String(v || '').match(/(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})/);
  if (!m) return 0;
  return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]).getTime();
}

function toast(msg, title) {
  try { ss().toast(msg, title || '샵 관리', 6); } catch (e) {}
}

function alert(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function sheetOf(name) {
  var s = ss().getSheetByName(name);
  if (!s) throw new Error('[' + name + '] 탭이 없습니다. 메뉴에서 "초기세팅"을 먼저 실행해 주세요.');
  return s;
}

function getOrCreateSheet(name) {
  var s = ss().getSheetByName(name);
  if (!s) s = ss().insertSheet(name);
  return s;
}

/* 데이터 시트 전체 읽기 */
function readData() {
  var sh = sheetOf(SH.DATA);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, D_HEAD.length).getValues();
}

/* 관리자에게 오류/알림 메일 */
function notifyAdmin(subject, body) {
  var to = cfgStr('관리자 이메일', '');
  if (!to) { try { to = Session.getEffectiveUser().getEmail(); } catch (e) { return; } }
  if (!to) return;
  try {
    MailApp.sendEmail(to, '[샵 CRM] ' + subject,
      body + '\n\n---\n스프레드시트 열기: ' + ss().getUrl());
  } catch (e) { Logger.log('알림메일 실패: ' + e); }
}

function logError(where, e) {
  var msg = where + ' : ' + (e && e.message ? e.message : e);
  Logger.log(msg);
  try { setCfg('마지막 오류', nowStr() + ' / ' + msg); } catch (ignore) {}
  notifyAdmin('오류가 발생했습니다', 
    '아래 작업에서 문제가 생겼어요.\n\n작업: ' + where + '\n내용: ' + msg +
    '\n\n대부분은 잠시 후 자동으로 다시 시도됩니다. 반복되면 개발자에게 이 메일을 보여주세요.');
}

/* =========================================================
 *  1. 초기세팅 (한 번만 실행)
 * ========================================================= */

function setup() {
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }
  if (ui) {
    var r = ui.alert('초기세팅',
      '필요한 탭·서식·자동실행·구글캘린더·사진폴더를 한 번에 만듭니다.\n\n' +
      '이미 있는 자료는 지우지 않습니다.\n계속할까요?', ui.ButtonSet.YES_NO);
    if (r !== ui.Button.YES) return;
  }
  try {
    try { ss().setSpreadsheetTimeZone('Asia/Seoul'); } catch (e) {}
    buildConfigSheet();
    _cfgCache = null;
    buildDataSheet();
    buildCustomerSheet();
    buildLogSheet();
    buildProfileSheet();
    buildSalesSheets();
    buildMenuSheet();
    buildSlogSheet();
    buildFormSheet();
    seedForms(false);
    buildConsentSheet();
    buildBlockSheet();
    buildCashSheet();
    buildPpSheet();
    buildTrashSheet();
    buildHomeSheet();
    createCalendarsForMonth(new Date());
    createCalendarsForMonth(nextMonthDate(new Date()));
    ensureGoogleCalendars();
    ensureDriveRoot();
    installTriggers();
    protectSheets();
    orderTabs();
    refreshHome();
    setCfg('마지막 오류', '');
    toast('초기세팅이 끝났습니다!', '완료');
    if (ui) ui.alert('초기세팅 완료',
      '준비가 끝났습니다.\n\n다음으로 할 일:\n' +
      '1) [설정] 탭에서 담당자 명단을 입력하세요.\n' +
      '2) 네이버 예약 메일이 이 지메일로 오도록 전달 설정을 해주세요.\n' +
      '3) 메뉴 → "지금 새 예약메일 확인"을 눌러 테스트해 보세요.', ui.ButtonSet.OK);
  } catch (e) {
    logError('초기세팅', e);
    if (ui) ui.alert('초기세팅 중 문제가 생겼습니다:\n\n' + e.message);
  }
}

function orderTabs() {
  var order = [SH.HOME];
  shops().forEach(function (s) {
    order.push(SH.CAL + s.short + '_' + fmt(new Date(), 'yyyy-MM'));
  });
  order.push(SH.DATA, SH.CUST, SH.PROFILE, SH.MENU, SH.SLOG, SH.FORM, SH.CONSENT);
  shops().forEach(function (s) { order.push(SH.SALES + s.short); });
  order.push(SH.LOG, SH.CONFIG);
  var pos = 1;
  order.forEach(function (n) {
    var sh = ss().getSheetByName(n);
    if (sh) { ss().setActiveSheet(sh); ss().moveActiveSheet(pos++); }
  });
  ss().setActiveSheet(ss().getSheetByName(SH.HOME));
}

/* ---------- 설정 시트 ---------- */

function defaultConfigRows() {
  return [
    ['■ 기본 설정', '', '이 탭의 값만 바꾸면 됩니다. 코드는 건드리지 않아도 됩니다.'],
    ['관리자 이메일', 'drbeautylab.official@gmail.com', '오류·알림을 받을 메일 주소'],
    ['샵1 이름', '당신을 그리다', '첫 번째 샵 이름 (메시지에 그대로 들어갑니다)'],
    ['샵1 약칭', '그리다', '탭 이름에 쓰는 짧은 이름 (한글 2~4자)'],
    ['샵1 식별키워드', '당신을 그리다,당신을그리다', '예약메일에서 이 샵을 알아볼 단어. 쉼표로 여러 개'],
    ['샵2 이름', '닥터 뷰티랩', ''],
    ['샵2 약칭', '뷰티랩', ''],
    ['샵2 식별키워드', '닥터 뷰티랩,닥터뷰티랩,뷰티랩', ''],
    ['담당자 명단', '원장', '캘린더 담당자 칸의 선택목록. 쉼표로 구분 (예: 원장,김실장,이실장)'],
    ['직원 인센티브(%)', '', '직원별 결산에서 인센티브 계산. 예: 원장:30,김실장:20'],
    ['지출 계정 목록', '재료비,월세,관리비,공과금,마케팅비,광고비,수수료,인건비,장비,교육,기타', '수입지출 화면에서 고르는 계정 목록'],
    ['네이버 수수료율(%)', '1.98', '네이버 결제 정산 수수료 (VAT 포함). 일반예약상품 1.98 / 네이버 주문상품 0.88 / 해외카드 3.85'],
    ['신용카드 수수료율(%)', '0.40', '카드 결제 수수료. 연매출 3억 이하 우대요율 0.40'],
    ['체크카드 수수료율(%)', '0.15', '체크카드 수수료. 연매출 3억 이하 우대요율 0.15'],
    ['하루 예약칸 수', '10', '캘린더에서 하루에 적을 수 있는 예약 줄 수'],
    ['영업 시작시각', '10:00', '캘린더 참고용'],
    ['영업 종료시각', '20:00', '캘린더 참고용'],
    ['예약 기본 소요시간(분)', '90', '구글캘린더에 등록될 예약 1건의 길이'],

    ['■ 예약메일 수집', '', ''],
    ['메일 자동수집 사용', '사용', '사용 / 미사용'],
    ['메일 검색조건', 'subject:"네이버 예약" newer_than:3d',
      '지메일 검색어. 메일이 안 잡히면 newer_than:7d 로 늘려보세요'],
    ['처리완료 라벨', 'CRM처리완료', '처리한 메일에 붙는 지메일 라벨 (중복처리 방지)'],
    ['한 번에 처리할 메일 수', '20', '너무 크게 하면 시간초과가 납니다'],

    ['■ 메일 읽는 규칙 (형식이 바뀌면 여기만 수정)', '', '정규식입니다. 잘 모르면 그대로 두세요'],
    ['규칙_제목_취소', '취소', '제목에 이 말이 있으면 취소 메일로 봅니다'],
    ['규칙_제목_변경', '변경|수정', ''],
    ['규칙_제목_확정', '확정|접수|신청|완료', ''],
    ['규칙_예약번호', '예약번호\\s*[:：]?\\s*([0-9]{6,})', ''],
    ['규칙_예약자명', '예약자명\\s*[:：]?\\s*([^\\n\\r]{1,30})', ''],
    ['규칙_이용일시', '이용일시\\s*[:：]?\\s*([^\\n\\r]{1,60})', ''],
    ['규칙_예약상품', '예약상품\\s*[:：]?\\s*([^\\n\\r]{1,100})', ''],
    ['규칙_선택메뉴', '선택메뉴\\s*[:：]?\\s*([^\\n\\r]{1,150})', ''],
    ['규칙_요청사항', '요청사항\\s*[:：]?\\s*([^\\n\\r]{0,300})', ''],
    ['규칙_연락처', '(?:연락처|휴대폰|전화번호|휴대전화)\\s*[:：]?\\s*(0[0-9\\-\\s]{8,16})', '메일에 없으면 비어있습니다'],
    ['규칙_결제상태', '결제상태\\s*[:：]?\\s*([^\\n\\r]{1,20})', ''],

    ['■ 앱 잠금 / 직원 로그인', '', '휴대폰에서 앱을 열 때 물어볼 번호입니다'],
    ['앱 비밀번호', '', '사장님 번호(전체 권한). 비워두면 잠그지 않습니다. 4자 이상. 앱의 [도구] 탭에서도 바꿀 수 있습니다'],
    ['직원 PIN', '', '직원용. 이름:번호 형식, 쉼표로 여러 명 (예: 김실장:482913,이실장:770124)'],
    ['직원 매출 보기', '숨김', '숨김 / 보임 — 직원에게 매출·수입지출·직원결산을 보여줄지'],

    ['■ 알림톡 / 문자 발송 (알리고)', '', '가입 전에는 비워두세요. 비어있으면 발송하지 않습니다'],
    ['알리고 사용자ID', '', 'smartsms.aligo.in 로그인 아이디'],
    ['알리고 API키', '', '알리고 사이트에서 발급'],
    ['발신번호', '', '알리고에 사전 등록한 번호 (예: 010-1234-5678)'],
    ['카카오채널 senderkey', '', '알림톡용. 없으면 문자(SMS)로만 나갑니다'],
    ['알림톡 템플릿코드_사전', '', ''],
    ['알림톡 템플릿코드_전일', '', ''],
    ['알림톡 템플릿코드_당일', '', ''],
    ['알림톡 템플릿코드_확정', '', ''],
    ['테스트 모드', '사용', '"사용"이면 실제로 보내지 않고 발송로그만 남깁니다. 실제 발송하려면 "미사용"'],
    ['샵1 발송 사용', '사용', ''],
    ['샵2 발송 사용', '사용', ''],
    ['사전 알림 사용', '사용', ''],
    ['사전 알림 며칠 전', '3', '예약일 며칠 전에 사전 안내를 보낼지 (예약 변경 기한에 맞추세요)'],
    ['사전 알림 시각', '14:00', '이 시각에 며칠 뒤 예약자에게 사전 안내 발송'],
    ['전일 알림 사용', '사용', ''],
    ['전일 알림 시각', '14:00', '전날 이 시각에 내일 예약자에게 발송'],
    ['당일 알림 사용', '사용', ''],
    ['당일 알림 시각', '09:00', '당일 이 시각에 오늘 예약자에게 발송'],
    ['확정 즉시 안내 사용', '미사용', '예약이 잡히자마자 안내문자 보내기'],
    ['테스트 수신번호', '', '"발송 테스트" 메뉴가 이 번호로 1건 보냅니다'],

    ['■ 문자 문구', '', '#{고객명} #{상점명} #{날짜요일} #{시각} #{예약일시} #{소요시간} #{상품} #{옵션} #{날짜} #{시간} 은 자동으로 바뀝니다'],
    ['문구_사전', '[#{샵}]\n#{이름}님 안녕하세요.\n#{예약일시} 예약 안내드립니다.\n· 시술: #{상품}\n변경이나 취소는 이 번호로 연락 주세요.', ''],
    ['문구_전일', '[#{샵}]\n#{이름}님 안녕하세요.\n내일 #{시간} 예약 안내드립니다.\n· 시술: #{상품}\n변경이나 취소는 이 번호로 연락 주세요.', ''],
    ['문구_당일', '[#{샵}]\n#{이름}님, 오늘 #{시간}에 뵙겠습니다.\n· 시술: #{상품}\n조심히 와주세요!', ''],
    ['문구_확정', '[#{샵}]\n#{이름}님, 예약이 확정되었습니다.\n· 일시: #{날짜} #{시간}\n· 시술: #{상품}\n감사합니다.', ''],

    ['샵1 문구_사전', '', '당신을 그리다 사전 안내. 비워두면 위의 공통 문구를 씁니다'],
    ['샵2 문구_사전', "#{고객명}님, 안녕하세요.\n#{예약일시} #{상품} 예약 안내드립니다.\n소요시간은 #{소요시간} 정도입니다.\n\n■ 먼저 확인해 주세요 (시술 불가)\n아래에 해당하시면 시술을 진행할 수 없습니다.\n매장으로 꼭 전화 주세요. 010-2722-3093\n이 경우 위약금 없이 변경·취소해 드립니다.\n\n임산부, 수유부, 켈로이드 피부, 아스피린 계열·고지혈증약 복용, 간염 보균, 심한 여드름성 피부, 이전 시술 잔흔이 강하게 남은 경우, 지루성 피부염, 트러블이 많은 두피\n\n그 밖의 질환이나 복용 중인 약이 있으시면 미리 전화로 말씀해 주세요. 당일에 확인되면 시술이 어려울 수 있습니다.\n\n■ 시술 전날까지\n· 음주는 피해주세요. 출혈과 붓기가 늘고 색소가 잘 자리 잡지 않습니다.\n· 시술 부위 각질제거, 필링, 레이저는 2주 전부터 피해주세요.\n\n■ 시술 당일\n· 카페인은 줄여주세요. 통증에 예민해집니다.\n· 충분히 주무시고 편한 컨디션으로 와주세요.\n· 두피 시술은 머리를 감고 와주세요.\n· 잔여 색소가 손이나 옷에 묻어날 수 있어 밝은 색이나 고가의 의상은 피해주세요. 의류 오염은 책임지지 않습니다.\n\n■ 디자인 안내\n자연스럽고 안전한 시술을 지향합니다. 전문가의 판단으로 부자연스럽거나 과한 디자인은 시술이 어려울 수 있습니다.\n\n■ 예약 안내\n· 100% 예약제 매장입니다. 정시 방문 부탁드립니다.\n· 그 외 사유의 예약 변경은 3일 전까지 1회 가능하며, 변경·취소 시 예약금은 환불되지 않습니다.\n· 15분 이상 지각 시 다음 예약 일정으로 인해 자동 취소되며, 노쇼 시 재예약이 제한됩니다.\n· 건물 지하 주차장을 2시간 무료로 이용하실 수 있습니다.\n\n안내를 확인하지 못해 생기는 불이익은 책임지기 어려우니 꼭 읽어봐 주세요.", '닥터 뷰티랩 사전 안내 (예약 3일 전). 전화번호를 실제 번호로 바꿔주세요'],
    ['샵1 문구_전일', "안녕하세요 #{고객명}님 :)\n당신을 그리다 마곡점입니다.\n예약 안내사항 보내드립니다.\n\n<< 예약일 : #{예약일시} >>\n\n*퍼스널컬러 진단*\n1.정확한 진단을 위해서 꼭 노메이크업으로 방문해 주세요. (메이크업 후 방문 시 세안하셔야 하며, 진단 시간에 포함됩니다.)\n2.정확한 눈동자 색상 확인을 위해 컬러렌즈 착용 불가합니다. (투명렌즈 착용 가능)\n3.안경은 진단 시 벗으셔야 하기 때문에 시력이 많이 안좋으신 경우 투명렌즈 착용을 추천드립니다.\n\n*메이크업 클래스*\n1.노메이크업으로 방문 해주세요.\n2.평소 사용하시는 메이크업 제품과 브러쉬를 가져와주세요. 제품이 없을 경우 매장에서 대여해드리며, 브러쉬는 매장에서 구매 가능합니다.\n3.브러쉬는 세척 후 방문해주세요. 미세척 시 위생 문제로 화장품 대여 불가합니다.\n\n**공통 안내사항**\n- 건물 지하 주차장에 주차 가능합니다. 컨설팅 시간만큼 무료 주차권 제공해 드립니다.\n- 예약시간 정시 방문을 꼭 부탁드립니다! 대기 공간이 없기 때문에 앞 타임 고객님이 계실 때 미리 도착하시면 외부에서 대기하셔야 합니다. 다음 예약 고객님을 위해 15분 이상 늦으실 경우 시술은 자동 예약 취소되며, 교육은 기존 시간에 맞춰서 종료됩니다.(당일 취소/노쇼 예약금 환불 불가)\n- 본 메세지는 자동 발송됩니다. 추가 문의는 네이버 톡톡으로 남겨주시면 확인하는대로 답변 드리겠습니다. 외부 강의&예약 일정으로 답변이 늦어질 수 있는 점 미리 양해 부탁드립니다!\n\n* [네이버지도]로 '당신을그리다' 검색 하시면 길 찾기 편리합니다! https://naver.me/GUTluIrH", '당신을 그리다 전용. 비워두면 위의 공통 문구를 씁니다'],
    ['샵2 문구_전일', "▶ 날짜 : #{날짜요일}\n▶ 시간 : #{시각}\n▶ 소요시간 : #{소요시간}\n\n안녕하세요, 닥터뷰티랩입니다.\n예약 안내 사항 보내드립니다. 사전 안내 미숙지로 인한 불이익은 책임지지 않습니다. 꼭 확인 부탁드립니다!\n\n[탈모 커버]\n- 시술 직후 두피에 잔여 색소가 남아 손이나 옷에 묻어날 수 있으니, 방문 시 고가의상·밝은 색 옷은 피해 주세요. 의류 오염은 책임지지 않으니 편한 복장으로 방문 부탁드립니다.\n\n[눈썹, 입술 커버]\n- 자연스럽고 안전한 시술을 지향합니다. 전문가의 판단으로 부자연스럽거나 과한 디자인은 시술이 불가할 수 있습니다.\n\n[공통 안내]\n- 100% 예약제 매장입니다. 정시 방문을 부탁드립니다. 예약 변경/취소 시 예약금 환불 불가하며, 예약 변경은 예약 3일 전까지 1회 변경 가능합니다. 15분 이상 지각 시 다음 예약 스케줄로 인해 자동 예약 취소되며, 노쇼 시 재예약이 제한됩니다.\n- 매장 건물 지하 주차장이 있습니다. (2시간 무료 이용 가능)\n- 시술 불가\n임산부/수유부/켈로이드피부/아스피린계열, 고지혈증약 복용/간염보균자/심한 여드름성피부/이전시술 잔흔이 강하게 남으신 분/지루성두피염/트러블이 많은 두피 고객님은 미리 연락 부탁드립니다. 기타 질환 및 질병은 대면 상담 시 꼭 말씀해 주세요!\n\n* [네이버지도]로 '닥터뷰티랩' 검색 하시면 길 찾기 편리합니다! https://naver.me/GL80O8yd", '닥터 뷰티랩 전용. 비워두면 위의 공통 문구를 씁니다'],

    ['■ 시스템 (자동으로 채워집니다 · 건드리지 마세요)', '', ''],
    ['사진 루트폴더ID', '', ''],
    ['샵1 구글캘린더ID', '', ''],
    ['샵2 구글캘린더ID', '', ''],
    ['마지막 메일확인 시각', '', ''],
    ['마지막 오류', '', '']
  ];
}

function buildConfigSheet() {
  var sh = getOrCreateSheet(SH.CONFIG);
  var rows = defaultConfigRows();
  var existing = {};
  if (sh.getLastRow() > 0) {
    var cur = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
    cur.forEach(function (r) {
      var k = String(r[0]).trim();
      if (k && k.indexOf('■') !== 0) existing[k] = r[1];
    });
  }
  sh.clear();
  var out = rows.map(function (r) {
    var k = String(r[0]).trim();
    var v = (existing[k] !== undefined && String(existing[k]).trim() !== '') ? existing[k] : r[1];
    return [r[0], v, r[2]];
  });
  sh.getRange(1, 1, out.length, 3).setValues(out);
  sh.setColumnWidth(1, 230).setColumnWidth(2, 380).setColumnWidth(3, 460);
  sh.getRange(1, 1, out.length, 3).setVerticalAlignment('middle').setWrap(true);
  // 섹션 머리글 서식
  for (var i = 0; i < out.length; i++) {
    if (String(out[i][0]).indexOf('■') === 0) {
      sh.getRange(i + 1, 1, 1, 3).setBackground('#1f3864').setFontColor('#ffffff').setFontWeight('bold');
    }
  }
  sh.getRange(1, 3, out.length, 1).setFontColor('#888888').setFontSize(9);
  sh.setFrozenColumns(1);
  return sh;
}

/* ---------- 데이터 시트 ---------- */

function buildDataSheet() {
  var sh = getOrCreateSheet(SH.DATA);
  if (sh.getLastRow() === 0) {
    sh.appendRow(D_HEAD);
  } else {
    sh.getRange(1, 1, 1, D_HEAD.length).setValues([D_HEAD]);
  }
  var h = sh.getRange(1, 1, 1, D_HEAD.length);
  h.setBackground('#1f3864').setFontColor('#ffffff').setFontWeight('bold')
   .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 34);
  var widths = [110, 100, 80, 90, 120, 120, 95, 70, 220, 240, 50, 200, 90, 100, 80, 90, 140, 140, 80, 240, 200, 200];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.hideColumns(D.CALID, 2); // 구글캘린더ID, 메일ID 숨김
  // 상태 조건부 서식
  var rng = sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), D_HEAD.length);
  var rules = [
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$C2="' + STATUS.CANCEL + '"')
      .setBackground('#f0f0f0').setFontColor('#999999').setStrikethrough(true).setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$C2="' + STATUS.CHECK + '"')
      .setBackground('#ffe0e0').setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($C2="' + STATUS.OK + '",$E2="",$F2="")')
      .setBackground('#fff4d6').setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$C2="' + STATUS.DONE + '"')
      .setBackground('#e8f5e9').setRanges([rng]).build()
  ];
  sh.setConditionalFormatRules(rules);
  // 상태 드롭다운
  var sv = SpreadsheetApp.newDataValidation()
    .requireValueInList([STATUS.OK, STATUS.CANCEL, STATUS.CHANGE, STATUS.DONE, STATUS.CHECK], true)
    .setAllowInvalid(true).build();
  sh.getRange(2, D.STATUS, Math.max(sh.getMaxRows() - 1, 1), 1).setDataValidation(sv);
  sh.getRange(2, D.DATE, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, D.TIME, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  sh.getRange(2, D.AMOUNT, Math.max(sh.getMaxRows() - 1, 1), 2).setNumberFormat('#,##0');
  return sh;
}

/* ---------- 고객DB ---------- */

function buildCustomerSheet() {
  var sh = getOrCreateSheet(SH.CUST);
  if (sh.getLastRow() === 0) sh.appendRow(C_HEAD);
  else sh.getRange(1, 1, 1, C_HEAD.length).setValues([C_HEAD]);
  sh.getRange(1, 1, 1, C_HEAD.length).setBackground('#7b3f61').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 34);
  var w = [90, 90, 120, 120, 90, 100, 100, 80, 200, 60, 60, 70, 300, 90, 220, 100,
           70, 100, 90, 100, 160, 160, 160, 200, 100, 160, 90];
  w.forEach(function (x, i) { sh.setColumnWidth(i + 1, x); });
  var maxR = Math.max(sh.getMaxRows() - 1, 1);
  var rng = sh.getRange(2, 1, maxR, C_HEAD.length);
  sh.setConditionalFormatRules([
    // 잔여회차 1 이하 → 재구매 관리 대상
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($J2<>"",$L2<=1)')
      .setBackground('#ffd9d9').setRanges([rng]).build(),
    // 연락처 없음 → 확인 필요
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($B2<>"",$C2="")')
      .setBackground('#fff4d6').setRanges([rng]).build()
  ]);
  var shopList = shops().map(function (s) { return s.name; });
  shopList.push('둘다');
  if (shopList.length > 1) {
    sh.getRange(2, C.SHOP, maxR, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(shopList, true).setAllowInvalid(true).build());
  }
  sh.getRange(2, C.FIRST, maxR, 2).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, C.NEXT, maxR, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, C.BIRTH, maxR, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, C.PREPAID, maxR, 1).setNumberFormat('#,##0');
  sh.getRange(2, C.GENDER, maxR, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['여', '남'], true).setAllowInvalid(true).build());
  sh.getRange(2, C.POINT, maxR, 1).setNumberFormat('#,##0');
  sh.getRange(2, C.GRADE, maxR, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(GRADES, true).setAllowInvalid(true).build());
  return sh;
}

/* ---------- 발송로그 ---------- */

function buildLogSheet() {
  var sh = getOrCreateSheet(SH.LOG);
  if (sh.getLastRow() === 0) sh.appendRow(L_HEAD);
  else sh.getRange(1, 1, 1, L_HEAD.length).setValues([L_HEAD]);
  sh.getRange(1, 1, 1, L_HEAD.length).setBackground('#3d5a3d').setFontColor('#ffffff').setFontWeight('bold');
  sh.setFrozenRows(1);
  [140, 110, 100, 90, 120, 80, 80, 80, 380, 220].forEach(function (x, i) { sh.setColumnWidth(i + 1, x); });
  return sh;
}

/* ---------- 고객프로필 ---------- */

function buildProfileSheet() {
  var sh = getOrCreateSheet(SH.PROFILE);
  sh.clear();
  sh.clearNotes();
  try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations(); } catch (e) {}

  sh.getRange('A1:H1').merge().setValue('👤 고객 프로필 조회')
    .setBackground('#7b3f61').setFontColor('#ffffff').setFontWeight('bold')
    .setFontSize(14).setHorizontalAlignment('center');
  sh.setRowHeight(1, 36);

  sh.getRange('A3').setValue('① 검색어').setFontWeight('bold');
  sh.getRange('B3').setValue('').setBackground('#fffbe6').setBorder(true, true, true, true, false, false);
  sh.getRange('C3').setValue('← 이름 또는 전화 뒤 4자리를 입력하고 Enter').setFontColor('#888888');

  sh.getRange('A4').setValue('② 고객 선택').setFontWeight('bold');
  sh.getRange('B4').setValue('').setBackground('#eef5ff').setBorder(true, true, true, true, false, false);
  sh.getRange('C4').setValue('← 목록에서 고르면 아래에 자동으로 표시됩니다').setFontColor('#888888');

  var labels = [
    ['이름', '연락처'],
    ['주이용샵', '안심번호'],
    ['총방문횟수', '최근방문일'],
    ['회차상품', '잔여회차 / 총회차'],
    ['다음 예약', '사진폴더'],
    ['메모', '']
  ];
  var r = 6;
  sh.getRange(r - 1, 1, 1, 8).merge().setValue('기본 정보')
    .setBackground('#f2e7ee').setFontWeight('bold');
  labels.forEach(function (pair, i) {
    sh.getRange(r + i, 1).setValue(pair[0]).setFontWeight('bold').setBackground('#faf6f8');
    if (pair[1]) {
      sh.getRange(r + i, 3).setValue(pair[1]).setFontWeight('bold').setBackground('#faf6f8');
      sh.getRange(r + i, 2).setBackground('#ffffff');
      sh.getRange(r + i, 4, 1, 5).merge().setBackground('#ffffff');
    } else {
      sh.getRange(r + i, 2, 1, 7).merge().setBackground('#ffffff').setWrap(true)
        .setVerticalAlignment('top');
      sh.setRowHeight(r + i, 48);
    }
  });

  var hr = r + labels.length + 1; // 방문 히스토리 헤더 행
  sh.getRange(hr, 1, 1, 8).merge().setValue('방문 히스토리 (최근 순)')
    .setBackground('#f2e7ee').setFontWeight('bold');
  var hh = ['방문일', '샵', '시술내용(예약상품)', '상세내용(선택메뉴)', '금액', '담당자', '상태', '사진'];
  sh.getRange(hr + 1, 1, 1, 8).setValues([hh])
    .setBackground('#7b3f61').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center');

  [110, 200, 110, 260, 100, 80, 80, 200].forEach(function (x, i) { sh.setColumnWidth(i + 1, x); });
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(4, 260);

  // 숨김 영역: 검색 후보 목록 (Z열)
  sh.getRange('Z1').setValue('검색후보(자동)');
  sh.hideColumns(26);

  PropertiesService.getScriptProperties().setProperty('PROFILE_HIST_ROW', String(hr + 2));
  PropertiesService.getScriptProperties().setProperty('PROFILE_INFO_ROW', String(r));
  return sh;
}

function profileRows() {
  var p = PropertiesService.getScriptProperties();
  return {
    info: Number(p.getProperty('PROFILE_INFO_ROW') || 6),
    hist: Number(p.getProperty('PROFILE_HIST_ROW') || 14)
  };
}

/* ---------- 시술메뉴 · 매출내역 ---------- */

var MENU_HEAD = ['샵', '시술명', '기본금액', '소요시간(분)'];
var SLOG_HEAD = ['등록시각', '예약번호', '샵', '고객ID', '고객명', '시술일', '시술명', '금액', '결제수단', '담당자', '네이버결제액'];

function buildMenuSheet() {
  var sh = getOrCreateSheet(SH.MENU);
  if (headerMismatch(sh, MENU_HEAD)) {
    sh.clear();
    sh.getRange(1, 1, 1, MENU_HEAD.length).setValues([MENU_HEAD])
      .setFontWeight('bold').setBackground('#1F3864').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 140); sh.setColumnWidth(2, 340);
    sh.setColumnWidth(3, 110); sh.setColumnWidth(4, 110);
    var seed = defaultMenuRows();
    if (seed.length) sh.getRange(2, 1, seed.length, MENU_HEAD.length).setValues(seed);
  }
  return sh;
}

function defaultMenuRows() {
  var list = shops();
  var s1 = list[0] ? list[0].name : '';
  var s2 = list[1] ? list[1].name : '';
  var out = [];
  var add = function (shop, arr) {
    if (!shop) return;
    arr.forEach(function (x) { out.push([shop, x[0], x[1], x[2] || 0]); });
  };
  /* 당신을 그리다 — 퍼스널컬러진단 */
  add(s1, [
    ['LIFE 1인', 300000], ['LIFE 2인', 500000], ['LIFE 3인', 690000],
    ['YOUR 1인', 200000], ['YOUR 2인', 300000], ['YOUR 3인', 390000],
    ['DRAW 1인', 110000], ['DRAW 2인', 209000],
    ['DRAW 3인', 240000], ['DRAW 4인', 300000],
    ['웨딩 컨설팅', 339000]
  ]);
  /* 당신을 그리다 — 메이크업클래스 */
  add(s1, [
    ['메이크업클래스 1인', 180000], ['메이크업클래스 2인', 320000],
    ['메이크업클래스 3인', 360000], ['메이크업클래스 4인', 440000]
  ]);
  /* 당신을 그리다 — 추가 옵션 */
  add(s1, [
    ['[추가] 헤어 디자인+컬러 추천', 20000],
    ['[추가] 아이브로우 디자인+컬러 추천', 20000],
    ['[추가] 패션 스타일링 추천+쥬얼리 추천', 20000],
    ['[추가] 안경 디자인+컬러 추천', 20000],
    ['[추가] 넥라인 추천', 20000],
    ['[추가] 얼굴 골격 분석', 50000],
    ['[추가] 체형 골격 분석', 50000],
    ['[추가] 메이크업클래스 1인', 150000],
    ['[추가] 메이크업 클래스 2인', 290000]
  ]);
  /* 당신을 그리다 — 제품 판매 */
  add(s1, [
    ['[제품] 브러쉬 파우더', 20000],
    ['[제품] 브러쉬 쉐딩', 20000],
    ['[제품] 브러쉬 치크', 20000],
    ['[제품] 브러쉬 섀도우', 10000],
    ['[제품] 브러쉬 하이라이트', 20000]
  ]);
  /* 닥터 뷰티랩 */
  add(s2, [
    ['퍼스널 브로우', 290000, 80],
    ['정수리 커버', 350000, 50],
    ['헤어라인 커버', 250000, 50],
    ['가르마 커버', 250000, 50],
    ['[추가] 커버 1부위 추가', 250000, 50],
    ['눈썹 리터치', 0, 50],
    ['두피 리터치', 0, 50],
    ['두피 생장 5회', 990000, 50],
    ['립 틴트', 360000, 110]
  ]);
  return out;
}

function headerMismatch(sh, head) {
  if (sh.getLastRow() === 0) return true;
  var cur = sh.getRange(1, 1, 1, head.length).getValues()[0];
  for (var i = 0; i < head.length; i++) {
    if (String(cur[i] || '').trim() !== head[i]) return true;
  }
  return false;
}

function buildSlogSheet() {
  var sh = getOrCreateSheet(SH.SLOG);
  /* 칸이 늘어난 경우: 기존 매출기록을 지우지 않고 새 칸만 덧붙입니다 */
  if (sh.getLastRow() > 0 && headerMismatch(sh, SLOG_HEAD)) {
    var oldHead = SLOG_HEAD.slice(0, 10);
    if (!headerMismatch(sh, oldHead)) {
      sh.getRange(1, 11, 1, SLOG_HEAD.length - 10)
        .setValues([SLOG_HEAD.slice(10)])
        .setFontWeight('bold').setBackground('#7B3F61').setFontColor('#ffffff');
      sh.setColumnWidth(11, 120);
    }
  }
  if (headerMismatch(sh, SLOG_HEAD)) {
    sh.clear();
    sh.getRange(1, 1, 1, SLOG_HEAD.length).setValues([SLOG_HEAD])
      .setFontWeight('bold').setBackground('#7B3F61').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    [150, 120, 140, 90, 100, 100, 340, 110, 110, 90, 120].forEach(function (x, i) {
      sh.setColumnWidth(i + 1, x);
    });
  }
  return sh;
}

var _menuCache = null;
function readMenu() {
  if (_menuCache) return _menuCache;
  var sh = buildMenuSheet();
  var last = sh.getLastRow();
  if (last < 2) { _menuCache = []; return _menuCache; }
  _menuCache = sh.getRange(2, 1, last - 1, MENU_HEAD.length).getValues()
    .filter(function (r) { return String(r[1] || '').trim(); })
    .map(function (r) {
      return { shop: String(r[0] || '').trim(), name: String(r[1]).trim(),
               price: Number(r[2] || 0), min: Number(r[3] || 0) };
    });
  return _menuCache;
}

function readSlog() {
  var sh = buildSlogSheet();
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, SLOG_HEAD.length).getValues()
    .filter(function (r) { return String(r[1] || '').trim(); })
    .map(function (r) {
      return { at: String(r[0] || ''), resno: String(r[1]).trim(), shop: String(r[2] || '').trim(),
               custId: String(r[3] || '').trim(), name: String(r[4] || '').trim(), date: ymd(r[5]),
               item: String(r[6] || '').trim(), amount: Number(r[7] || 0),
               pay: String(r[8] || '').trim(), staff: String(r[9] || '').trim(),
               npay: Number(r[10] || 0) };
    });
}

/* ---------- 매출통계 ---------- */

function buildSalesSheets() {
  shops().forEach(function (s) { buildSalesSheet(s); });
}

function buildSalesSheet(shop) {
  var name = SH.SALES + shop.short;
  var sh = getOrCreateSheet(name);
  var year = new Date().getFullYear();
  var prevYear = sh.getRange('B2').getValue();
  sh.clear();

  sh.getRange('A1:N1').merge().setValue('💰 ' + shop.name + ' 매출통계')
    .setBackground('#1f3864').setFontColor('#ffffff').setFontWeight('bold')
    .setFontSize(14).setHorizontalAlignment('center');
  sh.setRowHeight(1, 36);

  sh.getRange('A2').setValue('연도').setFontWeight('bold').setBackground('#eeeeee');
  sh.getRange('B2').setValue(prevYear || year).setBackground('#fffbe6')
    .setBorder(true, true, true, true, false, false).setNumberFormat('0');
  sh.getRange('C2').setValue('← 연도를 바꾸면 표가 자동으로 바뀝니다').setFontColor('#888888');

  sh.getRange('D2').setValue('연간 총매출').setFontWeight('bold').setBackground('#eeeeee');
  sh.getRange('E2').setFormula('=SUM(B5:M35)').setNumberFormat('₩#,##0').setFontWeight('bold');

  // 헤더: 일 / 1월~12월
  var head = ['일'];
  for (var m = 1; m <= 12; m++) head.push(m + '월');
  sh.getRange(4, 1, 1, 13).setValues([head])
    .setBackground('#1f3864').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center');

  var rows = [];
  for (var d = 1; d <= 31; d++) {
    var row = [d];
    for (var mm = 1; mm <= 12; mm++) {
      row.push('=IFERROR(IF(DAY(DATE($B$2,' + mm + ',$A' + (4 + d) + '))<>$A' + (4 + d) +
        ',"",SUMIFS(' + q(SH.DATA) + '!$N:$N,' + q(SH.DATA) + '!$B:$B,$A$3,' +
        q(SH.DATA) + '!$G:$G,DATE($B$2,' + mm + ',$A' + (4 + d) + '),' +
        q(SH.DATA) + '!$C:$C,"<>' + STATUS.CANCEL + '")),"")');
    }
    rows.push(row);
  }
  sh.getRange(5, 1, 31, 13).setValues(rows);

  // 샵 이름을 A3 에 (수식 참조용, 눈에 안 띄게)
  sh.getRange('A3').setValue(shop.name).setFontColor('#cccccc').setFontSize(8);

  // 월 합계
  sh.getRange(36, 1).setValue('월 합계').setFontWeight('bold').setBackground('#eeeeee');
  var sums = [];
  for (var k = 0; k < 12; k++) {
    var col = String.fromCharCode(66 + k); // B..M
    sums.push('=SUM(' + col + '5:' + col + '35)');
  }
  sh.getRange(36, 2, 1, 12).setValues([sums]).setFontWeight('bold').setBackground('#eeeeee');

  sh.getRange(37, 1).setValue('건수').setFontWeight('bold');
  var cnt = [];
  for (var k2 = 1; k2 <= 12; k2++) {
    cnt.push('=COUNTIFS(' + q(SH.DATA) + '!$B:$B,$A$3,' + q(SH.DATA) + '!$C:$C,"<>' + STATUS.CANCEL + '",' +
      q(SH.DATA) + '!$G:$G,">="&DATE($B$2,' + k2 + ',1),' +
      q(SH.DATA) + '!$G:$G,"<"&EDATE(DATE($B$2,' + k2 + ',1),1))');
  }
  sh.getRange(37, 2, 1, 12).setValues([cnt]);

  sh.getRange(5, 2, 32, 12).setNumberFormat('#,##0');
  sh.setColumnWidth(1, 60);
  for (var c = 2; c <= 13; c++) sh.setColumnWidth(c, 95);
  sh.setFrozenRows(4);
  sh.getRange(5, 1, 31, 1).setHorizontalAlignment('center').setBackground('#f7f7f7');
  return sh;
}

/* 시트 이름을 수식에 쓸 때 따옴표 처리 */
function q(name) { return "'" + String(name).replace(/'/g, "''") + "'"; }

/* ---------- 홈 ---------- */

function buildHomeSheet() {
  var sh = getOrCreateSheet(SH.HOME);
  sh.clear();
  sh.clearNotes();
  sh.getRange('A1:H1').merge().setValue('🌸  뷰티샵 예약·고객 관리')
    .setBackground('#1f3864').setFontColor('#ffffff').setFontWeight('bold')
    .setFontSize(18).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 52);

  sh.getRange('A2:H2').merge()
    .setValue('상단 메뉴  [📅 샵 관리]  에서 모든 기능을 실행합니다.')
    .setHorizontalAlignment('center').setFontColor('#666666');

  [140, 130, 130, 130, 130, 130, 130, 200].forEach(function (x, i) { sh.setColumnWidth(i + 1, x); });
  sh.setHiddenGridlines(true);
  refreshHome();
  return sh;
}

function refreshHome() {
  var sh = ss().getSheetByName(SH.HOME);
  if (!sh) return;
  var today = new Date();
  var t0 = ymd(today);
  var rows = readData();
  var list = shops();

  var box = function (row, title, color) {
    sh.getRange(row, 1, 1, 8).merge().setValue(title)
      .setBackground(color).setFontColor('#ffffff').setFontWeight('bold')
      .setVerticalAlignment('middle');
    sh.setRowHeight(row, 28);
  };

  sh.getRange(4, 1, 60, 8).clearContent().clearFormat().setBackground(null);

  var r = 4;
  box(r, '  📌 오늘 (' + t0 + ') 예약', '#7b3f61'); r++;

  list.forEach(function (s) {
    var todays = rows.filter(function (x) {
      return x[D.SHOP - 1] === s.name && ymd(x[D.DATE - 1]) === t0 && x[D.STATUS - 1] !== STATUS.CANCEL;
    }).sort(function (a, b) { return String(a[D.TIME - 1]).localeCompare(String(b[D.TIME - 1])); });

    sh.getRange(r, 1, 1, 8).merge()
      .setValue('▸ ' + s.name + '  —  ' + todays.length + '건')
      .setFontWeight('bold').setBackground('#f2e7ee');
    r++;
    if (todays.length === 0) {
      sh.getRange(r, 1, 1, 8).merge().setValue('   (오늘 예약 없음)').setFontColor('#999999');
      r++;
    } else {
      sh.getRange(r, 1, 1, 5).setValues([['시간', '예약자', '연락처', '시술', '담당자']])
        .setFontWeight('bold').setBackground('#eeeeee');
      r++;
      var vals = todays.map(function (x) {
        return [timeStr(x[D.TIME - 1]), x[D.NAME - 1], x[D.PHONE - 1] || '⚠︎ 없음',
                x[D.PRODUCT - 1], x[D.STAFF - 1] || ''];
      });
      sh.getRange(r, 1, vals.length, 5).setValues(vals);
      r += vals.length;
    }
    r++;
  });

  // 내일 예약 요약
  var tmr = new Date(today.getTime() + 86400000);
  var t1 = ymd(tmr);
  var tmrCount = rows.filter(function (x) {
    return ymd(x[D.DATE - 1]) === t1 && x[D.STATUS - 1] !== STATUS.CANCEL;
  }).length;

  box(r, '  ⚠️ 확인이 필요한 항목', '#b03a2e'); r++;
  var needCheck = rows.filter(function (x) { return x[D.STATUS - 1] === STATUS.CHECK; }).length;
  var noPhone = rows.filter(function (x) {
    return x[D.STATUS - 1] === STATUS.OK && !x[D.PHONE - 1] && !x[D.SAFE - 1] &&
           toDate(x[D.DATE - 1]) && toDate(x[D.DATE - 1]) >= toDate(t0);
  }).length;
  var cust = ss().getSheetByName(SH.CUST);
  var lowRound = 0;
  if (cust && cust.getLastRow() > 1) {
    cust.getRange(2, 1, cust.getLastRow() - 1, C_HEAD.length).getValues().forEach(function (x) {
      if (x[C.TOTAL - 1] !== '' && Number(x[C.LEFT - 1]) <= 1) lowRound++;
    });
  }
  var checks = [
    ['메일을 못 읽어 "확인필요"로 남은 예약', needCheck + ' 건', needCheck ? '→ 네이버예약_데이터 탭에서 직접 채워주세요' : '없음'],
    ['연락처가 비어있는 예약 (문자 발송 불가)', noPhone + ' 건', noPhone ? '→ 스마트플레이스에서 확인 후 입력' : '없음'],
    ['잔여회차 1회 이하 고객 (재구매 안내 대상)', lowRound + ' 명', lowRound ? '→ 고객DB 탭 붉은색 행' : '없음'],
    ['내일 예약', tmrCount + ' 건', '']
  ];
  sh.getRange(r, 1, checks.length, 3).setValues(checks);
  sh.getRange(r, 1, checks.length, 1).setFontWeight('bold');
  sh.getRange(r, 3, checks.length, 1).setFontColor('#888888');
  r += checks.length + 1;

  box(r, '  ⚙️ 상태', '#3d5a3d'); r++;
  var st = [
    ['마지막 메일 확인', cfgStr('마지막 메일확인 시각', '아직 없음'), ''],
    ['메일 자동수집', cfgOn('메일 자동수집 사용') ? '켜짐' : '꺼짐', ''],
    ['문자/알림톡', cfgStr('알리고 API키', '') ? (cfgOn('테스트 모드') ? '테스트 모드 (실제 발송 안 함)' : '실제 발송 켜짐') : '미설정 (설정 탭에 알리고 정보 입력 필요)', ''],
    ['마지막 오류', cfgStr('마지막 오류', '없음'), '']
  ];
  sh.getRange(r, 1, st.length, 3).setValues(st);
  sh.getRange(r, 1, st.length, 1).setFontWeight('bold');
  r += st.length + 1;

  sh.getRange(r, 1, 1, 8).merge()
    .setValue('업데이트: ' + nowStr()).setFontColor('#aaaaaa').setFontSize(9);
  sh.setHiddenGridlines(true);
}

/* =========================================================
 *  2. 예약캘린더 (월별)
 * ========================================================= */

var DOW = ['일', '월', '화', '수', '목', '금', '토'];

function slotCount() {
  var n = Number(cfg('하루 예약칸 수', 10));
  if (!n || n < 3) n = 3;
  if (n > 25) n = 25;
  return n;
}

function calSheetName(shop, date) {
  return SH.CAL + shop.short + '_' + fmt(date, 'yyyy-MM');
}

function nextMonthDate(d) { return new Date(d.getFullYear(), d.getMonth() + 1, 1); }

/* 날짜 → 캘린더 안의 위치 */
function calPos(date) {
  var first = new Date(date.getFullYear(), date.getMonth(), 1);
  var offset = first.getDay();                  // 0=일
  var idx = offset + date.getDate() - 1;
  var week = Math.floor(idx / 7);
  var dow = idx % 7;
  var N = slotCount();
  var dateRow = 3 + week * (2 + N);
  return {
    week: week, dow: dow,
    dateRow: dateRow,
    subRow: dateRow + 1,
    firstSlot: dateRow + 2,
    col: 1 + dow * 3,
    slots: N
  };
}

function weeksInMonth(date) {
  var first = new Date(date.getFullYear(), date.getMonth(), 1);
  var days = new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
  return Math.ceil((first.getDay() + days) / 7);
}

/* 캘린더 만들기 (이미 있으면 서식만 유지) */
function buildCalendarSheet(shop, date) {
  var name = calSheetName(shop, date);
  var exists = ss().getSheetByName(name);
  var sh = getOrCreateSheet(name);
  if (exists) return sh;

  var N = slotCount();
  var W = weeksInMonth(date);
  var totalRows = 2 + W * (2 + N) + 2;
  var totalCols = 21;

  if (sh.getMaxColumns() < totalCols) sh.insertColumnsAfter(sh.getMaxColumns(), totalCols - sh.getMaxColumns());
  if (sh.getMaxRows() < totalRows) sh.insertRowsAfter(sh.getMaxRows(), totalRows - sh.getMaxRows());
  sh.clear();
  sh.setHiddenGridlines(true);

  // 제목
  sh.getRange(1, 1, 1, totalCols).merge()
    .setValue(shop.name + '   ' + fmt(date, 'yyyy년 M월'))
    .setBackground(shop.idx === 1 ? '#7b3f61' : '#1f3864')
    .setFontColor('#ffffff').setFontWeight('bold').setFontSize(16)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 44);

  // 요일 헤더
  for (var d = 0; d < 7; d++) {
    var c = 1 + d * 3;
    sh.getRange(2, c, 1, 3).merge().setValue(DOW[d])
      .setHorizontalAlignment('center').setFontWeight('bold')
      .setBackground(d === 0 ? '#ffe5e5' : (d === 6 ? '#e5efff' : '#efefef'))
      .setFontColor(d === 0 ? '#c00000' : (d === 6 ? '#0044aa' : '#333333'));
  }
  sh.setRowHeight(2, 26);

  var daysInMonth = new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
  var staffList = cfgStr('담당자 명단', '').split(',').map(function (s) { return s.trim(); })
                    .filter(function (s) { return s; });
  var staffRule = staffList.length
    ? SpreadsheetApp.newDataValidation().requireValueInList(staffList, true).setAllowInvalid(true).build()
    : null;

  for (var w = 0; w < W; w++) {
    var dateRow = 3 + w * (2 + N);
    sh.setRowHeight(dateRow, 24);
    sh.setRowHeight(dateRow + 1, 20);
    for (var dd = 0; dd < 7; dd++) {
      var col = 1 + dd * 3;
      var dayNum = w * 7 + dd - new Date(date.getFullYear(), date.getMonth(), 1).getDay() + 1;
      var inMonth = dayNum >= 1 && dayNum <= daysInMonth;

      var dcell = sh.getRange(dateRow, col, 1, 3).merge();
      dcell.setValue(inMonth ? dayNum : '')
        .setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('middle')
        .setBackground(inMonth ? (dd === 0 ? '#ffeaea' : (dd === 6 ? '#eaf2ff' : '#f5f5f5')) : '#dddddd')
        .setFontColor(dd === 0 ? '#c00000' : (dd === 6 ? '#0044aa' : '#333333'));

      sh.getRange(dateRow + 1, col, 1, 3)
        .setValues([['시간', '예약자', '담당자']])
        .setFontSize(8).setFontColor('#888888').setHorizontalAlignment('center')
        .setBackground('#fafafa');

      var slotRange = sh.getRange(dateRow + 2, col, N, 3);
      slotRange.setFontSize(10).setVerticalAlignment('middle');
      if (!inMonth) slotRange.setBackground('#f0f0f0');
      if (staffRule) sh.getRange(dateRow + 2, col + 2, N, 1).setDataValidation(staffRule);
      sh.getRange(dateRow + 2, col, N, 1).setHorizontalAlignment('center').setNumberFormat('@');
    }
    // 주 구분선
    sh.getRange(dateRow, 1, 2 + N, 21).setBorder(true, true, true, true, false, null, '#bbbbbb', SpreadsheetApp.BorderStyle.SOLID);
  }
  // 열 너비
  for (var cc = 0; cc < 7; cc++) {
    sh.setColumnWidth(1 + cc * 3, 52);
    sh.setColumnWidth(2 + cc * 3, 96);
    sh.setColumnWidth(3 + cc * 3, 66);
  }
  for (var rr = 3; rr < 3 + W * (2 + N); rr++) {
    var off = (rr - 3) % (2 + N);
    if (off >= 2) sh.setRowHeight(rr, 21);
  }
  var footRow = 2 + W * (2 + N) + 1;
  sh.getRange(footRow, 1, 1, 21).merge()
    .setValue('빈 칸에 시간·예약자를 직접 적은 뒤 메뉴 → "수동예약 반영"을 누르면 고객DB와 문자 발송 대상에 들어갑니다.')
    .setFontColor('#999999').setFontSize(9);
  sh.setFrozenRows(2);
  return sh;
}

function createCalendarsForMonth(date) {
  shops().forEach(function (s) {
    buildCalendarSheet(s, date);
    refreshCalendarMonth(s, date);
  });
}

function createNextMonthCalendars() {
  try {
    var d = nextMonthDate(new Date());
    createCalendarsForMonth(d);
    var names = shops().map(function (s) { return calSheetName(s, d); }).join('\n');
    alert('다음달 캘린더를 만들었습니다.\n\n' + names);
  } catch (e) { logError('다음달 캘린더 생성', e); alert('문제가 생겼습니다: ' + e.message); }
}

function refreshThisMonthCalendars() {
  try {
    createCalendarsForMonth(new Date());
    toast('이번달 캘린더를 다시 그렸습니다.');
  } catch (e) { logError('이번달 캘린더 갱신', e); }
}

/* 한 달치 예약을 캘린더에 다시 그림 */
function refreshCalendarMonth(shop, date) {
  var sh = ss().getSheetByName(calSheetName(shop, date));
  if (!sh) return;
  var y = date.getFullYear(), m = date.getMonth();
  var rows = readData().filter(function (x) {
    var dt = toDate(x[D.DATE - 1]);
    return x[D.SHOP - 1] === shop.name && dt && dt.getFullYear() === y && dt.getMonth() === m;
  });
  var byDay = {};
  rows.forEach(function (x) {
    var dt = toDate(x[D.DATE - 1]);
    var k = dt.getDate();
    (byDay[k] = byDay[k] || []).push(x);
  });
  var daysInMonth = new Date(y, m + 1, 0).getDate();
  for (var d = 1; d <= daysInMonth; d++) {
    drawDay(sh, new Date(y, m, d), byDay[d] || []);
  }
}

/* 하루 칸 다시 그리기 */
function drawDay(sh, date, list) {
  var p = calPos(date);
  var N = p.slots;
  list = list.slice().sort(function (a, b) {
    return timeStr(a[D.TIME - 1]).localeCompare(timeStr(b[D.TIME - 1]));
  });
  var vals = [], notes = [];
  for (var i = 0; i < N; i++) {
    if (i < list.length) {
      var x = list[i];
      var mark = (x[D.STATUS - 1] === STATUS.CANCEL) ? '취소 ' : '';
      vals.push([timeStr(x[D.TIME - 1]), mark + (x[D.NAME - 1] || ''), x[D.STAFF - 1] || '']);
      notes.push(['', '예약번호:' + x[D.RESNO - 1] +
        '\n상품:' + (x[D.PRODUCT - 1] || '') +
        '\n연락처:' + (x[D.PHONE - 1] || x[D.SAFE - 1] || '없음') +
        '\n상태:' + x[D.STATUS - 1], '']);
    } else {
      vals.push(['', '', '']);
      notes.push(['', '', '']);
    }
  }
  var rng = sh.getRange(p.firstSlot, p.col, N, 3);
  rng.setValues(vals);
  rng.setNotes(notes);
  // 서식: 취소는 회색 취소선
  rng.setFontLine('none').setFontColor('#000000');
  for (var j = 0; j < list.length && j < N; j++) {
    if (list[j][D.STATUS - 1] === STATUS.CANCEL) {
      sh.getRange(p.firstSlot + j, p.col, 1, 3).setFontLine('line-through').setFontColor('#aaaaaa');
    } else if (list[j][D.STATUS - 1] === STATUS.DONE) {
      sh.getRange(p.firstSlot + j, p.col, 1, 3).setFontColor('#2e7d32');
    } else if (!list[j][D.PHONE - 1] && !list[j][D.SAFE - 1]) {
      sh.getRange(p.firstSlot + j, p.col + 1, 1, 1).setFontColor('#c26a00');
    }
  }
}

/* 예약 1건을 캘린더에 반영 (해당 날짜만 다시 그림) */
function syncCalendarForBooking(shopName, dateVal) {
  var shop = shopByName(shopName);
  var dt = toDate(dateVal);
  if (!shop || !dt) return;
  var sheetName = calSheetName(shop, dt);
  var sh = ss().getSheetByName(sheetName);
  if (!sh) { buildCalendarSheet(shop, dt); sh = ss().getSheetByName(sheetName); }
  if (!sh) return;
  var list = readData().filter(function (x) {
    return x[D.SHOP - 1] === shopName && ymd(x[D.DATE - 1]) === ymd(dt);
  });
  drawDay(sh, dt, list);
}

/* ---------- 수동예약 반영 ---------- */

function importManualBookings() {
  try {
    var added = 0;
    var months = [new Date(), nextMonthDate(new Date())];
    var dataSh = sheetOf(SH.DATA);
    shops().forEach(function (shop) {
      months.forEach(function (mdate) {
        var sh = ss().getSheetByName(calSheetName(shop, mdate));
        if (!sh) return;
        var y = mdate.getFullYear(), m = mdate.getMonth();
        var days = new Date(y, m + 1, 0).getDate();
        for (var d = 1; d <= days; d++) {
          var date = new Date(y, m, d);
          var p = calPos(date);
          var rng = sh.getRange(p.firstSlot, p.col, p.slots, 3);
          var vals = rng.getValues();
          var notes = rng.getNotes();
          for (var i = 0; i < vals.length; i++) {
            var nameCell = String(vals[i][1]).trim();
            var timeCell = String(vals[i][0]).trim();
            if (!nameCell) continue;
            if (notes[i][1] && notes[i][1].indexOf('예약번호:') === 0) continue; // 이미 등록됨
            var resno = 'M' + fmt(new Date(), 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
            var row = newDataRow();
            row[D.RESNO - 1]   = resno;
            row[D.SHOP - 1]    = shop.name;
            row[D.STATUS - 1]  = STATUS.OK;
            row[D.NAME - 1]    = nameCell;
            row[D.DATE - 1]    = date;
            row[D.TIME - 1]    = normTime(timeCell) || timeCell;
            row[D.STAFF - 1]   = String(vals[i][2]).trim();
            row[D.SOURCE - 1]  = '수동';
            row[D.PROCAT - 1]  = nowStr();
            row[D.MEMO - 1]    = '캘린더에서 직접 입력';
            dataSh.appendRow(row);
            added++;
          }
        }
      });
    });
    if (added > 0) {
      rematchCustomers();
      shops().forEach(function (s) {
        refreshCalendarMonth(s, new Date());
        refreshCalendarMonth(s, nextMonthDate(new Date()));
      });
      refreshHome();
    }
    alert(added > 0
      ? added + '건의 수동예약을 반영했습니다.\n\n연락처가 비어 있으니 [네이버예약_데이터] 또는 [고객DB] 탭에서 전화번호를 넣어주세요.\n(문자 발송에 필요합니다)'
      : '새로 반영할 수동예약이 없습니다.');
  } catch (e) { logError('수동예약 반영', e); alert('문제가 생겼습니다: ' + e.message); }
}

function newDataRow() {
  var a = [];
  for (var i = 0; i < D_HEAD.length; i++) a.push('');
  return a;
}

/* =========================================================
 *  3. 고객DB · 회차관리 · 고객프로필
 * ========================================================= */

function readCustomers() {
  var sh = sheetOf(SH.CUST);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, C_HEAD.length).getValues();
}

function newCustomerId() {
  var rows = readCustomers();
  var max = 0;
  rows.forEach(function (r) {
    var n = Number(String(r[C.ID - 1]).replace(/[^0-9]/g, ''));
    if (n > max) max = n;
  });
  return 'C' + ('0000' + (max + 1)).slice(-4);
}

/* 예약 1건 → 고객 찾기(없으면 생성). 고객ID 반환 */
function matchOrCreateCustomer(name, phone, safe, shopName, source) {
  name = String(name || '').trim();
  if (!name) return '';
  var sh = sheetOf(SH.CUST);
  var rows = readCustomers();
  var pd = onlyDigits(phone);
  var sd = onlyDigits(safe);

  var found = -1;
  if (pd) {
    for (var i = 0; i < rows.length; i++) {
      if (onlyDigits(rows[i][C.PHONE - 1]) === pd) { found = i; break; }
    }
  }
  if (found < 0 && sd) {
    for (var j = 0; j < rows.length; j++) {
      if (onlyDigits(rows[j][C.SAFE - 1]) === sd) { found = j; break; }
      if (String(rows[j][C.NAME - 1]).trim() === name && last4(rows[j][C.SAFE - 1]) === last4(sd)) { found = j; break; }
    }
  }
  if (found < 0 && !pd && !sd) {
    // 연락처가 아예 없는 경우: 같은 이름 + 연락처 미입력 고객에 붙임
    for (var k = 0; k < rows.length; k++) {
      if (String(rows[k][C.NAME - 1]).trim() === name && !onlyDigits(rows[k][C.PHONE - 1])) { found = k; break; }
    }
  }
  if (found < 0 && pd) {
    // 번호는 있는데 못 찾음 → 이름 같고 번호 비어있는 고객에 번호를 채워 넣음
    for (var m = 0; m < rows.length; m++) {
      if (String(rows[m][C.NAME - 1]).trim() === name && !onlyDigits(rows[m][C.PHONE - 1])) {
        sh.getRange(m + 2, C.PHONE).setValue(normPhone(phone));
        found = m; break;
      }
    }
  }

  if (found >= 0) {
    var row = found + 2;
    if (pd && !onlyDigits(rows[found][C.PHONE - 1])) sh.getRange(row, C.PHONE).setValue(normPhone(phone));
    if (sd && !onlyDigits(rows[found][C.SAFE - 1])) sh.getRange(row, C.SAFE).setValue(normPhone(safe));
    var curShop = String(rows[found][C.SHOP - 1] || '').trim();
    if (curShop && curShop !== shopName && curShop !== '둘다') sh.getRange(row, C.SHOP).setValue('둘다');
    else if (!curShop) sh.getRange(row, C.SHOP).setValue(shopName);
    return rows[found][C.ID - 1];
  }

  // 신규 생성
  var id = newCustomerId();
  var nr = [];
  for (var z = 0; z < C_HEAD.length; z++) nr.push('');
  nr[C.ID - 1] = id;
  nr[C.NAME - 1] = name;
  nr[C.PHONE - 1] = pd ? normPhone(phone) : '';
  nr[C.SAFE - 1] = sd ? normPhone(safe) : '';
  nr[C.SHOP - 1] = shopName;
  nr[C.SOURCE - 1] = source || '네이버자동';
  sh.appendRow(nr);
  var newRow = sh.getLastRow();
  sh.getRange(newRow, C.LEFT).setFormula('=IF($J' + newRow + '="","",$J' + newRow + '-$K' + newRow + ')');
  return id;
}

/* 전체 재매칭 + 통계 재계산 */
function rematchCustomers() {
  try {
    var dataSh = sheetOf(SH.DATA);
    var last = dataSh.getLastRow();
    if (last < 2) { toast('예약 데이터가 없습니다.'); return; }
    var rows = dataSh.getRange(2, 1, last - 1, D_HEAD.length).getValues();
    var idCol = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var id = r[D.CUSTID - 1];
      if (!id && r[D.NAME - 1]) {
        id = matchOrCreateCustomer(r[D.NAME - 1], r[D.PHONE - 1], r[D.SAFE - 1],
                                   r[D.SHOP - 1], r[D.SOURCE - 1] || '네이버자동');
      } else if (id && r[D.PHONE - 1]) {
        syncPhoneToCustomer(id, r[D.PHONE - 1]);
      }
      idCol.push([id || '']);
    }
    dataSh.getRange(2, D.CUSTID, idCol.length, 1).setValues(idCol);
    recomputeCustomerStats();
    refreshHome();
    toast('고객 매칭과 통계를 다시 계산했습니다.');
  } catch (e) { logError('고객 재매칭', e); alert('문제가 생겼습니다: ' + e.message); }
}

function syncPhoneToCustomer(custId, phone) {
  if (!phone) return;
  var sh = sheetOf(SH.CUST);
  var rows = readCustomers();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][C.ID - 1] === custId) {
      if (!onlyDigits(rows[i][C.PHONE - 1])) sh.getRange(i + 2, C.PHONE).setValue(normPhone(phone));
      return;
    }
  }
}

function recomputeCustomerStats() {
  var sh = sheetOf(SH.CUST);
  var rows = readCustomers();
  if (!rows.length) return;
  var data = readData();
  var today = toDate(ymd(new Date()));

  var agg = {};
  data.forEach(function (x) {
    var id = x[D.CUSTID - 1];
    if (!id) return;
    var st = x[D.STATUS - 1];
    var dt = toDate(x[D.DATE - 1]);
    var a = agg[id] || (agg[id] = { visits: 0, first: null, last: null, next: null, product: '', shops: {} });
    if (st === STATUS.CANCEL) return;
    a.shops[x[D.SHOP - 1]] = true;
    if (st === STATUS.DONE) {
      a.visits++;
      if (dt) {
        if (!a.first || dt < a.first) a.first = dt;
        if (!a.last || dt > a.last) a.last = dt;
      }
      if (x[D.PRODUCT - 1]) a.product = x[D.PRODUCT - 1];
    }
    if (st === STATUS.OK && dt && dt >= today) {
      if (!a.next || dt < a.next) a.next = dt;
    }
  });

  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var id = rows[i][C.ID - 1];
    var a = agg[id] || { visits: 0, first: null, last: null, next: null, product: '', shops: {} };
    var shopKeys = Object.keys(a.shops);
    var shopVal = rows[i][C.SHOP - 1];
    if (shopKeys.length > 1) shopVal = '둘다';
    else if (shopKeys.length === 1) shopVal = shopKeys[0];
    out.push([
      shopVal,
      a.first || rows[i][C.FIRST - 1] || '',
      a.last || '',
      a.visits
    ]);
  }
  sh.getRange(2, C.SHOP, out.length, 1).setValues(out.map(function (r) { return [r[0]]; }));
  sh.getRange(2, C.FIRST, out.length, 1).setValues(out.map(function (r) { return [r[1]]; }));
  sh.getRange(2, C.LAST,  out.length, 1).setValues(out.map(function (r) { return [r[2]]; }));
  sh.getRange(2, C.VISITS, out.length, 1).setValues(out.map(function (r) { return [r[3]]; }));

  // 사용회차 = 이용완료 건수, 다음예약일
  var used = [], nexts = [], prod = [];
  for (var j = 0; j < rows.length; j++) {
    var idj = rows[j][C.ID - 1];
    var aj = agg[idj] || { visits: 0, next: null, product: '' };
    used.push([aj.visits]);
    nexts.push([aj.next || '']);
    prod.push([rows[j][C.PRODUCT - 1] || aj.product || '']);
  }
  sh.getRange(2, C.USED, used.length, 1).setValues(used);
  sh.getRange(2, C.NEXT, nexts.length, 1).setValues(nexts);
  sh.getRange(2, C.PRODUCT, prod.length, 1).setValues(prod);
  // 잔여회차 수식 보정
  var f = [];
  for (var k = 0; k < rows.length; k++) {
    f.push(['=IF($J' + (k + 2) + '="","",$J' + (k + 2) + '-$K' + (k + 2) + ')']);
  }
  sh.getRange(2, C.LEFT, f.length, 1).setFormulas(f);
}

/* 지난 예약 이용완료 처리 */
function completePastBookings() {
  try {
    var sh = sheetOf(SH.DATA);
    var last = sh.getLastRow();
    if (last < 2) { alert('예약 데이터가 없습니다.'); return; }
    var rows = sh.getRange(2, 1, last - 1, D_HEAD.length).getValues();
    var today = toDate(ymd(new Date()));
    var n = 0;
    for (var i = 0; i < rows.length; i++) {
      var st = rows[i][D.STATUS - 1];
      var dt = toDate(rows[i][D.DATE - 1]);
      if (st === STATUS.OK && dt && dt < today) {
        sh.getRange(i + 2, D.STATUS).setValue(STATUS.DONE);
        n++;
      }
    }
    recomputeCustomerStats();
    refreshHome();
    alert(n + '건을 "이용완료"로 바꿨습니다.\n\n' +
      '· 회차 사용횟수가 1씩 올라갔습니다.\n' +
      '· 실제 받은 금액이 다르면 [네이버예약_데이터]의 "실결제금액"을 고쳐주세요.');
  } catch (e) { logError('이용완료 처리', e); alert('문제가 생겼습니다: ' + e.message); }
}

/* ---------- 고객프로필 화면 ---------- */

function refreshProfile() {
  try {
    var sh = sheetOf(SH.PROFILE);
    var term = String(sh.getRange('B3').getValue()).trim();
    var rows = readCustomers();
    var cands = rows.filter(function (r) {
      if (!term) return false;
      var nm = String(r[C.NAME - 1]);
      return nm.indexOf(term) >= 0 ||
             last4(r[C.PHONE - 1]) === term ||
             last4(r[C.SAFE - 1]) === term ||
             String(r[C.ID - 1]) === term;
    });
    var labels = cands.map(function (r) {
      return r[C.NAME - 1] + ' / ' + (last4(r[C.PHONE - 1]) || last4(r[C.SAFE - 1]) || '번호없음') + ' / ' + r[C.ID - 1];
    });
    // 후보 목록 Z열에 기록 + 드롭다운
    var zRows = Math.max(sh.getMaxRows() - 1, 1);
    sh.getRange(2, 26, zRows, 1).clearContent();
    if (labels.length) {
      sh.getRange(2, 26, labels.length, 1).setValues(labels.map(function (l) { return [l]; }));
      sh.getRange('B4').setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInRange(sh.getRange(2, 26, labels.length, 1), true)
          .setAllowInvalid(true).build());
      if (labels.length === 1) sh.getRange('B4').setValue(labels[0]);
    } else {
      sh.getRange('B4').clearDataValidations();
    }
    renderProfile();
  } catch (e) { logError('고객프로필 새로고침', e); }
}

function renderProfile() {
  var sh = sheetOf(SH.PROFILE);
  var pr = profileRows();
  var sel = String(sh.getRange('B4').getValue()).trim();

  // 이전 내용 지우기
  sh.getRange(pr.info, 2, 6, 1).clearContent();
  sh.getRange(pr.info, 4, 5, 1).clearContent();
  var clearRows = Math.max(sh.getMaxRows() - pr.hist + 1, 1);
  sh.getRange(pr.hist, 1, clearRows, 8).clearContent();

  if (!sel) return;
  var id = sel.split('/').pop().trim();
  var rows = readCustomers();
  var cust = null, custRow = -1;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][C.ID - 1]).trim() === id) { cust = rows[i]; custRow = i + 2; break; }
  }
  if (!cust) return;

  var left = cust[C.LEFT - 1] === '' ? '-' : cust[C.LEFT - 1];
  var total = cust[C.TOTAL - 1] === '' ? '-' : cust[C.TOTAL - 1];
  var vals1 = [
    [cust[C.NAME - 1]],
    [cust[C.SHOP - 1]],
    [cust[C.VISITS - 1]],
    [cust[C.PRODUCT - 1]],
    [cust[C.NEXT - 1] ? ymd(cust[C.NEXT - 1]) : '없음'],
    [cust[C.MEMO - 1]]
  ];
  var vals2 = [
    [cust[C.PHONE - 1] || '⚠︎ 미입력'],
    [cust[C.SAFE - 1] || '-'],
    [cust[C.LAST - 1] ? ymd(cust[C.LAST - 1]) : '-'],
    [left + ' / ' + total],
    [cust[C.FOLDER - 1] || '(메뉴 → 사진폴더 만들기)']
  ];
  sh.getRange(pr.info, 2, 6, 1).setValues(vals1);
  sh.getRange(pr.info, 4, 5, 1).setValues(vals2);

  // 방문 히스토리
  var hist = readData().filter(function (x) { return String(x[D.CUSTID - 1]).trim() === id; })
    .sort(function (a, b) { return (toDate(b[D.DATE - 1]) || 0) - (toDate(a[D.DATE - 1]) || 0); });
  if (!hist.length) return;
  var out = hist.map(function (x) {
    return [ymd(x[D.DATE - 1]) + ' ' + (x[D.TIME - 1] || ''),
            x[D.SHOP - 1], x[D.PRODUCT - 1], x[D.OPTION - 1],
            x[D.PAID - 1] || x[D.AMOUNT - 1] || '', x[D.STAFF - 1], x[D.STATUS - 1], ''];
  });
  sh.getRange(pr.hist, 1, out.length, 8).setValues(out);
  sh.getRange(pr.hist, 5, out.length, 1).setNumberFormat('#,##0');
  // 사진 폴더 링크
  var folder = cust[C.FOLDER - 1];
  if (folder) {
    sh.getRange(pr.hist, 8, out.length, 1)
      .setFormula('=HYPERLINK("' + folder + '","📷 사진 폴더 열기")');
  }
}

/* ---------- 드라이브 사진 폴더 ---------- */

function ensureDriveRoot() {
  var id = cfgStr('사진 루트폴더ID', '');
  if (id) { try { DriveApp.getFolderById(id); return id; } catch (e) {} }
  var f = DriveApp.createFolder('뷰티샵CRM_고객사진');
  setCfg('사진 루트폴더ID', f.getId());
  return f.getId();
}

function createFolderForSelectedCustomer() {
  try {
    var sh = ss().getActiveSheet();
    var id = '';
    if (sh.getName() === SH.PROFILE) {
      var sel = String(sh.getRange('B4').getValue()).trim();
      id = sel ? sel.split('/').pop().trim() : '';
    } else if (sh.getName() === SH.CUST) {
      var r = sh.getActiveRange().getRow();
      if (r > 1) id = String(sh.getRange(r, C.ID).getValue()).trim();
    }
    if (!id) { alert('[고객DB] 탭에서 고객 행을 클릭하거나, [고객프로필]에서 고객을 고른 뒤 실행해 주세요.'); return; }
    var url = ensureCustomerFolder(id);
    alert('사진 폴더를 만들었습니다.\n\n' + url +
      '\n\n구글 드라이브 앱에서 이 폴더에 before / after 사진을 올리면 됩니다.');
  } catch (e) { logError('사진폴더 생성', e); alert('문제가 생겼습니다: ' + e.message); }
}

function ensureCustomerFolder(custId) {
  var sh = sheetOf(SH.CUST);
  var rows = readCustomers();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][C.ID - 1]).trim() === custId) {
      if (rows[i][C.FOLDER - 1]) return rows[i][C.FOLDER - 1];
      var root = DriveApp.getFolderById(ensureDriveRoot());
      var fname = custId + '_' + rows[i][C.NAME - 1] + '_' + (last4(rows[i][C.PHONE - 1]) || last4(rows[i][C.SAFE - 1]) || '');
      var folder = root.createFolder(fname);
      var url = folder.getUrl();
      sh.getRange(i + 2, C.FOLDER).setValue(url);
      return url;
    }
  }
  throw new Error('고객을 찾지 못했습니다: ' + custId);
}

/* =========================================================
 *  4. 네이버 예약메일 자동 수집
 * ========================================================= */

function htmlToText(html) {
  if (!html) return '';
  var t = String(html);
  t = t.replace(/<style[\s\S]*?<\/style>/gi, ' ')
       .replace(/<script[\s\S]*?<\/script>/gi, ' ')
       .replace(/<br\s*\/?>/gi, '\n')
       .replace(/<\/(td|th|p|div|tr|li|h[1-6])>/gi, '\n')
       .replace(/<[^>]+>/g, ' ');
  t = t.replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<')
       .replace(/&gt;/gi, '>').replace(/&quot;/gi, '"').replace(/&#39;/gi, "'")
       .replace(/&[a-z]+;/gi, ' ');
  t = t.replace(/[ \t ]+/g, ' ')
       .replace(/\n\s*\n\s*\n+/g, '\n\n')
       .replace(/ *\n */g, '\n');
  return t.trim();
}

function mailText(msg) {
  var a = '';
  try { a = msg.getPlainBody() || ''; } catch (e) {}
  var b = '';
  try { b = htmlToText(msg.getBody() || ''); } catch (e) {}
  // 둘 다 붙여서 검색 (레이블-값이 어느 쪽에 붙어있든 잡히게)
  return (htmlToText(a) + '\n\n' + b).trim();
}

function pick(text, cfgKey) {
  var pat = cfgStr(cfgKey, '');
  if (!pat) return '';
  try {
    var m = new RegExp(pat).exec(text);
    return m && m[1] ? String(m[1]).trim() : '';
  } catch (e) { return ''; }
}

function testRe(text, cfgKey) {
  var pat = cfgStr(cfgKey, '');
  if (!pat) return false;
  try { return new RegExp(pat).test(text); } catch (e) { return false; }
}

/* '2026.08.27.(목) 오후 1:00' → {date, time} */
function parseKoreanDateTime(s) {
  var t = String(s || '').replace(/\s+/g, ' ').trim();
  var dm = t.match(/(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/);
  if (!dm) return null;
  var date = new Date(Number(dm[1]), Number(dm[2]) - 1, Number(dm[3]));
  var rest = t.slice(dm.index + dm[0].length);
  var time = normTime(rest) || normTime(t);
  return { date: date, time: time || '' };
}

function detectShop(text) {
  var list = shops();
  for (var i = 0; i < list.length; i++) {
    for (var k = 0; k < list[i].keywords.length; k++) {
      if (list[i].keywords[k] && text.indexOf(list[i].keywords[k]) >= 0) return list[i];
    }
  }
  return null;
}

function detectStatus(subject, body) {
  // 제목으로 판단하는 것이 가장 정확합니다. 제목이 비어있을 때만 본문 앞부분을 봅니다.
  var s = String(subject || '').trim();
  if (!s) s = String(body || '').slice(0, 300);
  if (testRe(s, '규칙_제목_취소')) return STATUS.CANCEL;
  if (testRe(s, '규칙_제목_변경')) return STATUS.CHANGE;
  if (testRe(s, '규칙_제목_확정')) return STATUS.OK;
  return STATUS.OK;
}

/* ---------- 요청사항 칸에서 실명·연락처 뽑기 ----------
   네이버는 예약자명을 '박*정'처럼 가려서 보내지만,
   고객이 직접 적은 '요청사항'은 가리지 않고 그대로 보냅니다.
   예: "박현정 01093156704"                                   */

function nameMatchesMask(real, mask) {
  real = String(real || '').trim();
  mask = String(mask || '').trim();
  if (!real || !mask) return false;
  if (real.length !== mask.length) return false;
  for (var i = 0; i < mask.length; i++) {
    var mc = mask.charAt(i);
    if (mc === '*') continue;
    if (mc !== real.charAt(i)) return false;
  }
  return true;
}

/* "성함", "연락처" 같은 안내 문구는 이름 후보에서 빼줍니다 */
function stripLabels(s) {
  return String(s || '').replace(
    /(예약자명|예약자|고객명|성함은|성함|성명|이름은|이름|연락처는|연락처|전화번호|휴대폰번호|휴대폰|핸드폰|전화|번호)/g, ' ');
}

function extractFromRequest(request, maskedName) {
  var out = { phone: '', name: '' };
  var t = String(request || '').trim();
  if (!t || t === '-') return out;

  // 휴대폰 번호 (하이픈·공백·점 허용)
  var pm = t.match(/(01[016789][\s\-.]?\d{3,4}[\s\-.]?\d{4})/);
  if (!pm) pm = t.match(/(050\d[\s\-.]?\d{3,4}[\s\-.]?\d{4})/);
  if (pm) out.phone = normPhone(pm[1]);

  var body = pm ? t.replace(pm[1], ' ') : t;
  var clean = stripLabels(body);

  // 실명: 가려진 이름(박*정)과 글자 수·보이는 글자가 맞는 한글 덩어리
  var names = (clean.match(/[가-힣]{2,5}/g) || []).concat(t.match(/[가-힣]{2,5}/g) || []);
  var ml = String(maskedName || '').trim().length;
  for (var i = 0; i < names.length; i++) {
    var c = names[i];
    if (nameMatchesMask(c, maskedName)) { out.name = c; break; }
    // "최수진이고"처럼 조사가 붙은 경우, 앞쪽 글자만 떼서 다시 맞춰봅니다
    if (ml && c.length > ml && nameMatchesMask(c.substring(0, ml), maskedName)) {
      out.name = c.substring(0, ml); break;
    }
  }
  // 가려진 이름과 못 맞췄지만 남은 한글 덩어리가 하나뿐이면 그걸 이름으로 씁니다
  if (!out.name && pm) {
    var only = clean.match(/[가-힣]{2,5}/g) || [];
    if (only.length === 1) out.name = only[0];
  }
  return out;
}

function parseBookingMail(msg) {
  var subject = '';
  try { subject = msg.getSubject() || ''; } catch (e) {}
  var body = mailText(msg);
  var all = subject + '\n' + body;

  var res = {
    ok: true, reason: '',
    subject: subject,
    mailId: msg.getId(),
    mailAt: msg.getDate(),
    shop: null,
    status: detectStatus(subject, body),
    resno: pick(body, '규칙_예약번호') || pick(all, '규칙_예약번호'),
    name: '',
    phone: '',
    safe: '',
    date: null, time: '',
    product: pick(body, '규칙_예약상품'),
    option: pick(body, '규칙_선택메뉴'),
    people: '',
    request: pick(body, '규칙_요청사항'),
    amount: ''
  };

  res.shop = detectShop(all);

  var nm = pick(body, '규칙_예약자명');
  nm = nm.replace(/님\s*$/, '').replace(/\s*\|.*$/, '').trim();
  res.name = nm;

  var ph = pick(body, '규칙_연락처');
  if (ph) {
    ph = normPhone(ph);
    if (isSafePhone(ph)) res.safe = ph; else res.phone = ph;
  }

  var dtRaw = pick(body, '규칙_이용일시');
  var dt = parseKoreanDateTime(dtRaw);
  if (dt) { res.date = dt.date; res.time = dt.time; }

  if (res.option) {
    var pm = res.option.match(/(\d+)\s*인/);
    if (pm) res.people = Number(pm[1]);
    var money = res.option.match(/([0-9][0-9,]*)\s*원\s*$/);
    if (!money) {
      var ms = res.option.match(/([0-9][0-9,]*)\s*원/g);
      if (ms && ms.length) money = ms[ms.length - 1].match(/([0-9][0-9,]*)/);
    }
    if (money) res.amount = Number(String(money[1]).replace(/,/g, ''));
  }
  if (res.request === '-' ) res.request = '';

  // 요청사항에 고객이 적어둔 실명·연락처가 있으면 끌어올립니다
  var ex = extractFromRequest(res.request, res.name);
  if (ex.phone && !res.phone && !res.safe) {
    if (isSafePhone(ex.phone)) res.safe = ex.phone; else res.phone = ex.phone;
  }
  if (ex.name) { res.maskedName = res.name; res.name = ex.name; }

  if (!res.resno) { res.ok = false; res.reason = '예약번호를 찾지 못함'; }
  else if (!res.shop) { res.ok = false; res.reason = '어느 샵인지 구분 못함'; }
  else if (!res.date && res.status !== STATUS.CANCEL) { res.ok = false; res.reason = '이용일시를 찾지 못함'; }
  else if (!res.name) { res.ok = false; res.reason = '예약자명을 찾지 못함'; }
  return res;
}

/* ---------- 메일 처리 본체 ---------- */

function checkMailNow() {
  var n = processBookingMails();
  refreshHome();
  alert(n.total === 0
    ? '새로 처리할 예약 메일이 없습니다.\n\n' +
      '· 메일이 지메일로 오고 있는지 확인해 주세요.\n' +
      '· [설정] 탭의 "메일 검색조건"을  newer_than:7d  로 늘려서 다시 시도해 보세요.'
    : '메일 ' + n.total + '통을 처리했습니다.\n\n' +
      '· 새 예약 ' + n.created + '건\n· 변경 ' + n.updated + '건\n· 취소 ' + n.canceled +
      '건\n· 확인필요 ' + n.failed + '건');
}

function processBookingMails() {
  var stat = { total: 0, created: 0, updated: 0, canceled: 0, failed: 0 };
  try { autoSettle(); } catch (e) { logError('자동 정산', e); }
  try {
    if (!cfgOn('메일 자동수집 사용')) return stat;
    var labelName = cfgStr('처리완료 라벨', 'CRM처리완료');
    var label = GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);
    var query = cfgStr('메일 검색조건', 'subject:"네이버 예약" newer_than:3d') + ' -label:' + labelName;
    var maxN = Number(cfg('한 번에 처리할 메일 수', 20)) || 20;

    var threads = GmailApp.search(query, 0, maxN);
    if (!threads.length) { setCfg('마지막 메일확인 시각', nowStr()); return stat; }

    var dataSh = sheetOf(SH.DATA);
    var seenMailIds = {};
    var idx = buildResnoIndex();
    var last = dataSh.getLastRow();
    if (last > 1) {
      dataSh.getRange(2, D.MAILID, last - 1, 1).getValues().forEach(function (r) {
        if (r[0]) seenMailIds[r[0]] = true;
      });
    }

    var touched = {};   // 'shop|yyyy-MM-dd' → true
    var newlyConfirmed = [];

    // 한 번에 여러 통이 들어올 때, 메일이 온 순서(오래된 것 → 최신)대로 처리해야
    // "확정 → 취소"가 뒤집혀서 취소가 사라지는 일이 없습니다.
    var pending = [];
    threads.forEach(function (th) {
      th.getMessages().forEach(function (msg) {
        if (!seenMailIds[msg.getId()]) pending.push(msg);
      });
    });
    pending.sort(function (a, b) { return a.getDate().getTime() - b.getDate().getTime(); });

    pending.forEach(function (msg) {
      {
        stat.total++;
        var p;
        try { p = parseBookingMail(msg); }
        catch (err) { p = { ok: false, reason: '읽기 오류: ' + err.message, mailId: msg.getId(), mailAt: msg.getDate(), subject: '' }; }

        if (!p.ok) {
          stat.failed++;
          var bad = newDataRow();
          bad[D.RESNO - 1]  = p.resno || ('X' + fmt(new Date(), 'yyMMddHHmmss'));
          bad[D.SHOP - 1]   = p.shop ? p.shop.name : '';
          bad[D.STATUS - 1] = STATUS.CHECK;
          bad[D.NAME - 1]   = p.name || '';
          bad[D.DATE - 1]   = p.date || '';
          bad[D.TIME - 1]   = p.time || '';
          bad[D.PRODUCT - 1]= p.product || '';
          bad[D.SOURCE - 1] = '네이버자동';
          bad[D.MAILAT - 1] = p.mailAt ? fmt(p.mailAt, 'yyyy-MM-dd HH:mm:ss') : '';
          bad[D.PROCAT - 1] = nowStr();
          bad[D.MAILID - 1] = p.mailId;
          bad[D.MEMO - 1]   = '자동읽기 실패(' + p.reason + ') / 제목: ' + (p.subject || '');
          dataSh.appendRow(bad);
          notifyAdmin('예약메일을 자동으로 읽지 못했습니다',
            '메일 제목: ' + (p.subject || '(없음)') + '\n이유: ' + p.reason +
            '\n\n[네이버예약_데이터] 탭 맨 아래에 "확인필요" 상태로 넣어두었습니다.\n' +
            '그 줄의 빈칸(예약자명·날짜·시간·상품)을 직접 채우고 상태를 "확정"으로 바꿔주세요.');
          return;
        }

        var key = String(p.resno);
        var row = idx[key];

        if (p.status === STATUS.CANCEL) {
          if (row) {
            dataSh.getRange(row.r, D.STATUS).setValue(STATUS.CANCEL);
            dataSh.getRange(row.r, D.MAILID).setValue(p.mailId);
            dataSh.getRange(row.r, D.PROCAT).setValue(nowStr());
            deleteGCalEvent(row.shop, dataSh.getRange(row.r, D.CALID).getValue());
            dataSh.getRange(row.r, D.CALID).setValue('');
            touched[row.shop + '|' + row.date] = true;
            stat.canceled++;
          } else {
            var cr = fillRow(newDataRow(), p);
            cr[D.STATUS - 1] = STATUS.CANCEL;
            dataSh.appendRow(cr);
            idx[key] = { r: dataSh.getLastRow(), shop: p.shop.name, date: ymd(p.date) };
            stat.canceled++;
          }
        } else if (row) {
          // 이미 취소 처리된 예약은, 그보다 먼저 온 확정 메일로 되살리지 않습니다
          var curStatus = String(dataSh.getRange(row.r, D.STATUS).getValue()).trim();
          var curStamp = parseStamp(dataSh.getRange(row.r, D.MAILAT).getValue());
          if (curStatus === STATUS.CANCEL && curStamp &&
              p.mailAt && curStamp > p.mailAt.getTime()) {
            seenMailIds[msg.getId()] = true;
            return;
          }
          // 변경 / 재확정
          var vals = fillRow(newDataRow(), p);
          var oldDate = row.date;
          [D.SHOP, D.NAME, D.DATE, D.TIME, D.PRODUCT, D.OPTION, D.PEOPLE, D.REQUEST, D.AMOUNT].forEach(function (col) {
            if (vals[col - 1] !== '' && vals[col - 1] !== null) dataSh.getRange(row.r, col).setValue(vals[col - 1]);
          });
          if (p.phone) dataSh.getRange(row.r, D.PHONE).setValue(p.phone);
          if (p.safe)  dataSh.getRange(row.r, D.SAFE).setValue(p.safe);
          dataSh.getRange(row.r, D.STATUS).setValue(STATUS.OK);
          dataSh.getRange(row.r, D.MAILID).setValue(p.mailId);
          dataSh.getRange(row.r, D.PROCAT).setValue(nowStr());
          upsertGCalEvent(dataSh, row.r);
          touched[p.shop.name + '|' + oldDate] = true;
          touched[p.shop.name + '|' + ymd(p.date)] = true;
          stat.updated++;
        } else {
          var nr = fillRow(newDataRow(), p);
          nr[D.CUSTID - 1] = matchOrCreateCustomer(p.name, p.phone, p.safe, p.shop.name, '네이버자동');
          dataSh.appendRow(nr);
          var newR = dataSh.getLastRow();
          idx[key] = { r: newR, shop: p.shop.name, date: ymd(p.date) };
          upsertGCalEvent(dataSh, newR);
          touched[p.shop.name + '|' + ymd(p.date)] = true;
          newlyConfirmed.push(newR);
          stat.created++;
        }
        seenMailIds[msg.getId()] = true;
      }
    });
    threads.forEach(function (th) { try { th.addLabel(label); } catch (e) {} });

    // 캘린더 갱신
    Object.keys(touched).forEach(function (k) {
      var parts = k.split('|');
      if (parts[1]) syncCalendarForBooking(parts[0], parts[1]);
    });

    if (stat.created || stat.updated) {
      recomputeCustomerStats();
    }
    // 예약확정 즉시 안내
    if (cfgOn('확정 즉시 안내 사용') && newlyConfirmed.length) {
      /* 밤 9시~아침 9시에 들어온 예약은 아침 9시 이후 자동 정산 때 보냅니다 (sendPendingConfirms) */
      if (!quietHour_()) newlyConfirmed.forEach(function (r) {
        try { sendForRow(dataSh, r, '확정'); } catch (e) { Logger.log(e); }
      });
    }
    setCfg('마지막 메일확인 시각', nowStr());
  } catch (e) {
    logError('예약메일 처리', e);
  }
  return stat;
}

function fillRow(row, p) {
  row[D.RESNO - 1]   = String(p.resno);
  row[D.SHOP - 1]    = p.shop ? p.shop.name : '';
  row[D.STATUS - 1]  = p.status;
  row[D.NAME - 1]    = p.name;
  row[D.PHONE - 1]   = p.phone || '';
  row[D.SAFE - 1]    = p.safe || '';
  row[D.DATE - 1]    = p.date || '';
  row[D.TIME - 1]    = p.time || '';
  row[D.PRODUCT - 1] = p.product || '';
  row[D.OPTION - 1]  = p.option || '';
  row[D.PEOPLE - 1]  = p.people || '';
  row[D.REQUEST - 1] = p.request || '';
  row[D.AMOUNT - 1]  = p.amount || '';
  row[D.PAID - 1]    = '';   // 매출은 시술 후 [결제 등록]으로 직접 넣습니다
  row[D.SOURCE - 1]  = '네이버자동';
  row[D.MAILAT - 1]  = p.mailAt ? fmt(p.mailAt, 'yyyy-MM-dd HH:mm:ss') : '';
  row[D.PROCAT - 1]  = nowStr();
  row[D.MAILID - 1]  = p.mailId || '';
  if (p.maskedName && p.maskedName !== p.name) {
    row[D.MEMO - 1] = '네이버 표기: ' + p.maskedName;
  }
  return row;
}

function buildResnoIndex() {
  var sh = sheetOf(SH.DATA);
  var last = sh.getLastRow();
  var idx = {};
  if (last < 2) return idx;
  var v = sh.getRange(2, 1, last - 1, D_HEAD.length).getValues();
  for (var i = 0; i < v.length; i++) {
    var k = String(v[i][D.RESNO - 1]).trim();
    if (k) idx[k] = { r: i + 2, shop: v[i][D.SHOP - 1], date: ymd(v[i][D.DATE - 1]) };
  }
  return idx;
}

/* =========================================================
 *  5. 구글캘린더 연동
 * ========================================================= */

function ensureGoogleCalendars() {
  shops().forEach(function (s) {
    var id = cfgStr(s.calIdKey, '');
    if (id) { try { if (CalendarApp.getCalendarById(id)) return; } catch (e) {} }
    var found = CalendarApp.getCalendarsByName(s.name + ' 예약');
    var cal = (found && found.length) ? found[0] : CalendarApp.createCalendar(s.name + ' 예약', {
      summary: s.name + ' 네이버 예약 자동연동',
      color: s.idx === 1 ? CalendarApp.Color.PINK : CalendarApp.Color.BLUE
    });
    setCfg(s.calIdKey, cal.getId());
  });
}

function calFor(shopName) {
  var s = shopByName(shopName);
  if (!s) return null;
  var id = cfgStr(s.calIdKey, '');
  if (!id) return null;
  try { return CalendarApp.getCalendarById(id); } catch (e) { return null; }
}

function upsertGCalEvent(dataSh, r) {
  try {
    var v = dataSh.getRange(r, 1, 1, D_HEAD.length).getValues()[0];
    if (v[D.STATUS - 1] === STATUS.CANCEL) return;
    var cal = calFor(v[D.SHOP - 1]);
    if (!cal) return;
    var d = toDate(v[D.DATE - 1]);
    var t = normTime(v[D.TIME - 1]) || '10:00';
    if (!d) return;
    var hhmm = t.split(':');
    var start = new Date(d.getFullYear(), d.getMonth(), d.getDate(), Number(hhmm[0]), Number(hhmm[1]));
    var mins = Number(cfg('예약 기본 소요시간(분)', 90)) || 90;
    var end = new Date(start.getTime() + mins * 60000);
    var title = [v[D.NAME - 1], v[D.PRODUCT - 1], v[D.OPTION - 1]].filter(function (x) { return x; }).join(' / ');
    var desc = '예약번호: ' + v[D.RESNO - 1] +
      '\n연락처: ' + (v[D.PHONE - 1] || v[D.SAFE - 1] || '미확인') +
      '\n인원: ' + (v[D.PEOPLE - 1] || '') +
      '\n요청사항: ' + (v[D.REQUEST - 1] || '') +
      '\n담당자: ' + (v[D.STAFF - 1] || '');
    var evId = v[D.CALID - 1];
    var ev = null;
    if (evId) { try { ev = cal.getEventById(evId); } catch (e) { ev = null; } }
    if (ev) {
      ev.setTitle(title);
      ev.setTime(start, end);
      ev.setDescription(desc);
    } else {
      ev = cal.createEvent(title, start, end, { description: desc });
      dataSh.getRange(r, D.CALID).setValue(ev.getId());
    }
  } catch (e) { Logger.log('구글캘린더 등록 실패: ' + e); }
}

function deleteGCalEvent(shopName, eventId) {
  if (!eventId) return;
  try {
    var cal = calFor(shopName);
    if (!cal) return;
    var ev = cal.getEventById(eventId);
    if (ev) ev.deleteEvent();
  } catch (e) { Logger.log('구글캘린더 삭제 실패: ' + e); }
}

/* =========================================================
 *  6. 알림톡 / 문자 발송 (알리고)
 * ========================================================= */

var ALIGO_SMS   = 'https://apis.aligo.in/send/';
var ALIGO_TOKEN = 'https://kakaoapi.aligo.in/akv10/token/create/30/s/';
var ALIGO_ATALK = 'https://kakaoapi.aligo.in/akv10/alimtalk/send/';

function aligoReady() {
  return !!(cfgStr('알리고 사용자ID', '') && cfgStr('알리고 API키', '') && cfgStr('발신번호', ''));
}

/* 문자에 쓸 호칭 정하기: 고객DB의 실명 > 마스킹 안 된 이름 > '고객' */
function displayName(rawName, custId) {
  var nm = String(rawName || '').trim();
  if (custId) {
    try {
      var rows = readCustomers();
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][C.ID - 1]).trim() === String(custId).trim()) {
          var cn = String(rows[i][C.NAME - 1] || '').trim();
          if (cn && cn.indexOf('*') < 0) return cn;
          break;
        }
      }
    } catch (e) {}
  }
  if (!nm || nm.indexOf('*') >= 0) return '고객';
  return nm;
}

function shopNo(name) {
  var list = shops();
  for (var i = 0; i < list.length; i++) if (list[i].name === name) return list[i].idx;
  return 1;
}

/* 90 -> "1시간 30분" */
function minText(m) {
  m = Number(m) || 0;
  var h = Math.floor(m / 60), r = m % 60;
  var s = (h ? h + '시간' : '') + (h && r ? ' ' : '') + (r ? r + '분' : '');
  return s || '0분';
}

/* 예약의 선택메뉴·예약상품을 시술메뉴와 맞춰 실제 소요시간을 찾습니다 */
function bookingMinutes(v) {
  var nosp = function (x) { return String(x || '').replace(/\s+/g, ''); };
  var hay = nosp(String(v[D.OPTION - 1] || '') + '|' + String(v[D.PRODUCT - 1] || ''));
  var shop = String(v[D.SHOP - 1] || '');
  var best = 0, bestLen = 0;
  try {
    readMenu().forEach(function (m) {
      if (m.shop && m.shop !== shop) return;
      if (!m.min || !m.name) return;
      var key = nosp(m.name);
      if (key && hay.indexOf(key) >= 0 && key.length > bestLen) {
        bestLen = key.length; best = m.min;
      }
    });
  } catch (e) {}
  if (best) return best;
  return Number(String(cfgStr('예약 기본 소요시간(분)', '90')).replace(/[^0-9]/g, '')) || 90;
}

/* 2026-08-31 -> "2026년 8월 31일(월)" */
function dayText(dateVal) {
  var dt = toDate(dateVal);
  if (!dt) return ymd(dateVal);
  return fmt(dt, 'yyyy년 M월 d일') + '(' + ['일', '월', '화', '수', '목', '금', '토'][dt.getDay()] + ')';
}

/* 13:00 -> "오후 1시",  13:30 -> "오후 1시 30분" */
function clockText(timeVal) {
  var m = String(timeVal || '').trim().match(/^(\d{1,2}):(\d{2})/);
  if (!m) return String(timeVal || '');
  var H = Number(m[1]), M = Number(m[2]);
  var ap = H < 12 ? '오전' : '오후';
  var h12 = H % 12; if (h12 === 0) h12 = 12;
  return ap + ' ' + h12 + '시' + (M ? ' ' + M + '분' : '');
}

/* 2026-08-31 + 13:00 -> "2026년 8월 31일(월) 오후 1:00" */
function whenText(dateVal, timeVal) {
  var dt = toDate(dateVal);
  var t = String(timeVal || '').trim();
  if (!dt) return (ymd(dateVal) + ' ' + t).trim();
  var s = fmt(dt, 'yyyy년 M월 d일') + '(' + ['일', '월', '화', '수', '목', '금', '토'][dt.getDay()] + ')';
  var m = t.match(/^(\d{1,2}):(\d{2})/);
  if (m) {
    var H = Number(m[1]);
    var ap = H < 12 ? '오전' : '오후';
    var h12 = H % 12; if (h12 === 0) h12 = 12;
    s += ' ' + ap + ' ' + h12 + ':' + m[2];
  }
  return s;
}

function buildMessage(v, type) {
  var base = { '사전': '문구_사전', '전일': '문구_전일', '당일': '문구_당일', '확정': '문구_확정', '개별': '문구_확정' }[type] || '문구_확정';
  var shop = v[D.SHOP - 1] || '';
  var tpl = String(cfgStr('샵' + shopNo(shop) + ' ' + base, '')).trim();   // 샵 전용 문구가 있으면 그것부터
  if (!tpl) tpl = String(cfgStr(base, '')).trim();
  if (!tpl) tpl = '[#{샵}] #{이름}님, #{날짜} #{시간} 예약 안내드립니다.';
  var name = v[D.NAME - 1] || '고객';
  var time = timeStr(v[D.TIME - 1]);
  return tpl
    .replace(/#\{샵\}/g, shop)
    .replace(/#\{상점명\}/g, shop)
    .replace(/#\{이름\}/g, name)
    .replace(/#\{고객명\}/g, name)
    .replace(/#\{날짜\}/g, ymd(v[D.DATE - 1]))
    .replace(/#\{시간\}/g, time)
    .replace(/#\{예약일시\}/g, whenText(v[D.DATE - 1], time))
    .replace(/#\{날짜요일\}/g, dayText(v[D.DATE - 1]))
    .replace(/#\{시각\}/g, clockText(time))
    .replace(/#\{소요시간\}/g, minText(bookingMinutes(v)))
    .replace(/#\{상품\}/g, v[D.PRODUCT - 1] || '')
    .replace(/#\{옵션\}/g, v[D.OPTION - 1] || '');
}

function alreadySent(resno, type) {
  var sh = sheetOf(SH.LOG);
  var last = sh.getLastRow();
  if (last < 2) return false;
  var v = sh.getRange(2, 2, last - 1, 7).getValues(); // B~H: 예약번호,샵,고객명,연락처,유형,채널,결과
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]) === String(resno) && String(v[i][4]) === type &&
        String(v[i][6]).indexOf('성공') === 0) {
      return true;
    }
  }
  return false;
}

function logSend(v, type, channel, result, message, resp) {
  sheetOf(SH.LOG).appendRow([
    nowStr(), v[D.RESNO - 1], v[D.SHOP - 1], v[D.NAME - 1],
    v[D.PHONE - 1] || v[D.SAFE - 1] || '', type, channel, result,
    message, typeof resp === 'string' ? resp.slice(0, 300) : JSON.stringify(resp).slice(0, 300)
  ]);
}

/* 예약 1행 발송 */
function sendForRow(dataSh, r, type, force) {
  var v = dataSh.getRange(r, 1, 1, D_HEAD.length).getValues()[0];
  if (v[D.STATUS - 1] === STATUS.CANCEL) return { ok: false, reason: '취소된 예약' };
  var shop = shopByName(v[D.SHOP - 1]);
  if (shop && !cfgOn(shop.sendKey)) return { ok: false, reason: '이 샵은 발송 꺼짐' };
  if (type !== '개별' && autoOffShop(v[D.SHOP - 1])) return { ok: false, reason: '이 샵 자동 안내는 팅커뷰에서 보냄' };
  var to = v[D.PHONE - 1] || v[D.SAFE - 1];
  if (!to) {
    logSend(v, type, '-', '실패', '', '연락처 없음');
    return { ok: false, reason: '연락처 없음' };
  }
  if (!force && alreadySent(v[D.RESNO - 1], type)) return { ok: false, reason: '이미 발송함' };

  // 네이버 메일의 이름은 '박*정'처럼 가려져 있습니다.
  // 고객DB에 실명을 적어두었으면 그 이름을, 아니면 '고객'으로 부릅니다.
  v[D.NAME - 1] = displayName(v[D.NAME - 1], v[D.CUSTID - 1]);

  var msg = buildMessage(v, type);

  if (cfgOn('테스트 모드')) {
    logSend(v, type, '테스트', '성공(테스트)', msg, '테스트 모드라 실제로 보내지 않았습니다');
    return { ok: true, test: true };
  }
  if (!aligoReady()) {
    logSend(v, type, '-', '실패', msg, '알리고 설정이 비어있습니다');
    return { ok: false, reason: '알리고 미설정' };
  }

  var tplCode = cfgStr('알림톡 템플릿코드_' + type, '');
  var senderKey = cfgStr('카카오채널 senderkey', '');
  var res;
  if (tplCode && senderKey) {
    res = sendAlimtalk(to, v[D.NAME - 1], msg, tplCode, senderKey);
    if (res.ok) { logSend(v, type, '알림톡', '성공', msg, res.raw); return { ok: true }; }
    // 실패 → 문자 대체
    var sres = sendSms(to, msg);
    logSend(v, type, sres.ok ? 'SMS(대체)' : 'SMS(대체)', sres.ok ? '성공' : '실패', msg,
      '알림톡실패:' + res.raw + ' / SMS:' + sres.raw);
    return { ok: sres.ok, reason: sres.ok ? '' : 'SMS도 실패' };
  }
  res = sendSms(to, msg);
  logSend(v, type, 'SMS', res.ok ? '성공' : '실패', msg, res.raw);
  return { ok: res.ok, reason: res.ok ? '' : res.raw };
}

function sendSms(to, msg) {
  try {
    var isLms = msg.length > 45;
    var payload = {
      key: cfgStr('알리고 API키', ''),
      user_id: cfgStr('알리고 사용자ID', ''),
      sender: cfgStr('발신번호', ''),
      receiver: onlyDigits(to),
      msg: msg,
      msg_type: isLms ? 'LMS' : 'SMS'
    };
    if (isLms) payload.title = '예약 안내';
    var r = UrlFetchApp.fetch(ALIGO_SMS, {
      method: 'post', payload: payload, muteHttpExceptions: true
    });
    var text = r.getContentText();
    var j = {};
    try { j = JSON.parse(text); } catch (e) {}
    return { ok: String(j.result_code) === '1', raw: text };
  } catch (e) {
    return { ok: false, raw: '통신오류: ' + e.message };
  }
}

function aligoToken() {
  try {
    var r = UrlFetchApp.fetch(ALIGO_TOKEN, {
      method: 'post',
      payload: { apikey: cfgStr('알리고 API키', ''), userid: cfgStr('알리고 사용자ID', '') },
      muteHttpExceptions: true
    });
    var j = JSON.parse(r.getContentText());
    return (j && j.token) ? j.token : '';
  } catch (e) { return ''; }
}

function sendAlimtalk(to, name, msg, tplCode, senderKey) {
  try {
    var token = aligoToken();
    if (!token) return { ok: false, raw: '알림톡 토큰 발급 실패' };
    var payload = {
      apikey: cfgStr('알리고 API키', ''),
      userid: cfgStr('알리고 사용자ID', ''),
      token: token,
      senderkey: senderKey,
      tpl_code: tplCode,
      sender: onlyDigits(cfgStr('발신번호', '')),
      receiver_1: onlyDigits(to),
      recvname_1: name || '고객',
      subject_1: '예약 안내',
      message_1: msg,
      failover: 'Y',
      fsubject_1: '예약 안내',
      fmessage_1: msg
    };
    var r = UrlFetchApp.fetch(ALIGO_ATALK, { method: 'post', payload: payload, muteHttpExceptions: true });
    var text = r.getContentText();
    var j = {};
    try { j = JSON.parse(text); } catch (e) {}
    return { ok: String(j.code) === '0', raw: text };
  } catch (e) {
    return { ok: false, raw: '통신오류: ' + e.message };
  }
}

/* ---------- 일괄 발송 ---------- */

function sendReminderForDate(dateObj, type) {
  var sh = sheetOf(SH.DATA);
  var last = sh.getLastRow();
  if (last < 2) return { sent: 0, skip: 0 };
  var target = ymd(dateObj);
  var v = sh.getRange(2, 1, last - 1, D_HEAD.length).getValues();
  var sent = 0, skip = 0;
  for (var i = 0; i < v.length; i++) {
    if (ymd(v[i][D.DATE - 1]) !== target) continue;
    if (v[i][D.STATUS - 1] !== STATUS.OK) continue;
    var res = sendForRow(sh, i + 2, type);
    if (res.ok) sent++; else skip++;
  }
  return { sent: sent, skip: skip };
}

function advanceDays() {
  var n = Math.round(Number(String(cfg('사전 알림 며칠 전', '3')).replace(/[^0-9]/g, '')) || 3);
  return Math.min(30, Math.max(1, n));
}

function sendReminderAdvance() {
  try {
    if (!cfgOn('사전 알림 사용')) return;
    var d = new Date(Date.now() + advanceDays() * 86400000);
    sendReminderForDate(d, '사전');
  } catch (e) { logError('사전 알림 발송', e); }
}

function sendReminderAdvanceManual() {
  var n = advanceDays();
  var r = sendReminderForDate(new Date(Date.now() + n * 86400000), '사전');
  alert(n + '일 뒤 예약자 발송 결과\n\n보냄: ' + r.sent + '건\n건너뜀: ' + r.skip + '건' +
    (cfgOn('테스트 모드') ? '\n\n※ 지금은 테스트 모드라 실제로 나가지 않았습니다.' : ''));
}

function sendReminderTomorrow() {
  try {
    if (!cfgOn('전일 알림 사용')) return;
    var d = new Date(Date.now() + 86400000);
    sendReminderForDate(d, '전일');
  } catch (e) { logError('전일 알림 발송', e); }
}

function sendReminderToday() {
  try {
    if (!cfgOn('당일 알림 사용')) return;
    sendReminderForDate(new Date(), '당일');
  } catch (e) { logError('당일 알림 발송', e); }
}

function sendReminderTomorrowManual() {
  var r = sendReminderForDate(new Date(Date.now() + 86400000), '전일');
  alert('내일 예약자 발송 결과\n\n보냄: ' + r.sent + '건\n건너뜀: ' + r.skip +
    '건 (연락처 없음·이미 발송·취소 등)\n\n자세한 내용은 [발송로그] 탭에서 볼 수 있습니다.' +
    (cfgOn('테스트 모드') ? '\n\n※ 지금은 테스트 모드라 실제로 나가지 않았습니다.' : ''));
}

function sendReminderTodayManual() {
  var r = sendReminderForDate(new Date(), '당일');
  alert('오늘 예약자 발송 결과\n\n보냄: ' + r.sent + '건\n건너뜀: ' + r.skip + '건' +
    (cfgOn('테스트 모드') ? '\n\n※ 지금은 테스트 모드라 실제로 나가지 않았습니다.' : ''));
}

/* 선택한 예약에 발송 */
function sendToSelected() {
  try {
    var sh = ss().getActiveSheet();
    var name = sh.getName();
    var resnos = [];
    if (name === SH.DATA) {
      var rl = sh.getActiveRangeList() ? sh.getActiveRangeList().getRanges() : [sh.getActiveRange()];
      rl.forEach(function (rg) {
        for (var r = rg.getRow(); r < rg.getRow() + rg.getNumRows(); r++) {
          if (r > 1) resnos.push(String(sh.getRange(r, D.RESNO).getValue()).trim());
        }
      });
    } else if (name.indexOf(SH.CAL) === 0) {
      var rg2 = sh.getActiveRange();
      for (var rr = rg2.getRow(); rr < rg2.getRow() + rg2.getNumRows(); rr++) {
        for (var cc = rg2.getColumn(); cc < rg2.getColumn() + rg2.getNumColumns(); cc++) {
          var base = cc - ((cc - 1) % 3); // 그 날짜칸의 시간 열
          var note = sh.getRange(rr, base + 1).getNote();
          var m = note && note.match(/예약번호:(\S+)/);
          if (m) resnos.push(m[1]);
        }
      }
    } else {
      alert('[네이버예약_데이터] 탭이나 [예약캘린더] 탭에서 보낼 예약을 선택한 뒤 실행해 주세요.');
      return;
    }
    resnos = resnos.filter(function (x, i, a) { return x && a.indexOf(x) === i; });
    if (!resnos.length) { alert('선택된 예약이 없습니다.'); return; }

    var ui = SpreadsheetApp.getUi();
    var conf = ui.alert('선택 발송', resnos.length + '건에 안내 메시지를 보냅니다.\n계속할까요?' +
      (cfgOn('테스트 모드') ? '\n\n※ 지금은 테스트 모드입니다 (실제 발송 안 됨).' : ''), ui.ButtonSet.YES_NO);
    if (conf !== ui.Button.YES) return;

    var dataSh = sheetOf(SH.DATA);
    var idx = buildResnoIndex();
    var ok = 0, ng = 0, reasons = [];
    resnos.forEach(function (no) {
      var row = idx[no];
      if (!row) { ng++; return; }
      var res = sendForRow(dataSh, row.r, '개별', true);
      if (res.ok) ok++; else { ng++; if (res.reason) reasons.push(no + ': ' + res.reason); }
    });
    alert('보냄 ' + ok + '건 / 실패 ' + ng + '건' +
      (reasons.length ? '\n\n' + reasons.slice(0, 8).join('\n') : ''));
  } catch (e) { logError('선택 발송', e); alert('문제가 생겼습니다: ' + e.message); }
}

function sendTestMessage() {
  var to = cfgStr('테스트 수신번호', '');
  if (!to) { alert('[설정] 탭의 "테스트 수신번호"에 본인 휴대폰번호를 먼저 넣어주세요.'); return; }
  if (cfgOn('테스트 모드')) {
    alert('지금은 "테스트 모드"라 실제로 보내지 않습니다.\n\n실제로 보내려면 [설정] 탭의 "테스트 모드"를 "미사용"으로 바꿔주세요.');
    return;
  }
  if (!aligoReady()) { alert('[설정] 탭에 알리고 사용자ID · API키 · 발신번호를 먼저 넣어주세요.'); return; }
  var r = sendSms(to, '[샵 CRM] 발송 테스트입니다. 이 문자가 보이면 설정이 정상입니다.');
  alert(r.ok ? '보냈습니다! 휴대폰을 확인해 주세요.' : '실패했습니다.\n\n' + r.raw);
}

/* =========================================================
 *  7. 자동실행(트리거) · 편집 감지 · 보호 · 진단
 * ========================================================= */

var TRIGGER_FNS = ['processBookingMails', 'sendReminderAdvance', 'sendReminderTomorrow', 'sendReminderToday',
                   'dailyMaintenance', 'onEditInstallable'];

function installTriggers() {
  try {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (TRIGGER_FNS.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
    });

    ScriptApp.newTrigger('processBookingMails').timeBased().everyMinutes(5).create();

    var t0 = parseHm(cfgStr('사전 알림 시각', '14:00'));
    ScriptApp.newTrigger('sendReminderAdvance').timeBased()
      .atHour(t0.h).nearMinute(t0.m).everyDays(1).create();

    var t1 = parseHm(cfgStr('전일 알림 시각', '18:00'));
    ScriptApp.newTrigger('sendReminderTomorrow').timeBased()
      .atHour(t1.h).nearMinute(t1.m).everyDays(1).create();

    var t2 = parseHm(cfgStr('당일 알림 시각', '09:00'));
    ScriptApp.newTrigger('sendReminderToday').timeBased()
      .atHour(t2.h).nearMinute(t2.m).everyDays(1).create();

    ScriptApp.newTrigger('dailyMaintenance').timeBased().atHour(1).nearMinute(15).everyDays(1).create();

    ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(ss()).onEdit().create();

    toast('자동실행 시간을 다시 맞췄습니다.');
  } catch (e) { logError('자동실행 설정', e); }
}

function parseHm(s) {
  var m = String(s).match(/(\d{1,2})\s*[:시]\s*(\d{1,2})?/);
  if (!m) return { h: 9, m: 0 };
  var h = Math.min(23, Math.max(0, Number(m[1])));
  var mi = Math.min(59, Math.max(0, Number(m[2] || 0)));
  // nearMinute 는 15분 단위 권장
  mi = Math.round(mi / 15) * 15;
  if (mi > 45) mi = 45;
  return { h: h, m: mi };
}

function dailyMaintenance() {
  try {
    recomputeCustomerStats();
    // 이번달/다음달 캘린더 확보
    createCalendarsForMonth(new Date());
    var nm = nextMonthDate(new Date());
    shops().forEach(function (s) {
      if (!ss().getSheetByName(calSheetName(s, nm))) buildCalendarSheet(s, nm);
    });
    createCalendarsForMonth(nm);
    refreshHome();
  } catch (e) { logError('일일 정리작업', e); }
}

/* ---------- 편집 감지 ---------- */

function onEditInstallable(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    var name = sh.getName();
    var row = e.range.getRow(), col = e.range.getColumn();

    // 1) 캘린더에서 담당자 입력 → 데이터 시트에 반영
    if (name.indexOf(SH.CAL) === 0) {
      var offset = (col - 1) % 3;         // 0=시간 1=예약자 2=담당자
      var baseCol = col - offset;
      var note = sh.getRange(row, baseCol + 1).getNote();
      var m = note && note.match(/예약번호:(\S+)/);
      if (m && offset === 2) {
        var idx = buildResnoIndex();
        var target = idx[m[1]];
        if (target) {
          var dsh = sheetOf(SH.DATA);
          dsh.getRange(target.r, D.STAFF).setValue(e.range.getValue());
        }
      }
      return;
    }

    // 2) 고객프로필 검색
    if (name === SH.PROFILE) {
      if (row === 3 && col === 2) { refreshProfile(); }
      else if (row === 4 && col === 2) { renderProfile(); }
      return;
    }

    // 3) 데이터 시트 직접 수정
    if (name === SH.DATA && row > 1) {
      if (col === D.PHONE) {
        var v = sh.getRange(row, 1, 1, D_HEAD.length).getValues()[0];
        var cid = v[D.CUSTID - 1];
        if (!cid) {
          cid = matchOrCreateCustomer(v[D.NAME - 1], v[D.PHONE - 1], v[D.SAFE - 1], v[D.SHOP - 1], v[D.SOURCE - 1] || '수동');
          sh.getRange(row, D.CUSTID).setValue(cid);
        } else {
          syncPhoneToCustomer(cid, v[D.PHONE - 1]);
        }
        syncCalendarForBooking(v[D.SHOP - 1], v[D.DATE - 1]);
      } else if (col === D.STATUS || col === D.DATE || col === D.TIME || col === D.NAME) {
        var v2 = sh.getRange(row, 1, 1, D_HEAD.length).getValues()[0];
        if (v2[D.STATUS - 1] === STATUS.CANCEL) {
          deleteGCalEvent(v2[D.SHOP - 1], v2[D.CALID - 1]);
          sh.getRange(row, D.CALID).setValue('');
        } else {
          upsertGCalEvent(sh, row);
        }
        syncCalendarForBooking(v2[D.SHOP - 1], v2[D.DATE - 1]);
      }
      return;
    }

    // 4) 설정 변경 시 캐시 비우기
    if (name === SH.CONFIG) { _cfgCache = null; }
  } catch (err) {
    Logger.log('onEdit 오류: ' + err);
  }
}

/* ---------- 시트 보호 ---------- */

function protectSheets() {
  try {
    var me = Session.getEffectiveUser();
    [SH.DATA, SH.CUST, SH.LOG].forEach(function (n) {
      var sh = ss().getSheetByName(n);
      if (!sh) return;
      sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
        if (p.getDescription() === '자동기록(경고)') p.remove();
      });
      var p = sh.getRange(1, 1, 1, sh.getMaxColumns()).protect()
        .setDescription('자동기록(경고)').setWarningOnly(true);
    });
    var log = ss().getSheetByName(SH.LOG);
    if (log) {
      log.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
      log.protect().setDescription('발송로그(경고)').setWarningOnly(true);
    }
  } catch (e) { Logger.log('보호 설정 실패: ' + e); }
}

/* ---------- 매출통계 ---------- */

function refreshAllSales() {
  try {
    shops().forEach(function (s) {
      if (!ss().getSheetByName(SH.SALES + s.short)) buildSalesSheet(s);
    });
    SpreadsheetApp.flush();
    toast('매출통계를 새로고침했습니다. (수식이 자동 계산됩니다)');
  } catch (e) { logError('매출통계 새로고침', e); }
}

/* ---------- 진단 ---------- */

function runDiagnostics() {
  var lines = [];
  var ok = function (b) { return b ? '✅' : '❌'; };
  try {
    lines.push('■ 탭 확인');
    [SH.HOME, SH.DATA, SH.CUST, SH.PROFILE, SH.LOG, SH.CONFIG].forEach(function (n) {
      lines.push('  ' + ok(!!ss().getSheetByName(n)) + ' ' + n);
    });
    shops().forEach(function (s) {
      lines.push('  ' + ok(!!ss().getSheetByName(calSheetName(s, new Date()))) + ' ' + calSheetName(s, new Date()));
      lines.push('  ' + ok(!!ss().getSheetByName(SH.SALES + s.short)) + ' ' + SH.SALES + s.short);
    });

    lines.push('');
    lines.push('■ 자동실행');
    var fns = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
    ['processBookingMails', 'sendReminderAdvance', 'sendReminderTomorrow', 'sendReminderToday', 'onEditInstallable'].forEach(function (f) {
      lines.push('  ' + ok(fns.indexOf(f) >= 0) + ' ' + f);
    });

    lines.push('');
    lines.push('■ 연동');
    lines.push('  ' + ok(cfgOn('메일 자동수집 사용')) + ' 메일 자동수집 켜짐');
    var q = cfgStr('메일 검색조건', '');
    var cnt = 0;
    try { cnt = GmailApp.search(q, 0, 5).length; } catch (e) {}
    lines.push('  ' + ok(cnt > 0) + ' 검색조건으로 찾은 메일: ' + cnt + '통  (조건: ' + q + ')');
    shops().forEach(function (s) {
      lines.push('  ' + ok(!!cfgStr(s.calIdKey, '')) + ' ' + s.name + ' 구글캘린더');
    });
    lines.push('  ' + ok(!!cfgStr('사진 루트폴더ID', '')) + ' 사진 저장 폴더');
    lines.push('  ' + ok(aligoReady()) + ' 알리고 문자 설정' + (cfgOn('테스트 모드') ? ' (테스트 모드)' : ''));

    lines.push('');
    lines.push('■ 데이터');
    var rows = readData();
    lines.push('  예약 ' + rows.length + '건 / 고객 ' + readCustomers().length + '명');
    lines.push('  확인필요 ' + rows.filter(function (x) { return x[D.STATUS - 1] === STATUS.CHECK; }).length + '건');
    lines.push('  연락처 없는 예약 ' + rows.filter(function (x) {
      return !x[D.PHONE - 1] && !x[D.SAFE - 1] && x[D.STATUS - 1] !== STATUS.CANCEL; }).length + '건');
    var err = cfgStr('마지막 오류', '');
    lines.push('');
    lines.push('■ 마지막 오류: ' + (err || '없음'));
  } catch (e) {
    lines.push('진단 중 오류: ' + e.message);
  }
  alert(lines.join('\n'));
}

/* ---------- 메일 파싱 테스트 (개발/점검용) ---------- */

function testParseLatestMail() {
  var q = cfgStr('메일 검색조건', 'subject:"네이버 예약" newer_than:30d');
  var th = GmailApp.search(q, 0, 1);
  if (!th.length) { alert('검색조건에 맞는 메일이 없습니다.\n조건: ' + q); return; }
  var msgs = th[0].getMessages();
  var msg = msgs[msgs.length - 1];
  var p = parseBookingMail(msg);
  alert(
    '제목: ' + p.subject + '\n\n' +
    '읽기 성공: ' + (p.ok ? '예' : '아니오 (' + p.reason + ')') + '\n' +
    '샵: ' + (p.shop ? p.shop.name : '(못 찾음)') + '\n' +
    '상태: ' + p.status + '\n' +
    '예약번호: ' + p.resno + '\n' +
    '예약자명: ' + p.name + '\n' +
    '연락처: ' + (p.phone || p.safe || '(메일에 없음)') + '\n' +
    '이용일: ' + (p.date ? ymd(p.date) : '') + ' ' + p.time + '\n' +
    '예약상품: ' + p.product + '\n' +
    '선택메뉴: ' + p.option + '\n' +
    '인원: ' + p.people + '\n' +
    '금액: ' + p.amount + '\n' +
    '요청사항: ' + p.request);
}

function showMailBodyForDebug() {
  var q = cfgStr('메일 검색조건', 'subject:"네이버 예약" newer_than:30d');
  var th = GmailApp.search(q, 0, 1);
  if (!th.length) { alert('메일이 없습니다.'); return; }
  var msgs = th[0].getMessages();
  var t = mailText(msgs[msgs.length - 1]);
  alert(t.slice(0, 1200));
}

/* =========================================================
 *  8. 웹앱 (홈페이지·휴대폰 앱 화면)
 *     배포 → 새 배포 → 웹 앱 으로 게시하면 주소가 나옵니다.
 * ========================================================= */

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || 'app';
  var out = HtmlService.createTemplateFromFile('WebApp').evaluate();
  out = out
    .setTitle('뷰티샵 예약 관리')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  /* 홈 화면에 추가했을 때 주소창 없이 앱처럼 열리게 합니다 */
  try { out.addMetaTag('mobile-web-app-capable', 'yes'); } catch (e) {}
  try { out.addMetaTag('apple-mobile-web-app-capable', 'yes'); } catch (e) {}
  return out;
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* 웹앱 주소 보여주기 */
function showWebAppUrl() {
  var url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) {}
  if (!url) {
    alert('아직 웹앱으로 배포되지 않았습니다.\n\n' +
      'Apps Script 편집기에서 오른쪽 위 [배포] → [새 배포] →\n' +
      '유형 선택(⚙️) → [웹 앱] → 액세스 권한 "나만" → [배포]\n' +
      '를 누르면 주소가 나옵니다.');
    return;
  }
  alert('앱 주소입니다. 휴대폰 브라우저에서 열고\n"홈 화면에 추가"를 누르면 앱처럼 쓸 수 있습니다.\n\n' + url);
}

/* ---------- 웹앱용 공통 변환 ---------- */

function rowToObj(v) {
  return {
    resno:   String(v[D.RESNO - 1] || ''),
    shop:    String(v[D.SHOP - 1] || ''),
    status:  String(v[D.STATUS - 1] || ''),
    name:    String(v[D.NAME - 1] || ''),
    phone:   String(v[D.PHONE - 1] || ''),
    safe:    String(v[D.SAFE - 1] || ''),
    date:    ymd(v[D.DATE - 1]),
    time:    timeStr(v[D.TIME - 1]),
    product: String(v[D.PRODUCT - 1] || ''),
    option:  String(v[D.OPTION - 1] || ''),
    people:  String(v[D.PEOPLE - 1] || ''),
    request: String(v[D.REQUEST - 1] || ''),
    amount:  Number(v[D.AMOUNT - 1] || 0),
    paid:    Number(v[D.PAID - 1] || 0),
    staff:   String(v[D.STAFF - 1] || ''),
    source:  String(v[D.SOURCE - 1] || ''),
    custId:  String(v[D.CUSTID - 1] || ''),
    memo:    String(v[D.MEMO - 1] || '')
  };
}

function custToObj(v) {
  return {
    id:      String(v[C.ID - 1] || ''),
    name:    String(v[C.NAME - 1] || ''),
    phone:   String(v[C.PHONE - 1] || ''),
    safe:    String(v[C.SAFE - 1] || ''),
    shop:    String(v[C.SHOP - 1] || ''),
    first:   ymd(v[C.FIRST - 1]),
    last:    ymd(v[C.LAST - 1]),
    visits:  Number(v[C.VISITS - 1] || 0),
    product: String(v[C.PRODUCT - 1] || ''),
    total:   v[C.TOTAL - 1] === '' ? null : Number(v[C.TOTAL - 1]),
    used:    Number(v[C.USED - 1] || 0),
    left:    v[C.LEFT - 1] === '' ? null : Number(v[C.LEFT - 1]),
    memo:    String(v[C.MEMO - 1] || ''),
    folder:  String(v[C.FOLDER - 1] || ''),
    next:    ymd(v[C.NEXT - 1]),
    gender:  String(v[C.GENDER - 1] || ''),
    birth:   ymd(v[C.BIRTH - 1]),
    staff:   String(v[C.STAFF - 1] || ''),
    prepaid: Number(v[C.PREPAID - 1] || 0),
    rel:     String(v[C.REL - 1] || ''),
    needle:  String(v[C.NEEDLE - 1] || ''),
    pigment: String(v[C.PIGMENT - 1] || ''),
    tone:    String(v[C.TONE - 1] || ''),
    grade:   String(v[C.GRADE - 1] || ''),
    tag:     String(v[C.TAG - 1] || ''),
    point:   Number(v[C.POINT - 1] || 0),
    lang:    String(v[C.LANG - 1] || ''),
    cway:    String(v[C.CWAY - 1] || ''),
    cid:     String(v[C.CID - 1] || ''),
    caution: String(v[C.CAUTION - 1] || ''),
    route:   String(v[C.ROUTE - 1] || '')
  };
}

/* ---------- 앱 시작할 때 한 번 ---------- */

/* ---------- 앱 잠금 (비밀번호) ---------- */

function appPassword() { return String(cfgStr('앱 비밀번호', '')).trim(); }

/* 직원 PIN: "김실장:482913, 이실장:770124" 형식 */
function staffPins() {
  var out = [];
  String(cfgStr('직원 PIN', '')).split(',').forEach(function (chunk) {
    var t = String(chunk).trim();
    if (!t) return;
    var i = t.lastIndexOf(':');
    if (i < 1) return;
    var nm = t.slice(0, i).trim(), pin = t.slice(i + 1).trim();
    if (nm && pin) out.push({ name: nm, pin: pin });
  });
  return out;
}
function appLocked() { return !!(appPassword() || staffPins().length); }
function staffSeesSales() {
  return String(cfgStr('직원 매출 보기', '숨김')).indexOf('보임') >= 0;
}

/* 접속 기록 */
function logAuth(name, role, result) {
  try {
    var sh = ss().getSheetByName(SH.AUTH);
    if (!sh) {
      sh = ss().insertSheet(SH.AUTH);
      sh.getRange(1, 1, 1, AUTH_HEAD.length).setValues([AUTH_HEAD]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.appendRow([fmt(new Date(), 'yyyy-MM-dd HH:mm:ss'), name || '', role || '', result || '']);
    var last = sh.getLastRow();
    if (last > 501) sh.deleteRows(2, last - 501);
  } catch (e) {}
}

/* 무작정 대입 막기 (10분 안에 10번 실패하면 잠시 막습니다) */
function loginFails() {
  try { return Number(CacheService.getScriptCache().get('LOGINFAIL') || 0); } catch (e) { return 0; }
}
function bumpLoginFail() {
  try {
    var c = CacheService.getScriptCache();
    c.put('LOGINFAIL', String(loginFails() + 1), 600);
  } catch (e) {}
}
function clearLoginFail() {
  try { CacheService.getScriptCache().remove('LOGINFAIL'); } catch (e) {}
}

function tokenStore() { return PropertiesService.getScriptProperties(); }
function newToken(name, role) {
  var t = Utilities.getUuid().replace(/-/g, '');
  tokenStore().setProperty('TK_' + t,
    [String(Date.now() + 30 * 86400000), role || 'owner', name || '사장님'].join('|'));
  return t;
}
function clearTokens(role) {
  var st = tokenStore(), all = st.getProperties();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('TK_') !== 0) return;
    if (role && String(all[k]).split('|')[1] !== role) return;
    st.deleteProperty(k);
  });
}
function authInfo(token) {
  if (!appLocked()) return { ok: true, role: 'owner', name: '사장님' };
  token = String(token || '').trim();
  if (!token) return { ok: false };
  var st = tokenStore();
  var v = st.getProperty('TK_' + token);
  if (!v) return { ok: false };
  var p = String(v).split('|');
  if (Number(p[0]) < Date.now()) { st.deleteProperty('TK_' + token); return { ok: false }; }
  var role = p[1] || 'owner', name = p[2] || '사장님';
  /* 명단에서 지운 직원은 남아있는 토큰도 바로 막습니다 */
  if (role === 'staff') {
    var live = staffPins(), found = false;
    for (var i = 0; i < live.length; i++) if (live[i].name === name) { found = true; break; }
    if (!found) { st.deleteProperty('TK_' + token); return { ok: false }; }
  }
  return { ok: true, role: role, name: name };
}
function authOK(token) { return authInfo(token).ok; }

function api_login(pw) {
  var owner = appPassword(), staff = staffPins();
  if (!owner && !staff.length) {
    return { ok: true, token: 'open', role: 'owner', name: '사장님' };
  }
  if (loginFails() >= 10) {
    return { ok: false, msg: '틀린 시도가 많습니다. 10분 뒤에 다시 해주세요.' };
  }
  var v = String(pw || '').trim();
  if (owner && v === owner) {
    clearLoginFail(); logAuth('사장님', 'owner', '성공');
    return { ok: true, token: newToken('사장님', 'owner'), role: 'owner', name: '사장님' };
  }
  for (var i = 0; i < staff.length; i++) {
    if (v === staff[i].pin) {
      clearLoginFail(); logAuth(staff[i].name, 'staff', '성공');
      return { ok: true, token: newToken(staff[i].name, 'staff'),
               role: 'staff', name: staff[i].name };
    }
  }
  bumpLoginFail(); logAuth('', '', '실패');
  Utilities.sleep(1500);                         // 무작정 대입하는 것을 늦춥니다
  return { ok: false, msg: '번호가 맞지 않습니다.' };
}

/* 사장님만: 직원 PIN 명단 저장 */
function api_savePins(p) {
  var list = (p && p.staff) || [];
  var seen = {}, out = [];
  for (var i = 0; i < list.length; i++) {
    var nm = String(list[i].name || '').replace(/[,:]/g, ' ').trim();
    var pin = String(list[i].pin || '').trim();
    if (!nm || !pin) continue;
    if (!/^[A-Za-z0-9]{4,20}$/.test(pin)) {
      return { ok: false, msg: nm + ' 번호는 영문·숫자 4~20자로 넣어주세요. (띄어쓰기·기호 불가)' };
    }
    if (pin === appPassword()) return { ok: false, msg: nm + ' 번호가 사장님 번호와 같습니다.' };
    if (seen[pin]) return { ok: false, msg: '같은 번호를 두 사람에게 줄 수 없습니다: ' + pin };
    seen[pin] = 1;
    out.push(nm + ':' + pin);
  }
  if (out.length && !appPassword()) {
    return { ok: false, msg: '사장님 번호를 먼저 정해주세요. 안 그러면 사장님이 못 들어옵니다.' };
  }
  setCfg('직원 PIN', out.join(','));
  setCfg('직원 매출 보기', (p && p.sales) ? '보임' : '숨김');
  _cfgCache = null;
  return { ok: true, count: out.length };
}

function api_setPassword(p) {
  if (!authOK(p && p.token)) return { ok: false, authFail: true, msg: '다시 로그인해 주세요.' };
  var np = String((p && p.pw) || '').trim();
  if (np.length < 4) return { ok: false, msg: '번호는 4자 이상으로 정해주세요.' };
  if (!/^[A-Za-z0-9]{4,20}$/.test(np)) {
    return { ok: false, msg: '영문·숫자 4~20자로 정해주세요. (띄어쓰기·기호는 로그인 화면에서 입력할 수 없습니다)' };
  }
  var sp = staffPins();
  for (var i = 0; i < sp.length; i++) {
    if (sp[i].pin === np) return { ok: false, msg: '직원 번호와 같은 번호는 쓸 수 없습니다.' };
  }
  setCfg('앱 비밀번호', np);
  _cfgCache = null;
  clearTokens('owner');                           // 사장님으로 로그인된 기기만 해제
  return { ok: true, token: newToken('사장님', 'owner') };
}

/* 화면에서 오는 모든 요청은 이 문을 거칩니다 */
var API_OPEN = { api_login: true };

/* 사장님만 쓸 수 있는 기능 */
var OWNER_ONLY = {
  api_sales: 1, api_salesReport: 1, api_cash: 1, api_saveCash: 1, api_deleteCash: 1,
  api_staffReport: 1, api_trash: 1, api_restoreTrash: 1, api_deleteCustomer: 1,
  api_setPassword: 1, api_savePins: 1, api_saveMenu: 1, api_saveForms: 1
};
/* 그 중 매출 관련 — [설정]에서 "직원 매출 보기 = 보임"이면 직원도 볼 수 있습니다 */
var SALES_API = {
  api_sales: 1, api_salesReport: 1, api_cash: 1, api_saveCash: 1,
  api_deleteCash: 1, api_staffReport: 1
};
function blockedForStaff(name) {
  if (!OWNER_ONLY[name]) return false;
  if (SALES_API[name] && staffSeesSales()) return false;
  return true;
}

/* 지금 요청을 보낸 사람 (api()가 채워 넣습니다) */
var CUR = { role: 'owner', name: '사장님' };

function api(name, args, token) {
  try {
    name = String(name || '');
    if (name.indexOf('api_') !== 0) return { ok: false, msg: '알 수 없는 요청입니다.' };
    var fn = globalThis[name];
    if (typeof fn !== 'function') return { ok: false, msg: '알 수 없는 요청입니다: ' + name };
    if (API_OPEN[name]) return fn.apply(null, args || []);
    var who = authInfo(token);
    if (!who.ok) return { ok: false, authFail: true, msg: '다시 로그인해 주세요.' };
    CUR = { role: who.role, name: who.name };
    if (who.role !== 'owner' && blockedForStaff(name)) {
      return { ok: false, msg: '이 기능은 사장님만 쓸 수 있습니다.' };
    }
    return fn.apply(null, args || []);
  } catch (e) {
    logError('앱 요청(' + name + ')', e);
    return { ok: false, msg: e.message };
  }
}

function api_bootstrap() {
  var list = shops().map(function (s) { return { name: s.name, short: s.short, idx: s.idx }; });
  return {
    ok: true,
    today: ymd(new Date()),
    shops: list,
    staff: cfgStr('담당자 명단', '').split(',').map(function (x) { return x.trim(); })
             .filter(function (x) { return x; }),
    testMode: cfgOn('테스트 모드'),
    smsReady: aligoReady(),
    lastMail: cfgStr('마지막 메일확인 시각', ''),
    locked: appLocked(),
    role: CUR.role,
    me: CUR.name,
    canSales: (CUR.role === 'owner') || staffSeesSales(),
    staffSales: staffSeesSales(),
    pins: (CUR.role === 'owner')
      ? staffPins().map(function (x) { return { name: x.name, pin: x.pin }; })
      : [],
    sheetUrl: (CUR.role === 'owner') ? ss().getUrl() : ''
  };
}

/* ---------- 홈 화면 ---------- */

function api_home() {
  var rows = readData();
  var t0 = ymd(new Date());
  var t1 = ymd(new Date(Date.now() + 86400000));
  var today = [], tomorrow = [];
  var needCheck = 0, noPhone = 0;

  rows.forEach(function (v) {
    var o = rowToObj(v);
    if (o.status === STATUS.CHECK) needCheck++;
    if (o.status === STATUS.OK && !o.phone && !o.safe && o.date >= t0) noPhone++;
    if (o.status === STATUS.CANCEL) return;
    if (o.date === t0) today.push(o);
    else if (o.date === t1) tomorrow.push(o);
  });
  var byTime = function (a, b) { return String(a.time).localeCompare(String(b.time)); };
  today.sort(byTime); tomorrow.sort(byTime);

  var lowRound = [];
  readCustomers().forEach(function (v) {
    var c = custToObj(v);
    if (c.total !== null && c.left !== null && c.left <= 1) lowRound.push(c);
  });

  // 이번달 매출
  var now = new Date();
  var ym = fmt(now, 'yyyy-MM');
  var monthSales = {}, todaySales = {};
  shops().forEach(function (s) { monthSales[s.name] = 0; todaySales[s.name] = 0; });
  readSlog().forEach(function (g) {
    if (String(g.date).indexOf(t0) === 0) {
      if (todaySales[g.shop] === undefined) todaySales[g.shop] = 0;
      todaySales[g.shop] += g.amount;
    }
    if (String(g.date).indexOf(ym) !== 0) return;
    if (monthSales[g.shop] === undefined) monthSales[g.shop] = 0;
    monthSales[g.shop] += g.amount;
  });
  // 시술은 끝났는데 결제가 아직 안 들어간 예약
  var noPaid = 0;
  var paidSet = {};
  readSlog().forEach(function (g) { paidSet[g.resno] = true; });
  rows.forEach(function (v) {
    var o = rowToObj(v);
    if (o.status === STATUS.DONE && !paidSet[o.resno]) noPaid++;
  });

  return {
    ok: true, today: t0, tomorrow: t1,
    todayList: today, tomorrowList: tomorrow,
    needCheck: needCheck, noPhone: noPhone,
    lowRound: lowRound.slice(0, 20),
    monthSales: monthSales,
    todaySales: todaySales,
    twoLeft: (function () { try { return twoLeftList_(); } catch (e) { return []; } })(),
    browRetouch: (function () { try { return browRetouchHome_(); } catch (e) { return []; } })(),
    noPaid: noPaid,
    lastMail: cfgStr('마지막 메일확인 시각', '')
  };
}

/* ---------- 날짜별 예약 ---------- */

function api_day(dateStr) {
  var rows = readData();
  var out = rows.map(rowToObj).filter(function (o) { return o.date === dateStr; });
  out.sort(function (a, b) { return String(a.time).localeCompare(String(b.time)); });
  return { ok: true, date: dateStr, list: out };
}

/* 한 달 요약: 날짜별 건수 (달력에 점 찍기용) */
function api_range(from, to) {
  try {
    var f = ymd(toDate(from)), t = ymd(toDate(to));
    var out = readData().map(rowToObj).filter(function (o) {
      return o.date && o.date >= f && o.date <= t;
    });
    out.sort(function (a, b) {
      if (a.date !== b.date) return a.date < b.date ? -1 : 1;
      return String(a.time).localeCompare(String(b.time));
    });
    return { ok: true, from: f, to: t, list: out };
  } catch (e) { return { ok: false, msg: e.message, list: [] }; }
}

function api_month(ym) {
  var counts = {};
  readData().forEach(function (v) {
    var o = rowToObj(v);
    if (o.status === STATUS.CANCEL) return;
    if (o.date.indexOf(ym) !== 0) return;
    counts[o.date] = counts[o.date] || { n: 0, shops: {} };
    counts[o.date].n++;
    counts[o.date].shops[o.shop] = true;
  });
  var out = {};
  Object.keys(counts).forEach(function (k) {
    out[k] = { n: counts[k].n, shops: Object.keys(counts[k].shops) };
  });
  return { ok: true, ym: ym, days: out };
}

/* ---------- 예약 수정 ---------- */

function api_updateBooking(p) {
  try {
    var sh = sheetOf(SH.DATA);
    var idx = buildResnoIndex();
    var t = idx[String(p.resno)];
    if (!t) return { ok: false, msg: '예약을 찾지 못했습니다.' };
    var r = t.r;
    var before = sh.getRange(r, 1, 1, D_HEAD.length).getValues()[0];

    if (p.phone !== undefined) {
      var ph = normPhone(p.phone);
      if (isSafePhone(ph)) sh.getRange(r, D.SAFE).setValue(ph);
      else sh.getRange(r, D.PHONE).setValue(ph);
    }
    if (p.staff !== undefined)  sh.getRange(r, D.STAFF).setValue(p.staff);
    if (p.status !== undefined) sh.getRange(r, D.STATUS).setValue(p.status);
    if (p.paid !== undefined && p.paid !== '')  sh.getRange(r, D.PAID).setValue(Number(p.paid));
    if (p.time !== undefined && p.time)  sh.getRange(r, D.TIME).setValue(normTime(p.time) || p.time);
    if (p.date !== undefined && p.date)  sh.getRange(r, D.DATE).setValue(toDate(p.date));
    if (p.memo !== undefined)   sh.getRange(r, D.MEMO).setValue(p.memo);

    var after = sh.getRange(r, 1, 1, D_HEAD.length).getValues()[0];
    var cid = after[D.CUSTID - 1];
    if (!cid && after[D.NAME - 1]) {
      cid = matchOrCreateCustomer(after[D.NAME - 1], after[D.PHONE - 1], after[D.SAFE - 1],
                                  after[D.SHOP - 1], after[D.SOURCE - 1] || '수동');
      sh.getRange(r, D.CUSTID).setValue(cid);
    } else if (cid && after[D.PHONE - 1]) {
      syncPhoneToCustomer(cid, after[D.PHONE - 1]);
    }

    if (after[D.STATUS - 1] === STATUS.CANCEL) {
      deleteGCalEvent(after[D.SHOP - 1], after[D.CALID - 1]);
      sh.getRange(r, D.CALID).setValue('');
    } else {
      upsertGCalEvent(sh, r);
    }
    syncCalendarForBooking(before[D.SHOP - 1], before[D.DATE - 1]);
    if (ymd(before[D.DATE - 1]) !== ymd(after[D.DATE - 1])) {
      syncCalendarForBooking(after[D.SHOP - 1], after[D.DATE - 1]);
    }
    recomputeCustomerStats();
    return { ok: true, booking: rowToObj(sh.getRange(r, 1, 1, D_HEAD.length).getValues()[0]) };
  } catch (e) {
    logError('앱 예약수정', e);
    return { ok: false, msg: e.message };
  }
}

/* ---------- 수동 예약 추가 ---------- */

function api_addBooking(p) {
  try {
    if (!p.shop || !p.date || !p.name) return { ok: false, msg: '샵·날짜·이름은 반드시 넣어주세요.' };
    var sh = sheetOf(SH.DATA);
    var resno = 'M' + fmt(new Date(), 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    var row = newDataRow();
    row[D.RESNO - 1]   = resno;
    row[D.SHOP - 1]    = p.shop;
    row[D.STATUS - 1]  = STATUS.OK;
    row[D.NAME - 1]    = String(p.name).trim();
    var ph = normPhone(p.phone || '');
    if (isSafePhone(ph)) row[D.SAFE - 1] = ph; else row[D.PHONE - 1] = ph;
    row[D.DATE - 1]    = toDate(p.date);
    row[D.TIME - 1]    = normTime(p.time || '') || (p.time || '');
    row[D.PRODUCT - 1] = p.product || '';
    row[D.OPTION - 1]  = p.option || '';
    row[D.AMOUNT - 1]  = p.amount ? Number(p.amount) : '';
    row[D.PAID - 1]    = '';
    row[D.STAFF - 1]   = p.staff || '';
    row[D.REQUEST - 1] = p.request || '';
    row[D.SOURCE - 1]  = '수동';
    row[D.PROCAT - 1]  = nowStr();
    row[D.CUSTID - 1]  = matchOrCreateCustomer(row[D.NAME - 1], row[D.PHONE - 1], row[D.SAFE - 1], p.shop, '수동');
    sh.appendRow(row);
    var r = sh.getLastRow();
    upsertGCalEvent(sh, r);
    syncCalendarForBooking(p.shop, row[D.DATE - 1]);
    recomputeCustomerStats();
    return { ok: true, resno: resno };
  } catch (e) {
    logError('앱 예약추가', e);
    return { ok: false, msg: e.message };
  }
}

/* ---------- 고객 ---------- */

function api_searchCustomers(term, filter) {
  term = String(term || '').trim();
  var f = String(filter || '').trim();
  var rows = readCustomers().map(custToObj);
  var today = ymd(new Date());
  var mmdd = today.slice(5);

  if (f) {
    rows = rows.filter(function (c) {
      if (f === '생일') return c.birth && String(c.birth).slice(5) === mmdd;
      if (f === '이번달생일') return c.birth && String(c.birth).slice(5, 7) === today.slice(5, 7);
      if (f === '선불보유') return Number(c.prepaid || 0) > 0;
      if (f === '번호없음') return !c.phone && !c.safe && !c.cid;
      if (f === '외국인') return !!c.lang && c.lang !== '한국어';
      if (f === '휴면') {
        if (c.grade === '휴면') return true;
        if (!c.last) return false;
        return dayDiff(c.last, today) >= 180;
      }
      if (GRADES.indexOf(f) >= 0) return c.grade === f;
      return String(c.tag || '').indexOf(f) >= 0;
    });
  }

  if (!term) {
    rows.sort(function (a, b) { return String(b.last).localeCompare(String(a.last)); });
    return { ok: true, list: rows.slice(0, 60), total: rows.length,
             grades: GRADES, tags: allTags() };
  }
  /* 영문 이름은 대소문자 구분 없이 찾습니다 (ph → Phionna) */
  var lt = term.toLowerCase();
  var hit = rows.filter(function (c) {
    return String(c.name).toLowerCase().indexOf(lt) >= 0 ||
           last4(c.phone) === term || last4(c.safe) === term ||
           onlyDigits(c.phone).indexOf(onlyDigits(term)) >= 0 && onlyDigits(term).length >= 3 ||
           String(c.tag || '').toLowerCase().indexOf(lt) >= 0 ||
           (c.cid && String(c.cid).toLowerCase().replace(/^@/, '').indexOf(lt.replace(/^@/, '')) >= 0) ||
           c.id === term;
  });
  return { ok: true, list: hit.slice(0, 60), total: hit.length,
           grades: GRADES, tags: allTags() };
}

function allTags() {
  var set = {};
  readCustomers().forEach(function (v) {
    String(v[C.TAG - 1] || '').split(',').forEach(function (t) {
      t = t.trim(); if (t) set[t] = true;
    });
  });
  return Object.keys(set).sort();
}

function api_customer(id) {
  var rows = readCustomers();
  var found = null;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][C.ID - 1]).trim() === String(id).trim()) { found = custToObj(rows[i]); break; }
  }
  if (!found) return { ok: false, msg: '고객을 찾지 못했습니다.' };
  var hist = readData().map(rowToObj).filter(function (o) { return o.custId === found.id; });
  hist.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
  var sales = 0;
  try {
    readSlog().forEach(function (g) { if (g.custId === found.id) sales += g.amount; });
  } catch (e) {}
  return { ok: true, customer: found, history: hist, sales: sales,
           staff: cfgStr('담당자 명단', '').split(',').map(function (x) { return x.trim(); })
                    .filter(function (x) { return x; }) };
}

/* ── 연결 고객 양쪽 맞추기: A에 B를 넣으면 B에도 A가 들어가고, A에서 B를 빼면 B에서도 A가 빠집니다 ── */
function relEntries(v) {
  return String(v || '').split(',').map(function (x) { return x.trim(); }).filter(function (x) { return x; })
    .map(function (x) {
      var m = /^([^(]+?)\s*\((\d{4})\)\s*(?:\(([^)]*)\))?/.exec(x);
      return m ? { name: m[1].trim(), p4: m[2], rel: (m[3] || '').trim(), raw: x } : { raw: x };
    });
}
function relReverse(r) {
  if (r === '소개해 준 분') return '소개받은 분';
  if (r === '소개받은 분') return '소개해 준 분';
  if (['배우자', '가족', '친구', '지인', '직장 동료'].indexOf(r) >= 0) return r;
  return '';  /* '언니'처럼 직접 쓴 관계는 반대말을 알 수 없어 이름만 연결 */
}
function syncRelations(rows, me, oldRel, newRel, newName, newPhone) {
  var sh = sheetOf(SH.CUST);
  var myName = String(newName || rows[me][C.NAME - 1]).trim();
  var myP4 = last4(onlyDigits(newPhone !== undefined ? newPhone : (rows[me][C.PHONE - 1] || rows[me][C.SAFE - 1])));
  if (!myName || !myP4) return;
  var myKey = myName + '(' + myP4 + ')';
  function findRow(e) {
    for (var k = 0; k < rows.length; k++) {
      if (k === me) continue;
      if (String(rows[k][C.NAME - 1]).trim() === e.name &&
          (last4(onlyDigits(rows[k][C.PHONE - 1])) === e.p4 || last4(onlyDigits(rows[k][C.SAFE - 1])) === e.p4)) return k;
    }
    return -1;
  }
  var olds = relEntries(oldRel).filter(function (e) { return e.name; });
  var news = relEntries(newRel).filter(function (e) { return e.name; });
  var keyOf = function (e) { return e.name + '(' + e.p4 + ')'; };
  var newKeys = news.map(keyOf);
  /* 새로 넣었거나 관계가 바뀐 사람: 상대 쪽에 나를 넣기 */
  news.forEach(function (e) {
    var k = findRow(e); if (k < 0) return;
    var theirs = relEntries(rows[k][C.REL - 1]);
    var back = relReverse(e.rel);
    var line = myKey + (back ? ' (' + back + ')' : '');
    var at = -1;
    theirs.forEach(function (t, j) { if (t.name && keyOf(t) === myKey) at = j; });
    if (at >= 0) { if (!back || theirs[at].rel === back) return; theirs[at] = { raw: line }; }
    else theirs.push({ raw: line });
    var v = theirs.map(function (t) { return t.raw; }).join(', ');
    sh.getRange(k + 2, C.REL).setValue(v); rows[k][C.REL - 1] = v;
  });
  /* 뺀 사람: 상대 쪽에서도 나를 빼기 */
  olds.forEach(function (e) {
    if (newKeys.indexOf(keyOf(e)) >= 0) return;
    var k = findRow(e); if (k < 0) return;
    var theirs = relEntries(rows[k][C.REL - 1]).filter(function (t) { return !(t.name && keyOf(t) === myKey); });
    var v = theirs.map(function (t) { return t.raw; }).join(', ');
    sh.getRange(k + 2, C.REL).setValue(v); rows[k][C.REL - 1] = v;
  });
}

function api_saveCustomer(p) {
  try {
    var sh = sheetOf(SH.CUST);
    var rows = readCustomers();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][C.ID - 1]).trim() === String(p.id).trim()) {
        var r = i + 2;
        if (p.name !== undefined && p.name)  sh.getRange(r, C.NAME).setValue(String(p.name).trim());
        if (p.phone !== undefined) {
          var ph = normPhone(p.phone);
          if (isSafePhone(ph)) sh.getRange(r, C.SAFE).setValue(ph);
          else sh.getRange(r, C.PHONE).setValue(ph);
        }
        if (p.memo !== undefined)    sh.getRange(r, C.MEMO).setValue(p.memo);
        if (p.product !== undefined) sh.getRange(r, C.PRODUCT).setValue(p.product);
        if (p.total !== undefined)   sh.getRange(r, C.TOTAL).setValue(p.total === '' ? '' : Number(p.total));
        if (p.gender !== undefined)  sh.getRange(r, C.GENDER).setValue(p.gender);
        if (p.birth !== undefined)   sh.getRange(r, C.BIRTH).setValue(p.birth ? toDate(p.birth) : '');
        if (p.staff !== undefined)   sh.getRange(r, C.STAFF).setValue(p.staff);
        if (p.prepaid !== undefined) sh.getRange(r, C.PREPAID).setValue(p.prepaid === '' ? '' : Number(p.prepaid));
        if (p.rel !== undefined) {
          var oldRel = String(rows[i][C.REL - 1] || '');
          sh.getRange(r, C.REL).setValue(p.rel);
          try { syncRelations(rows, i, oldRel, String(p.rel || ''), p.name, p.phone); } catch (er) { logError('연결고객 맞추기', er); }
        }
        if (p.needle !== undefined)  sh.getRange(r, C.NEEDLE).setValue(p.needle);
        if (p.pigment !== undefined) sh.getRange(r, C.PIGMENT).setValue(p.pigment);
        if (p.tone !== undefined)    sh.getRange(r, C.TONE).setValue(p.tone);
        if (p.grade !== undefined)   sh.getRange(r, C.GRADE).setValue(p.grade);
        if (p.tag !== undefined)     sh.getRange(r, C.TAG).setValue(p.tag);
        if (p.lang !== undefined)    sh.getRange(r, C.LANG).setValue(p.lang);
        if (p.cway !== undefined)    sh.getRange(r, C.CWAY).setValue(p.cway);
        if (p.cid !== undefined)     sh.getRange(r, C.CID).setValue(String(p.cid).trim());
        if (p.caution !== undefined) sh.getRange(r, C.CAUTION).setValue(p.caution);
        sh.getRange(r, C.LEFT).setFormula('=IF($J' + r + '="","",$J' + r + '-$K' + r + ')');
        SpreadsheetApp.flush();
        return { ok: true, customer: custToObj(sh.getRange(r, 1, 1, C_HEAD.length).getValues()[0]) };
      }
    }
    return { ok: false, msg: '고객을 찾지 못했습니다.' };
  } catch (e) {
    logError('앱 고객저장', e);
    return { ok: false, msg: e.message };
  }
}

function api_addCustomer(p) {
  try {
    var name = String((p && p.name) || '').trim();
    if (!name) return { ok: false, msg: '이름을 넣어주세요.' };
    var ph = normPhone(p.phone || '');
    var phone = '', safe = '';
    if (ph) { if (isSafePhone(ph)) safe = ph; else phone = ph; }
    var shopName = p.shop || (shops()[0] ? shops()[0].name : '');
    var id = matchOrCreateCustomer(name, phone, safe, shopName, '수동');
    if (!id) return { ok: false, msg: '고객을 만들지 못했습니다.' };
    var q = { id: id, name: name };
    if (ph) q.phone = ph;
    if (p.product !== undefined) q.product = p.product || '';
    if (p.total !== undefined)   q.total = (p.total === '' ? '' : p.total);
    if (p.memo !== undefined)    q.memo = p.memo || '';
    var r = api_saveCustomer(q);
    return { ok: true, id: id, customer: r && r.customer };
  } catch (e) {
    logError('앱 고객추가', e);
    return { ok: false, msg: e.message };
  }
}

function api_deleteCustomer(id) {
  try {
    var sh = sheetOf(SH.CUST);
    var rows = readCustomers();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][C.ID - 1]).trim() === String(id).trim()) {
        toTrash('고객', id, rows[i][C.NAME - 1] + ' ' + (rows[i][C.PHONE - 1] || rows[i][C.SAFE - 1] || ''), rows[i]);
        sh.deleteRow(i + 2);
        SpreadsheetApp.flush();
        return { ok: true };
      }
    }
    return { ok: false, msg: '고객을 찾지 못했습니다.' };
  } catch (e) {
    logError('앱 고객삭제', e);
    return { ok: false, msg: e.message };
  }
}

function api_customerFolder(id) {
  try { return { ok: true, url: ensureCustomerFolder(String(id)) }; }
  catch (e) { return { ok: false, msg: e.message }; }
}

/* ---------- 매출 ---------- */

function api_sales(year) {
  year = Number(year) || new Date().getFullYear();
  var out = { ok: true, year: year, shops: {}, months: [], byItem: [], total: 0,
              fee: 0, net: 0, feeMonths: [], feeRate: npayRate() };
  shops().forEach(function (s) {
    out.shops[s.name] = { total: 0, months: [] };
    for (var m = 0; m < 12; m++) out.shops[s.name].months.push(0);
  });
  for (var i = 0; i < 12; i++) { out.months.push(0); out.feeMonths.push(0); }
  var items = {};
  var byResM = {};
  readSlog().forEach(function (g) {
    if (!g.date || Number(String(g.date).slice(0, 4)) !== year) return;
    var m = Number(String(g.date).slice(5, 7)) - 1;
    if (m < 0 || m > 11) return;
    if (!out.shops[g.shop]) {
      out.shops[g.shop] = { total: 0, months: [] };
      for (var k = 0; k < 12; k++) out.shops[g.shop].months.push(0);
    }
    out.shops[g.shop].months[m] += g.amount;
    out.shops[g.shop].total += g.amount;
    out.months[m] += g.amount;
    out.total += g.amount;
    var rk = String(m) + '|' + g.resno;
    var br = byResM[rk] || (byResM[rk] = { m: m, pay: '', amount: 0, npay: 0 });
    if (g.pay) br.pay = g.pay;
    br.amount += g.amount;
    if (Number(g.npay) > br.npay) br.npay = Number(g.npay);
    var key = g.shop + '|' + g.item;
    if (!items[key]) items[key] = { shop: g.shop, name: g.item, total: 0, n: 0 };
    items[key].total += g.amount;
    if (g.item !== '할인') items[key].n++;
  });
  Object.keys(byResM).forEach(function (k) {
    var b = byResM[k];
    var one = {}; one[k] = b;
    var f = feeOf(one).fee;
    out.feeMonths[b.m] += f;
    out.fee += f;
  });
  out.net = out.total - out.fee;
  out.byItem = Object.keys(items).map(function (k) { return items[k]; })
    .sort(function (a, b) { return b.total - a.total; });
  return out;
}

/* ---------- 시술메뉴 ---------- */

function api_menu() {
  try { return { ok: true, list: readMenu(), shops: shops().map(function (s) { return s.name; }) }; }
  catch (e) { return { ok: false, msg: e.message, list: [], shops: [] }; }
}

function api_saveMenu(list) {
  try {
    var sh = buildMenuSheet();
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, MENU_HEAD.length).clearContent();
    var rows = (list || []).filter(function (m) { return String(m.name || '').trim(); })
      .map(function (m) {
        return [String(m.shop || '').trim(), String(m.name).trim(),
                Number(m.price || 0), Number(m.min || 0)];
      });
    if (rows.length) sh.getRange(2, 1, rows.length, MENU_HEAD.length).setValues(rows);
    _menuCache = null;
    SpreadsheetApp.flush();
    return { ok: true, n: rows.length };
  } catch (e) { logError('앱 메뉴저장', e); return { ok: false, msg: e.message }; }
}

/* ---------- 결제(매출) 등록 ---------- */

function api_payment(resno) {
  try {
    var found = null;
    readData().forEach(function (v) {
      var o = rowToObj(v);
      if (o.resno === String(resno)) found = o;
    });
    if (!found) return { ok: false, msg: '예약을 찾지 못했습니다.' };
    var items = [], discount = 0, has = false, pay = '', staff = '', npay = 0;
    readSlog().forEach(function (g) {
      if (g.resno !== String(resno)) return;
      has = true;
      if (g.pay) pay = g.pay;
      if (g.staff) staff = g.staff;
      if (Number(g.npay) > npay) npay = Number(g.npay);
      if (g.item === '할인') discount += -g.amount;
      else items.push({ name: g.item, amount: g.amount });
    });
    if (!has && found.product) items.push({ name: found.product, amount: found.amount || 0 });
    var menu = readMenu().filter(function (m) { return !m.shop || m.shop === found.shop; });
    return { ok: true, booking: found, items: items, discount: discount, menu: menu, saved: has,
             pay: pay, staff: staff || found.staff || '', npay: npay, feeRate: npayRate(),
             payList: ['신용카드', '체크카드', '해외카드', '현금', '계좌이체', '네이버페이', '신용카드+네이버페이', '체크카드+네이버페이', '해외카드+네이버페이', '기타'],
             creditRate: creditRate(), debitRate: debitRate(), foreignRate: foreignRate(),
             scalp: { one: Number(cfgStr('두피 부위당 1회 금액', '250000')) || 250000,
                      twoPre: Number(cfgStr('두피 부위당 2회 함께 결제 금액', '450000')) || 450000 },
             staffList: cfgStr('담당자 명단', '').split(',').map(function (x) { return x.trim(); })
                          .filter(function (x) { return x; }) };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function api_savePayment(p) {
  try {
    var resno = String((p && p.resno) || '').trim();
    if (!resno) return { ok: false, msg: '예약번호가 없습니다.' };
    var dataSh = sheetOf(SH.DATA);
    var rows = readData();
    var idx = -1, o = null;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][D.RESNO - 1]).trim() === resno) { idx = i + 2; o = rowToObj(rows[i]); break; }
    }
    if (idx < 0) return { ok: false, msg: '예약을 찾지 못했습니다.' };
    var preTotal = 0;
    (p.items || []).forEach(function (it) { if (String(it.name || '').trim()) preTotal += Number(it.amount || 0); });
    preTotal -= Number(p.discount || 0);
    if (preTotal < 0) return { ok: false, msg: '할인이 시술 금액보다 커서 실결제금액이 0원보다 작아요. 금액을 다시 확인해 주세요.' };

    var sh = buildSlogSheet();
    var last = sh.getLastRow();
    if (last > 1) {
      var vals = sh.getRange(2, 1, last - 1, SLOG_HEAD.length).getValues();
      for (var r = vals.length - 1; r >= 0; r--) {
        if (String(vals[r][1]).trim() === resno) sh.deleteRow(r + 2);
      }
    }
    var items = (p.items || []).filter(function (it) { return String(it.name || '').trim(); });
    var disc = Number(p.discount || 0);
    var pay = String(p.pay || '').trim();
    var staff = String(p.staff || o.staff || '').trim();
    var cid = String(o.custId || '');
    var npay = Math.max(0, Number(p.npay || 0));
    var total = 0, out = [], stamp = nowStr(), dt = toDate(o.date);
    items.forEach(function (it) {
      var amt = Number(it.amount || 0);
      total += amt;
      out.push([stamp, resno, o.shop, cid, o.name, dt, String(it.name).trim(), amt, pay, staff, '']);
    });
    if (disc) {
      total -= disc;
      out.push([stamp, resno, o.shop, cid, o.name, dt, '할인', -disc, pay, staff, '']);
    }
    /* 네이버 결제액은 예약당 한 번만 (첫 줄에) 적습니다 */
    if (out.length && npay) out[0][10] = npay;
    if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, SLOG_HEAD.length).setValues(out);

    dataSh.getRange(idx, D.PAID).setValue(total);
    dataSh.getRange(idx, D.STATUS).setValue(STATUS.DONE);
    if (staff) dataSh.getRange(idx, D.STAFF).setValue(staff);
    SpreadsheetApp.flush();
    recomputeCustomerStats();
    refreshHome();
    return { ok: true, total: total };
  } catch (e) { logError('앱 결제등록', e); return { ok: false, msg: e.message }; }
}

/* ---------- 동의서 양식·기록 ---------- */

var FORM_HEAD = ['샵', '양식명', '순서', '문항', '유형', '보기', '필수'];
var CONSENT_HEAD = ['등록시각', '고객ID', '고객명', '샵', '양식명', '응답', '서명링크'];

function buildFormSheet() {
  var sh = getOrCreateSheet(SH.FORM);
  if (headerMismatch(sh, FORM_HEAD)) {
    sh.clear();
    sh.getRange(1, 1, 1, FORM_HEAD.length).setValues([FORM_HEAD])
      .setFontWeight('bold').setBackground('#1F3864').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    [140, 170, 60, 420, 90, 420, 70].forEach(function (x, i) { sh.setColumnWidth(i + 1, x); });
  }
  return sh;
}

function buildConsentSheet() {
  var sh = getOrCreateSheet(SH.CONSENT);
  if (headerMismatch(sh, CONSENT_HEAD)) {
    sh.clear();
    sh.getRange(1, 1, 1, CONSENT_HEAD.length).setValues([CONSENT_HEAD])
      .setFontWeight('bold').setBackground('#7B3F61').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    [150, 90, 100, 140, 160, 520, 220].forEach(function (x, i) { sh.setColumnWidth(i + 1, x); });
  }
  return sh;
}

/* ==================== 선불·포인트 이력 ==================== */

function buildPpSheet() {
  var sh = getOrCreateSheet(SH.PPLOG);
  if (headerMismatch(sh, PP_HEAD)) {
    sh.clear();
    sh.getRange(1, 1, 1, PP_HEAD.length).setValues([PP_HEAD])
      .setFontWeight('bold').setBackground('#8A5B2A').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    [150, 90, 100, 80, 80, 110, 110, 240, 100].forEach(function (x, i) { sh.setColumnWidth(i + 1, x); });
  }
  return sh;
}

function readPp(custId) {
  var sh = buildPpSheet();
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, PP_HEAD.length).getValues()
    .filter(function (r) { return String(r[1] || '').trim(); })
    .map(function (r) {
      return { at: String(r[0]), custId: String(r[1]).trim(), name: String(r[2] || ''),
               type: String(r[3] || ''), way: String(r[4] || ''),
               amount: Number(r[5] || 0), left: Number(r[6] || 0),
               memo: String(r[7] || ''), staff: String(r[8] || '') };
    })
    .filter(function (g) { return !custId || g.custId === String(custId).trim(); });
}

/* type: 선불 | 포인트   way: 충전 | 사용 | 조정 */
function api_savePp(p) {
  try {
    var cid = String((p && p.custId) || '').trim();
    if (!cid) return { ok: false, msg: '고객을 찾지 못했습니다.' };
    var amt = Math.abs(Number(p.amount || 0));
    if (!amt) return { ok: false, msg: '금액을 넣어주세요.' };
    var type = (p.type === '포인트') ? '포인트' : '선불';
    var way = (p.way === '사용') ? '사용' : (p.way === '조정' ? '조정' : '충전');

    var sh = sheetOf(SH.CUST);
    var rows = readCustomers();
    var idx = -1;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][C.ID - 1]).trim() === cid) { idx = i; break; }
    }
    if (idx < 0) return { ok: false, msg: '고객을 찾지 못했습니다.' };

    var col = (type === '포인트') ? C.POINT : C.PREPAID;
    var cur = Number(rows[idx][col - 1] || 0);
    var delta = (way === '사용') ? -amt : amt;
    if (way === '조정') delta = amt - cur;
    var next = cur + delta;
    if (next < 0) return { ok: false, msg: '잔액보다 많이 쓸 수 없습니다. (현재 ' + cur + '원)' };

    sheetOf(SH.CUST).getRange(idx + 2, col).setValue(next);
    buildPpSheet().appendRow([nowStr(), cid, rows[idx][C.NAME - 1], type, way,
                              (way === '조정' ? delta : amt), next,
                              p.memo || '', p.staff || '']);
    SpreadsheetApp.flush();
    return { ok: true, left: next, type: type };
  } catch (e) { logError('앱 선불포인트', e); return { ok: false, msg: e.message }; }
}

function api_ppLog(custId) {
  try {
    var list = readPp(custId);
    list.reverse();
    return { ok: true, list: list.slice(0, 60) };
  } catch (e) { return { ok: false, msg: e.message, list: [] }; }
}

/* ==================== 직원별 결산 ==================== */

function staffList() {
  return cfgStr('담당자 명단', '').split(',')
    .map(function (x) { return x.trim(); })
    .filter(function (x) { return x; });
}

function incentiveMap() {
  /* 설정: "원장:30,김실장:20"  (퍼센트) */
  var raw = String(cfgStr('직원 인센티브(%)', '')).trim();
  var m = {};
  raw.split(',').forEach(function (pair) {
    var kv = pair.split(':');
    if (kv.length === 2 && kv[0].trim()) m[kv[0].trim()] = Number(kv[1]) || 0;
  });
  return m;
}

function api_staffReport(from, to, shop) {
  try {
    var t = ymd(new Date());
    var f = ymd(toDate(from) || toDate(t.slice(0, 8) + '01'));
    var e = ymd(toDate(to) || toDate(t));
    var inc = incentiveMap();
    var PAYS = ['카드', '현금', '계좌이체', '네이버페이', '기타'];

    var by = {};
    function slot(name) {
      if (!by[name]) {
        by[name] = { staff: name, total: 0, n: 0, cust: {}, pay: {}, items: {} };
        PAYS.forEach(function (k) { by[name].pay[k] = 0; });
        by[name].pay['미입력'] = 0;
      }
      return by[name];
    }

    readSlog().forEach(function (g) {
      if (!g.date || g.date < f || g.date > e) return;
      if (shop && g.shop !== shop) return;
      var name = g.staff || '지정 안 함';
      var b = slot(name);
      b.total += g.amount;
      if (g.item !== '할인') {
        b.n++;
        b.items[g.item] = (b.items[g.item] || 0) + g.amount;
      }
      var pk = g.pay && PAYS.indexOf(g.pay) >= 0 ? g.pay : (g.pay ? '기타' : '미입력');
      b.pay[pk] += g.amount;
      if (g.custId) b.cust[g.custId] = true;
    });

    var list = Object.keys(by).map(function (k) {
      var b = by[k];
      var rate = inc[k] || 0;
      return { staff: b.staff, total: b.total, n: b.n,
               customers: Object.keys(b.cust).length,
               pay: b.pay, rate: rate,
               incentive: Math.floor(b.total * rate / 100),
               top: Object.keys(b.items).map(function (i) { return { name: i, total: b.items[i] }; })
                      .sort(function (x, y) { return y.total - x.total; }).slice(0, 5) };
    }).sort(function (a, b) { return b.total - a.total; });

    var grand = 0; list.forEach(function (x) { grand += x.total; });
    return { ok: true, from: f, to: e, shop: shop || '', list: list, total: grand,
             pays: PAYS.concat(['미입력']),
             staffList: staffList(), shops: shops().map(function (x) { return x.name; }) };
  } catch (e) { logError('앱 직원결산', e); return { ok: false, msg: e.message }; }
}

/* ==================== 이용률 분석 ==================== */

function api_usage(from, to, shop) {
  try {
    var t = ymd(new Date());
    var f = ymd(toDate(from) || toDate(shiftDate(t, -29)));
    var e = ymd(toDate(to) || toDate(t));

    function hourOf(v, dflt) {
      var m = String(v == null ? '' : v).match(/(\d{1,2})\s*[:시]/);
      if (m) return Number(m[1]);
      var n = Number(String(v).replace(/[^0-9]/g, ''));
      return n > 0 && n < 24 ? n : dflt;
    }
    var openH = hourOf(cfgStr('영업 시작시각', '10:00'), 10);
    var closeH = hourOf(cfgStr('영업 종료시각', '20:00'), 20) + 1;
    if (closeH <= openH) closeH = openH + 1;
    if (closeH > 24) closeH = 24;

    var grid = {}, dowTot = [0,0,0,0,0,0,0], hourTot = {}, dayCount = [0,0,0,0,0,0,0];
    for (var h = openH; h < closeH; h++) hourTot[h] = 0;

    var seenDay = {};
    var cur = f;
    while (cur <= e) {
      var dd = toDate(cur);
      dayCount[dd.getDay()]++;
      cur = shiftDate(cur, 1);
    }

    var total = 0;
    readData().map(rowToObj).forEach(function (o) {
      if (!o.date || o.date < f || o.date > e) return;
      if (o.status === STATUS.CANCEL) return;
      if (shop && o.shop !== shop) return;
      var d = toDate(o.date); if (!d) return;
      var m = String(o.time || '').match(/(\d{1,2})\s*[:시]/);
      var hh = m ? Number(m[1]) : -1;
      if (hh < openH || hh >= closeH) hh = -1;
      total++;
      dowTot[d.getDay()]++;
      if (hh >= 0) {
        hourTot[hh]++;
        var k = d.getDay() + '|' + hh;
        grid[k] = (grid[k] || 0) + 1;
      }
      seenDay[o.date] = true;
    });

    var maxCell = 0;
    Object.keys(grid).forEach(function (k) { if (grid[k] > maxCell) maxCell = grid[k]; });

    var hours = [];
    for (var h2 = openH; h2 < closeH; h2++) hours.push(h2);

    return { ok: true, from: f, to: e, shop: shop || '',
             open: openH, close: closeH, hours: hours,
             grid: grid, maxCell: maxCell,
             dow: dowTot, dowDays: dayCount, hourTot: hourTot,
             total: total, busyDays: Object.keys(seenDay).length,
             shops: shops().map(function (x) { return x.name; }) };
  } catch (e) { logError('앱 이용률', e); return { ok: false, msg: e.message }; }
}

/* ==================== 간단일정 (시간 막기) ==================== */

function buildBlockSheet() {
  var sh = getOrCreateSheet(SH.BLOCK);
  if (headerMismatch(sh, BLOCK_HEAD)) {
    sh.clear();
    sh.getRange(1, 1, 1, BLOCK_HEAD.length).setValues([BLOCK_HEAD])
      .setFontWeight('bold').setBackground('#5B6B7B').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    [90, 140, 100, 80, 80, 220, 260, 150].forEach(function (x, i) { sh.setColumnWidth(i + 1, x); });
  }
  return sh;
}

function readBlocks() {
  var sh = buildBlockSheet();
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, BLOCK_HEAD.length).getValues()
    .filter(function (r) { return String(r[0] || '').trim(); })
    .map(function (r) {
      return { id: String(r[0]), shop: String(r[1] || ''), date: ymd(r[2]),
               start: hhmm(r[3]), end: hhmm(r[4]),
               title: String(r[5] || ''), memo: String(r[6] || '') };
    });
}

function hhmm(v) {
  if (v instanceof Date) return fmt(v, 'HH:mm');
  var t = String(v || '');
  var m = t.match(/(\d{1,2})\s*[:시]\s*(\d{1,2})?/);
  if (!m) return '';
  var h = Number(m[1]), mi = Number(m[2] || 0);
  if (/오후|PM|pm/.test(t) && h < 12) h += 12;
  if (/오전|AM|am/.test(t) && h === 12) h = 0;
  if (h > 23) h = 23;
  return ('0' + h).slice(-2) + ':' + ('0' + mi).slice(-2);
}

function api_blocks(from, to) {
  try {
    var f = ymd(toDate(from)), t = ymd(toDate(to || from));
    var list = readBlocks().filter(function (b) { return b.date >= f && b.date <= t; });
    list.sort(function (a, b) {
      if (a.date !== b.date) return a.date < b.date ? -1 : 1;
      return String(a.start).localeCompare(String(b.start));
    });
    return { ok: true, list: list, shops: shops().map(function (x) { return x.name; }) };
  } catch (e) { return { ok: false, msg: e.message, list: [] }; }
}

function api_saveBlock(p) {
  try {
    var date = ymd(toDate(p && p.date));
    if (!date) return { ok: false, msg: '날짜를 골라주세요.' };
    var title = String((p && p.title) || '').trim() || '일정';
    var st = hhmm(p.start), en = hhmm(p.end);
    if (!st) return { ok: false, msg: '시작 시간을 넣어주세요.' };
    if (en && en <= st) return { ok: false, msg: '끝 시간이 시작보다 빨라요.' };
    var sh = buildBlockSheet();

    if (p.id) {
      var rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), BLOCK_HEAD.length).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][0]) === String(p.id)) {
          sh.getRange(i + 2, 2, 1, 6).setValues([[p.shop || '', date, st, en, title, p.memo || '']]);
          SpreadsheetApp.flush();
          return { ok: true, id: p.id };
        }
      }
      return { ok: false, msg: '일정을 찾지 못했습니다.' };
    }

    var id = 'B' + Date.now();
    sh.appendRow([id, p.shop || '', date, st, en, title, p.memo || '', nowStr()]);
    SpreadsheetApp.flush();
    return { ok: true, id: id };
  } catch (e) { logError('앱 일정저장', e); return { ok: false, msg: e.message }; }
}

function api_deleteBlock(id) {
  try {
    var sh = buildBlockSheet();
    var last = sh.getLastRow();
    if (last < 2) return { ok: false, msg: '일정이 없습니다.' };
    var rows = sh.getRange(2, 1, last - 1, BLOCK_HEAD.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        toTrash('간단일정', id, rows[i][5] + ' ' + ymd(rows[i][2]), rows[i]);
        sh.deleteRow(i + 2);
        SpreadsheetApp.flush();
        return { ok: true };
      }
    }
    return { ok: false, msg: '일정을 찾지 못했습니다.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

/* ==================== 휴지통 ==================== */

function buildTrashSheet() {
  var sh = getOrCreateSheet(SH.TRASH);
  if (headerMismatch(sh, TRASH_HEAD)) {
    sh.clear();
    sh.getRange(1, 1, 1, TRASH_HEAD.length).setValues([TRASH_HEAD])
      .setFontWeight('bold').setBackground('#7A7A7A').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    [150, 100, 130, 320, 600].forEach(function (x, i) { sh.setColumnWidth(i + 1, x); });
  }
  return sh;
}

function toTrash(kind, key, summary, rowArr) {
  try {
    var sh = buildTrashSheet();
    var payload = JSON.stringify((rowArr || []).map(function (v) {
      return (v instanceof Date) ? fmt(v, 'yyyy-MM-dd HH:mm:ss') : v;
    }));
    if (payload.length > 45000) payload = payload.slice(0, 45000);
    sh.appendRow([nowStr(), kind, String(key), String(summary || ''), payload]);
    purgeTrash();
  } catch (e) { logError('휴지통 저장', e); }
}

function purgeTrash() {
  var sh = buildTrashSheet();
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var vals = sh.getRange(2, 1, last - 1, 1).getValues();
  var cut = ymd(shiftDate(ymd(new Date()), -30));
  var del = 0;
  for (var i = vals.length - 1; i >= 0; i--) {
    var d = ymd(toDate(String(vals[i][0]).slice(0, 10)));
    if (d && d < cut) { sh.deleteRow(i + 2); del++; }
  }
  return del;
}

function api_trash(kind) {
  try {
    var sh = buildTrashSheet();
    var last = sh.getLastRow();
    if (last < 2) return { ok: true, list: [], kinds: [] };
    var v = sh.getRange(2, 1, last - 1, TRASH_HEAD.length).getValues();
    var kinds = {};
    var out = [];
    v.forEach(function (r, i) {
      kinds[String(r[1])] = true;
      if (kind && String(r[1]) !== kind) return;
      out.push({ row: i + 2, at: String(r[0]), kind: String(r[1]),
                 key: String(r[2]), summary: String(r[3]) });
    });
    out.reverse();
    return { ok: true, list: out.slice(0, 100), total: out.length,
             kinds: Object.keys(kinds) };
  } catch (e) { return { ok: false, msg: e.message, list: [] }; }
}

function api_restoreTrash(rowNum) {
  try {
    var sh = buildTrashSheet();
    var r = Number(rowNum);
    if (!(r >= 2) || r > sh.getLastRow()) return { ok: false, msg: '기록을 찾지 못했습니다.' };
    var v = sh.getRange(r, 1, 1, TRASH_HEAD.length).getValues()[0];
    var kind = String(v[1]);
    var arr;
    try { arr = JSON.parse(v[4]); } catch (e) { return { ok: false, msg: '내용이 깨져 복구할 수 없습니다.' }; }
    if (!(arr instanceof Array)) return { ok: false, msg: '내용이 깨져 복구할 수 없습니다.' };

    var target = null;
    if (kind === '고객') target = buildCustomerSheet();
    else if (kind === '예약') target = sheetOf(SH.DATA);
    else if (kind === '매출') target = buildSlogSheet();
    else if (kind === '간단일정') target = buildBlockSheet();
    else if (kind === '수입지출') target = buildCashSheet();
    if (!target) return { ok: false, msg: '복구할 수 없는 종류입니다.' };

    target.appendRow(arr);
    sh.deleteRow(r);
    SpreadsheetApp.flush();
    return { ok: true, kind: kind };
  } catch (e) { logError('휴지통 복구', e); return { ok: false, msg: e.message }; }
}

/* ==================== 수입·지출 ==================== */

function buildCashSheet() {
  var sh = getOrCreateSheet(SH.CASH);
  if (headerMismatch(sh, CASH_HEAD)) {
    sh.clear();
    sh.getRange(1, 1, 1, CASH_HEAD.length).setValues([CASH_HEAD])
      .setFontWeight('bold').setBackground('#2F5D50').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    [90, 100, 70, 120, 240, 110, 110, 130, 240, 150].forEach(function (x, i) { sh.setColumnWidth(i + 1, x); });
  }
  return sh;
}

function CASH_ACCOUNTS() {
  var v = cfgStr('지출 계정 목록', '재료비,월세,관리비,공과금,마케팅비,광고비,수수료,인건비,장비,교육,기타');
  return String(v).split(',').map(function (x) { return x.trim(); }).filter(function (x) { return x; });
}

function readCash() {
  var sh = buildCashSheet();
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, CASH_HEAD.length).getValues()
    .filter(function (r) { return String(r[0] || '').trim(); })
    .map(function (r, i) {
      return { row: i + 2, id: String(r[0]), date: ymd(r[1]), kind: String(r[2] || ''),
               acct: String(r[3] || ''), note: String(r[4] || ''),
               amount: Number(r[5] || 0), pay: String(r[6] || ''),
               shop: String(r[7] || ''), memo: String(r[8] || '') };
    });
}

function api_cash(from, to) {
  try {
    var t = ymd(new Date());
    var f = ymd(toDate(from) || toDate(t.slice(0, 8) + '01'));
    var e = ymd(toDate(to) || toDate(t));
    var all = readCash().filter(function (g) { return g.date >= f && g.date <= e; });
    var inc = 0, exp = 0, byAcct = {};
    all.forEach(function (g) {
      if (g.kind === '수입') inc += g.amount;
      else {
        exp += g.amount;
        byAcct[g.acct || '기타'] = (byAcct[g.acct || '기타'] || 0) + g.amount;
      }
    });
    var cur = salesSum(readSlog(), f, e, '');
    var sales = cur.total, fee = cur.fee;
    all.sort(function (a, b) { return a.date < b.date ? 1 : (a.date > b.date ? -1 : 0); });
    return { ok: true, from: f, to: e, income: inc, expense: exp,
             sales: sales, fee: fee, feeRate: npayRate(),
             naverFee: cur.naverFee, cardFee: cur.cardFee,
             creditRate: creditRate(), debitRate: debitRate(), npayMixed: cur.npayMixed,
             net: sales - fee + inc - exp,
             byAcct: byAcct, list: all.slice(0, 200),
             accounts: CASH_ACCOUNTS(),
             payList: ['카드', '현금', '계좌이체', '자동이체', '기타'],
             shops: shops().map(function (x) { return x.name; }) };
  } catch (e) { logError('앱 수입지출', e); return { ok: false, msg: e.message }; }
}

function api_saveCash(p) {
  try {
    var date = ymd(toDate(p && p.date));
    if (!date) return { ok: false, msg: '날짜를 골라주세요.' };
    var amt = Number(p.amount || 0);
    if (!amt) return { ok: false, msg: '금액을 넣어주세요.' };
    var kind = (p.kind === '수입') ? '수입' : '지출';
    var sh = buildCashSheet();
    var id = 'X' + Date.now();
    sh.appendRow([id, toDate(date), kind, p.acct || '', p.note || '',
                  Math.abs(amt), p.pay || '', p.shop || '', p.memo || '', nowStr()]);
    SpreadsheetApp.flush();
    return { ok: true, id: id };
  } catch (e) { logError('앱 수입지출저장', e); return { ok: false, msg: e.message }; }
}

function api_deleteCash(id) {
  try {
    var sh = buildCashSheet();
    var last = sh.getLastRow();
    if (last < 2) return { ok: false, msg: '내역이 없습니다.' };
    var rows = sh.getRange(2, 1, last - 1, CASH_HEAD.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        toTrash('수입지출', id, rows[i][2] + ' ' + rows[i][4] + ' ' + rows[i][5], rows[i]);
        sh.deleteRow(i + 2);
        SpreadsheetApp.flush();
        return { ok: true };
      }
    }
    return { ok: false, msg: '내역을 찾지 못했습니다.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

/* ---------- 팅커뷰 고객 가져오기 (1회용) ---------- */

function tinkerviewRows() { return []; } /* 이관용 개인정보 데이터 — 공개 저장소라 비움 */
/* 팅커뷰 '담당자' 값(LIFE/YOUR/DRAW/메이크업클래스/닥뷰/웨딩)은 상품라인이라
   샵 판별에만 쓰고, 원래 값은 태그로 남깁니다. */
function lineToShop(line) {
  var sl = shops();
  function nm(i, d) { for (var k = 0; k < sl.length; k++) if (sl[k].idx === i) return sl[k].name; return d; }
  var v = String(line || '').trim();
  if (!v) return '';
  if (v === '닥뷰' || v.indexOf('닥터') >= 0 || v.indexOf('반영구') >= 0) return nm(2, '닥터 뷰티랩');
  return nm(1, '당신을 그리다');
}

function importTinkerview() {
  var sh = buildCustomerSheet();
  var existing = {};
  readCustomers().forEach(function (v) {
    var ph = onlyDigits(v[C.PHONE - 1]) || onlyDigits(v[C.SAFE - 1]);
    if (ph) existing[ph] = v[C.ID - 1];
    var nm = String(v[C.NAME - 1] || '').trim();
    if (nm) existing['N:' + nm] = v[C.ID - 1];
  });

  var maxId = 0;
  readCustomers().forEach(function (v) {
    var n = Number(String(v[C.ID - 1]).replace(/[^0-9]/g, ''));
    if (n > maxId) maxId = n;
  });

  var add = [], skip = 0;
  tinkerviewRows().forEach(function (r) {
    var name = r[0].trim(), phone = normPhone(r[1]);
    var key = onlyDigits(phone);
    if ((key && existing[key]) || (!key && existing['N:' + name])) { skip++; return; }
    if (key) existing[key] = 1;

    var line = r[2].trim();
    var tags = [];
    if (line) tags.push(line);
    if (r[9].trim()) tags.push(r[9].trim());

    var row = new Array(C_HEAD.length).fill('');
    maxId++;
    row[C.ID - 1]      = 'C' + ('0000' + maxId).slice(-4);
    row[C.NAME - 1]    = name;
    if (isSafePhone(phone)) row[C.SAFE - 1] = phone; else row[C.PHONE - 1] = phone;
    row[C.SHOP - 1]    = lineToShop(line);
    row[C.LAST - 1]    = r[7] ? toDate(r[7]) : '';
    row[C.VISITS - 1]  = Number(r[4] || 0);
    row[C.MEMO - 1]    = r[10] || '';
    row[C.SOURCE - 1]  = '팅커뷰';
    row[C.BIRTH - 1]   = r[8] ? toDate(r[8]) : '';
    row[C.PREPAID - 1] = Number(r[5] || 0);
    row[C.GRADE - 1]   = r[3].trim();
    row[C.TAG - 1]     = tags.join(', ');
    add.push(row);
  });

  if (add.length) {
    var startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, add.length, C_HEAD.length).setValues(add);
    for (var q = 0; q < add.length; q++) {
      var rr = startRow + q;
      sh.getRange(rr, C.LEFT).setFormula('=IF($J' + rr + '="","",$J' + rr + '-$K' + rr + ')');
    }
    SpreadsheetApp.flush();
  }
  return { ok: true, added: add.length, skipped: skip };
}

function A_팅커뷰_고객_가져오기() {
  var r = importTinkerview();
  Logger.log('가져옴 ' + r.added + '명 / 이미 있어서 건너뜀 ' + r.skipped + '명');
  return r;
}

/* ---------- 동의서 기본 문항 (팅커뷰 양식 그대로) ---------- */

function defaultForms() {
  var sl = shops();
  function nm(i, dflt) {
    for (var k = 0; k < sl.length; k++) if (sl[k].idx === i) return sl[k].name;
    return dflt;
  }
  var S1 = nm(1, '당신을 그리다'), S2 = nm(2, '닥터 뷰티랩');
  var F1 = '사전 상담지', F2 = '시술 동의서';
  var L = [];
  function q(shop, form, text, type, opts, need) {
    L.push({ shop: shop, form: form, q: text, type: type, opts: opts || [], need: !!need });
  }

  /* ===== 당신을 그리다 · 사전 상담지 ===== */
  q(S1, F1, '생년월일을 알려주세요. (예.1988년 12월 17일)', '단답', [], true);
  q(S1, F1, '주소를 알려주세요 (예. 서울시 강서구)', '단답', [], false);
  q(S1, F1, '해당 직업에 체크해주세요.', '단일선택',
    ['사무직', '서비스직', '판매직', '전문직', '관리직', '군인', '무직/취준생'], false);
  q(S1, F1, '방문 경로를 체크해주세요. (중복 가능)', '복수선택',
    ['네이버 플레이스', '네이버 지도', '네이버 프로그램', '인스타그램', '블로그', '지인추천',
     '유튜브(T1 오너선수,롯데백화점 등)'], false);
  q(S1, F1, '검색 키워드 (서울 퍼스널컬러, 메이크업 컨설팅 직접 검색한 단어)', '단답', [], false);
  q(S1, F1, '퍼스널컬러 진단/메이크업 컨설팅 경험', '복수선택',
    ['퍼스널컬러 진단 처음이에요.', '퍼스널컬러 재진단이에요.',
     '메이크업 수업 처음이에요.', '메이크업 수업 들어본 적 있어요.'], false);
  q(S1, F1, "진단 받은 시기와 결과를 알려주세요. 첫 진단일 경우 '없음'", '단답', [], false);
  q(S1, F1, '방문 목적이 무엇인가요? (예. 면접/웨딩 등 중요 일정, 이미지 변신)', '단답', [], false);
  q(S1, F1, '가장 고민되는 부분을 체크해 주세요. 솔루션을 제안해 드립니다.', '복수선택',
    ['피부톤에 딱 맞는 파운데이션 궁금해요', '색조 메이크업을 하면 촌스러워져요.',
     '내 얼굴에 어울리는 눈썹이 궁금해요.', '눈화장을 하면 답답해보여요.',
     '눈썹때문에 첫인상이 흐릿해보여요.', '이마 라인이 넓어서 얼굴이 커보여요.',
     '가르마정수리가 비어보여서 신경쓰여요.', '얼굴이 작아보이는 윤곽화장 궁금해요.',
     '나에게 어울리는 립 컬러 궁금해요.'], false);
  q(S1, F1, '이미지 보완을 위해 당신을그리다 반영구 센터를 연계해 드립니다.', '복수선택',
    ['※ 당일 예약 시 특별 혜택 적용', '퍼스널 브로우:자연스러운 눈썹 교정',
     '탈모커버:두피채움으로 얻는 동안이미지', '헤어라인교정:작아보이는 얼굴윤곽 교정',
     '퍼스널 립틴트:혈색있는 립 솔루션', '필요 없음'], false);
  q(S1, F1, '요청 사항이나 고민되는 부분이 있다면 상세히 적어주세요.', '서술', [], false);
  q(S1, F1, '\u26a0\ufe0f 컨설팅 보안 및 저작권 동의\n' +
    '[당신을 그리다]의 컨설팅은 고유의 노하우가 담긴 지적 재산입니다. ' +
    '원활한 서비스 제공을 위해 아래 사항에 대한 동의가 필요합니다.\n\n' +
    '1. 본 컨설팅 과정에 대한 무단 음성 녹음 및 영상 녹화를 엄격히 금지합니다.\n\n' +
    '2. 제공되는 진단 결과지 및 가이드 자료를 제3자에게 공유하거나 온라인상에 무단 배포할 수 없습니다.\n\n' +
    '3. 위 사항 위반 시 관련 법적 책임이 발생할 수 있음을 확인하였습니다.\n\n' +
    '본인은 위 내용을 충분히 숙지하였으며, 이에 동의합니다.', '동의', [], true);

  /* ===== 닥터 뷰티랩 · 시술 동의서 ===== */
  q(S2, F2, '생년월일을 작성해주세요. (예.1988년 12월 17일)', '단답', [], true);
  q(S2, F2, '방문 경로를 체크해주세요.', '복수선택',
    ['네이버 플레이스', '네이버 지도', '네이버 블로그', '인스타그램', '공식 홈페이지', '지인 소개'], false);
  q(S2, F2, '과거 눈썹 문신 경험이 있으신가요? (예. 총 2회, 리터치 포함 4회)', '단답', [], false);
  q(S2, F2, '민감성 피부 혹은 예민한 피부 타입이신가요?', '단일선택', ['예', '아니오'], false);
  q(S2, F2, '피부 질환이나 알레르기가 있으신가요? (켈로이드,묘기증,금속 알러지 등)', '복수선택',
    ['켈로이드', '아토피, 지루성 피부염, 건선', '헤르페스 (단순포진)',
     '간염 및 혈액 매개 감염 질환', '피부묘기증', '금속 알러지', '없음'], false);
  q(S2, F2, '현재 앓고 있는 질환이 있으신가요? (예. 당뇨, 고혈압, 임신, 없음)', '단답', [], false);
  q(S2, F2, '현재 복용 중인 약물이 있으신가요?', '복수선택',
    ['항응고제 / 아스피린', '여드름 치료제 (이소트레티노인)', '면역억제제 / 스테로이드',
     '오메가3, 고용량 비타민 E', '기타 (꼭 말씀해주세요!)'], false);
  q(S2, F2, '기타 특이사항이 있으면 알려주세요.', '서술', [], false);
  q(S2, F2, '시술 동의사항\n' +
    '1. 시술 과정 중 전후 사진 및 영상 촬영이 이루어질 수 있으며, 촬영된 자료는 우선적으로 시술 기록 ' +
    '보관을 목적으로 진행되며, 이를 통해 시술 전후의 상태를 정확히 관리하고 추후 관리에 도움을 ' +
    '드리고자 합니다. 필요시 SNS를 통한 업로드될 수 있으며, 이에 동의하지 않으실 경우 시술 전에 ' +
    '미리 말씀해 주시면 감사하겠습니다.\n\n' +
    '2. 본인은 상기 시술을 시행하기에 앞서 매우 드물게 특이 체질 등으로 예측할 수 없는 부작용과 ' +
    '후유증이 발생될 가능성이 있는 점을 염두에 두며 시술 관련 안내에 대해 구체적으로 충분히 ' +
    '안내받았음을 확인합니다. 본인이 지침 위반하여 발생하는 문제에 대해서 업체는 책임지지 않습니다. ' +
    '본인은 시술에 자발적 동의하며, 시술 결과에 대해 사업자의 객관적 중대한 책임 외 이의 제기하지 ' +
    '않음을 확인합니다.\n\n' +
    '3. 개인정보의 제공 및 활용에 동의합니다.\n' +
    '(수집한 개인 정보 항목 : 이름, 전화번호, 생년월일 등 기재된 사항)', '동의', [], true);

  return L;
}

function seedForms(force) {
  var sh = buildFormSheet();
  if (!force && sh.getLastRow() > 1) return 0;
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, FORM_HEAD.length).clearContent();
  var rows = defaultForms().map(function (f, i) {
    return [f.shop, f.form, i + 1, f.q, f.type, (f.opts || []).join(' | '), f.need ? 'Y' : ''];
  });
  sh.getRange(2, 1, rows.length, FORM_HEAD.length).setValues(rows);
  sh.getRange(2, 1, rows.length, FORM_HEAD.length).setVerticalAlignment('top').setWrap(true);
  SpreadsheetApp.flush();
  return rows.length;
}

function A_동의서문항_기본값으로_다시만들기() {
  var n = seedForms(true);
  Logger.log('동의서 문항 ' + n + '개를 넣었습니다.');
  return n;
}

function readForms() {
  var sh = buildFormSheet();
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, FORM_HEAD.length).getValues()
    .filter(function (r) { return String(r[3] || '').trim(); })
    .map(function (r) {
      return { shop: String(r[0] || '').trim(), form: String(r[1] || '').trim(),
               order: Number(r[2] || 0), q: String(r[3]).trim(),
               type: String(r[4] || '단답').trim(),
               opts: String(r[5] || '').split('|').map(function (x) { return x.trim(); })
                       .filter(function (x) { return x; }),
               need: String(r[6] || '') === 'Y' };
    })
    .sort(function (a, b) { return a.order - b.order; });
}

function api_forms(shop) {
  try {
    seedForms(false);
    var all = readForms();
    var list = shop ? all.filter(function (f) { return !f.shop || f.shop === shop; }) : all;
    return { ok: true, list: list, all: all,
             shops: shops().map(function (x) { return x.name; }),
             types: ['단일선택', '복수선택', '단답', '서술', '안내', '동의'] };
  } catch (e) { return { ok: false, msg: e.message, list: [], all: [] }; }
}

function api_saveForms(list) {
  try {
    var sh = buildFormSheet();
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, FORM_HEAD.length).clearContent();
    var rows = (list || []).filter(function (f) { return String(f.q || '').trim(); })
      .map(function (f, i) {
        return [String(f.shop || '').trim(), String(f.form || '상담지').trim(), i + 1,
                String(f.q).trim(), String(f.type || '단답').trim(),
                (f.opts || []).join(' | '), f.need ? 'Y' : ''];
      });
    if (rows.length) sh.getRange(2, 1, rows.length, FORM_HEAD.length).setValues(rows);
    SpreadsheetApp.flush();
    return { ok: true, n: rows.length };
  } catch (e) { logError('앱 양식저장', e); return { ok: false, msg: e.message }; }
}

function api_saveConsent(p) {
  try {
    var cid = String((p && p.custId) || '').trim();
    if (!cid) return { ok: false, msg: '고객을 찾지 못했습니다.' };
    var cust = api_customer(cid);
    if (!cust.ok) return { ok: false, msg: '고객을 찾지 못했습니다.' };
    var c = cust.customer;
    var url = '';
    if (p.sign) {
      var f = savePhotoToCustomer(cid, '동의서', p.sign, '서명_' + fmt(new Date(), 'yyyyMMdd_HHmmss') + '.png');
      url = f.url || '';
    }
    var lines = (p.answers || []).map(function (a) {
      return '· ' + a.q + ' → ' + (a.v === true ? '동의' : (a.v === false ? '동의안함' : String(a.v || '')));
    }).join('\n');
    buildConsentSheet().appendRow([nowStr(), cid, c.name, p.shop || c.shop || '',
                                   p.form || '상담지', lines, url]);
    SpreadsheetApp.flush();
    return { ok: true, url: url };
  } catch (e) { logError('앱 동의서저장', e); return { ok: false, msg: e.message }; }
}

function api_consents(custId) {
  try {
    var sh = buildConsentSheet();
    var last = sh.getLastRow();
    if (last < 2) return { ok: true, list: [] };
    var out = [];
    sh.getRange(2, 1, last - 1, CONSENT_HEAD.length).getValues().forEach(function (r) {
      if (String(r[1] || '').trim() !== String(custId).trim()) return;
      out.push({ at: String(r[0] || ''), shop: String(r[3] || ''), form: String(r[4] || ''),
                 body: String(r[5] || ''), url: String(r[6] || '') });
    });
    out.reverse();
    return { ok: true, list: out };
  } catch (e) { return { ok: false, msg: e.message, list: [] }; }
}

/* ---------- 사진 (시술 전·후) ---------- */

function savePhotoToCustomer(custId, kind, dataUrl, name) {
  var url = ensureCustomerFolder(String(custId));
  var id = String(url).match(/folders\/([^/?]+)/);
  if (!id) throw new Error('사진 폴더를 찾지 못했습니다.');
  var root = DriveApp.getFolderById(id[1]);
  var subs = root.getFoldersByName(kind);
  var sub = subs.hasNext() ? subs.next() : root.createFolder(kind);
  var m = String(dataUrl || '').match(/^data:([^;]+);base64,(.+)$/);
  if (!m) throw new Error('사진 형식을 읽지 못했습니다.');
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name || ('photo_' + Date.now()));
  var file = sub.createFile(blob);
  return { url: file.getUrl(), folder: sub.getUrl() };
}

function api_uploadPhoto(p) {
  try {
    if (!p || !p.custId || !p.data) return { ok: false, msg: '사진이 없습니다.' };
    var kind = (p.kind === '후') ? '시술후' : '시술전';
    var name = kind + '_' + fmt(new Date(), 'yyyyMMdd_HHmmss') + '.jpg';
    var r = savePhotoToCustomer(p.custId, kind, p.data, name);
    return { ok: true, url: r.url, folder: r.folder, kind: kind };
  } catch (e) { logError('앱 사진업로드', e); return { ok: false, msg: e.message }; }
}

/* ---------- 문자 발송 내역 ---------- */

function api_sendLog(limit) {
  try {
    var sh = sheetOf(SH.LOG);
    var last = sh.getLastRow();
    if (last < 2) return { ok: true, list: [], total: 0 };
    var n = Math.min(Number(limit) || 60, last - 1);
    var v = sh.getRange(last - n + 1, 1, n, L_HEAD.length).getValues();
    var out = v.map(function (r) {
      return { at: String(r[0] || ''), resno: String(r[1] || ''), shop: String(r[2] || ''),
               name: String(r[3] || ''), phone: String(r[4] || ''), type: String(r[5] || ''),
               ch: String(r[6] || ''), result: String(r[7] || ''),
               msg: String(r[8] || '').slice(0, 400) };
    });
    out.reverse();
    return { ok: true, list: out, total: last - 1 };
  } catch (e) { return { ok: false, msg: e.message, list: [], total: 0 }; }
}

/* ---------- 매출 분석 ---------- */

function dayDiff(a, b) { return Math.round((toDate(b) - toDate(a)) / 86400000); }
function shiftDate(d, days) {
  var t = toDate(d); t.setDate(t.getDate() + days); return ymd(t);
}

/* ---------- 네이버 결제 수수료 ---------- */

function rateOf(key, dflt) {
  var n = parseFloat(String(cfg(key, dflt)).replace(/[^0-9.]/g, ''));
  return (isNaN(n) || n < 0) ? 0 : n;
}
/* 네이버 정산 수수료율(%). 일반예약상품 1.98 / 주문상품 0.88 (VAT 포함) */
function npayRate() { return rateOf('네이버 수수료율(%)', '1.98'); }
/* 카드 수수료율(%). 연매출 3억 이하 우대: 신용 0.40 / 체크 0.15 */
function creditRate() { return rateOf('신용카드 수수료율(%)', '0.40'); }
function debitRate()  { return rateOf('체크카드 수수료율(%)', '0.15'); }
/* 해외 발급 카드는 우대요율 대상이 아니라 따로 받습니다 (하나카드 기준 3.6) */
function foreignRate() { return rateOf('해외카드 수수료율(%)', '3.6'); }
function cardRateOf(pay) {
  var p = String(pay || '');
  if (p.indexOf('카드') < 0) return 0;
  if (p.indexOf('해외') >= 0) return foreignRate();
  return (p.indexOf('체크') >= 0) ? debitRate() : creditRate();
}
function npayFee(base) { return Math.round(Number(base || 0) * npayRate() / 100); }

/* 예약 단위로 수수료를 계산합니다.
   - 네이버 단독: 전액에 네이버 요율
   - 카드 단독: 전액에 카드 요율 (체크면 체크 요율)
   - 섞인 결제: 네이버결제액을 적었으면 나눠서, 안 적었으면 빼고 mixed 로 알려줍니다 */
function feeOf(byRes) {
  var r = { naverBase: 0, cardBase: 0, naverFee: 0, cardFee: 0, fee: 0, mixed: 0 };
  Object.keys(byRes).forEach(function (k) {
    var b = byRes[k];
    var amt = Number(b.amount || 0), np = Number(b.npay || 0);
    var pay = String(b.pay || '');
    var hasN = pay.indexOf('네이버') >= 0, hasC = pay.indexOf('카드') >= 0;
    var cr = cardRateOf(pay);
    if (hasN && hasC) {
      if (np <= 0) { r.mixed++; return; }
      var cardPart = Math.max(0, amt - np);
      r.naverBase += np;  r.naverFee += Math.round(np * npayRate() / 100);
      r.cardBase  += cardPart; r.cardFee += Math.round(cardPart * cr / 100);
    } else if (hasN) {
      var nb = np > 0 ? np : amt;
      r.naverBase += nb; r.naverFee += Math.round(nb * npayRate() / 100);
    } else if (hasC) {
      r.cardBase += amt; r.cardFee += Math.round(amt * cr / 100);
    }
  });
  r.fee = r.naverFee + r.cardFee;
  return r;
}

function salesSum(rows, from, to, shop) {
  var res = { total: 0, count: 0, cust: {}, byShop: {}, byItem: {}, byDow: [0,0,0,0,0,0,0],
              byPay: {}, byDay: {}, list: [] };
  var byRes = {};
  rows.forEach(function (g) {
    if (!g.date || g.date < from || g.date > to) return;
    if (shop && g.shop !== shop) return;
    res.total += g.amount;
    res.byShop[g.shop] = (res.byShop[g.shop] || 0) + g.amount;
    var ik = g.shop + '|' + g.item;
    if (!res.byItem[ik]) res.byItem[ik] = { shop: g.shop, name: g.item, total: 0, n: 0 };
    res.byItem[ik].total += g.amount;
    if (g.item !== '할인') res.byItem[ik].n++;
    var dt = toDate(g.date);
    if (dt) res.byDow[dt.getDay()] += g.amount;
    res.byDay[g.date] = (res.byDay[g.date] || 0) + g.amount;
    var pk = g.pay || '미입력';
    var pparts = payParts_(pk);
    if (pparts && g.item !== '할인') {
      /* 나눠 받은 결제는 첫 줄에서 수단별로 한 번만 더합니다 */
      if (!res.__payDone) res.__payDone = {};
      if (!res.__payDone[g.resno]) {
        res.__payDone[g.resno] = true;
        pparts.forEach(function (x) { res.byPay[x.m] = (res.byPay[x.m] || 0) + x.amt; });
      }
    } else if (!pparts) {
      res.byPay[pk] = (res.byPay[pk] || 0) + g.amount;
    }
    if (g.custId) res.cust[g.custId] = (res.cust[g.custId] || 0) + 1;
    if (!byRes[g.resno]) {
      byRes[g.resno] = { resno: g.resno, date: g.date, shop: g.shop, name: g.name,
                         staff: g.staff, pay: g.pay, items: [], amount: 0, npay: 0 };
      res.list.push(byRes[g.resno]);
    }
    if (g.item !== '할인') byRes[g.resno].items.push(g.item);
    byRes[g.resno].amount += g.amount;
    if (Number(g.npay) > Number(byRes[g.resno].npay)) byRes[g.resno].npay = Number(g.npay);
  });
  var fo = feeOf(byRes);
  res.npayBase = fo.naverBase;
  res.cardBase = fo.cardBase;
  res.naverFee = fo.naverFee;
  res.cardFee = fo.cardFee;
  res.npayMixed = fo.mixed;
  res.fee = fo.fee;
  res.net = res.total - res.fee;
  delete res.__payDone;
  res.count = res.list.length;
  res.list.sort(function (a, b) { return a.date < b.date ? 1 : (a.date > b.date ? -1 : 0); });
  return res;
}

function api_salesReport(from, to, shop) {
  try {
    var t = ymd(new Date());
    from = ymd(toDate(from) || toDate(t.slice(0, 8) + '01'));
    to = ymd(toDate(to) || toDate(t));
    var rows = readSlog();
    var cur = salesSum(rows, from, to, shop);

    var span = Math.max(dayDiff(from, to) + 1, 1);
    var pFrom = shiftDate(from, -span), pTo = shiftDate(from, -1);
    var prev = salesSum(rows, pFrom, pTo, shop);

    // 재방문: 이 기간 이전에도 결제 기록이 있는 고객
    var before = {};
    rows.forEach(function (g) {
      if (g.date && g.date < from && g.custId && (!shop || g.shop === shop)) before[g.custId] = true;
    });
    var ids = Object.keys(cur.cust);
    var repeat = ids.filter(function (id) { return before[id]; }).length;

    var items = Object.keys(cur.byItem).map(function (k) { return cur.byItem[k]; })
      .sort(function (a, b) { return b.total - a.total; });

    return {
      ok: true, from: from, to: to, shop: shop || '',
      prevFrom: pFrom, prevTo: pTo,
      total: cur.total, prevTotal: prev.total,
      fee: cur.fee, net: cur.net, feeRate: npayRate(),
      naverFee: cur.naverFee, cardFee: cur.cardFee,
      creditRate: creditRate(), debitRate: debitRate(),
      npayBase: cur.npayBase, cardBase: cur.cardBase, npayMixed: cur.npayMixed,
      prevFee: prev.fee, prevNet: prev.net,
      count: cur.count, prevCount: prev.count,
      customers: ids.length, prevCustomers: Object.keys(prev.cust).length,
      repeat: repeat,
      avg: ids.length ? Math.round(cur.total / ids.length) : 0,
      byShop: cur.byShop, byItem: items, byDow: cur.byDow, byPay: cur.byPay,
      byDay: cur.byDay,
      list: cur.list.slice(0, 200),
      shops: shops().map(function (x) { return x.name; })
    };
  } catch (e) { logError('앱 매출분석', e); return { ok: false, msg: e.message }; }
}

/* ---------- 도구 ---------- */

function api_checkMail() {
  var s = processBookingMails();
  refreshHome();
  return { ok: true, stat: s, lastMail: cfgStr('마지막 메일확인 시각', '') };
}

function api_send(resnos, type) {
  var sh = sheetOf(SH.DATA);
  var idx = buildResnoIndex();
  var ok = 0, ng = 0, msgs = [];
  (resnos || []).forEach(function (no) {
    var t = idx[String(no)];
    if (!t) { ng++; return; }
    var r = sendForRow(sh, t.r, type || '개별', true);
    if (r.ok) ok++; else { ng++; if (r.reason) msgs.push(no + ' · ' + r.reason); }
  });
  return { ok: true, sent: ok, failed: ng, reasons: msgs.slice(0, 8), testMode: cfgOn('테스트 모드') };
}

function api_completePast() {
  var sh = sheetOf(SH.DATA);
  var last = sh.getLastRow();
  if (last < 2) return { ok: true, n: 0 };
  var rows = sh.getRange(2, 1, last - 1, D_HEAD.length).getValues();
  var today = toDate(ymd(new Date()));
  var n = 0;
  for (var i = 0; i < rows.length; i++) {
    var dt = toDate(rows[i][D.DATE - 1]);
    if (rows[i][D.STATUS - 1] === STATUS.OK && dt && dt < today) {
      sh.getRange(i + 2, D.STATUS).setValue(STATUS.DONE);
      n++;
    }
  }
  recomputeCustomerStats();
  return { ok: true, n: n };
}

function api_needAttention() {
  var rows = readData().map(rowToObj);
  var t0 = ymd(new Date());
  return {
    ok: true,
    check: rows.filter(function (o) { return o.status === STATUS.CHECK; }),
    noPhone: rows.filter(function (o) {
      return o.status === STATUS.OK && !o.phone && !o.safe && o.date >= t0; }),
    noPaid: (function () {
      var paidSet = {};
      readSlog().forEach(function (g) { paidSet[g.resno] = true; });
      return rows.filter(function (o) {
        return o.status === STATUS.DONE && !paidSet[o.resno]; });
    })(),
    pastOpen: rows.filter(function (o) {
      return o.status === STATUS.OK && o.date && o.date < t0; })
  };
}


/* ---------- 유지보수: 모든 예약/고객 데이터 지우고 메일에서 다시 가져오기 ---------- */
function resetAllData() {
  var sh = sheetOf(SH.DATA);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, D_HEAD.length).clearContent();
  var c = sheetOf(SH.CUST);
  if (c.getLastRow() > 1) c.getRange(2, 1, c.getLastRow() - 1, C_HEAD.length).clearContent();
  var lg = sheetOf(SH.LOG);
  if (lg.getLastRow() > 1) lg.getRange(2, 1, lg.getLastRow() - 1, L_HEAD.length).clearContent();
  var labName = cfgStr('처리완료 라벨', 'CRM처리완료');
  var lab = GmailApp.getUserLabelByName(labName);
  if (lab) {
    var th = lab.getThreads(0, 200);
    for (var i = 0; i < th.length; i++) { try { th[i].removeLabel(lab); } catch (e) {} }
  }
  shops().forEach(function (s) {
    var cal = calFor(s.name);
    if (!cal) return;
    var evs = cal.getEvents(new Date(Date.now() - 120 * 86400000), new Date(Date.now() + 365 * 86400000));
    evs.forEach(function (e) { try { e.deleteEvent(); } catch (x) {} });
  });
  shops().forEach(function (s) {
    refreshCalendarMonth(s, new Date());
    refreshCalendarMonth(s, nextMonthDate(new Date()));
  });
  refreshHome();
}

function resetAndReimport() {
  resetAllData();
  var st = processBookingMails();
  recomputeCustomerStats();
  refreshHome();
  Logger.log(JSON.stringify(st));
  return st;
}


/* ---------- 이미 들어온 예약의 요청사항을 다시 훑어 실명·연락처 채우기 ---------- */

function updateCustomerIdentity(custId, realName, phone) {
  if (!custId) return;
  var sh = sheetOf(SH.CUST);
  var rows = readCustomers();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][C.ID - 1]).trim() !== String(custId).trim()) continue;
    if (realName && String(rows[i][C.NAME - 1]).indexOf('*') >= 0) {
      sh.getRange(i + 2, C.NAME).setValue(realName);
    }
    if (phone) {
      if (isSafePhone(phone)) {
        if (!onlyDigits(rows[i][C.SAFE - 1])) sh.getRange(i + 2, C.SAFE).setValue(phone);
      } else if (!onlyDigits(rows[i][C.PHONE - 1])) {
        sh.getRange(i + 2, C.PHONE).setValue(phone);
      }
    }
    return;
  }
}

function rescanRequests() {
  var sh = sheetOf(SH.DATA);
  var last = sh.getLastRow();
  if (last < 2) return { ok: true, n: 0 };
  var v = sh.getRange(2, 1, last - 1, D_HEAD.length).getValues();
  var n = 0;
  var touched = {};
  for (var i = 0; i < v.length; i++) {
    var ex = extractFromRequest(v[i][D.REQUEST - 1], v[i][D.NAME - 1]);
    var changed = false;
    if (ex.phone && !v[i][D.PHONE - 1] && !v[i][D.SAFE - 1]) {
      if (isSafePhone(ex.phone)) sh.getRange(i + 2, D.SAFE).setValue(ex.phone);
      else sh.getRange(i + 2, D.PHONE).setValue(ex.phone);
      changed = true;
    }
    if (ex.name && String(v[i][D.NAME - 1]).indexOf('*') >= 0) {
      sh.getRange(i + 2, D.NAME).setValue(ex.name);
      changed = true;
    }
    if (changed) {
      updateCustomerIdentity(v[i][D.CUSTID - 1], ex.name, ex.phone);
      touched[v[i][D.SHOP - 1] + '|' + ymd(v[i][D.DATE - 1])] = true;
      n++;
    }
  }
  if (n) {
    SpreadsheetApp.flush();
    Object.keys(touched).forEach(function (k) {
      var p = k.split('|');
      if (p[1]) syncCalendarForBooking(p[0], p[1]);
    });
    recomputeCustomerStats();
    refreshHome();
  }
  return { ok: true, n: n };
}

function api_rescanRequests() {
  try { return rescanRequests(); }
  catch (e) { logError('요청사항 재검사', e); return { ok: false, msg: e.message }; }
}

/* 스프레드시트 통째로 백업 사본 만들기 */
function A_백업본_만들기() {
  var f = DriveApp.getFileById(SHEET_ID);
  var name = f.getName() + ' (백업 ' + fmt(new Date(), 'yyyy-MM-dd HH:mm') + ')';
  var copy = f.makeCopy(name);
  Logger.log('백업 완료: ' + name + '\n' + copy.getUrl());
  return copy.getUrl();
}

/* ==================== 팅커뷰 과거기록 가져오기 (1회용) ==================== */
/* 고객 정보: [이름, 연락처, 최근방문일, 생년월일, 메모] */
function tvCustRows() { return []; } /* 이관용 개인정보 데이터 — 공개 저장소라 비움 */
/* 매출/예약: [날짜, 시각, 담당라인, 고객명, 시술명, 결제수단, 금액, 시술메모] */
function tvSaleRows() { return []; } /* 이관용 개인정보 데이터 — 공개 저장소라 비움 */
function importTinkerviewHistory() {
  var custSheet = buildCustomerSheet();
  var dataSheet = sheetOf(SH.DATA);
  var slogSheet = buildSlogSheet();

  /* ---- 기존 고객 색인 (연락처 숫자 / 이름) ---- */
  var rows = readCustomers();
  var byPhone = {}, byName = {};
  for (var i = 0; i < rows.length; i++) {
    var p = String(rows[i][C.PHONE - 1] || '').replace(/[^0-9]/g, '');
    var nm = String(rows[i][C.NAME - 1] || '').trim();
    if (p) byPhone[p] = i + 2;                 // 시트 행번호
    if (nm && !byName[nm]) byName[nm] = i + 2;
  }
  function rowOf(name, phone) {
    var p = String(phone || '').replace(/[^0-9]/g, '');
    if (p && byPhone[p]) return byPhone[p];
    var nm = String(name || '').trim();
    return byName[nm] || 0;
  }
  function custIdAt(r) {
    return r ? String(custSheet.getRange(r, C.ID).getValue() || '').trim() : '';
  }

  /* ---- 1) 고객 정보 채우기 (빈 칸만, 메모는 덧붙임) ---- */
  var cFilled = 0, cMemo = 0;
  tvCustRows().forEach(function (t) {
    var r = rowOf(t[0], t[1]);
    if (!r) return;
    if (t[2]) {
      var cur = custSheet.getRange(r, C.LAST).getValue();
      if (!cur) { custSheet.getRange(r, C.LAST).setValue(t[2]); cFilled++; }
    }
    if (t[3]) {
      var b = custSheet.getRange(r, C.BIRTH).getValue();
      if (!b) { custSheet.getRange(r, C.BIRTH).setValue(t[3]); cFilled++; }
    }
    if (t[4]) {
      var m = String(custSheet.getRange(r, C.MEMO).getValue() || '');
      if (m.indexOf(t[4]) < 0) {
        custSheet.getRange(r, C.MEMO).setValue(m ? (m + '\n' + t[4]) : t[4]);
        cMemo++;
      }
    }
  });

  /* ---- 2) 매출 + 예약 기록 ---- */
  var haveRes = {};
  var last = dataSheet.getLastRow();
  if (last > 1) {
    dataSheet.getRange(2, D.RESNO, last - 1, 1).getValues()
      .forEach(function (v) { haveRes[String(v[0] || '').trim()] = 1; });
  }
  var haveSlog = {};
  var sl = slogSheet.getLastRow();
  if (sl > 1) {
    slogSheet.getRange(2, 2, sl - 1, 1).getValues()
      .forEach(function (v) { haveSlog[String(v[0] || '').trim()] = 1; });
  }

  var now = fmt(new Date(), 'yyyy-MM-dd HH:mm:ss');
  var addRes = [], addSlog = [], seq = {}, sMemo = 0;

  tvSaleRows().forEach(function (t) {
    var date = t[0], time = t[1], line = t[2], name = t[3];
    var item = t[4], pay = t[5], amt = Number(t[6] || 0), memo = t[7];
    var shop = lineToShop(line);
    var key = date.replace(/-/g, '');
    seq[key] = (seq[key] || 0) + 1;
    var resno = 'TV' + key + '-' + seq[key];

    var r = rowOf(name, '');
    var cid = custIdAt(r);

    if (!haveRes[resno]) {
      var d = new Array(D_HEAD.length).fill('');
      d[D.RESNO - 1] = resno;
      d[D.SHOP - 1] = shop;
      d[D.STATUS - 1] = STATUS.DONE;
      d[D.NAME - 1] = name;
      d[D.DATE - 1] = date;
      d[D.TIME - 1] = time;
      d[D.PRODUCT - 1] = item;
      d[D.AMOUNT - 1] = amt;
      d[D.PAID - 1] = amt;
      d[D.STAFF - 1] = line;
      d[D.SOURCE - 1] = '팅커뷰';
      d[D.CUSTID - 1] = cid;
      d[D.PROCAT - 1] = now;
      d[D.MEMO - 1] = memo;
      addRes.push(d);
    }
    if (!haveSlog[resno]) {
      addSlog.push([now, resno, shop, cid, name, date, item, amt, pay, line, '']);
    }
    /* 시술메모 → 고객메모에 날짜와 함께 덧붙임 */
    if (memo && r) {
      var tag = '[' + date + ' 시술] ' + memo;
      var m = String(custSheet.getRange(r, C.MEMO).getValue() || '');
      if (m.indexOf(memo) < 0) {
        custSheet.getRange(r, C.MEMO).setValue(m ? (m + '\n' + tag) : tag);
        sMemo++;
      }
    }
  });

  if (addRes.length) {
    dataSheet.getRange(dataSheet.getLastRow() + 1, 1, addRes.length, D_HEAD.length).setValues(addRes);
  }
  if (addSlog.length) {
    slogSheet.getRange(slogSheet.getLastRow() + 1, 1, addSlog.length, SLOG_HEAD.length).setValues(addSlog);
  }

  return { ok: true, 고객채움: cFilled, 고객메모: cMemo, 시술메모: sMemo,
           예약추가: addRes.length, 매출추가: addSlog.length };
}

function A_팅커뷰_과거기록_가져오기() {
  var r = importTinkerviewHistory();
  Logger.log('예약 ' + r.예약추가 + '건 / 매출 ' + r.매출추가 + '건 / 고객정보 ' + r.고객채움 +
             '칸 / 고객메모 ' + r.고객메모 + '건 / 시술메모 ' + r.시술메모 + '건 넣었습니다.');
  return r;
}

/* ===================== 토스 사장님 과거기록 가져오기 ===================== */

/* 고객: [이름, 연락처, 등급, 고객메모, 포인트] */
function tossCustRows() { return []; } /* 이관용 개인정보 데이터 — 공개 저장소라 비움 */
/* 예약: [날짜, 시각, 고객명, 연락처, 시술명, 금액, 상태, 예약메모] */
function tossResRows() { return []; } /* 이관용 개인정보 데이터 — 공개 저장소라 비움 */
function importTossHistory() {
  var custSheet = buildCustomerSheet();
  var dataSheet = sheetOf(SH.DATA);
  var slogSheet = buildSlogSheet();
  var shop = lineToShop('닥뷰');
  var STAFF = '제이 대표원장';

  /* 상태 옮기기: 토스 → 이 앱 */
  function toStatus(s) {
    if (s === '방문완료') return STATUS.DONE;
    if (s === '예정') return STATUS.OK;
    return STATUS.CANCEL;            /* 방문미완료, 예약취소노쇼, 예약삭제 */
  }

  /* ---- 1) 고객 (없으면 만들고, 빈 칸만 채웁니다) ---- */
  var cid = {}, cRow = {}, cNew = 0, cFilled = 0, cMemo = 0;
  tossCustRows().forEach(function (t) {
    var name = t[0], phone = t[1], grade = t[2], memo = t[3], point = t[4];
    var before = readCustomers().length;
    var id = matchOrCreateCustomer(name, phone, '', shop, '토스');
    if (!id) return;
    if (readCustomers().length > before) cNew++;

    var rows = readCustomers(), r = 0;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][C.ID - 1]).trim() === String(id).trim()) { r = i + 2; break; }
    }
    if (!r) return;
    cid[phone] = id; cRow[phone] = r;

    if (grade && !String(custSheet.getRange(r, C.GRADE).getValue() || '').trim()) {
      custSheet.getRange(r, C.GRADE).setValue(grade); cFilled++;
    }
    if (point && !Number(custSheet.getRange(r, C.POINT).getValue() || 0)) {
      custSheet.getRange(r, C.POINT).setValue(Number(point)); cFilled++;
    }
    if (memo) {
      var m = String(custSheet.getRange(r, C.MEMO).getValue() || '');
      if (m.indexOf(memo) < 0) {
        custSheet.getRange(r, C.MEMO).setValue(m ? (m + '\n' + memo) : memo);
        cMemo++;
      }
    }
  });

  /* ---- 2) 예약 + 매출 ---- */
  var haveRes = {};
  var last = dataSheet.getLastRow();
  if (last > 1) {
    dataSheet.getRange(2, D.RESNO, last - 1, 1).getValues()
      .forEach(function (v) { haveRes[String(v[0] || '').trim()] = 1; });
  }
  var haveSlog = {};
  var sl = slogSheet.getLastRow();
  if (sl > 1) {
    slogSheet.getRange(2, 2, sl - 1, 1).getValues()
      .forEach(function (v) { haveSlog[String(v[0] || '').trim()] = 1; });
  }

  var now = fmt(new Date(), 'yyyy-MM-dd HH:mm:ss');
  var addRes = [], addSlog = [], seq = {}, sMemo = 0;

  tossResRows().forEach(function (t) {
    var date = t[0], time = t[1], name = t[2], phone = t[3];
    var item = t[4], amt = Number(t[5] || 0), state = t[6], memo = t[7];
    var st = toStatus(state);
    var key = date.replace(/-/g, '');
    seq[key] = (seq[key] || 0) + 1;
    var resno = 'TS' + key + '-' + seq[key];
    var id = cid[phone] || '';

    if (!haveRes[resno]) {
      var d = new Array(D_HEAD.length).fill('');
      d[D.RESNO - 1]   = resno;
      d[D.SHOP - 1]    = shop;
      d[D.STATUS - 1]  = st;
      d[D.NAME - 1]    = name;
      d[D.PHONE - 1]   = phone;
      d[D.DATE - 1]    = date;
      d[D.TIME - 1]    = time;
      d[D.PRODUCT - 1] = item;
      d[D.AMOUNT - 1]  = amt;
      d[D.PAID - 1]    = (st === STATUS.DONE) ? amt : 0;
      d[D.STAFF - 1]   = STAFF;
      d[D.SOURCE - 1]  = '토스';
      d[D.CUSTID - 1]  = id;
      d[D.PROCAT - 1]  = now;
      d[D.MEMO - 1]    = memo;
      addRes.push(d);
    }
    /* 매출은 실제 방문한 건만 */
    if (st === STATUS.DONE && !haveSlog[resno]) {
      addSlog.push([now, resno, shop, id, name, date, item, amt, '', STAFF, '']);
    }
    /* 예약메모 → 고객메모에 날짜와 함께 덧붙임 */
    if (memo && cRow[phone]) {
      var tag = '[' + date + ' 시술] ' + memo;
      var m = String(custSheet.getRange(cRow[phone], C.MEMO).getValue() || '');
      if (m.indexOf(memo) < 0) {
        custSheet.getRange(cRow[phone], C.MEMO).setValue(m ? (m + '\n' + tag) : tag);
        sMemo++;
      }
    }
  });

  if (addRes.length) {
    dataSheet.getRange(dataSheet.getLastRow() + 1, 1, addRes.length, D_HEAD.length).setValues(addRes);
  }
  if (addSlog.length) {
    slogSheet.getRange(slogSheet.getLastRow() + 1, 1, addSlog.length, SLOG_HEAD.length).setValues(addSlog);
  }

  recomputeCustomerStats();

  return { ok: true, 고객신규: cNew, 고객채움: cFilled, 고객메모: cMemo,
           시술메모: sMemo, 예약추가: addRes.length, 매출추가: addSlog.length };
}

function A_토스_과거기록_가져오기() {
  var r = importTossHistory();
  Logger.log('신규고객 ' + r.고객신규 + '명 / 예약 ' + r.예약추가 + '건 / 매출 ' + r.매출추가 +
             '건 / 고객정보 ' + r.고객채움 + '칸 / 고객메모 ' + r.고객메모 +
             '건 / 시술메모 ' + r.시술메모 + '건 넣었습니다.');
  return r;
}

/* ===================== 지출·수수료 설정 업데이트 (1회용) ===================== */

function A_지출_수수료_설정_업데이트() {
  /* 1) 설정 시트에 새 항목 추가 (기존 값은 그대로 둡니다) */
  buildConfigSheet();
  _cfgCache = null;

  /* 2) 지출 계정 목록에 빠진 계정만 덧붙입니다 */
  var want = ['재료비', '월세', '관리비', '공과금', '마케팅비', '광고비',
              '수수료', '인건비', '장비', '교육', '기타'];
  var cur = CASH_ACCOUNTS();
  var added = [];
  want.forEach(function (w) { if (cur.indexOf(w) < 0) { cur.push(w); added.push(w); } });
  if (added.length) setCfg('지출 계정 목록', cur.join(','));

  /* 3) 매출내역 시트에 네이버결제액 칸 만들기 (기존 기록은 건드리지 않습니다) */
  buildSlogSheet();

  _cfgCache = null;
  var msg = '계정 추가: ' + (added.length ? added.join(', ') : '없음') +
            ' / 네이버 ' + npayRate() + '% / 신용 ' + creditRate() + '% / 체크 ' + debitRate() + '%';
  Logger.log(msg);
  return { ok: true, added: added, naver: npayRate(), credit: creditRate(), debit: debitRate() };
}


/* ===================== 사전 안내(3일 전) 설정 (1회용) ===================== */

function A_사전안내_설정_업데이트() {
  /* 1) 설정 시트에 사전 안내 항목 추가 (기존 값은 그대로 둡니다) */
  buildConfigSheet();
  _cfgCache = null;

  /* 2) 닥터 뷰티랩 전일 문구가 예전 긴 안내면 짧은 리마인더로 교체합니다.
        (긴 안내는 '샵2 문구_사전'으로 옮겨졌습니다) */
  var cur = String(cfg('샵2 문구_전일', ''));
  var swapped = false;
  if (cur.length > 400) {
    setCfg('샵2 문구_전일', "#{고객명}님, 내일 뵙겠습니다.\n\n▶ 일시 : #{예약일시}\n▶ 시술 : #{상품}\n\n· 오늘 음주는 피해주세요.\n· 내일 카페인은 줄여주세요.\n· 두피 시술은 머리를 감고 와주세요.\n· 색소가 묻어날 수 있어 밝은 색이나 고가의 의상은 피해주세요.\n· 건물 지하 주차장을 2시간 무료로 이용하실 수 있습니다.\n\n15분 이상 늦으시면 다음 예약 일정으로 인해 자동 취소됩니다. 늦으실 것 같으면 미리 전화 주세요. 010-2722-3093");
    swapped = true;
  }
  _cfgCache = null;

  /* 3) 자동실행 시간표를 다시 깔아 사전 안내 트리거를 만듭니다 */
  installTriggers();

  var msg = '사전 안내 ' + advanceDays() + '일 전 / ' + cfgStr('사전 알림 시각', '') +
            ' 발송 · 닥뷰 전일 문구 교체: ' + (swapped ? '했음' : '안 함(이미 짧거나 직접 수정하신 상태)');
  Logger.log(msg);
  return { ok: true, days: advanceDays(), swapped: swapped };
}


/* ===================== 토스에서 온 '방문 미완료' 정리 (1회용) ===================== */

function A_토스_미방문_정리() {
  var sh = sheetOf(SH.DATA);
  var last = sh.getLastRow();
  if (last < 2) return { ok: true, fixed: 0 };
  var v = sh.getRange(2, 1, last - 1, D_HEAD.length).getValues();
  var fixed = [];
  for (var i = 0; i < v.length; i++) {
    var no = String(v[i][D.RESNO - 1] || '');
    if (no.indexOf('TS') !== 0) continue;              /* 토스에서 넘어온 기록만 */
    if (v[i][D.STATUS - 1] !== STATUS.CHECK) continue;
    sh.getRange(i + 2, D.STATUS).setValue(STATUS.CANCEL);
    fixed.push(no + ' ' + v[i][D.NAME - 1]);
  }
  if (fixed.length) { SpreadsheetApp.flush(); recomputeCustomerStats(); refreshHome(); }
  Logger.log('정리한 예약: ' + (fixed.length ? fixed.join(', ') : '없음'));
  return { ok: true, fixed: fixed.length, list: fixed };
}

/* ---------- 자동 정산 (2026-11-15) ----------
   이용 시간이 끝난 '확정' 예약을 자동으로 '이용완료'로 바꿉니다.
   - 당신을그리다(이름에 '그리다'가 들어간 샵): 예약 메일의 메뉴 금액으로 매출 자동 등록,
     결제수단 네이버페이(수수료는 매출 화면에서 자동 반영)
   - 닥터뷰티랩 등 그 밖의 샵: 이용완료만 하고 금액은 사장님이 직접 입력
     (홈 화면 '금액 입력 대기'에 모입니다)
   - 취소·삭제는 건드리지 않습니다. 이미 매출이 등록된 예약은 다시 등록하지 않습니다.
   - 메일 확인(5분마다) 때 같이 돌지만 실제 처리는 15분에 한 번만 합니다. */
function isAutoSalesShop(shopName) {
  var key = cfgStr('자동매출 샵 키워드', '그리다');
  return !!key && String(shopName || '').replace(/\s+/g, '').indexOf(key.replace(/\s+/g, '')) >= 0;
}

function menuNameFor(v) {
  var nosp = function (x) { return String(x || '').replace(/\s+/g, ''); };
  var hay = nosp(String(v[D.OPTION - 1] || '') + '|' + String(v[D.PRODUCT - 1] || ''));
  var shop = String(v[D.SHOP - 1] || '');
  var best = '', bestLen = 0;
  try {
    readMenu().forEach(function (m) {
      if (m.shop && m.shop !== shop) return;
      var key = nosp(m.name);
      if (key && hay.indexOf(key) >= 0 && key.length > bestLen) { bestLen = key.length; best = m.name; }
    });
  } catch (e) {}
  return best || String(v[D.PRODUCT - 1] || '').trim() || '시술';
}

function autoSettle(force) {
  var props = PropertiesService.getScriptProperties();
  var lastRun = Number(props.getProperty('AUTO_SETTLE_AT') || 0);
  if (!force && Date.now() - lastRun < 15 * 60 * 1000) return { skipped: true };
  props.setProperty('AUTO_SETTLE_AT', String(Date.now()));

  var dataSh = sheetOf(SH.DATA);
  var last = dataSh.getLastRow();
  if (last < 2) return { done: 0, sales: 0 };
  var rows = dataSh.getRange(2, 1, last - 1, D_HEAD.length).getValues();
  var now = new Date();
  // 이 날짜 이전의 예약은 손대지 않습니다 (예전에 밀린 건이 한꺼번에 매출로 잡히는 것 방지)
  var since = cfgStr('자동정산 시작일', '2026-11-15');

  var paidSet = {};
  readSlog().forEach(function (g) { paidSet[g.resno] = true; });

  var slogSh = null, out = [], stamp = nowStr(), done = 0;
  for (var i = 0; i < rows.length; i++) {
    var v = rows[i];
    if (v[D.STATUS - 1] !== STATUS.OK) continue;
    var d0 = toDate(v[D.DATE - 1]);
    if (!d0 || ymd(d0) < since) continue;
    var start = new Date(d0.getTime());
    var t = timeStr(v[D.TIME - 1]).match(/(\d{1,2}):(\d{2})/);
    if (t) start.setHours(Number(t[1]), Number(t[2]), 0, 0); else start.setHours(23, 59, 0, 0);
    var end = new Date(start.getTime() + bookingMinutes(v) * 60000);
    if (end > now) continue;

    var r = i + 2;
    var resno = String(v[D.RESNO - 1] || '').trim();
    dataSh.getRange(r, D.STATUS).setValue(STATUS.DONE);
    done++;

    var amt = Number(v[D.AMOUNT - 1] || 0);
    if (isAutoSalesShop(v[D.SHOP - 1]) && resno && !paidSet[resno] && amt > 0) {
      out.push([stamp, resno, String(v[D.SHOP - 1]), String(v[D.CUSTID - 1] || ''),
                String(v[D.NAME - 1] || ''), d0, menuNameFor(v), amt, '네이버페이',
                String(v[D.STAFF - 1] || ''), amt]);
      dataSh.getRange(r, D.PAID).setValue(amt);
      paidSet[resno] = true;
    }
  }
  if (out.length) {
    slogSh = buildSlogSheet();
    slogSh.getRange(slogSh.getLastRow() + 1, 1, out.length, SLOG_HEAD.length).setValues(out);
  }
  if (done) {
    SpreadsheetApp.flush();
    recomputeCustomerStats();
    refreshHome();
  }
  return { done: done, sales: out.length };
}

function A_자동정산_지금실행() { var r = autoSettle(true); Logger.log(JSON.stringify(r)); return r; }


/* ===== 클라우드플레어용 덮어쓰기 ===== */
/* 원본 Code.gs 함수 중 웹앱에 필요 없는(스프레드시트 화면용·구글 캘린더용) 것을 비웁니다.
   같은 이름의 함수 선언은 뒤에 있는 것이 이깁니다. */
function refreshHome() {}
function syncCalendarForBooking() {}
function refreshCalendarMonth() {}
function createCalendarsForMonth() {}
function createNextMonthCalendars() {}
function refreshThisMonthCalendars() {}
function buildCalendarSheet() {}
function drawDay() {}
function refreshAllSales() {}
function buildSalesSheets() {}
function buildSalesSheet() {}
function refreshProfile() {}
function renderProfile() {}
function upsertGCalEvent() {}
function deleteGCalEvent() {}
function ensureGoogleCalendars() {}
function protectSheets() {}
function orderTabs() {}
function installTriggers() {}
function toast(msg) { Logger.log(msg); }
function alert(msg) { Logger.log(msg); }
function notifyAdmin(subject, body) { Logger.log('[알림] ' + subject + ' / ' + body); }
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function dailyMaintenance() {
  try { recomputeCustomerStats(); } catch (e) { logError('일일 정리작업', e); }
}

/* 앱 요청 입구: 원래 globalThis 에서 찾던 것을 __API 표에서 찾습니다 */
function api(name, args, token) {
  try {
    name = String(name || '');
    if (name.indexOf('api_') !== 0) return { ok: false, msg: '알 수 없는 요청입니다.' };
    var fn = __API[name];
    if (typeof fn !== 'function') return { ok: false, msg: '알 수 없는 요청입니다: ' + name };
    if (API_OPEN[name]) return fn.apply(null, args || []);
    var who = authInfo(token);
    if (!who.ok) return { ok: false, authFail: true, msg: '다시 로그인해 주세요.' };
    CUR = { role: who.role, name: who.name };
    if (who.role !== 'owner' && blockedForStaff(name)) {
      return { ok: false, msg: '이 기능은 사장님만 쓸 수 있습니다.' };
    }
    return fn.apply(null, args || []);
  } catch (e) {
    logError('앱 요청(' + name + ')', e);
    return { ok: false, msg: e.message };
  }
}

/* 사진 폴더(구글 드라이브)는 2단계에서 클라우드플레어 저장소로 바뀝니다 */
function api_customerFolder() { return { ok: false, msg: '사진 저장소는 곧 연결됩니다.' }; }
function api_uploadPhoto() { return { ok: false, msg: '사진 저장소는 곧 연결됩니다.' }; }

/* =========================================================
 *  전자 상담지·동의서 링크 (손님이 휴대폰에서 작성·서명)
 * ========================================================= */
var LINK_SHEET = '동의서링크';
var LINK_HEAD = ['토큰', '생성시각', '만료시각', '상태', '고객ID', '고객명', '예약번호', '샵', '양식명', '제출시각', '만든사람'];
var ESIGN_SHEET = '전자서명기록';
var ESIGN_HEAD = ['제출시각', '토큰', '고객ID', '고객명', '예약번호', '샵', '양식명', '응답', '서명', '기기정보', '접속주소'];

function sheetWithHead(name, head) {
  var s = getOrCreateSheet(name);
  if (s.getLastRow() < 1) s.appendRow(head);
  return s;
}
function linkRows() {
  var s = sheetWithHead(LINK_SHEET, LINK_HEAD), last = s.getLastRow();
  return { s: s, v: last < 2 ? [] : s.getRange(2, 1, last - 1, LINK_HEAD.length).getValues() };
}
function formNamesFor(shop) {
  var seen = {}, out = [];
  readForms().forEach(function (f) {
    if (f.shop && f.shop !== shop) return;
    if (!seen[f.form]) { seen[f.form] = true; out.push(f.form); }
  });
  return out;
}
function defaultFormFor(shop) {
  var names = formNamesFor(shop);
  return names.length ? names[0] : '';
}

/* 링크 만들기: p = { resno } 또는 { custId, shop }, form(선택) */
function api_consentLink(p) {
  try {
    p = p || {};
    var custId = String(p.custId || ''), name = '', shop = String(p.shop || ''), resno = String(p.resno || '');
    if (resno) {
      var row = null;
      readData().forEach(function (v) { if (String(v[D.RESNO - 1]) === resno) row = rowToObj(v); });
      if (!row) return { ok: false, msg: '예약을 찾지 못했습니다.' };
      custId = row.custId; name = row.name; shop = row.shop; p.product = row.product;
    }
    if (custId && !name) {
      var c = api_customer(custId);
      if (c.ok) { name = c.customer.name; if (!shop) shop = c.customer.shop; }
    }
    if (!shop) return { ok: false, msg: '어느 샵 양식인지 알 수 없습니다.' };
    if (String(p.form || '') === MEMBER_FORM) ensureMemberForm_();
    var form = String(p.form || '') || defaultFormFor(shop, p.product);
    if (!form) return { ok: false, msg: shop + ' 에 등록된 상담지·동의서 양식이 없습니다.' };
    /* 열려 있는 기간: 예약이 있으면 예약일 밤 11:59까지, 없으면 보낸 날 밤 11:59까지 */
    var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    var now = new Date(), today = fmt(now, 'yyyy-MM-dd'), resDay = '', resTime = '';
    if (resno && row) { resDay = String(row.date || '').slice(0, 10); resTime = String(row.time || ''); }
    if (!resDay && custId) {
      readData().forEach(function (v) {
        var o = rowToObj(v), d = String(o.date || '').slice(0, 10);
        if (o.custId === custId && (!shop || o.shop === shop) && o.status !== '취소' && d >= today && (!resDay || d < resDay)) { resDay = d; resTime = String(o.time || ''); }
      });
    }
    var exp;
    if (/^\d{4}-\d{2}-\d{2}$/.test(resDay) && resDay >= today) {
      exp = new Date(Number(resDay.slice(0, 4)), Number(resDay.slice(5, 7)) - 1, Number(resDay.slice(8, 10)), 23, 59, 59);
    } else {
      exp = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);
    }
    linkRows().s.appendRow([token, nowStr(), fmt(exp, 'yyyy-MM-dd HH:mm:ss'), '대기', custId, name, resno, shop, form, '', CUR.name]);
    var path = '/s/' + token;
    var text = (name ? String(name).replace(/님$/, '') + '님, ' : '') + '방문 전 ' + form + ' 작성 부탁드립니다.\n' + '#{링크}' + '\n(' + (exp.getMonth() + 1) + '월 ' + exp.getDate() + '일까지 작성해 주세요)';
    /* 닥터뷰티랩: 원장님이 토스로 직접 보내는 시술 전 공통 안내 + 링크 (부위 상관없이 공통) */
    if (shop === lineToShop('닥뷰')) {
      var nm = name ? String(name).replace(/님$/, '') + '님, ' : '';
      var when = '';
      if (/^\d{4}-\d{2}-\d{2}$/.test(resDay)) {
        var dd = new Date(Number(resDay.slice(0, 4)), Number(resDay.slice(5, 7)) - 1, Number(resDay.slice(8, 10)));
        when = (dd.getMonth() + 1) + '월 ' + dd.getDate() + '일(' + '일월화수목금토'.charAt(dd.getDay()) + ')' + (resTime ? ' ' + resTime : '') + ' ';
      }
      text = '[닥터뷰티랩] ' + nm + when + '예약 안내드려요.\n\n' +
        '· 아스피린·혈전약 등 피를 묽게 하는 약을 드시면 꼭 미리 알려 주세요. 시술이 어려울 수 있어요.\n' +
        '· 시술 전날은 술을 피해 주세요.\n' +
        '· 두피 시술은 당일 샴푸 후 오시고, 흑채·왁스는 쓰지 말아 주세요.\n' +
        '· 건물 지하 주차 2시간 무료예요.\n\n' +
        '상담지 미리 작성 (매장에서 써도 돼요)\n#{링크}\n\n' +
        '늦으시면 미리 연락 주세요 010-2722-3093';
    }
    if (form === MEMBER_FORM) {
      text = '[닥터뷰티랩] ' + (name ? String(name).replace(/님$/, '') + '님, ' : '') + '멤버십 이용 동의서입니다.\n내용을 확인하시고 서명 부탁드립니다.\n#{링크}\n(' + (exp.getMonth() + 1) + '월 ' + exp.getDate() + '일까지 작성하실 수 있습니다)';
      return { ok: true, token: token, path: path, form: form, forms: [form], solo: true, text: text };
    }
    return { ok: true, token: token, path: path, form: form, forms: formNamesFor(shop), text: text };
  } catch (e) { logError('동의서 링크', e); return { ok: false, msg: e.message }; }
}

function findLink(token) {
  token = String(token || '').trim();
  if (!/^[a-f0-9]{40}$/.test(token)) return null;
  var L = linkRows();
  for (var i = 0; i < L.v.length; i++) if (String(L.v[i][0]) === token) return { r: i + 2, v: L.v[i], s: L.s };
  return null;
}

/* 손님 화면용: 비밀번호 없이 토큰으로만 열림 */
function publicForm(token) {
  var L = findLink(token);
  if (!L) return { ok: false, msg: '링크가 올바르지 않습니다.' };
  var st = String(L.v[3]);
  if (st === '완료') return { ok: false, done: true, msg: '이미 작성을 마친 링크입니다. 감사합니다.' };
  if (st === '취소') return { ok: false, msg: '사용할 수 없는 링크입니다. 매장에 문의해 주세요.' };
  if (String(L.v[2]) < nowStr()) return { ok: false, msg: '링크 사용 기간이 지났습니다. 매장에 새 링크를 요청해 주세요.' };
  var shop = String(L.v[7]), form = String(L.v[8]);
  var qs = readForms().filter(function (f) { return (!f.shop || f.shop === shop) && f.form === form; })
    .map(function (f) { return { q: f.q, type: f.type, opts: f.opts, need: f.need }; });
  return { ok: true, shop: shop, form: form, name: String(L.v[5] || ''), questions: qs };
}

function publicSubmit(token, body, meta) {
  var f = publicForm(token);
  if (!f.ok) return f;
  body = body || {};
  var ans = Array.isArray(body.answers) ? body.answers : [];
  for (var i = 0; i < f.questions.length; i++) {
    var q = f.questions[i];
    if (!q.need || q.type === '안내') continue;
    var a = ans[i] ? ans[i].v : '';
    var empty = (a === '' || a === null || a === undefined || a === false || (Array.isArray(a) && !a.length));
    if (empty) return { ok: false, msg: '필수 항목을 작성해 주세요: ' + q.q };
  }
  var sign = String(body.sign || '');
  if (!/^data:image\/png;base64,/.test(sign) || sign.length < 200) return { ok: false, msg: '서명을 해주세요.' };
  if (sign.length > 400000) return { ok: false, msg: '서명 이미지가 너무 큽니다. 다시 서명해 주세요.' };
  var L = findLink(token), v = L.v;
  var clean = f.questions.map(function (q, i) { return { q: q.q, type: q.type, v: ans[i] ? ans[i].v : '' }; });
  var stamp = nowStr();
  sheetWithHead(ESIGN_SHEET, ESIGN_HEAD).appendRow([stamp, token, v[4], v[5], v[6], v[7], v[8],
    JSON.stringify(clean), sign, String((meta && meta.ua) || '').slice(0, 300), String((meta && meta.ip) || '')]);
  L.s.getRange(L.r, 4).setValue('완료');
  L.s.getRange(L.r, 10).setValue(stamp);
  /* 기존 고객 화면의 상담지 목록에도 요약을 남깁니다 */
  try {
    var lines = clean.filter(function (x) { return x.type !== '안내'; }).map(function (x) {
      var val = Array.isArray(x.v) ? x.v.join(', ') : (x.v === true ? '동의' : (x.v === false ? '동의안함' : String(x.v || '')));
      return '· ' + x.q + ' → ' + val;
    }).join('\n');
    buildConsentSheet().appendRow([stamp, v[4], v[5], v[7], v[8] + ' (전자서명' + (LN ? ' · ' + LN : '') + ')', lines, 'esign:' + token]);
  } catch (e) {}
  return { ok: true };
}

/* 앱 안에서 보기 (로그인 필요) */
function api_esignList(custId) {
  try {
    var s = sheetWithHead(ESIGN_SHEET, ESIGN_HEAD), last = s.getLastRow(), out = [];
    if (last >= 2) s.getRange(2, 1, last - 1, ESIGN_HEAD.length).getValues().forEach(function (r) {
      if (String(r[2]) !== String(custId)) return;
      out.push({ at: String(r[0]), token: String(r[1]), resno: String(r[4]), shop: String(r[5]), form: String(r[6]) });
    });
    out.reverse();
    var pend = [];
    linkRows().v.forEach(function (r) {
      if (String(r[4]) === String(custId) && String(r[3]) === '대기' && String(r[2]) >= nowStr())
        pend.push({ token: String(r[0]), form: String(r[8]), made: String(r[1]), exp: String(r[2]) });
    });
    return { ok: true, list: out, pending: pend };
  } catch (e) { return { ok: false, msg: e.message, list: [], pending: [] }; }
}
function api_esignView(token) {
  var s = sheetWithHead(ESIGN_SHEET, ESIGN_HEAD), last = s.getLastRow();
  if (last < 2) return { ok: false, msg: '기록이 없습니다.' };
  var v = s.getRange(2, 1, last - 1, ESIGN_HEAD.length).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][1]) === String(token)) {
    var r = v[i], answers = [];
    try { answers = JSON.parse(String(r[7] || '[]')); } catch (e) {}
    if (!CUR || CUR.role !== 'owner') answers = answers.map(function (a) {
      return ownerOnlyQ_(a.q) ? { q: '(대표만 볼 수 있는 항목)', v: '—', type: a.type } : a;
    });
    return { ok: true, at: String(r[0]), name: String(r[3]), resno: String(r[4]), shop: String(r[5]), form: String(r[6]),
             answers: answers, sign: String(r[8]), device: String(r[9]) };
  }
  return { ok: false, msg: '기록을 찾지 못했습니다.' };
}

/* =========================================================
 *  시술 사진 (원본은 클라우드플레어 R2, 목록은 이 시트)
 * ========================================================= */
var PHOTO_SHEET = '시술사진';
var PHOTO_HEAD = ['등록시각', '고객ID', '시술일', '구분', '칸', '원본키', '미리보기키', '크기', '찍은사람', '삭제시각', '파일종류'];
function photoSheet() { return sheetWithHead(PHOTO_SHEET, PHOTO_HEAD); }
function authOf(token) { return authInfo(token); }
function addPhotoRow(o) {
  photoSheet().appendRow([nowStr(), o.custId, o.date || ymd(new Date()), o.kind || '', o.slot || '', o.key, o.thumb || '',
    Number(o.size || 0), o.by || '', '', o.type || '']);
  return { ok: true };
}
function api_photos(custId) {
  var s = photoSheet(), last = s.getLastRow(), out = [];
  if (last >= 2) s.getRange(2, 1, last - 1, PHOTO_HEAD.length).getValues().forEach(function (r) {
    if (String(r[1]) !== String(custId) || String(r[9] || '')) return;
    out.push({ at: String(r[0]), date: ymd(r[2]) || String(r[2]), kind: String(r[3]), slot: String(r[4]),
               key: String(r[5]), thumb: String(r[6]), type: String(r[10] || '') });
  });
  return { ok: true, list: out };
}
/* 지우기 = 휴지통 표시 (30일 뒤 실제 삭제) */
function api_deletePhoto(key) {
  var s = photoSheet(), last = s.getLastRow();
  if (last < 2) return { ok: false, msg: '사진을 찾지 못했습니다.' };
  var v = s.getRange(2, 6, last - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === String(key)) {
    s.getRange(i + 2, 10).setValue(nowStr());
    return { ok: true };
  }
  return { ok: false, msg: '사진을 찾지 못했습니다.' };
}
/* 30일 지난 휴지통 사진 목록 (실제 파일 삭제는 바깥에서) */
function expiredPhotos() {
  var s = photoSheet(), last = s.getLastRow(), out = [], limit = fmt(new Date(Date.now() - 30 * 86400000), 'yyyy-MM-dd HH:mm:ss');
  if (last < 2) return out;
  var v = s.getRange(2, 1, last - 1, PHOTO_HEAD.length).getValues();
  for (var i = v.length - 1; i >= 0; i--) {
    var del = String(v[i][9] || '');
    if (del && del < limit) { out.push(String(v[i][5])); if (v[i][6]) out.push(String(v[i][6])); s.deleteRow(i + 2); }
  }
  return out;
}

/* 방문 경로·검색 키워드 통계 (전자 상담지 답변에서 모음) */
function api_sourceStats(from, to) {
  var s = sheetWithHead(ESIGN_SHEET, ESIGN_HEAD), last = s.getLastRow();
  var src = {}, kw = {}, ref = {}, n = 0;
  if (last >= 2) s.getRange(2, 1, last - 1, ESIGN_HEAD.length).getValues().forEach(function (r) {
    var d = String(r[0]).slice(0, 10);
    if ((from && d < from) || (to && d > to)) return;
    var ans = []; try { ans = JSON.parse(String(r[7] || '[]')); } catch (e) {}
    n++;
    ans.forEach(function (a) {
      var q = String(a.q || '').replace(/\s+/g, '');
      var vals = Array.isArray(a.v) ? a.v : (a.v ? [a.v] : []);
      if (q.indexOf('방문경로') >= 0 || q.indexOf('알게') >= 0) {
        vals.forEach(function (x) { x = String(x).trim(); if (x) src[x] = (src[x] || 0) + 1; });
      } else if (q.indexOf('키워드') >= 0 || q.indexOf('검색') >= 0) {
        vals.forEach(function (x) {
          String(x).split(/[,\/·]|\n/).forEach(function (k) { k = k.trim().replace(/\s+/g, ' '); if (k && k.length < 40) kw[k] = (kw[k] || 0) + 1; });
        });
      } else if (q.indexOf('소개') >= 0) {
        vals.forEach(function (x) { x = String(x).trim(); if (x) ref[x] = (ref[x] || 0) + 1; });
      }
    });
  });
  function top(o) { return Object.keys(o).map(function (k) { return { name: k, n: o[k] }; }).sort(function (a, b) { return b.n - a.n; }); }
  var refc = []; try { refc = topReferrers_(); } catch (e) {}
  return { ok: true, count: n, sources: top(src), keywords: top(kw).slice(0, 30), referrers: top(ref).slice(0, 20), refCustomers: refc };
}

/* ---------- 팅커뷰 매출 옮겨 넣기 (대표만) ----------
   rows: [[날짜 yyyy-MM-dd, 시각 HH:mm, 담당라인, 고객명, 시술명, 결제수단, 금액, 네이버결제액], ...]
   같은 예약번호(TVB날짜-순번)가 이미 있으면 건너뜁니다. */
function api_importTvSales(rows) {
  try {
    if (CUR.role !== 'owner') return { ok: false, msg: '대표만 쓸 수 있습니다.' };
    var dataSheet = sheetOf(SH.DATA);
    var slogSheet = buildSlogSheet();
    var haveRes = {};
    readData().forEach(function (v) { haveRes[String(v[D.RESNO - 1] || '').trim()] = 1; });
    var now = nowStr(), seq = {}, addRes = [], addSlog = [], skipped = 0, newCust = 0;
    (rows || []).forEach(function (t) {
      var date = String(t[0]), time = String(t[1] || ''), line = String(t[2] || '');
      var name = String(t[3] || '').trim(), item = String(t[4] || ''), pay = String(t[5] || '');
      var amt = Number(t[6] || 0), npay = Number(t[7] || 0);
      var shop = lineToShop(line);
      var key = date.replace(/-/g, '');
      seq[key] = (seq[key] || 0) + 1;
      var resno = 'TVB' + key + '-' + seq[key];
      if (haveRes[resno]) { skipped++; return; }
      var before = readCustomers().length;
      var cid = matchOrCreateCustomer(name, '', '', shop, '팅커뷰');
      if (readCustomers().length > before) newCust++;
      var d = new Array(D_HEAD.length).fill('');
      d[D.RESNO - 1] = resno; d[D.SHOP - 1] = shop; d[D.STATUS - 1] = STATUS.DONE;
      d[D.NAME - 1] = name; d[D.DATE - 1] = date; d[D.TIME - 1] = time;
      d[D.PRODUCT - 1] = item; d[D.AMOUNT - 1] = amt; d[D.PAID - 1] = amt;
      d[D.STAFF - 1] = line; d[D.SOURCE - 1] = '팅커뷰'; d[D.CUSTID - 1] = cid; d[D.PROCAT - 1] = now;
      addRes.push(d);
      addSlog.push([now, resno, shop, cid, name, toDate(date), item, amt, pay, line, npay || '']);
    });
    if (addRes.length) dataSheet.getRange(dataSheet.getLastRow() + 1, 1, addRes.length, D_HEAD.length).setValues(addRes);
    if (addSlog.length) slogSheet.getRange(slogSheet.getLastRow() + 1, 1, addSlog.length, SLOG_HEAD.length).setValues(addSlog);
    recomputeCustomerStats();
    return { ok: true, added: addSlog.length, skipped: skipped, newCustomers: newCust };
  } catch (e) { logError('팅커뷰 매출 넣기', e); return { ok: false, msg: e.message }; }
}

/* ---------- 동의서 양식: 조건부 문항 · 도움말 · 공통 보호자 동의 ----------
   동의서양식 시트 8열 '조건'(예: 방문경로=네이버|인스타, 나이<19), 9열 '도움말' */
var FORM_COMMON = '보호자 동의(공통)';
function readForms() {
  var sh = buildFormSheet();
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, FORM_HEAD.length + 2).getValues()
    .filter(function (r) { return String(r[3] || '').trim(); })
    .map(function (r) {
      return { shop: String(r[0] || '').trim(), form: String(r[1] || '').trim(),
               order: Number(r[2] || 0), q: String(r[3]).trim(),
               type: String(r[4] || '단답').trim(),
               opts: String(r[5] || '').split('|').map(function (x) { return x.trim(); })
                       .filter(function (x) { return x; }),
               need: String(r[6] || '') === 'Y',
               cond: String(r[7] || '').trim(), help: String(r[8] || '').trim() };
    })
    .sort(function (a, b) { return a.order - b.order; });
}
function formRow_(f, i) {
  return [String(f.shop || '').trim(), String(f.form || '상담지').trim(), i + 1,
          String(f.q).trim(), String(f.type || '단답').trim(),
          (f.opts || []).join(' | '), f.need ? 'Y' : '', String(f.cond || ''), String(f.help || '')];
}
function writeForms_(list) {
  var sh = buildFormSheet();
  sh.getRange(1, FORM_HEAD.length + 1, 1, 2).setValues([['조건', '도움말']]);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, FORM_HEAD.length + 2).clearContent();
  var rows = list.filter(function (f) { return String(f.q || '').trim(); }).map(formRow_);
  if (rows.length) sh.getRange(2, 1, rows.length, FORM_HEAD.length + 2).setValues(rows);
  return rows.length;
}
function api_saveForms(list) {
  try {
    /* 앱 편집 화면이 조건·도움말을 모르고 보내도 기존 값을 지킵니다 */
    var old = {};
    readForms().forEach(function (f) { old[f.shop + '\u0001' + f.form + '\u0001' + f.q] = f; });
    var merged = (list || []).map(function (f) {
      var o = old[String(f.shop || '').trim() + '\u0001' + String(f.form || '상담지').trim() + '\u0001' + String(f.q || '').trim()];
      if (o && f.cond === undefined) f.cond = o.cond;
      if (o && f.help === undefined) f.help = o.help;
      return f;
    });
    return { ok: true, n: writeForms_(merged) };
  } catch (e) { logError('앱 양식저장', e); return { ok: false, msg: e.message }; }
}

function formQuestions_(shop, form) {
  var all = readForms();
  var qs = all.filter(function (f) { return (!f.shop || f.shop === shop) && f.form === form; });
  /* 생년월일 문항이 있는 양식에는 공통 보호자 동의를 뒤에 붙입니다 */
  if (form !== FORM_COMMON && qs.some(function (f) { return f.type === '생년월일'; })) {
    qs = qs.concat(all.filter(function (f) { return f.form === FORM_COMMON; }));
  }
  return qs.map(function (f) { return { q: f.q, type: f.type, opts: f.opts, need: f.need, cond: f.cond, help: f.help }; });
}
function ageOf_(v) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(v || ''));
  if (!m) return null;
  var n = new Date(), a = n.getFullYear() - Number(m[1]);
  if (n.getMonth() + 1 < Number(m[2]) || (n.getMonth() + 1 === Number(m[2]) && n.getDate() < Number(m[3]))) a--;
  return a;
}
function shownQ_(qs, ans, i, depth) {
  var c = qs[i].cond;
  if (!c || (depth || 0) > 5) return true;
  if (/^lang\s*=/.test(c)) return true;
  var m = /^나이\s*<\s*(\d+)$/.exec(c);
  if (m) {
    for (var k = 0; k < qs.length; k++) if (qs[k].type === '생년월일') {
      var a = ageOf_(ans[k] ? ans[k].v : ''); return a !== null && a < Number(m[1]);
    }
    return false;
  }
  var p = c.indexOf('=');
  if (p < 0) return true;
  var t = c.slice(0, p).trim(), want = c.slice(p + 1).split('|').map(function (x) { return x.trim(); });
  for (var j = 0; j < qs.length; j++) if (qs[j].q === t) {
    if (!shownQ_(qs, ans, j, (depth || 0) + 1)) return false;
    var v = ans[j] ? ans[j].v : '';
    if (Array.isArray(v)) return v.some(function (x) { return want.indexOf(x) >= 0; });
    return want.indexOf(v) >= 0;
  }
  return false;
}

function publicForm(token) {
  var L = findLink(token);
  if (!L) return { ok: false, msg: '링크가 올바르지 않습니다.' };
  var st = String(L.v[3]);
  if (st === '완료') return { ok: false, done: true, msg: '이미 작성을 마친 링크입니다. 감사합니다.' };
  if (st === '취소') return { ok: false, msg: '사용할 수 없는 링크입니다. 매장에 문의해 주세요.' };
  if (String(L.v[2]) < nowStr()) return { ok: false, msg: '링크 사용 기간이 지났습니다. 매장에 새 링크를 요청해 주세요.' };
  var shop = String(L.v[7]), form = String(L.v[8]);
  return { ok: true, shop: shop, form: form, name: String(L.v[5] || ''), questions: formQuestions_(shop, form) };
}

function publicSubmit(token, body, meta) {
  var f = publicForm(token);
  if (!f.ok) return f;
  body = body || {};
  var ans = Array.isArray(body.answers) ? body.answers : [];
  var qs = f.questions;
  for (var i = 0; i < qs.length; i++) {
    var q = qs[i];
    var on = shownQ_(qs, ans, i);
    if (!on) { ans[i] = { v: '' }; continue; }
    if (q.type === '서명') {
      var sv = String(ans[i] ? ans[i].v : '');
      if (sv && (!/^data:image\/png;base64,/.test(sv) || sv.length > 400000)) return { ok: false, msg: '보호자 서명을 다시 해주세요.' };
    }
    if (!q.need || q.type === '안내') continue;
    var a = ans[i] ? ans[i].v : '';
    var empty = (a === '' || a === null || a === undefined || a === false || (Array.isArray(a) && !a.length));
    if (empty) return { ok: false, msg: '필수 항목을 작성해 주세요: ' + q.q };
  }
  var sign = String(body.sign || '');
  if (!/^data:image\/png;base64,/.test(sign) || sign.length < 200) return { ok: false, msg: '서명을 해주세요.' };
  if (sign.length > 400000) return { ok: false, msg: '서명 이미지가 너무 큽니다. 다시 서명해 주세요.' };
  var L = findLink(token), v = L.v;
  var clean = qs.map(function (q, i) { return { q: q.q, type: q.type, v: ans[i] ? ans[i].v : '' }; });
  var LN = { en: '영어', ja: '일본어', zh: '중국어' }[meta && meta.lang];
  if (LN) clean.push({ q: '작성 언어', type: '단답', v: LN });
  var stamp = nowStr();
  sheetWithHead(ESIGN_SHEET, ESIGN_HEAD).appendRow([stamp, token, v[4], v[5], v[6], v[7], v[8],
    JSON.stringify(clean), sign, String((meta && meta.ua) || '').slice(0, 300), String((meta && meta.ip) || '')]);
  L.s.getRange(L.r, 4).setValue('완료');
  L.s.getRange(L.r, 10).setValue(stamp);
  try {
    var lines = clean.filter(function (x, i) { return x.type !== '안내' && (i >= qs.length || shownQ_(qs, ans, i)); }).map(function (x) {
      var val = x.type === '서명' ? (x.v ? '서명 완료' : '') :
        (Array.isArray(x.v) ? x.v.join(', ') : (x.v === true ? '동의' : (x.v === false ? '동의안함' : String(x.v || ''))));
      return '· ' + x.q + ' → ' + val;
    }).join('\n');
    buildConsentSheet().appendRow([stamp, v[4], v[5], v[7], v[8] + ' (전자서명)', lines, 'esign:' + token]);
  } catch (e) {}
  try { fillChartFromForm_(String(v[4] || ''), clean, qs, ans, meta && meta.lang, stamp); } catch (e) { logError('상담지→고객차트', e); }
  return { ok: true };
}

/* ---------- 상담지·동의서 → 고객 차트 ----------
   인적사항(성함·연락처·생년월일·성별·언어)은 차트가 비어 있을 때만 채우고,
   방문경로와 주의사항(건강 질문에 '예'라고 한 것만)은 가장 최근 서식 기준으로 바꿔 넣어요. */
var HEALTH_CONFIRM_Q = '건강 상태 답변 확인';
var HEALTH_CONFIRM_A = '위 건강 상태 질문에 사실대로 답했음을 확인합니다. 사실과 다르게 적거나 알리지 않아 생긴 문제는 본인에게 책임이 있음을 이해합니다.';
/* 감염병 문항은 대표만 봅니다 (직원 화면·고객 차트에는 내용 대신 '대표 확인 필요'만) */
function ownerOnlyQ_(q) { return String(q || '').indexOf('감염병') >= 0; }
var CAUTION_LABEL = {
  '과거에': '이전 문신 이력', '이전 시술 뒤': '이전 시술 후 이상', '문신 색소에 알레르기': '색소 알레르기',
  '민감하거나': '민감한 피부', '시술 부위에 상처': '시술 부위 피부 이상', '최근 3개월 안에 피부과': '최근 피부과 치료',
  '흉터가 붉게': '켈로이드·잘 붓는 피부', '알레르기가 있는 물질': '알레르기', 'B형·C형 간염': '혈액 감염병',
  '실신한 적이': '실신·뇌신경계 질환', '지금 복용 중인 약': '복용 약', '오늘 아픈 곳': '오늘 컨디션 안 좋음',
  '최근 24시간 안에 술': '24시간 내 음주', '임신 중이거나': '임신·수유 중', '입술 헤르페스': '입술 헤르페스',
  '모발 이식': '모발 이식', '두피염': '두피염·두피 질환'
};
function fillChartFromForm_(cid, clean, qs, ans, lang, stamp) {
  if (!cid) return;
  var sh = sheetOf(SH.CUST), rows = readCustomers(), r = -1;
  for (var i = 0; i < rows.length; i++) if (String(rows[i][C.ID - 1]).trim() === cid) { r = i; break; }
  if (r < 0) return;
  var row = rows[r], R = r + 2;
  function shown(i) {
    if (qs) return i < qs.length && shownQ_(qs, ans, i);
    var x = clean[i] && clean[i].v; return !(x === '' || x === null || x === undefined || (Array.isArray(x) && !x.length));
  }
  function val(x) { return Array.isArray(x) ? x.join(', ') : (x === true || x === false ? '' : String(x || '').trim()); }
  function byQ(q) { for (var i = 0; i < clean.length; i++) if (clean[i].q === q && shown(i)) return val(clean[i].v); return ''; }
  function empty(col) { return String(row[col - 1] || '').trim() === ''; }

  /* 인적사항: 비어 있는 칸만 */
  var name = byQ('성함'), cur = String(row[C.NAME - 1] || '');
  if (name && (cur.indexOf('*') >= 0 || cur === '고객' || !cur)) { try { renameEverywhere(cid, '', name); } catch (e) {} }
  var ph = normPhone(byQ('연락처'));
  if (ph && empty(C.PHONE) && empty(C.SAFE)) sh.getRange(R, isSafePhone(ph) ? C.SAFE : C.PHONE).setValue(ph);
  var bd = byQ('생년월일');
  if (/^\d{4}-\d{2}-\d{2}$/.test(bd) && empty(C.BIRTH)) sh.getRange(R, C.BIRTH).setValue(toDate(bd));
  var gd = byQ('성별');
  if ((gd === '여' || gd === '남') && empty(C.GENDER)) sh.getRange(R, C.GENDER).setValue(gd);
  var LN = { en: '영어', ja: '일본어', zh: '중국어', ko: '한국어' }[lang || 'ko'];
  if (LN && empty(C.LANG)) sh.getRange(R, C.LANG).setValue(LN);

  /* 방문경로: 경로 + 키워드/소개자 */
  var route = byQ('방문경로');
  if (route) {
    var kw = byQ('어떤 키워드로 검색해서 오셨어요?'), ref = byQ('소개해 주신 분 성함'), etc = byQ('어떤 경로로 오셨는지 알려 주세요');
    var extra = kw || ref || etc;
    sh.getRange(R, C.ROUTE).setValue(route + (extra ? ' (' + extra + ')' : ''));
  }

  /* 주의사항: 건강 확인 구간에서 '예'·질병 선택·특이사항만 */
  var start = -1, notes = [];
  for (var k = 0; k < clean.length; k++) {
    var q = clean[k].q;
    if (q === '건강 상태 확인') { start = k; continue; }
    if (start < 0 || !shown(k)) continue;
    if (q === '그 밖에 원장님께 알려 주실 내용') { var t = val(clean[k].v); if (t) notes.push('특이사항: ' + t); break; }
    var a = clean[k].v;
    if (q.indexOf('진단받았거나 지금 치료 중인 질병') === 0) {
      var L = (Array.isArray(a) ? a : [a]).filter(function (x) { return x && x !== '해당 없음' && x !== '기타'; });
      var other = byQ('어떤 질병인지 알려 주세요');
      if (other) L.push(other);
      if (L.length) notes.push('질병: ' + L.join(', '));
      continue;
    }
    if (a !== '예' || clean[k].type !== '단일선택') continue;
    if (ownerOnlyQ_(q)) { notes.push('대표 확인 필요 항목 있음'); continue; }
    var label = q;
    Object.keys(CAUTION_LABEL).some(function (key) { if (q.indexOf(key) === 0) { label = CAUTION_LABEL[key]; return true; } return false; });
    /* 바로 다음 칸이 이 질문의 '자세히' 칸이면 같이 붙여요 */
    var nx = qs ? qs[k + 1] : null, nc = clean[k + 1];
    var isDet = qs ? (nx && nx.cond === q + '=예') : (nc && nc.type === '단답');
    var det = (isDet && shown(k + 1)) ? val(nc.v) : '';
    notes.push(label + (det ? ': ' + det : ''));
  }
  if (start >= 0) sh.getRange(R, C.CAUTION).setValue(notes.length ? '[' + String(stamp).slice(0, 10) + ' 상담지] ' + notes.join(' · ') : '');
  SpreadsheetApp.flush();
}

/* 이미 받은 전자 상담지·동의서를 오래된 것부터 차트에 다시 반영 (대표만) */
function api_backfillCharts() {
  if (CUR.role !== 'owner') return { ok: false, msg: '대표만 쓸 수 있습니다.' };
  var s = sheetWithHead(ESIGN_SHEET, ESIGN_HEAD), last = s.getLastRow();
  if (last < 2) return { ok: true, n: 0, people: 0 };
  var rows = s.getRange(2, 1, last - 1, ESIGN_HEAD.length).getValues();
  rows.sort(function (a, b) { return String(a[0]).localeCompare(String(b[0])); });
  var n = 0, seen = {};
  rows.forEach(function (r) {
    var cid = String(r[2] || '').trim(); if (!cid) return;
    var clean = []; try { clean = JSON.parse(String(r[7] || '[]')); } catch (e) { return; }
    var langRow = clean.filter(function (x) { return x.q === '작성 언어'; })[0];
    var lang = langRow ? ({ '영어': 'en', '일본어': 'ja', '중국어': 'zh' }[langRow.v] || 'ko') : 'ko';
    try { fillChartFromForm_(cid, clean, null, null, lang, String(r[0])); n++; seen[cid] = 1; } catch (e) { logError('차트 다시 반영', e); }
  });
  try { recomputeCustomerStats(); } catch (e) {}
  return { ok: true, n: n, people: Object.keys(seen).length };
}

/* ---------- 닥터뷰티랩 눈썹·입술·두피 상담·동의서 + 공통 보호자 동의 (대표만 실행) ----------
   원장님 캔바 상담지(1~3쪽) + 문신시술 표준지침(2026.7) 서식 1·3·4호 항목을 합친 기본 문구 */
function pmuForms_(shop) {
  var YN = ['예', '아니요'];
  function Q(form, q, type, opts, need, cond, help) {
    return { shop: shop, form: form, q: q, type: type, opts: opts || [], need: !!need, cond: cond || '', help: help || '' };
  }
  function top(form, part) {
    return [
      Q(form, part + ' 시술 전 상담지·동의서예요.\n작성해 주신 내용은 안전한 시술을 위한 기록으로, 관련 법령에 따라 5년간 보관돼요.', '안내'),
      Q(form, '성함', '단답', [], 1),
      Q(form, '연락처', '전화', [], 1),
      Q(form, '생년월일', '생년월일', [], 1, '', '만 19세 미만이면 아래에 보호자 동의 칸이 함께 나와요.'),
      Q(form, '성별', '단일선택', ['여', '남'], 1),
      Q(form, '주소', '단답', [], 0, '', '시·구·동까지만 적어 주셔도 괜찮아요.'),
      Q(form, '응급 시 연락드릴 분', '단답', [], 1, '', '성함 · 연락처 · 관계를 적어 주세요. 예: 홍길동 010-0000-0000 배우자'),
      Q(form, '방문 구분', '단일선택', ['신규', '재방문'], 1),
      Q(form, '방문경로', '단일선택', ['네이버', '인스타', '블로그', '당신을그리다', '소개', '기타'], 1),
      Q(form, '어떤 키워드로 검색해서 오셨어요?', '단답', [], 0, '방문경로=네이버|인스타|블로그', '예: 마곡 눈썹문신 (선택)'),
      Q(form, '소개해 주신 분 성함', '단답', [], 0, '방문경로=소개', '예: 김○○ 님 (선택)'),
      Q(form, '어떤 경로로 오셨는지 알려 주세요', '단답', [], 0, '방문경로=기타')
    ];
  }
  function yn(form, q, detail, help) {
    var out = [Q(form, q, '단일선택', YN, 1, '', help)];
    if (detail) out.push(Q(form, detail, '단답', [], 1, q + '=예'));
    return out;
  }
  function health(form, part, prior, extra) {
    var priorQ = '과거에 ' + prior + '을 받으신 적이 있나요?';
    var L = [Q(form, '건강 상태 확인', '안내')];
    L = L.concat(yn(form, priorQ, '리터치를 포함한 총 횟수와 마지막 시술 연도를 알려 주세요'));
    L = L.concat(yn(form, '이전 시술 뒤 염증·부기·변색 같은 이상 증상이 있었나요?', '어떤 증상이었는지 알려 주세요'));
    L[L.length - 2].cond = priorQ + '=예';
    L = L.concat(yn(form, '문신 색소에 알레르기 반응(발진·가려움·부기)을 겪은 적이 있나요?', '어떤 색상이었는지 알려 주세요'));
    L = L.concat(extra || []);
    L = L.concat(yn(form, '민감하거나 예민한 피부이신가요?'));
    L = L.concat(yn(form, '시술 부위에 상처, 염증, 발진 등 피부 이상이 있나요?', '어떤 상태인지 알려 주세요'));
    L = L.concat(yn(form, '최근 3개월 안에 피부과 치료(염증, 건선, 습진, 아토피 등)를 받으셨나요?', '어떤 치료였는지 알려 주세요'));
    L = L.concat(yn(form, '흉터가 붉게 부풀어 오르거나(켈로이드) 피부가 잘 붓는 편인가요?'));
    L = L.concat(yn(form, '알레르기가 있는 물질(금속·약·알코올·라텍스 등)이 있나요?', '무엇인지 알려 주세요'));
    L.push(Q(form, '진단받았거나 지금 치료 중인 질병이 있나요?', '복수선택',
      ['뇌졸중(중풍)', '심근경색·협심증', '고혈압', '당뇨병', '고지혈증', '폐결핵', '기타', '해당 없음'], 1));
    L.push(Q(form, '어떤 질병인지 알려 주세요', '단답', [], 1, '진단받았거나 지금 치료 중인 질병이 있나요?=기타'));
    L = L.concat(yn(form, 'B형·C형 간염, 후천성면역결핍증(AIDS) 등 혈액으로 옮을 수 있는 감염병으로 진단받았거나 치료 중이신가요?', ''));
    L = L.concat(yn(form, '실신한 적이 있거나 뇌전증 등 뇌신경계 질환이 있나요?', ''));
    L = L.concat(yn(form, '지금 복용 중인 약이 있나요? (항응고제·아스피린·면역억제제·탈모약 등)', '약 이름을 알려 주세요'));
    L = L.concat(yn(form, '오늘 아픈 곳이 있거나 열, 오한, 심한 피로감이 있나요?'));
    L = L.concat(yn(form, '최근 24시간 안에 술을 드셨나요?'));
    L = L.concat(yn(form, '임신 중이거나 수유 중이신가요?'));
    L.push(Q(form, '그 밖에 원장님께 알려 주실 내용', '서술', [], 0));
    L.push(Q(form, HEALTH_CONFIRM_Q, '동의', [HEALTH_CONFIRM_A], 1));
    return L;
  }
  function tail(form, care, gift) {
    var agreeText = '본인은 상기 시술을 시행하기에 앞서 매우 드물게 특이 체질 등으로 예측할 수 없는 부작용과 후유증이 발생될 가능성이 있는 점을 염두에 두며 시술 관련 안내에 대해 구체적으로 충분히 안내받았음을 확인합니다. 본인이 지침 위반하여 발생하는 문제에 대해서 업체는 책임지지 않습니다. 본인은 시술에 자발적 동의하며, 시술 결과에 대해 사업자의 객관적 중대한 책임 외 이의 제기하지 않음을 확인합니다.';
    return [
      Q(form, '시술 후 주의사항\n\n' + care.map(function (x) { return '· ' + x; }).join('\n'), '안내'),
      Q(form, '시술 후 주의사항', '동의', ['위 주의사항을 읽고 이해했어요.'], 1),
      Q(form, '시술 동의사항', '안내'),
      Q(form, '시술 방법과 과정', '동의', ['문신행위의 방법과 과정에 대해 충분한 설명을 들었습니다.'], 1),
      Q(form, '부작용·후유증·사후관리', '동의', ['시술 후 생길 수 있는 부작용, 후유증, 사후관리에 대해 충분한 설명을 들었습니다.'], 1),
      Q(form, '시술 전후 유의사항', '동의', ['시술 전후 지켜야 할 유의사항에 대해 충분한 설명을 들었습니다.'], 1),
      Q(form, '자발적 동의', '동의', ['이번 시술은 본인의 자발적인 의사에 따릅니다.'], 1),
      Q(form, agreeText, '안내'),
      Q(form, '시술 동의 확인', '동의', ['위 내용을 확인했으며 동의합니다.'], 1),
      Q(form, '기록용 사진·영상 촬영', '동의', ['시술 전후 상태를 정확히 기록하고 관리하기 위한 사진·영상 촬영에 동의합니다.'], 1),
      Q(form, '시술 사진을 매장 SNS에 올려도 될까요? (선택)', '단일선택', ['동의해요', '동의하지 않아요'], 0, '',
        '동의하지 않아도 시술받는 데 아무 문제 없어요.'),
      /* 리뷰 혜택은 한국어 화면에서만 보여요 (외국어 화면에서는 숨김) */
      Q(form, '리뷰 작성 시 혜택이 있어요. 참여하시겠어요? (선택)', '단일선택', ['참여할게요', '다음에 할게요'], 0, 'lang=ko',
        '시술 후 리뷰를 남겨 주시면 ' + gift + ' 드려요.'),
      Q(form, '개인정보 수집·이용', '동의', ['이름, 연락처, 생년월일 등 작성한 개인정보의 수집·이용에 동의하며, 관련 법령에 따라 일정 기간 보관됨을 설명 들었습니다.'], 1),
      Q(form, '동의서 사본이 필요하시면 언제든 말씀해 주세요.', '안내')
    ];
  }
  var EB = '눈썹 상담·동의서', LP = '입술 상담·동의서', SC = '두피 상담·동의서';
  var out = [];
  out = out.concat(top(EB, '눈썹'), health(EB, '눈썹', '눈썹 문신(반영구)'), tail(EB, [
    '일주일간 시술 부위에 물이나 마찰이 닿지 않게 주의해 주세요. 눈썹 화장, 클렌징오일, 유분기 많은 화장품은 일주일 정도 피해 주세요.',
    '재생크림을 아침·저녁 세안 시 얇게 발라 주세요. 면봉에 색소가 묻어 나올 수 있어요. (너무 많이 바르면 색소 안착에 방해가 돼요.)',
    '회복 기간(최소 1주) 동안 음주와 흡연, 사우나, 땀이 많이 나는 운동은 염증을 일으킬 수 있어 자제해 주세요.',
    '딱지를 억지로 떼지 마세요. 억지로 떼면 색소가 고르게 안착되지 않거나 색이 덜 나올 수 있어요.',
    '1달 동안 시간을 두고 지켜봐 주세요. 점점 예쁘게 색이 올라와요.'
  ], '재생크림을'));
  out = out.concat(top(LP, '입술'), health(LP, '입술', '입술 문신(반영구)',
    yn(LP, '입술 헤르페스(구순포진) 수포가 올라온 적이 있나요?', '마지막으로 올라온 시기를 알려 주세요')), tail(LP, [
    '시술 후 각질이 떨어지는 동안 입술이 많이 건조해요. 립밤을 자주 발라 보습해 주세요. (매우 중요)',
    '시술 부위가 물에 닿지 않게 주의해 주세요. 세안·양치할 때 입술에 비누와 물이 닿지 않게 하고, 샤워할 때도 조심해 주세요.',
    '각질은 저절로 떨어지게 두세요. 손으로 떼거나 문지르면 색이 고르지 않게 되거나 염증이 생길 수 있어요.',
    '맵거나 뜨거운 음식, 와인처럼 입술에 자극이 되는 음식은 피하고, 음료는 빨대로 드시는 걸 권해요.',
    '립스틱·틴트 같은 색조 화장은 상처가 다 아문 뒤에 써 주세요. (컬러 립밤은 7일 뒤, 착색이 강한 틴트는 1달 뒤)',
    '붓기는 사람마다 달라 2~3일 정도 생길 수 있어요. 헤르페스 수포가 올라오면 바로 피부과에 가 주세요.',
    '회복 기간(최소 1주) 동안 음주와 흡연, 사우나, 땀이 많이 나는 운동은 염증을 일으킬 수 있어 자제해 주세요.',
    '색을 오래 유지하려면 평소에도 립밤으로 보습을 꾸준히 해 주세요.',
    '1달 동안 색이 변하는 건 자연스러운 과정이에요. 시간을 두고 지켜봐 주세요.'
  ], '입술 보습용 바셀린을'));
  out = out.concat(top(SC, '두피'), health(SC, '두피', '두피 문신(반영구)',
    yn(SC, '모발 이식을 받으신 적이 있나요?', '횟수와 시기를 알려 주세요').concat(
    yn(SC, '두피염, 지루성 피부염, 두피 감염이 있나요?', '어떤 상태인지 알려 주세요'))), tail(SC, [
    '시술 24시간 뒤부터 샴푸해 주세요. 손에서 거품을 충분히 낸 뒤 자극 없이 감고 시원한 물로 헹궈 주세요.',
    '회복 기간(최소 1주) 동안 음주와 흡연, 사우나, 땀이 많이 나는 운동은 염증을 일으킬 수 있어 자제해 주세요.',
    '각질은 저절로 떨어지게 두세요. 손으로 떼거나 문지르면 색이 고르지 않게 되거나 염증이 생길 수 있어요.',
    '첫 샴푸 때 머리카락에 묻어 있던 색소가 염색물처럼 빠질 수 있어요. 두피에 넣은 색소가 빠지는 게 아니니 걱정 마시고, 어두운 수건을 써 주세요.',
    '두피에 남은 색소가 베개에 묻을 수 있어요. 어두운 수건을 깔고 주무세요.',
    '염색이나 펌은 2주 이상 피해 주세요. (가능하면 리터치 시술 1달 뒤부터)',
    '1달 동안 색이 변하는 건 자연스러운 과정이에요. 시간을 두고 지켜봐 주세요.'
  ], '색소가 묻지 않게 막아 주는 방수 베개 커버를'));
  var G = FORM_COMMON, M = '나이<19';
  out = out.concat([
    { shop: '', form: G, q: '보호자 동의 (만 19세 미만)\n만 19세 미만 고객님은 보호자 동의가 꼭 필요해요. (「문신사법」 제22조)\n방문하실 때 고객님과 보호자 분의 신분증, 그리고 가족관계증명서처럼 관계를 확인할 수 있는 서류를 함께 가져와 주세요.', type: '안내', opts: [], need: false, cond: M, help: '' },
    { shop: '', form: G, q: '보호자 성함', type: '단답', opts: [], need: true, cond: M, help: '' },
    { shop: '', form: G, q: '보호자 생년월일', type: '단답', opts: [], need: true, cond: M, help: '예: 1980-01-01' },
    { shop: '', form: G, q: '고객님과의 관계', type: '단일선택', opts: ['부', '모', '법정대리인(후견인)', '기타'], need: true, cond: M, help: '' },
    { shop: '', form: G, q: '보호자 연락처', type: '전화', opts: [], need: true, cond: M, help: '' },
    { shop: '', form: G, q: '보호자 주소', type: '단답', opts: [], need: false, cond: M, help: '' },
    { shop: '', form: G, q: '보호자 동의', type: '동의', opts: ['본인은 미성년자의 보호자로서, 이 시술의 영구성, 제거의 어려움, 생길 수 있는 부작용에 대해 충분한 설명을 듣고 이해했으며, 미성년자에게 시술하는 것에 동의합니다.'], need: true, cond: M, help: '' },
    { shop: '', form: G, q: '보호자 서명', type: '서명', opts: [], need: true, cond: M, help: '' }
  ]);
  return out;
}
function api_seedPmuForms() {
  try {
    if (CUR.role !== 'owner') return { ok: false, msg: '대표만 쓸 수 있습니다.' };
    var shop = lineToShop('닥뷰');
    var fresh = pmuForms_(shop);
    var names = {}; fresh.forEach(function (f) { names[f.form] = 1; });
    var keep = readForms().filter(function (f) { return !names[f.form]; });
    var n = writeForms_(keep.concat(fresh));
    return { ok: true, shop: shop, total: n, added: fresh.length, kept: keep.length };
  } catch (e) { logError('동의서 양식 넣기', e); return { ok: false, msg: e.message }; }
}

/* 링크 양식 목록: 공통 보호자 동의는 따로 고르지 않습니다(자동으로 붙음) */
function formNamesFor(shop) {
  var seen = {}, out = [];
  readForms().forEach(function (f) {
    if (f.form === FORM_COMMON || f.form === MEMBER_FORM) return;
    if (f.shop && f.shop !== shop) return;
    if (!seen[f.form]) { seen[f.form] = true; out.push(f.form); }
  });
  /* 새 시술별 양식을 앞으로 */
  out.sort(function (a, b) { return (/상담·동의서$/.test(b) ? 1 : 0) - (/상담·동의서$/.test(a) ? 1 : 0); });
  return out;
}
/* 예약 메뉴에 눈썹·입술·두피가 들어 있으면 그 양식을 먼저 고릅니다 */
function defaultFormFor(shop, product) {
  var names = formNamesFor(shop), p = String(product || '');
  var keys = [['두피', '두피'], ['헤어라인', '두피'], ['입술', '입술'], ['립', '입술'], ['눈썹', '눈썹'], ['브로우', '눈썹']];
  for (var i = 0; i < keys.length; i++) if (p.indexOf(keys[i][0]) >= 0) {
    for (var j = 0; j < names.length; j++) if (names[j].indexOf(keys[i][1]) === 0) return names[j];
  }
  return names.length ? names[0] : '';
}

/* ---------- 두 매장 같은 시간 막기·풀기 할 일 ----------
   원장님 한 분이 두 매장을 보므로, 한 매장에 예약이 들어오면 다른 매장의 같은 시간을 네이버에서 막아야 합니다.
   네이버 예약을 직접 바꿀 수는 없어서, 막을(풀) 시간을 홈 화면 맨 위에 할 일로 띄웁니다. */
var BLOCK_KEY = '매장간 막기 확인', UNBLOCK_KEY = '매장간 풀기 확인';
function ackSet_(key) {
  try { var a = JSON.parse(PropertiesService.getScriptProperties().getProperty(key) || '[]'); return Array.isArray(a) ? a : []; }
  catch (e) { return []; }
}
function durFor_(o) {
  var best = 0, p = String(o.product || '') + ' ' + String(o.option || '');
  readMenu().forEach(function (m) {
    if (m.min && (!m.shop || m.shop === o.shop) && p.indexOf(m.name) >= 0 && m.min > best) best = m.min;
  });
  return best || Number(cfg('겹침 기본 소요시간(분)', 120)) || 120;
}
function addMin_(hhmm, min) {
  var m = /^(\d{1,2}):(\d{2})/.exec(String(hhmm || ''));
  if (!m) return '';
  var t = Number(m[1]) * 60 + Number(m[2]) + min;
  return ('0' + Math.floor(t / 60)).slice(-2) + ':' + ('0' + (t % 60)).slice(-2);
}
function otherShop_(shop) {
  var sl = shops().map(function (x) { return x.name; });
  for (var i = 0; i < sl.length; i++) if (sl[i] !== shop) return sl[i];
  return '';
}
function blockTodo_(rows, t0) {
  if (shops().length < 2) return { block: [], unblock: [] };
  var acked = ackSet_(BLOCK_KEY), un = ackSet_(UNBLOCK_KEY);
  function item(o) {
    var d = durFor_(o);
    var viaNaver = String(o.source || '').indexOf('네이버') >= 0;
    /* 네이버로 들어온 예약은 다른 매장만, 앱에서 직접 넣은 예약(전화·인스타 등)은 두 매장 모두 막아야 합니다 */
    return { resno: o.resno, shop: o.shop, other: viaNaver ? otherShop_(o.shop) : '두 매장 모두', name: o.name, date: o.date,
             from: o.time, to: addMin_(o.time, d), min: d, product: o.product };
  }
  var block = rows.filter(function (o) {
    return o.status === STATUS.OK && o.date >= t0 && o.time && acked.indexOf(o.resno) < 0;
  }).map(item);
  var unblock = rows.filter(function (o) {
    return o.status === STATUS.CANCEL && o.date >= t0 && acked.indexOf(o.resno) >= 0 && un.indexOf(o.resno) < 0;
  }).map(item);
  var by = function (a, b) { return (a.date + a.from).localeCompare(b.date + b.from); };
  return { block: block.sort(by), unblock: unblock.sort(by) };
}
function api_ackBlock(resno, kind) {
  var key = kind === 'unblock' ? UNBLOCK_KEY : BLOCK_KEY;
  var a = ackSet_(key);
  if (a.indexOf(String(resno)) < 0) a.push(String(resno));
  if (a.length > 2000) a = a.slice(-2000);
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(a));
  return { ok: true };
}
/* 처음 켤 때: 이미 지난 확인을 일일이 누르지 않도록, 지금 있는 예약은 '막음'으로 표시 (대표만) */
function api_ackAllBlocks() {
  if (CUR.role !== 'owner') return { ok: false, msg: '대표만 쓸 수 있습니다.' };
  var rows = readData().map(rowToObj), t0 = ymd(new Date());
  var a = ackSet_(BLOCK_KEY);
  rows.forEach(function (o) { if (o.status === STATUS.OK && o.date >= t0 && a.indexOf(o.resno) < 0) a.push(o.resno); });
  PropertiesService.getScriptProperties().setProperty(BLOCK_KEY, JSON.stringify(a));
  return { ok: true, n: a.length };
}
function api_needAttention() {
  var rows = readData().map(rowToObj);
  var t0 = ymd(new Date());
  var paidSet = {};
  readSlog().forEach(function (g) { paidSet[g.resno] = true; });
  var bt = blockTodo_(rows, t0);
  return {
    ok: true,
    check: rows.filter(function (o) { return o.status === STATUS.CHECK; }),
    noPhone: rows.filter(function (o) { return o.status === STATUS.OK && !o.phone && !o.safe && o.date >= t0; }),
    noPaid: rows.filter(function (o) { return o.status === STATUS.DONE && !paidSet[o.resno]; }),
    pastOpen: rows.filter(function (o) { return o.status === STATUS.OK && o.date && o.date < t0; }),
    block: bt.block, unblock: bt.unblock
  };
}

/* 전화번호 정리: '+' 또는 '00'으로 시작하면 외국 번호로 보고 '+숫자' 그대로 저장
   (예: +65 9123 4567 → +6591234567). 한국 번호는 예전처럼 하이픈 넣기 */
function normPhone(s) {
  var raw = String(s || '').trim();
  var d = onlyDigits(raw);
  if (!d) return '';
  if (/^\+/.test(raw)) return '+' + d;
  if (/^00/.test(d)) return '+' + d.slice(2);
  if (d.length === 11) return d.slice(0, 3) + '-' + d.slice(3, 7) + '-' + d.slice(7);
  if (d.length === 10) return d.slice(0, 3) + '-' + d.slice(3, 6) + '-' + d.slice(6);
  return raw;
}

/* ---------- 이름 고치기: 고객DB·예약·매출내역·선불내역 이름을 한꺼번에 ----------
   네이버가 가린 이름(PH*** 등)을 실명으로 바꿀 때 씁니다. */
function renameEverywhere(cid, resno, newName) {
  newName = String(newName || '').trim();
  if (!newName) return { ok: false, msg: '이름을 넣어 주세요.' };
  if (newName.indexOf('*') >= 0) return { ok: false, msg: '별표(*) 없이 실명을 넣어 주세요.' };
  var n = { cust: 0, book: 0, sales: 0, pp: 0 };

  /* 예약: 같은 고객의 모든 예약 (고객ID가 없으면 그 예약 하나만) */
  var dsh = sheetOf(SH.DATA), dl = dsh.getLastRow();
  if (dl >= 2) {
    var v = dsh.getRange(2, 1, dl - 1, D_HEAD.length).getValues();
    for (var i = 0; i < v.length; i++) {
      var hit = cid ? String(v[i][D.CUSTID - 1]).trim() === String(cid).trim()
                    : String(v[i][D.RESNO - 1]) === String(resno);
      if (!hit) continue;
      var old = String(v[i][D.NAME - 1] || '').trim();
      if (old === newName) continue;
      dsh.getRange(i + 2, D.NAME).setValue(newName);
      var memo = String(v[i][D.MEMO - 1] || '');
      if (old.indexOf('*') >= 0 && memo.indexOf('네이버 표기') < 0) {
        dsh.getRange(i + 2, D.MEMO).setValue((memo ? memo + ' / ' : '') + '네이버 표기: ' + old);
      }
      n.book++;
    }
  }
  if (cid) {
    /* 고객DB */
    var csh = sheetOf(SH.CUST), rows = readCustomers();
    for (var j = 0; j < rows.length; j++) {
      if (String(rows[j][C.ID - 1]).trim() !== String(cid).trim()) continue;
      if (String(rows[j][C.NAME - 1]).trim() !== newName) { csh.getRange(j + 2, C.NAME).setValue(newName); n.cust++; }
      break;
    }
    /* 매출내역(5열 고객명), 선불포인트내역(3열 고객명) */
    [[SH.SLOG, SLOG_HEAD, 4, 5, 'sales'], [SH.PPLOG, PP_HEAD, 2, 3, 'pp']].forEach(function (t) {
      try {
        var sh = ss().getSheetByName(t[0]);
        if (!sh) return;
        var last = sh.getLastRow();
        if (last < 2) return;
        var vals = sh.getRange(2, 1, last - 1, t[1].length).getValues();
        for (var k = 0; k < vals.length; k++) {
          if (String(vals[k][t[2] - 1]).trim() !== String(cid).trim()) continue;
          if (String(vals[k][t[3] - 1]).trim() === newName) continue;
          sh.getRange(k + 2, t[3]).setValue(newName);
          n[t[4]]++;
        }
      } catch (e) { logError('이름 고치기 ' + t[0], e); }
    });
  }
  SpreadsheetApp.flush();
  return { ok: true, n: n };
}

function api_renameBooking(resno, name) {
  try {
    var dsh = sheetOf(SH.DATA);
    var t = buildResnoIndex()[String(resno)];
    if (!t) return { ok: false, msg: '예약을 찾지 못했습니다.' };
    var row = dsh.getRange(t.r, 1, 1, D_HEAD.length).getValues()[0];
    var cid = row[D.CUSTID - 1];
    if (!cid) {
      cid = matchOrCreateCustomer(String(name).trim(), row[D.PHONE - 1], row[D.SAFE - 1],
                                  row[D.SHOP - 1], row[D.SOURCE - 1] || '수동');
      dsh.getRange(t.r, D.CUSTID).setValue(cid);
    }
    return renameEverywhere(cid, resno, name);
  } catch (e) {
    logError('앱 예약 이름 고치기', e);
    return { ok: false, msg: e.message };
  }
}

function api_renameCustomer(id, name) {
  try { return renameEverywhere(id, '', name); }
  catch (e) { logError('앱 고객 이름 고치기', e); return { ok: false, msg: e.message }; }
}

/* ---------- 문자(알리고) 설정: 대표만 ----------
   키는 화면에 다시 보여주지 않고 뒷 4자리만 보여줍니다. 빈칸으로 저장하면 기존 값을 그대로 둡니다. */
function smsCfgView_() {
  var key = cfgStr('알리고 API키', '');
  var sk = cfgStr('솔라피 API키', '');
  return { ok: true, userId: cfgStr('알리고 사용자ID', ''), sender: cfgStr('발신번호', ''),
           keyTail: key ? key.slice(-4) : '', test: cfgOn('테스트 모드'), ready: aligoReady(),
           provider: solapiReady_() ? '솔라피' : '알리고', solKeyTail: sk ? sk.slice(-4) : '', solSecretSet: !!cfgStr('솔라피 시크릿', ''),
           auto: { advance: cfgOn('사전 알림 사용'), advDays: advanceDays(), advTime: cfgStr('사전 알림 시각', '14:00'),
                   confirm: cfgOn('확정 즉시 안내 사용'), tomorrow: cfgOn('전일 알림 사용'), today: cfgOn('당일 알림 사용'),
                   retouchYear: cfgOn('1년 리터치 안내 사용') } };
}
function api_smsCfg() {
  if (CUR.role !== 'owner') return { ok: false, msg: '대표만 볼 수 있어요.' };
  return smsCfgView_();
}
function api_saveSmsCfg(p) {
  try {
    if (CUR.role !== 'owner') return { ok: false, msg: '대표만 바꿀 수 있어요.' };
    p = p || {};
    if (p.userId !== undefined && String(p.userId).trim()) setCfg('알리고 사용자ID', String(p.userId).trim());
    if (p.key !== undefined && String(p.key).trim()) setCfg('알리고 API키', String(p.key).trim());
    if (p.sender !== undefined && String(p.sender).trim()) setCfg('발신번호', String(p.sender).trim());
    if (p.test !== undefined) setCfg('테스트 모드', p.test ? '사용' : '미사용');
    if (p.solKey !== undefined && String(p.solKey).trim()) setCfg('솔라피 API키', String(p.solKey).trim());
    if (p.solSecret !== undefined && String(p.solSecret).trim()) setCfg('솔라피 시크릿', String(p.solSecret).trim());
    if (p.auto) {
      var A = { advance: '사전 알림 사용', tomorrow: '전일 알림 사용', today: '당일 알림 사용', confirm: '확정 즉시 안내 사용', retouchYear: '1년 리터치 안내 사용' };
      Object.keys(A).forEach(function (k) { if (p.auto[k] !== undefined) setCfg(A[k], p.auto[k] ? '사용' : '미사용'); });
    }
    _cfgCache = null;
    return smsCfgView_();
  } catch (e) { logError('문자 설정 저장', e); return { ok: false, msg: e.message }; }
}
/* 시험 발송: 테스트 모드와 상관없이 입력한 번호로 한 통 실제 발송하고, 알리고 답을 그대로 보여줍니다 */
function api_testSms(to) {
  try {
    if (CUR.role !== 'owner') return { ok: false, msg: '대표만 쓸 수 있어요.' };
    if (!aligoReady()) return { ok: false, msg: '문자 업체 키와 발신번호를 먼저 저장해 주세요.' };
    var d = onlyDigits(to);
    if (d.length < 10) return { ok: false, msg: '받을 번호를 확인해 주세요.' };
    var r = sendSms(d, '[뷰로그] 문자 연결 시험 발송입니다.');
    var j = {}; try { j = JSON.parse(r.raw); } catch (e) {}
    return { ok: r.ok, msg: r.ok ? '보냈어요. 휴대폰에서 받았는지 확인해 주세요.' : ('알리고 답: ' + (j.message || r.raw || '알 수 없음')) };
  } catch (e) { return { ok: false, msg: e.message }; }
}

/* =========================================================
 *  자동 안내 문자 (2026-11-20)
 *  - 당신을그리다: 팅커뷰가 보내므로 자동 안내(확정·사전·전일·당일)는 건너뜁니다. 개별 발송은 그대로.
 *  - 닥터뷰티랩: 예약 확정 직후 + 전날 오후 2시, 아래 문구로 보냅니다.
 * ========================================================= */
var APP_BASE_URL = 'https://crm.drawyourlife.co.kr';
function autoOffShop(shopName) { return shopNo(shopName) === 1; }

var DBL_TPL = {
  '확정': "[닥터뷰티랩] 예약이 확인되었습니다\n\n#{고객명}님, 예약해 주셔서 감사합니다.\n▶ 일시: #{날짜요일} #{시각}\n같은 시간에 예약이 겹친 경우에는 따로 연락드리겠습니다.\n\n■ 상담지 (방문 전 작성 필수)\n보건복지부 「문신시술 표준지침」에 따라 시술 전 건강 상태를 확인합니다.\n방문 전에 아래 상담지를 꼭 작성해 주십시오.\n확인이 필요한 경우 따로 연락드리겠습니다.\n#{상담지링크}\n링크 작성이 어려우신 분은 매장에서 작성하셔도 됩니다.\n\n■ 시술 부위에 예전 문신 자국(잔흔)이 있으신가요?\n처음 시술받으시는 분은 해당되지 않습니다.\n잔흔이 있으시면 시술 전에 반드시 사전 상담이 필요합니다.\n시술 부위 사진을 네이버 톡톡으로 보내 주십시오.\n사진 보내기가 어려우시면 전화 주셔도 됩니다.\n톡톡: https://talk.naver.com/ct/w5xote\n전화: 010-2722-3093\n\n■ 아래에 해당되시면 미리 연락 주십시오\n· 임신 중이거나 수유 중인 경우\n· 피를 묽게 하는 약을 드시는 경우\n· 당뇨로 상처가 잘 아물지 않는 경우\n· 켈로이드 피부, 여드름이 심한 피부, 트러블이 많은 두피\n\n■ 예약 안내\n· 100% 예약제로 운영합니다. 손님 한 분을 위해 시간을 비워 두고 준비합니다.\n· 예약 변경은 예약일 3일 전까지 1회 가능합니다.\n· 취소 시 환불은 네이버 예약 취소·환불 규정에 따릅니다.\n  - 7일 전까지: 전액 환불\n  - 6~5일 전: 30% 차감\n  - 4~3일 전: 50% 차감\n  - 2일 전부터 당일: 취소 불가\n· 연락 없이 오지 않으시면 예약이 취소되며, 예약금은 환불되지 않습니다.\n· 15분 이상 늦으시면 다음 예약 손님을 위해 예약이 자동 취소되며, 예약금은 환불되지 않습니다.\n· 잔흔이나 위 내용을 미리 알려 주시지 않고 방문하셔서 당일 시술이 어려운 경우에도 예약금은 돌려드리기 어렵습니다.\n\n■ 예약일이 가까운 날 예약하신 분\n예약일 7일 안으로 예약하셨다면, 이 문자를 받으신 당일 안에 연락 주시면 위약금 없이 변경·취소해 드립니다.\n\n■ 디자인 안내\n닥터뷰티랩은 얼굴에 자연스럽게 어울리는 시술을 지향합니다.\n과하거나 결과가 좋지 않을 것으로 판단되는 디자인은 시술이 어려운 점 양해 부탁드립니다.\n상담을 통해 가장 잘 어울리는 디자인을 함께 찾아 드리겠습니다.\n\n시술 전날 주의사항을 한 번 더 보내 드리겠습니다.",
  '전일': "[닥터뷰티랩] 내일 예약 안내\n\n#{고객명}님, 내일 뵙겠습니다.\n▶ 일시: #{날짜요일} #{시각}\n\n■ 오늘 밤\n· 음주는 삼가시고 푹 주무시고 오십시오.\n\n■ 내일 오시기 전에\n· 커피 등 카페인은 줄여 주십시오.\n· 두피 시술은 머리를 감고, 어두운 색의 편한 옷을 입고 오십시오.\n\n■ 상담지\n아직 작성하지 않으셨다면 방문 전에 작성해 주십시오.\n#{상담지링크}\n\n■ 주차\n· 건물 지하 주차장 2시간 무료\n· 2인 이상이거나 2부위 이상 시술하시면 이너매스마곡1에 주차해 주십시오.\n\n■ 오시는 길\nhttps://naver.me/GL80O8yd\n문의 010-2722-3093"
};

/* 예약 1건의 상담지 링크: 아직 안 쓴 링크가 살아 있으면 그대로, 없으면 새로 만듭니다 */
function consentUrlFor_(v) {
  try {
    var resno = String(v[D.RESNO - 1] || '');
    if (!resno) return '';
    var now = nowStr(), found = '';
    linkRows().v.forEach(function (r) {
      if (String(r[6]) === resno && String(r[3]) === '대기' && String(r[2]) >= now) found = String(r[0]);
    });
    if (found) return APP_BASE_URL + '/s/' + found;
    var saved = CUR; if (!CUR || !CUR.name) CUR = { role: 'owner', name: '자동발송' };
    var res = api_consentLink({ resno: resno });
    CUR = saved;
    if (res && res.ok) {
      var m = String(res.url || res.path || res.link || '').match(/\/s\/[A-Za-z0-9]+/);
      if (m) return APP_BASE_URL + m[0];
    }
    /* 응답 형식이 달라도 방금 만든 줄에서 토큰을 찾습니다 */
    var tok = '';
    linkRows().v.forEach(function (r) { if (String(r[6]) === resno && String(r[3]) === '대기') tok = String(r[0]); });
    return tok ? APP_BASE_URL + '/s/' + tok : '';
  } catch (e) { logError('상담지 링크 만들기', e); return ''; }
}

function buildMessage(v, type) {
  var base = { '사전': '문구_사전', '전일': '문구_전일', '당일': '문구_당일', '확정': '문구_확정', '개별': '문구_확정' }[type] || '문구_확정';
  var shop = v[D.SHOP - 1] || '';
  var tpl = (shopNo(shop) === 2 && DBL_TPL[type]) ? DBL_TPL[type] : '';
  if (!tpl) tpl = String(cfgStr('샵' + shopNo(shop) + ' ' + base, '')).trim();
  if (!tpl) tpl = String(cfgStr(base, '')).trim();
  if (!tpl) tpl = '[#{샵}] #{이름}님, #{날짜} #{시간} 예약 안내드립니다.';
  var name = v[D.NAME - 1] || '고객';
  var time = timeStr(v[D.TIME - 1]);
  if (tpl.indexOf('#{상담지링크}') >= 0) {
    var url = consentUrlFor_(v);
    tpl = tpl.replace(/#\{상담지링크\}/g, url || '(링크는 매장에서 따로 보내 드립니다)');
  }
  return tpl
    .replace(/#\{샵\}/g, shop)
    .replace(/#\{상점명\}/g, shop)
    .replace(/#\{이름\}/g, name)
    .replace(/#\{고객명\}/g, name)
    .replace(/#\{날짜\}/g, ymd(v[D.DATE - 1]))
    .replace(/#\{시간\}/g, time)
    .replace(/#\{예약일시\}/g, whenText(v[D.DATE - 1], time))
    .replace(/#\{날짜요일\}/g, dayText(v[D.DATE - 1]))
    .replace(/#\{시각\}/g, clockText(time))
    .replace(/#\{소요시간\}/g, minText(bookingMinutes(v)))
    .replace(/#\{상품\}/g, v[D.PRODUCT - 1] || '')
    .replace(/#\{옵션\}/g, v[D.OPTION - 1] || '');
}

/* =========================================================
 *  2026-11-20 추가: 2회차 남음 · 눈썹 무료 리터치 기간 · 소개 손님 수 · 1년 리터치 문자 · 결제수단 나눠 넣기
 * ========================================================= */
function daysBetween_(a, b) {   /* 'yyyy-MM-dd' 두 날짜 사이 일수 (b - a) */
  var x = toDate(a), y = toDate(b);
  if (!x || !y) return null;
  return Math.round((Date.UTC(y.getFullYear(), y.getMonth(), y.getDate()) - Date.UTC(x.getFullYear(), x.getMonth(), x.getDate())) / 86400000);
}
function addDays_(d, n) { var t = toDate(d); t.setDate(t.getDate() + n); return ymd(t); }
function drShop_() { var n = ''; shops().forEach(function (x) { if (x.idx === 2) n = x.name; }); return n; }
/* 닥터뷰티랩 매출만 날짜순으로 (당신을그리다 '아이브로우' 같은 항목이 섞이지 않게) */
function slogByDate_() { var dr = drShop_(); return readSlog().filter(function (g) { return g.shop === dr; }).sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); }); }

/* 2회를 함께 결제했는데 아직 2회차 결제(0원 포함)가 없는 손님 */
function twoLeftList_() {
  var map = {};
  slogByDate_().forEach(function (g) {
    if (!g.custId || g.item === '할인') return;
    var it = String(g.item || '');
    if (/2회차/.test(it)) { if (map[g.custId]) map[g.custId].done = true; return; }
    if (/부위\s*2회/.test(it) || /2회\s*함께/.test(it)) map[g.custId] = { custId: g.custId, name: g.name, date: g.date, item: it, done: false };
  });
  return Object.keys(map).map(function (k) { return map[k]; }).filter(function (x) { return !x.done; })
    .map(function (x) { x.days = daysBetween_(x.date, ymd(new Date())); return x; })
    .sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });
}

/* 눈썹 무료 리터치: 첫 시술 후 4~8주(28~56일) 안에만 가능 */
var BROW_RT_FROM = 28, BROW_RT_TO = 56;
function isBrowRetouch_(it) { return /^\s*(눈썹|브로우)\s*리터치/.test(String(it || '')); }
function isBrowMain_(it) { it = String(it || ''); return !isBrowRetouch_(it) && /(브로우|눈썹)/.test(it); }
function browInfoAll_() {
  var map = {};
  slogByDate_().forEach(function (g) {
    if (!g.custId) return;
    if (isBrowMain_(g.item)) map[g.custId] = { custId: g.custId, name: g.name, browDate: g.date, retouched: '' };
    else if (isBrowRetouch_(g.item) && map[g.custId] && g.date >= map[g.custId].browDate) map[g.custId].retouched = g.date;
  });
  var today = ymd(new Date());
  return Object.keys(map).map(function (k) {
    var x = map[k];
    x.start = addDays_(x.browDate, BROW_RT_FROM); x.end = addDays_(x.browDate, BROW_RT_TO);
    x.days = daysBetween_(x.browDate, today);
    x.left = daysBetween_(today, x.end);
    x.state = x.retouched ? '리터치 완료' : (x.days < BROW_RT_FROM ? '기간 전' : (x.days <= BROW_RT_TO ? '기간 중' : '기간 지남'));
    return x;
  });
}
/* 홈에 띄울 손님: 리터치 안 했고, 기간 시작 1주 전부터 기간 끝난 뒤 2주까지 */
function browRetouchHome_() {
  return browInfoAll_().filter(function (x) { return !x.retouched && x.days >= BROW_RT_FROM - 7 && x.days <= BROW_RT_TO + 14; })
    .sort(function (a, b) { return String(a.end).localeCompare(String(b.end)); });
}

/* 소개: 연결 고객에 '소개받은 분'으로 들어간 사람 수 */
function referList_(c) {
  return relEntries(c.rel).filter(function (e) { return e.rel === '소개받은 분'; }).map(function (e) { return e.name; });
}
function topReferrers_() {
  return readCustomers().map(custToObj).map(function (c) { var l = referList_(c); return { id: c.id, name: c.name, n: l.length, names: l }; })
    .filter(function (x) { return x.n > 0; }).sort(function (a, b) { return b.n - a.n; }).slice(0, 30);
}

/* 고객 화면 덧붙임 정보 */
function api_custExtra(id) {
  try {
    id = String(id || '').trim();
    var c = null;
    readCustomers().forEach(function (v) { if (String(v[C.ID - 1]).trim() === id) c = custToObj(v); });
    if (!c) return { ok: false, msg: '고객을 찾지 못했습니다.' };
    var brow = null; browInfoAll_().forEach(function (x) { if (x.custId === id) brow = x; });
    var two = null; twoLeftList_().forEach(function (x) { if (x.custId === id) two = x; });
    var refs = referList_(c);
    return { ok: true, brow: brow, two: two, referred: refs.length, referredNames: refs };
  } catch (e) { return { ok: false, msg: e.message }; }
}

/* ---------- 결제수단 나눠 넣기 ----------
   결제수단 칸에 '네이버페이:180000+계좌이체:20000' 처럼 적어 둡니다 (예전 기록은 그대로 읽힘) */
function payParts_(pay) {
  pay = String(pay || '');
  if (pay.indexOf(':') < 0) return null;
  return pay.split('+').map(function (s) {
    var m = s.split(':'); return { m: String(m[0] || '').trim(), amt: Number(String(m[1] || '').replace(/[^0-9.-]/g, '')) || 0 };
  }).filter(function (x) { return x.m; });
}
function feeOf(byRes) {
  var r = { naverBase: 0, cardBase: 0, naverFee: 0, cardFee: 0, fee: 0, mixed: 0 };
  Object.keys(byRes).forEach(function (k) {
    var b = byRes[k];
    var amt = Number(b.amount || 0), np = Number(b.npay || 0);
    var pay = String(b.pay || '');
    var parts = payParts_(pay);
    if (parts) {
      parts.forEach(function (p) {
        if (p.m.indexOf('네이버') >= 0) { r.naverBase += p.amt; r.naverFee += Math.round(p.amt * npayRate() / 100); }
        else if (p.m.indexOf('카드') >= 0) { r.cardBase += p.amt; r.cardFee += Math.round(p.amt * cardRateOf(p.m) / 100); }
      });
      return;
    }
    var hasN = pay.indexOf('네이버') >= 0, hasC = pay.indexOf('카드') >= 0;
    var cr = cardRateOf(pay);
    if (hasN && hasC) {
      if (np <= 0) { r.mixed++; return; }
      var cardPart = Math.max(0, amt - np);
      r.naverBase += np;  r.naverFee += Math.round(np * npayRate() / 100);
      r.cardBase  += cardPart; r.cardFee += Math.round(cardPart * cr / 100);
    } else if (hasN) {
      var nb = np > 0 ? np : amt;
      r.naverBase += nb; r.naverFee += Math.round(nb * npayRate() / 100);
    } else if (hasC) {
      r.cardBase += amt; r.cardFee += Math.round(amt * cr / 100);
    }
  });
  r.fee = r.naverFee + r.cardFee;
  return r;
}

/* ---------- 1년 리터치 안내 문자 (닥터뷰티랩, 대표가 켠 경우만) ----------
   마지막 닥터뷰티랩 시술일로부터 정확히 1년 되는 날 오후 2시에 한 번 보냅니다. */
var RETOUCH_YEAR_TPL = "(광고)[닥터뷰티랩]\n#{고객명}님, 안녕하세요.\n시술받으신 지 1년이 되었습니다.\n색이 옅어졌다면 리터치로 처음처럼 또렷하게 관리하실 수 있습니다.\n\n상담·예약: 네이버 톡톡 https://talk.naver.com/ct/w5xote\n문의 010-2722-3093\n\n수신을 원하지 않으시면 이 번호로 '수신거부'라고 답장 주십시오.";
function sendRetouchYear() {
  try {
    if (!cfgOn('1년 리터치 안내 사용')) return { sent: 0 };
    var dr = null; shops().forEach(function (s) { if (s.idx === 2) dr = s.name; });
    if (!dr) return { sent: 0 };
    var target = ymd(new Date(Date.now() - 365 * 86400000));
    var lastBy = {};
    readSlog().forEach(function (g) { if (g.shop === dr && g.custId && (!lastBy[g.custId] || g.date > lastBy[g.custId])) lastBy[g.custId] = g.date; });
    var sent = 0;
    readCustomers().map(custToObj).forEach(function (c) {
      if (lastBy[c.id] !== target) return;
      if (/수신거부/.test(c.tag + ' ' + c.memo)) return;
      var to = c.phone || c.safe;
      var key = 'RY-' + c.id + '-' + target;
      var v = new Array(D_HEAD.length).fill('');
      v[D.RESNO - 1] = key; v[D.SHOP - 1] = dr; v[D.NAME - 1] = c.name; v[D.PHONE - 1] = to;
      if (!to) { logSend(v, '1년리터치', '-', '실패', '', '연락처 없음'); return; }
      if (alreadySent(key, '1년리터치')) return;
      var msg = RETOUCH_YEAR_TPL.replace(/#\{고객명\}/g, String(c.name || '고객').replace(/님$/, ''));
      if (cfgOn('테스트 모드')) { logSend(v, '1년리터치', '테스트', '성공(테스트)', msg, '테스트 모드'); return; }
      var res = sendSms(to, msg);
      logSend(v, '1년리터치', 'SMS', res.ok ? '성공' : '실패', msg, res.raw);
      if (res.ok) sent++;
    });
    return { sent: sent };
  } catch (e) { logError('1년 리터치 안내', e); return { sent: 0 }; }
}


/* ---------- 밤에 들어온 예약 확정 문자는 아침 9시에 (2026-11-20) ---------- */
function quietHour_() { var h = new Date().getHours(); return h >= 21 || h < 9; }
/* 15분마다 도는 자동 정산 때 같이 실행: 아직 확정 문자를 못 보낸 앞으로의 예약(최근 이틀 안에 들어온 것)을 보냅니다 */
function sendPendingConfirms() {
  try {
    if (quietHour_() || !cfgOn('확정 즉시 안내 사용')) return { sent: 0 };
    var sh = sheetOf(SH.DATA), last = sh.getLastRow();
    if (last < 2) return { sent: 0 };
    var v = sh.getRange(2, 1, last - 1, D_HEAD.length).getValues();
    var today = ymd(new Date()), since = fmt(new Date(Date.now() - 2 * 86400000), 'yyyy-MM-dd HH:mm:ss'), sent = 0;
    for (var i = 0; i < v.length; i++) {
      var o = v[i];
      if (o[D.STATUS - 1] !== STATUS.OK) continue;
      if (ymd(o[D.DATE - 1]) < today) continue;
      if (String(o[D.SOURCE - 1] || '').indexOf('네이버') < 0 && !o[D.MAILAT - 1]) continue;   /* 손으로 넣은 예약은 제외 */
      var at = String(o[D.PROCAT - 1] || '');
      if (!at || at < since) continue;
      var hh = Number(at.slice(11, 13));
      if (!(hh >= 21 || hh < 9)) continue;   /* 밤에 들어온 것만 */
      if (alreadySent(o[D.RESNO - 1], '확정')) continue;
      var r = sendForRow(sh, i + 2, '확정');
      if (r && r.ok) sent++;
    }
    return { sent: sent };
  } catch (e) { logError('밤 예약 확정 문자', e); return { sent: 0 }; }
}

/* 이미 저장된 눈썹·입술·두피 양식에 2026-11-20 변경을 한 번에 반영 (대표만, 여러 번 눌러도 같은 결과)
   - 소견서 안내 문구 빼기 (예를 골라도 화면에 따로 뜨는 말 없음)
   - 건강 문항 끝에 '건강 상태 답변 확인' 필수 동의 넣기 */
function api_patchHealthForms() {
  try {
    if (CUR.role !== 'owner') return { ok: false, msg: '대표만 쓸 수 있습니다.' };
    var list = readForms(), out = [], help = 0, added = 0;
    for (var i = 0; i < list.length; i++) {
      var f = list[i];
      if (/상담·동의서$/.test(f.form) && /소견서/.test(f.help || '')) { f.help = ''; help++; }
      if (f.form && /상담·동의서$/.test(f.form) && f.q === HEALTH_CONFIRM_Q) continue;   /* 다시 넣으므로 기존 것 빼기 */
      out.push(f);
      if (/상담·동의서$/.test(f.form) && f.q === '그 밖에 원장님께 알려 주실 내용') {
        out.push({ shop: f.shop, form: f.form, q: HEALTH_CONFIRM_Q, type: '동의', opts: [HEALTH_CONFIRM_A], need: true, cond: '', help: '' });
        added++;
      }
    }
    writeForms_(out);
    return { ok: true, helpRemoved: help, confirmAdded: added, total: out.length };
  } catch (e) { logError('건강 문항 반영', e); return { ok: false, msg: e.message }; }
}

/* 알리고가 실제로 거절한 발송을 발송로그에서 '실패'로 바로잡습니다 (store.js 가 호출) */
function markSendFailures(list) {
  var sh = sheetOf(SH.LOG), last = sh.getLastRow();
  if (last < 2) return 0;
  var v = sh.getRange(2, 1, last - 1, 10).getValues(), n = 0;
  (list || []).forEach(function (b) {
    var to = onlyDigits(b.to);
    for (var i = v.length - 1; i >= 0; i--) {
      if (onlyDigits(v[i][4]) === to && String(v[i][7]) === '성공') {
        v[i][7] = '실패'; v[i][9] = '알리고 거절: ' + b.why;
        sh.getRange(i + 2, 8).setValue('실패'); sh.getRange(i + 2, 10).setValue(v[i][9]); n++; break;
      }
    }
  });
  return n;
}

/* =========================================================
 *  솔라피 문자 (2026-11-21)
 *  알리고는 고정 IP만 허용해서 클라우드플레어에서 보낼 수 없습니다.
 *  솔라피 API키·시크릿을 넣으면 솔라피로 보냅니다(서명은 store.js 가 보낼 때 붙임).
 * ========================================================= */
var SOLAPI_SEND = 'https://api.solapi.com/messages/v4/send';
function solapiReady_() { return !!(cfgStr('솔라피 API키', '') && cfgStr('솔라피 시크릿', '') && cfgStr('발신번호', '')); }
function aligoReady() {
  if (solapiReady_()) return true;
  return !!(cfgStr('알리고 사용자ID', '') && cfgStr('알리고 API키', '') && cfgStr('발신번호', ''));
}
function sendSms(to, msg) {
  try {
    if (solapiReady_()) {
      var isLms = msg.length > 45;
      var m = { to: onlyDigits(to), from: onlyDigits(cfgStr('발신번호', '')), text: msg, type: isLms ? 'LMS' : 'SMS' };
      if (isLms) m.subject = '예약 안내';
      UrlFetchApp.fetch(SOLAPI_SEND, { method: 'post', contentType: 'application/json',
        solapi: { key: cfgStr('솔라피 API키', ''), secret: cfgStr('솔라피 시크릿', '') },
        payload: JSON.stringify({ message: m }), receiver: m.to });
      return { ok: true, raw: '솔라피로 보냄' };
    }
    var isLms2 = msg.length > 45;
    var payload = { key: cfgStr('알리고 API키', ''), user_id: cfgStr('알리고 사용자ID', ''), sender: cfgStr('발신번호', ''),
                    receiver: onlyDigits(to), msg: msg, msg_type: isLms2 ? 'LMS' : 'SMS' };
    if (isLms2) payload.title = '예약 안내';
    var r = UrlFetchApp.fetch(ALIGO_SMS, { method: 'post', payload: payload, muteHttpExceptions: true });
    var text = r.getContentText(); var j = {}; try { j = JSON.parse(text); } catch (e) {}
    return { ok: String(j.result_code) === '1', raw: text };
  } catch (e) { return { ok: false, raw: '통신오류: ' + e.message }; }
}


/* =========================================================
 *  멤버십 이용 동의서 (2026-11-23)
 *  닥터뷰티랩 선결제 멤버십을 결제할 때 받는 약관 동의·서명 양식입니다.
 *  - 기존 상담지·동의서와 같은 링크·서명 화면을 그대로 씁니다(고객 기록에 날짜별로 쌓임).
 *  - 일반 상담지 목록에는 섞이지 않고, 고객 화면의 멤버십 버튼으로만 엽니다.
 *  - 환불 조항은 공정위 「소비자분쟁해결기준」 미용업(소비자 사정 해지: 이용 금액 + 총 이용금액 10%)과
 *    방문판매법 계속거래 해지 규정에 맞춰 적었습니다. 문구를 바꾸면 MEMBER_FORM_VER 를 올리세요.
 * ========================================================= */
var MEMBER_FORM = '멤버십 이용 동의서';
var MEMBER_FORM_VER = '2026-11-23f';
function memberForm_(shop) {
  var F = MEMBER_FORM;
  function Q(q, type, opts, need, help) {
    return { shop: shop, form: F, q: q, type: type, opts: opts || [], need: !!need, cond: '', help: help || '' };
  }
  var OK = ['위 내용을 확인했으며 동의합니다.'];
  return [
    Q('닥터뷰티랩 멤버십 이용 동의서입니다.\n멤버십은 미리 결제하신 뒤, 시술을 받으실 때마다 금액을 차감해 사용하는 방식입니다. 아래 내용을 확인하시고 서명해 주시기 바랍니다.', '안내'),
    Q('성함', '단답', [], 1),
    Q('연락처', '전화', [], 1),
    Q('가입하시는 멤버십', '단일선택', [
      'ESSENTIAL · 200만원 결제 → 230만원 사용 (혜택 금액 30만원)',
      'SIGNATURE · 300만원 결제 → 360만원 사용 (혜택 금액 60만원)',
      'PRESTIGE · 500만원 결제 → 625만원 사용 (혜택 금액 125만원)'], 1),
    Q('함께 사용하실 가족 성함', '단답', [], 0, '가족과 함께 사용하실 수 있습니다. 함께 사용하실 분의 성함을 적어 주세요. (선택)'),

    Q('1. 이용 방법\n· 시술 금액표의 1회 정가로 차감됩니다.\n· 결제하신 금액부터 먼저 사용되며, 혜택 금액은 다음 방문부터 사용하실 수 있습니다.\n· 눈썹·입술·두피·리터치 등 닥터뷰티랩의 모든 시술에 사용하실 수 있습니다.\n· 멤버십은 2회 결제 할인, 재방문 할인 등 다른 할인과 함께 적용되지 않습니다.\n· 적어 주신 가족과 함께 사용하실 수 있으며, 사용하실 때마다 남은 금액을 가입하신 분께 안내해 드립니다.', '안내'),
    Q('이용 방법 확인', '동의', OK, 1),

    Q('2. 사용 기한\n· ESSENTIAL · SIGNATURE: 결제일로부터 1년\n· PRESTIGE: 결제일로부터 2년\n· 기한이 지나면 남은 금액은 소멸됩니다. 기한 안에 모두 사용해 주시기 바랍니다.', '안내'),
    Q('사용 기한 확인', '동의', OK, 1),

    Q('3. 환불 기준\n· 언제든지 해지하실 수 있습니다.\n· 환불 금액 = 결제하신 금액 − 이용하신 금액(1회 정가 기준) − 위약금(결제하신 금액의 10%)\n· 혜택 금액은 환불 대상에 포함되지 않습니다.\n· 예) 200만원 결제 후 정가 100만원 상당을 이용하고 해지하시는 경우: 200만원(결제 금액) − 100만원(이용 금액) − 20만원(위약금 10%) = 80만원 환불\n· 이용하신 금액과 위약금의 합이 결제하신 금액보다 많은 경우 환불 금액은 없습니다.\n· 환불은 결제하신 수단으로, 요청하신 날부터 영업일 3일 이내에 처리해 드립니다.\n(공정거래위원회 「소비자분쟁해결기준」 미용업 기준을 따릅니다.)', '안내'),
    Q('환불 기준 확인', '동의', OK, 1),

    Q('4. 그 밖의 사항\n· 멤버십은 가입하신 분과 적어 주신 가족만 사용하실 수 있으며, 다른 분께 양도하실 수 없습니다. 더 이용하지 않으실 경우 위 환불 기준에 따라 환불해 드립니다.\n· 시술 가격이 변경되더라도 이미 결제하신 멤버십은 가입 당시 금액표 기준으로 차감됩니다.\n· 여기에 적히지 않은 사항은 관련 법령과 「소비자분쟁해결기준」을 따릅니다.', '안내'),
    Q('위 내용을 모두 읽고 이해했으며, 멤버십 가입에 동의합니다.', '동의', ['모두 확인했으며 동의합니다.'], 1)
  ];
}
/* 양식이 없거나 문구 판이 바뀌었으면 멤버십 양식만 새로 넣습니다(다른 양식은 그대로) */
function ensureMemberForm_() {
  var props = PropertiesService.getScriptProperties();
  var all = readForms();
  var has = all.some(function (f) { return f.form === MEMBER_FORM; });
  if (has && props.getProperty('멤버십 양식 판') === MEMBER_FORM_VER) return false;
  var keep = all.filter(function (f) { return f.form !== MEMBER_FORM; });
  /* 매장 이름 표기가 달라도 열리도록 매장 칸은 비워 둡니다(멤버십 버튼은 닥터뷰티랩에만 있음) */
  writeForms_(keep.concat(memberForm_('')));
  props.setProperty('멤버십 양식 판', MEMBER_FORM_VER);
  return true;
}


  const __API = { api_login: api_login, api_savePins: api_savePins, api_setPassword: api_setPassword, api_bootstrap: api_bootstrap, api_home: api_home, api_day: api_day, api_range: api_range, api_month: api_month, api_updateBooking: api_updateBooking, api_addBooking: api_addBooking, api_searchCustomers: api_searchCustomers, api_customer: api_customer, api_saveCustomer: api_saveCustomer, api_addCustomer: api_addCustomer, api_deleteCustomer: api_deleteCustomer, api_customerFolder: api_customerFolder, api_sales: api_sales, api_menu: api_menu, api_saveMenu: api_saveMenu, api_payment: api_payment, api_savePayment: api_savePayment, api_savePp: api_savePp, api_ppLog: api_ppLog, api_staffReport: api_staffReport, api_usage: api_usage, api_blocks: api_blocks, api_saveBlock: api_saveBlock, api_deleteBlock: api_deleteBlock, api_trash: api_trash, api_restoreTrash: api_restoreTrash, api_cash: api_cash, api_saveCash: api_saveCash, api_deleteCash: api_deleteCash, api_forms: api_forms, api_saveForms: api_saveForms, api_saveConsent: api_saveConsent, api_consents: api_consents, api_uploadPhoto: api_uploadPhoto, api_sendLog: api_sendLog, api_salesReport: api_salesReport, api_checkMail: api_checkMail, api_send: api_send, api_completePast: api_completePast, api_needAttention: api_needAttention, api_rescanRequests: api_rescanRequests, api_consentLink: api_consentLink, api_esignList: api_esignList, api_esignView: api_esignView, api_photos: api_photos, api_deletePhoto: api_deletePhoto, api_sourceStats: api_sourceStats, api_importTvSales: api_importTvSales, api_backfillCharts: api_backfillCharts, api_seedPmuForms: api_seedPmuForms, api_ackBlock: api_ackBlock, api_ackAllBlocks: api_ackAllBlocks, api_renameBooking: api_renameBooking, api_renameCustomer: api_renameCustomer, api_smsCfg: api_smsCfg, api_saveSmsCfg: api_saveSmsCfg, api_testSms: api_testSms, api_custExtra: api_custExtra, api_patchHealthForms: api_patchHealthForms };
  return { __API, api, processBookingMails, autoSettle, sendReminderAdvance, sendReminderTomorrow,
           sendReminderToday, sendRetouchYear, sendPendingConfirms, markSendFailures, dailyMaintenance, setup, publicForm, publicSubmit, authOf, addPhotoRow, expiredPhotos, CUR: () => CUR };
}

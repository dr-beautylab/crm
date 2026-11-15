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

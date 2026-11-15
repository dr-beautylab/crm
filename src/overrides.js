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
      custId = row.custId; name = row.name; shop = row.shop;
    }
    if (custId && !name) {
      var c = api_customer(custId);
      if (c.ok) { name = c.customer.name; if (!shop) shop = c.customer.shop; }
    }
    if (!shop) return { ok: false, msg: '어느 샵 양식인지 알 수 없습니다.' };
    var form = String(p.form || '') || defaultFormFor(shop);
    if (!form) return { ok: false, msg: shop + ' 에 등록된 상담지·동의서 양식이 없습니다.' };
    var days = Number(cfg('동의서 링크 유효일수', 7)) || 7;
    var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    var now = new Date(), exp = new Date(now.getTime() + days * 86400000);
    linkRows().s.appendRow([token, nowStr(), fmt(exp, 'yyyy-MM-dd HH:mm:ss'), '대기', custId, name, resno, shop, form, '', CUR.name]);
    var path = '/s/' + token;
    var text = (name ? name + '님, ' : '') + '방문 전 ' + form + ' 작성 부탁드립니다.\n' + '#{링크}' + '\n(' + days + '일 안에 작성해 주세요)';
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
    buildConsentSheet().appendRow([stamp, v[4], v[5], v[7], v[8] + ' (전자서명)', lines, 'esign:' + token]);
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

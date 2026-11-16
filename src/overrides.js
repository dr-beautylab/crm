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
    var form = String(p.form || '') || defaultFormFor(shop, p.product);
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
  return { ok: true, count: n, sources: top(src), keywords: top(kw).slice(0, 30), referrers: top(ref).slice(0, 20) };
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
  var stamp = nowStr();
  sheetWithHead(ESIGN_SHEET, ESIGN_HEAD).appendRow([stamp, token, v[4], v[5], v[6], v[7], v[8],
    JSON.stringify(clean), sign, String((meta && meta.ua) || '').slice(0, 300), String((meta && meta.ip) || '')]);
  L.s.getRange(L.r, 4).setValue('완료');
  L.s.getRange(L.r, 10).setValue(stamp);
  try {
    var lines = clean.filter(function (x, i) { return x.type !== '안내' && shownQ_(qs, ans, i); }).map(function (x) {
      var val = x.type === '서명' ? (x.v ? '서명 완료' : '') :
        (Array.isArray(x.v) ? x.v.join(', ') : (x.v === true ? '동의' : (x.v === false ? '동의안함' : String(x.v || ''))));
      return '· ' + x.q + ' → ' + val;
    }).join('\n');
    buildConsentSheet().appendRow([stamp, v[4], v[5], v[7], v[8] + ' (전자서명)', lines, 'esign:' + token]);
  } catch (e) {}
  return { ok: true };
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
    L = L.concat(yn(form, 'B형·C형 간염, 후천성면역결핍증(AIDS) 등 혈액으로 옮을 수 있는 감염병으로 진단받았거나 치료 중이신가요?', '',
      '해당되면 병원 소견서를 확인한 뒤 시술 여부를 정해요.'));
    L = L.concat(yn(form, '실신한 적이 있거나 뇌전증 등 뇌신경계 질환이 있나요?', '', '해당되면 병원 소견서를 확인한 뒤 시술 여부를 정해요.'));
    L = L.concat(yn(form, '지금 복용 중인 약이 있나요? (항응고제·아스피린·면역억제제·탈모약 등)', '약 이름을 알려 주세요',
      '항응고제·아스피린처럼 출혈 위험이 있는 약은 병원 소견서를 확인한 뒤 시술 여부를 정해요.'));
    L = L.concat(yn(form, '오늘 아픈 곳이 있거나 열, 오한, 심한 피로감이 있나요?'));
    L = L.concat(yn(form, '최근 24시간 안에 술을 드셨나요?'));
    L = L.concat(yn(form, '임신 중이거나 수유 중이신가요?'));
    L.push(Q(form, '그 밖에 원장님께 알려 주실 내용', '서술', [], 0));
    return L;
  }
  function tail(form, care) {
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
  ]));
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
  ]));
  out = out.concat(top(SC, '두피'), health(SC, '두피', '두피 문신이나 반영구',
    yn(SC, '모발 이식을 받으신 적이 있나요?', '횟수와 시기를 알려 주세요').concat(
    yn(SC, '두피염, 지루성 피부염, 두피 감염이 있나요?', '어떤 상태인지 알려 주세요'))), tail(SC, [
    '시술 24시간 뒤부터 샴푸해 주세요. 손에서 거품을 충분히 낸 뒤 자극 없이 감고 시원한 물로 헹궈 주세요.',
    '회복 기간(최소 1주) 동안 음주와 흡연, 사우나, 땀이 많이 나는 운동은 염증을 일으킬 수 있어 자제해 주세요.',
    '각질은 저절로 떨어지게 두세요. 손으로 떼거나 문지르면 색이 고르지 않게 되거나 염증이 생길 수 있어요.',
    '첫 샴푸 때 머리카락에 묻어 있던 색소가 염색물처럼 빠질 수 있어요. 두피에 넣은 색소가 빠지는 게 아니니 걱정 마시고, 어두운 수건을 써 주세요.',
    '두피에 남은 색소가 베개에 묻을 수 있어요. 어두운 수건을 깔고 주무세요.',
    '염색이나 펌은 2주 이상 피해 주세요. (가능하면 리터치 시술 1달 뒤부터)',
    '1달 동안 색이 변하는 건 자연스러운 과정이에요. 시간을 두고 지켜봐 주세요.'
  ]));
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
    if (f.form === FORM_COMMON) return;
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

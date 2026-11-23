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
      text = '[닥터뷰티랩] ' + (name ? String(name).replace(/님$/, '') + '님, ' : '') + '멤버십 이용 동의서예요.\n내용을 읽어 보시고 서명해 주세요.\n#{링크}\n(' + (exp.getMonth() + 1) + '월 ' + exp.getDate() + '일까지 열려 있어요)';
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
var MEMBER_FORM_VER = '2026-11-23a';
function memberForm_(shop) {
  var F = MEMBER_FORM;
  function Q(q, type, opts, need, help) {
    return { shop: shop, form: F, q: q, type: type, opts: opts || [], need: !!need, cond: '', help: help || '' };
  }
  var OK = ['위 내용을 확인했으며 동의합니다.'];
  return [
    Q('닥터뷰티랩 멤버십 이용 동의서예요.\n멤버십은 미리 결제하시고, 시술 받으실 때마다 금액을 차감해 쓰는 방식이에요. 아래 내용을 읽고 서명해 주세요.', '안내'),
    Q('성함', '단답', [], 1),
    Q('연락처', '전화', [], 1),
    Q('가입하시는 멤버십', '단일선택', [
      'ESSENTIAL · 200만원 결제 → 230만원 사용 (30만원 추가 제공)',
      'SIGNATURE · 300만원 결제 → 360만원 사용 (60만원 추가 제공)',
      'PRESTIGE · 500만원 결제 → 625만원 사용 (125만원 추가 제공)'], 1),
    Q('함께 쓰실 가족 성함', '단답', [], 0, '가족과 나눠 쓰실 수 있어요. 함께 쓰실 분 성함을 적어 주세요. (선택)'),

    Q('1. 이용 방법\n· 시술 금액표의 1회 정가로 차감돼요.\n· 결제하신 금액부터 먼저 쓰이고, 추가로 드린 금액은 다음 방문부터 쓰실 수 있어요.\n· 눈썹·입술·두피·리터치 등 닥터뷰티랩의 모든 시술에 쓰실 수 있어요.\n· 멤버십은 2회 결제 할인, 재방문 할인 등 다른 할인과 함께 적용되지 않아요.\n· 적어 주신 가족과 함께 쓰실 수 있고, 쓰실 때마다 남은 금액을 가입하신 분께 알려 드려요.', '안내'),
    Q('이용 방법 확인', '동의', OK, 1),

    Q('2. 사용 기한\n· ESSENTIAL · SIGNATURE: 결제일로부터 1년\n· PRESTIGE: 결제일로부터 2년\n· 기한이 지나면 추가로 드린 금액은 사라져요. 결제하신 금액 중 남은 금액은 아래 환불 기준대로 돌려 드려요.', '안내'),
    Q('사용 기한 확인', '동의', OK, 1),

    Q('3. 환불 기준\n· 언제든지 해지하실 수 있어요.\n· 환불 금액 = 결제하신 금액 − 이용하신 금액(1회 정가 기준) − 위약금(결제하신 금액의 10%)\n· 추가로 드린 금액은 혜택으로 드린 것이라 환불 금액에 들어가지 않아요.\n· 예) 200만원 결제 후 정가 100만원어치를 이용하고 해지하시면: 200만원 − 100만원 − 20만원 = 80만원 환불\n· 이용하신 금액과 위약금을 더한 금액이 결제하신 금액보다 많으면 환불 금액은 없어요.\n· 매장 사정으로 더 이용하실 수 없게 되면 위약금 없이 남은 금액을 돌려 드리고, 결제하신 금액의 10%를 더해 드려요.\n· 환불은 결제하신 수단으로, 요청하신 날부터 영업일 3일 안에 처리해 드려요.\n(공정거래위원회 「소비자분쟁해결기준」 미용업 기준을 따릅니다.)', '안내'),
    Q('환불 기준 확인', '동의', OK, 1),

    Q('4. 그 밖에\n· 다른 분께 넘기실 때는 매장에 미리 알려 주세요.\n· 시술 가격이 바뀌어도 이미 결제하신 멤버십은 가입 당시 금액표 기준으로 차감돼요.\n· 여기 적히지 않은 내용은 관련 법령과 「소비자분쟁해결기준」을 따릅니다.', '안내'),
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

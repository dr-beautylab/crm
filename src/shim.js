/* 구글 Apps Script 의 스프레드시트·서비스를 클라우드플레어(D1)에서 흉내 내는 층.
   기존 Code.gs 로직을 거의 그대로 돌리기 위한 것입니다. */

const OFF = 9 * 3600000; // 한국 시간
const _D = globalThis.Date;

/* 한국 시간 기준으로 동작하는 Date (Apps Script 의 스크립트 시간대 = Asia/Seoul 과 같게) */
export class KDate extends _D {
  constructor(...a) {
    if (a.length === 0) super();
    else if (a.length === 1) {
      const x = a[0];
      if (x instanceof _D) super(x.getTime());
      else if (typeof x === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(x.trim())) super(_D.parse(x.trim() + 'T00:00:00+09:00'));
      else super(x);
    } else {
      super(_D.UTC(a[0], a[1], a[2] ?? 1, a[3] ?? 0, a[4] ?? 0, a[5] ?? 0, a[6] ?? 0) - OFF);
    }
  }
  _k() { return new _D(this.getTime() + OFF); }
  getFullYear() { return this._k().getUTCFullYear(); }
  getMonth() { return this._k().getUTCMonth(); }
  getDate() { return this._k().getUTCDate(); }
  getDay() { return this._k().getUTCDay(); }
  getHours() { return this._k().getUTCHours(); }
  getMinutes() { return this._k().getUTCMinutes(); }
  getSeconds() { return this._k().getUTCSeconds(); }
  getTimezoneOffset() { return -540; }
  _set(fn) { const k = this._k(); fn(k); return this.setTime(k.getTime() - OFF); }
  setHours(h, m, s, ms) { return this._set(k => k.setUTCHours(h, m ?? k.getUTCMinutes(), s ?? k.getUTCSeconds(), ms ?? k.getUTCMilliseconds())); }
  setMinutes(m, s, ms) { return this._set(k => k.setUTCMinutes(m, s ?? k.getUTCSeconds(), ms ?? k.getUTCMilliseconds())); }
  setDate(d) { return this._set(k => k.setUTCDate(d)); }
  setMonth(mo, d) { return this._set(k => k.setUTCMonth(mo, d ?? k.getUTCDate())); }
  setFullYear(y, mo, d) { return this._set(k => k.setUTCFullYear(y, mo ?? k.getUTCMonth(), d ?? k.getUTCDate())); }
  static now() { return _D.now(); }
  static parse(s) { return _D.parse(s); }
  static UTC(...a) { return _D.UTC(...a); }
}

const pad = (n, w = 2) => String(n).padStart(w, '0');
export function formatDate(d, _tz, pattern) {
  const k = new _D(new _D(d).getTime() + OFF);
  const Y = k.getUTCFullYear(), M = k.getUTCMonth() + 1, D = k.getUTCDate();
  const h = k.getUTCHours(), m = k.getUTCMinutes(), s = k.getUTCSeconds();
  return String(pattern).replace(/yyyy|yy|MM|M|dd|d|HH|H|mm|ss|E/g, t => ({
    yyyy: Y, yy: pad(Y % 100), MM: pad(M), M: M, dd: pad(D), d: D, HH: pad(h), H: h, mm: pad(m), ss: pad(s),
    E: ['일', '월', '화', '수', '목', '금', '토'][k.getUTCDay()]
  })[t]);
}

/* 어떤 메서드를 불러도 조용히 자기 자신을 돌려주는 객체 (서식·캘린더 등 웹앱에 필요 없는 기능용) */
export const NOOP = new Proxy(function () {}, {
  get: (t, k) => (k === Symbol.toPrimitive ? () => '' : k === 'then' ? undefined : k === 'length' ? 0 : NOOP),
  apply: () => NOOP
});
function chain(obj) {
  return new Proxy(obj, {
    get(t, k) {
      if (k in t) { const v = t[k]; return typeof v === 'function' ? v.bind(t) : v; }
      if (typeof k === 'symbol' || k === 'then') return undefined;
      return () => chain(t); // 서식 지정류: 무시하고 이어서 쓸 수 있게
    }
  });
}

function a1(ref) {
  const m = String(ref).toUpperCase().match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
  if (!m) throw new Error('A1 표기 해석 불가: ' + ref);
  const col = s => s.split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  const r1 = +m[2], c1 = col(m[1]);
  const r2 = m[4] ? +m[4] : r1, c2 = m[3] ? col(m[3]) : c1;
  return [r1, c1, r2 - r1 + 1, c2 - c1 + 1];
}

const blank = v => v === '' || v === null || v === undefined;

class Range {
  constructor(sh, r, c, nr, nc) { this.sh = sh; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
  getRow() { return this.r; }
  getColumn() { return this.c; }
  getNumRows() { return this.nr; }
  getNumColumns() { return this.nc; }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const src = this.sh.rows[this.r - 1 + i] || [];
      const row = [];
      for (let j = 0; j < this.nc; j++) { const v = src[this.c - 1 + j]; row.push(blank(v) ? '' : v); }
      out.push(row);
    }
    return out;
  }
  getDisplayValues() { return this.getValues().map(r => r.map(v => (v instanceof _D ? formatDate(v, '', 'yyyy-MM-dd') : String(v)))); }
  getValue() { return this.getValues()[0][0]; }
  getDisplayValue() { return this.getDisplayValues()[0][0]; }
  setValues(vals) {
    for (let i = 0; i < this.nr; i++) {
      const ri = this.r - 1 + i;
      while (this.sh.rows.length <= ri) this.sh.rows.push([]);
      const row = this.sh.rows[ri];
      for (let j = 0; j < this.nc; j++) {
        const v = vals[i] ? vals[i][j] : '';
        while (row.length < this.c - 1 + j) row.push('');
        row[this.c - 1 + j] = (v instanceof _D) ? new KDate(v.getTime()) : (v === undefined || v === null ? '' : v);
      }
    }
    this.sh.touch();
    return chain(this);
  }
  setValue(v) { return this.setValues(Array.from({ length: this.nr }, () => Array(this.nc).fill(v))); }
  clearContent() { return this.setValue(''); }
  clear() { return this.clearContent(); }
  getNote() { return ''; }
  getFormula() { return ''; }
  setFormula() { return chain(this); }
  merge() { return chain(this); }
}

class Sheet {
  constructor(book, name, rows) { this.book = book; this.name = name; this.rows = rows || []; }
  touch() { this.book.dirty.add(this.name); }
  getName() { return this.name; }
  getSheetId() { return this.name; }
  getLastRow() {
    for (let i = this.rows.length - 1; i >= 0; i--) {
      const r = this.rows[i];
      if (r && r.some(v => !blank(v))) return i + 1;
    }
    return 0;
  }
  getLastColumn() { return this.rows.reduce((m, r) => { let n = 0; (r || []).forEach((v, j) => { if (!blank(v)) n = j + 1; }); return Math.max(m, n); }, 0); }
  getMaxRows() { return Math.max(this.rows.length, 1000); }
  getMaxColumns() { return Math.max(26, this.getLastColumn()); }
  getRange(a, b, c, d) {
    if (typeof a === 'string') { const [r, cc, nr, nc] = a1(a); return chain(new Range(this, r, cc, nr, nc)); }
    return chain(new Range(this, a, b, c ?? 1, d ?? 1));
  }
  getDataRange() { return this.getRange(1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }
  appendRow(arr) {
    this.rows.length = this.getLastRow();
    this.rows.push(arr.map(v => (v instanceof _D ? new KDate(v.getTime()) : (v ?? ''))));
    this.touch();
    return chain(this);
  }
  deleteRow(r) { this.rows.splice(r - 1, 1); this.touch(); return chain(this); }
  deleteRows(r, n) { this.rows.splice(r - 1, n); this.touch(); return chain(this); }
  insertRowsAfter() { return chain(this); }
  insertRowAfter() { return chain(this); }
  clear() { this.rows = []; this.touch(); return chain(this); }
  clearContents() { return this.clear(); }
  getActiveRange() { return null; }
  getActiveRangeList() { return null; }
  getProtections() { return []; }
  getParent() { return this.book.api; }
}

export class Book {
  constructor(data) {
    this.sheets = new Map();
    this.dirty = new Set();
    for (const [name, rows] of Object.entries(data || {})) this.sheets.set(name, chain(new Sheet(this, name, rows)));
    const self = this;
    this.api = chain({
      getSheetByName: n => self.sheets.get(n) || null,
      insertSheet: n => { const s = chain(new Sheet(self, n, [])); self.sheets.set(n, s); self.dirty.add(n); return s; },
      deleteSheet: s => { const n = s.getName(); self.sheets.delete(n); self.dirty.add(n); },
      getSheets: () => [...self.sheets.values()],
      getUrl: () => '',
      getId: () => 'd1',
      getName: () => '뷰티샵 CRM',
      setActiveSheet: s => s,
      getActiveSheet: () => null,
      toast: () => {}
    });
  }
  /* 시트 이름 → 2차원 배열 (D1 저장용) */
  rowsOf(name) { const s = this.sheets.get(name); return s ? s.rows : null; }
}

/* JSON 저장 시 날짜를 표시해두고, 불러올 때 KDate 로 되살립니다 */
export function encodeRows(rows) {
  return JSON.stringify(rows, function (k, v) {
    const raw = this[k];
    return raw instanceof _D ? { $d: raw.getTime() } : v;
  });
}
/* 구글 시트의 '시간만 있는 칸'은 1899-12-30 날짜로 오는데, 그 시절 서울 표준시(+8:27:52)로 계산돼
   +9시간으로 읽으면 32분이 밀립니다(13:00 → 13:32). 1905년 이전 값은 그 시각 기준 'HH:mm' 글자로 바꿉니다. */
const OLD_LIMIT = Date.UTC(1905, 0, 1), LMT_SEOUL = (8 * 3600 + 27 * 60 + 52) * 1000;
function oldTime(ms) {
  const d = new Date(ms + LMT_SEOUL), p = n => String(n).padStart(2, '0');
  return p(d.getUTCHours()) + ':' + p(d.getUTCMinutes());
}
export function decodeRows(text) {
  return JSON.parse(text, (k, v) => {
    if (v && typeof v === 'object' && '$d' in v && Object.keys(v).length === 1) {
      return v.$d < OLD_LIMIT ? oldTime(v.$d) : new KDate(v.$d);
    }
    return v;
  });
}

/* Apps Script 전역 서비스 흉내 */
export function makeServices(book, io) {
  const PROPS = '__props';
  const propSheet = () => book.api.getSheetByName(PROPS) || book.api.insertSheet(PROPS);
  const props = {
    getProperty(k) { const r = propSheet().rows.find(x => x[0] === k); return r ? String(r[1]) : null; },
    setProperty(k, v) {
      const s = propSheet(); const r = s.rows.find(x => x[0] === k);
      if (r) r[1] = String(v); else s.rows.push([k, String(v)]);
      s.touch(); return props;
    },
    deleteProperty(k) { const s = propSheet(); const i = s.rows.findIndex(x => x[0] === k); if (i >= 0) { s.rows.splice(i, 1); s.touch(); } return props; },
    getProperties() { const o = {}; propSheet().rows.forEach(x => { if (x[0]) o[x[0]] = String(x[1]); }); return o; },
    getKeys() { return propSheet().rows.map(x => x[0]); }
  };
  const cache = {
    get: k => props.getProperty('CACHE_' + k),
    put: (k, v) => props.setProperty('CACHE_' + k, v),
    remove: k => props.deleteProperty('CACHE_' + k)
  };
  const SpreadsheetApp = new Proxy({
    getActiveSpreadsheet: () => book.api,
    openById: () => book.api,
    flush: () => {},
    getUi: () => { throw new Error('웹앱에서는 화면 알림을 쓸 수 없습니다'); }
  }, { get: (t, k) => (k in t ? t[k] : NOOP) });

  const Utilities = {
    formatDate,
    getUuid: () => crypto.randomUUID(),
    sleep: () => {},
    base64Decode: s => Uint8Array.from(atob(s), c => c.charCodeAt(0)),
    base64Encode: b => btoa(String.fromCharCode(...new Uint8Array(b))),
    newBlob: (bytes, type, name) => ({ bytes, type, name, getBytes: () => bytes, getContentType: () => type, getName: () => name })
  };

  /* 외부 호출(문자 발송 등)은 모아뒀다가 요청 끝에 실제로 보냅니다 */
  const UrlFetchApp = {
    fetch(url, opt) {
      io.fetches.push({ url, opt });
      const body = JSON.stringify({ result_code: '1', code: 0, message: '발송 대기열에 넣었습니다' });
      return { getContentText: () => body, getResponseCode: () => 200 };
    }
  };

  /* 지메일: 구글 쪽 전달 스크립트가 보내준 메일만 보입니다 */
  const msgObj = m => ({
    getId: () => m.id, getSubject: () => m.subject || '', getDate: () => new KDate(m.date || Date.now()),
    getPlainBody: () => m.plain || '', getBody: () => m.html || m.plain || '', getFrom: () => m.from || '',
    getThread: () => thread
  });
  const msgs = (io.mails || []).map(msgObj);
  const thread = { getMessages: () => msgs, addLabel: () => thread, getId: () => 'ingest' };
  const GmailApp = {
    search: () => (msgs.length ? msgs.map(x => ({ getMessages: () => [x], addLabel: () => {} })) : []),
    getUserLabelByName: () => NOOP, createLabel: () => NOOP
  };

  const Logger = { log: (...a) => { io.logs.push(a.join(' ')); } };

  return {
    SpreadsheetApp, Utilities, UrlFetchApp, GmailApp, Logger,
    PropertiesService: { getScriptProperties: () => props, getUserProperties: () => props, getDocumentProperties: () => props },
    CacheService: { getScriptCache: () => cache, getUserCache: () => cache },
    ScriptApp: NOOP, CalendarApp: NOOP, DriveApp: NOOP, MailApp: NOOP, HtmlService: NOOP,
    Session: { getEffectiveUser: () => ({ getEmail: () => '' }), getScriptTimeZone: () => 'Asia/Seoul' },
    LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock: () => {}, releaseLock: () => {} }) }
  };
}

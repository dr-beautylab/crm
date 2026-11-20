/* D1 저장소: 시트 하나 = 여러 조각(행 800줄씩)으로 저장 */
import { Book, encodeRows, decodeRows } from './shim.js';
import { makeGas } from '../functions/_gas.js';

const CHUNK = 800;

export const SCHEMA = `CREATE TABLE IF NOT EXISTS sheets (
  name TEXT NOT NULL, part INTEGER NOT NULL, data TEXT NOT NULL, updated TEXT,
  PRIMARY KEY (name, part))`;

export async function loadBook(db) {
  await db.prepare(SCHEMA).run();
  const { results } = await db.prepare('SELECT name, part, data FROM sheets ORDER BY name, part').all();
  const data = {};
  for (const r of results) (data[r.name] ||= []).push(...decodeRows(r.data));
  return new Book(data);
}

export async function saveBook(db, book) {
  if (!book.dirty.size) return;
  const now = new Date().toISOString();
  const stmts = [];
  for (const name of book.dirty) {
    stmts.push(db.prepare('DELETE FROM sheets WHERE name = ?').bind(name));
    const rows = book.rowsOf(name);
    if (!rows) continue; // 지워진 시트
    let last = rows.length;
    while (last > 0 && !(rows[last - 1] || []).some(v => v !== '' && v != null)) last--;
    const trimmed = rows.slice(0, last);
    for (let i = 0, p = 0; i < Math.max(trimmed.length, 1); i += CHUNK, p++) {
      stmts.push(db.prepare('INSERT INTO sheets (name, part, data, updated) VALUES (?, ?, ?, ?)')
        .bind(name, p, encodeRows(trimmed.slice(i, i + CHUNK)), now));
    }
  }
  await db.batch(stmts);
}

/* 한 번의 요청 = 불러오기 → 원래 로직 실행 → 바뀐 시트만 저장 → 모아둔 외부호출 실행 */
export async function run(env, fn, extra = {}) {
  const book = await loadBook(env.DB);
  const io = { fetches: [], logs: [], mails: extra.mails || [] };
  const gas = makeGas(book, io);
  const result = fn(gas);
  await saveBook(env.DB, book);
  const calls = [];
  for (const f of io.fetches) {
    try {
      const o = f.opt || {};
      const body = o.payload && typeof o.payload === 'object' ? new URLSearchParams(o.payload) : o.payload;
      const headers = {};
      if (o.contentType) headers['content-type'] = o.contentType;
      if (o.solapi) headers['Authorization'] = await solapiAuth(o.solapi.key, o.solapi.secret);
      const res = await fetch(f.url, { method: (o.method || 'get').toUpperCase(), body, headers });
      const text = await res.text();
      calls.push({ url: f.url, status: res.status, text: text.slice(0, 500) });
    } catch (e) { io.logs.push('외부호출 실패: ' + e.message); calls.push({ url: f.url, status: 0, text: '통신오류: ' + e.message }); }
  }
  /* 알리고가 거절한 문자는 발송로그의 '성공'을 '실패'로 바로잡습니다 */
  const isSend = u => (/aligo/.test(u) && /send/.test(u)) || /solapi\.com\/messages/.test(u);
  const sends = io.fetches.filter(f => isSend(f.url));
  const bad = calls.filter(c => isSend(c.url)).map((c, i) => ({ c, f: sends[i] }))
    .filter(x => !smsOk(x.c))
    .map(x => { const o = (x.f && x.f.opt) || {}; const p = (o.payload && typeof o.payload === 'object') ? o.payload : {};
                return { to: String(o.receiver || p.receiver || p.receiver_1 || ''), why: smsWhy(x.c) }; });
  if (bad.length) {
    try {
      const book2 = await loadBook(env.DB);
      const gas2 = makeGas(book2, { fetches: [], logs: [], mails: [] });
      gas2.markSendFailures(bad);
      await saveBook(env.DB, book2);
    } catch (e) { io.logs.push('발송로그 고치기 실패: ' + e.message); }
  }
  return { result, logs: io.logs, calls };
}

export const json = (o, status = 200) => new Response(JSON.stringify(o), {
  status, headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' }
});

/* 구글 쪽 스크립트(메일 전달·예약 알림 시계)만 쓰는 비밀 열쇠 확인 */
export function keyOK(request, env) {
  const k = request.headers.get('x-crm-key') || '';
  return !!env.CRM_KEY && k === env.CRM_KEY;
}

/* 문자 업체 응답 판단 */
export function smsOk(c) {
  let j = {}; try { j = JSON.parse(c.text); } catch (e) {}
  if (/solapi/.test(c.url)) return c.status >= 200 && c.status < 300 && (!j.statusCode || String(j.statusCode) === '2000');
  return String(j.result_code) === '1' || String(j.code) === '0';
}
export function smsWhy(c) {
  let j = {}; try { j = JSON.parse(c.text); } catch (e) {}
  return j.message || j.errorMessage || j.statusMessage || c.text || ('HTTP ' + c.status);
}
/* 솔라피 HMAC-SHA256 서명 */
async function solapiAuth(key, secret) {
  const date = new Date().toISOString();
  const salt = crypto.randomUUID().replace(/-/g, '');
  const k = await crypto.subtle.importKey('raw', new TextEncoder().encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  const sig = await crypto.subtle.sign('HMAC', k, new TextEncoder().encode(date + salt));
  const hex = [...new Uint8Array(sig)].map(b => b.toString(16).padStart(2, '0')).join('');
  return `HMAC-SHA256 apiKey=${key}, date=${date}, salt=${salt}, signature=${hex}`;
}

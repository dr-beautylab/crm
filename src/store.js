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
  for (const f of io.fetches) {
    try {
      const o = f.opt || {};
      const body = o.payload && typeof o.payload === 'object' ? new URLSearchParams(o.payload) : o.payload;
      await fetch(f.url, { method: (o.method || 'get').toUpperCase(), body });
    } catch (e) { io.logs.push('외부호출 실패: ' + e.message); }
  }
  return { result, logs: io.logs };
}

export const json = (o, status = 200) => new Response(JSON.stringify(o), {
  status, headers: { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' }
});

/* 구글 쪽 스크립트(메일 전달·예약 알림 시계)만 쓰는 비밀 열쇠 확인 */
export function keyOK(request, env) {
  const k = request.headers.get('x-crm-key') || '';
  return !!env.CRM_KEY && k === env.CRM_KEY;
}

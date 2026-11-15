import { json, keyOK, loadBook, saveBook } from '../../src/store.js';
import { KDate } from '../../src/shim.js';

/* 구글 스프레드시트 → 새 앱으로 데이터 통째로 옮기기 (시트 단위로 덮어씀) */
export async function onRequestPost({ request, env }) {
  if (!keyOK(request, env)) return json({ ok: false, msg: '권한 없음' }, 403);
  const { sheets } = await request.json(); // { 시트이름: [[...], ...] }, 날짜는 {"$d": ms}
  const book = await loadBook(env.DB);
  const out = {};
  for (const [name, rows] of Object.entries(sheets || {})) {
    const revived = rows.map(r => r.map(v => (v && typeof v === 'object' && '$d' in v ? new KDate(v.$d) : v)));
    let s = book.api.getSheetByName(name) || book.api.insertSheet(name);
    s.rows.length = 0; s.rows.push(...revived); s.touch();
    out[name] = revived.length;
  }
  await saveBook(env.DB, book);
  return json({ ok: true, imported: out });
}

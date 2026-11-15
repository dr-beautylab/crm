import { run, json } from '../../src/store.js';

/* 사진 보기: 로그인한 사람만 */
export async function onRequestGet({ request, env }) {
  if (!env.PHOTOS) return json({ ok: false, msg: '사진 저장소가 아직 연결되지 않았습니다.' }, 503);
  const u = new URL(request.url);
  const key = u.searchParams.get('k') || '';
  const token = u.searchParams.get('t') || '';
  if (!/^photos\/[A-Za-z0-9_-]+\/\d{4}-\d{2}-\d{2}\/[A-Za-z0-9_]+\.[a-z0-9]+$/.test(key)) return json({ ok: false }, 400);
  const { result: who } = await run(env, g => g.authOf(token));
  if (!who || !who.ok) return json({ ok: false, authFail: true }, 401);
  const obj = await env.PHOTOS.get(key);
  if (!obj) return json({ ok: false, msg: '없는 사진입니다.' }, 404);
  const h = new Headers();
  obj.writeHttpMetadata(h);
  h.set('cache-control', 'private, max-age=86400');
  if (u.searchParams.get('dl')) h.set('content-disposition', `attachment; filename="${key.split('/').pop()}"`);
  return new Response(obj.body, { headers: h });
}

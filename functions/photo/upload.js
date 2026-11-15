import { run, json } from '../../src/store.js';

/* 사진 올리기: 몸통 = 파일 그대로(원본 보존), 정보는 주소 뒤에 */
export async function onRequestPost({ request, env }) {
  if (!env.PHOTOS) return json({ ok: false, msg: '사진 저장소가 아직 연결되지 않았습니다.' }, 503);
  const u = new URL(request.url);
  const q = k => (u.searchParams.get(k) || '').slice(0, 60);
  const token = request.headers.get('x-crm-token') || '';
  const { result: who } = await run(env, g => g.authOf(token));
  if (!who || !who.ok) return json({ ok: false, authFail: true, msg: '다시 로그인해 주세요.' }, 401);
  const cust = q('cust').replace(/[^A-Za-z0-9_-]/g, '');
  if (!cust) return json({ ok: false, msg: '고객 정보가 없습니다.' }, 400);
  const type = (request.headers.get('content-type') || 'application/octet-stream').split(';')[0];
  if (!/^(image|video)\//.test(type)) return json({ ok: false, msg: '사진이나 영상만 올릴 수 있습니다.' }, 400);
  const body = await request.arrayBuffer();
  if (body.byteLength > 95 * 1024 * 1024) return json({ ok: false, msg: '파일이 너무 큽니다(95MB 이하).' }, 413);
  const role = q('role') === 'thumb' ? 'thumb' : 'orig';
  const date = /^\d{4}-\d{2}-\d{2}$/.test(q('date')) ? q('date') : new Date(Date.now() + 9 * 3600e3).toISOString().slice(0, 10);
  const id = q('id').replace(/[^A-Za-z0-9]/g, '') || crypto.randomUUID().replace(/-/g, '');
  const ext = (type.split('/')[1] || 'bin').replace('jpeg', 'jpg').replace('quicktime', 'mov').slice(0, 5);
  const key = `photos/${cust}/${date}/${id}${role === 'thumb' ? '_s.jpg' : '.' + ext}`;
  await env.PHOTOS.put(key, body, { httpMetadata: { contentType: type }, customMetadata: { by: who.name || '', at: new Date().toISOString() } });
  if (role === 'orig') {
    await run(env, g => g.addPhotoRow({ custId: cust, date, kind: q('kind'), slot: q('slot'), key,
      thumb: q('thumb') ? `photos/${cust}/${date}/${id}_s.jpg` : '', size: body.byteLength, by: who.name, type }));
  }
  return json({ ok: true, key });
}

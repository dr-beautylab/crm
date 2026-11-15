import { run, json } from '../src/store.js';

export async function onRequestPost({ request, env }) {
  let body;
  try { body = await request.json(); } catch (e) { return json({ ok: false, msg: '잘못된 요청입니다.' }, 400); }
  const { fn, args, token } = body || {};
  try {
    const { result } = await run(env, gas => gas.api(fn, args || [], token || ''));
    return json(result === undefined ? null : result);
  } catch (e) {
    return json({ ok: false, msg: '서버 오류: ' + e.message }, 500);
  }
}

import { run, json, keyOK } from '../src/store.js';

/* 구글(지메일) 쪽 전달 스크립트가 네이버 예약 메일을 보내주는 입구 */
export async function onRequestPost({ request, env }) {
  if (!keyOK(request, env)) return json({ ok: false, msg: '권한 없음' }, 403);
  const { mails } = await request.json();
  const { result, logs } = await run(env, gas => gas.processBookingMails(), { mails: mails || [] });
  return json({ ok: true, stat: result, logs });
}

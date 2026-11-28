import { run } from '../../src/store.js';
import { onWebhook } from '../../src/talk.js';

/* 네이버 톡톡이 손님 메시지를 보내는 입구: /talk/1 (당신을 그리다), /talk/2 (닥터뷰티랩)
   네이버는 5초 안에 답을 받아야 하므로 바로 200을 돌려주고, 답은 뒤에서 따로 보냅니다. */
export async function onRequestPost({ request, env, params, waitUntil }) {
  const shop = Number(params.shop);
  if (shop !== 1 && shop !== 2) return new Response('not found', { status: 404 });
  const key = new URL(request.url).searchParams.get('k') || '';
  let body = {};
  try { body = await request.json(); } catch (e) { return new Response('bad', { status: 400 }); }
  const ip = request.headers.get('cf-connecting-ip') || '';
  const notify = (to, msg) => run(env, g => g.sendSms(to, msg));
  const r = await onWebhook(env, shop, key, body, ip, notify);
  if (r.later) waitUntil(r.later().catch(e => console.log('톡톡 처리 실패', e && e.message)));
  return new Response('', { status: r.status });
}
export async function onRequestGet() { return new Response('ok'); }

import { run, json, keyOK } from '../../src/store.js';

/* 구글 쪽 시계(트리거)가 정해진 시각에 부르는 입구 */
const JOBS = {
  settle: g => g.autoSettle(true),
  /* 매일 오후 2시: 사전 안내 + 전날 안내(닥터뷰티랩) 함께 */
  advance: g => { g.sendReminderAdvance(); g.sendReminderTomorrow(); },
  tomorrow: g => g.sendReminderTomorrow(),
  today: g => g.sendReminderToday(),
  daily: g => g.dailyMaintenance(),
  setup: g => g.setup()
};
export async function onRequestPost({ request, env, params }) {
  if (!keyOK(request, env)) return json({ ok: false, msg: '권한 없음' }, 403);
  const job = JOBS[params.job];
  if (!job) return json({ ok: false, msg: '알 수 없는 작업' }, 404);
  const { result, logs } = await run(env, job);
  /* 매일 정리할 때 휴지통에서 30일 지난 사진 파일도 지웁니다 */
  if (params.job === 'daily' && env.PHOTOS) {
    const { result: keys } = await run(env, g => g.expiredPhotos());
    for (const k of keys || []) { try { await env.PHOTOS.delete(k); } catch (e) {} }
  }
  return json({ ok: true, result, logs });
}

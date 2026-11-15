import { run, json, keyOK } from '../../src/store.js';

/* 구글 쪽 시계(트리거)가 정해진 시각에 부르는 입구 */
const JOBS = {
  settle: g => g.autoSettle(true),
  advance: g => g.sendReminderAdvance(),
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
  return json({ ok: true, result, logs });
}

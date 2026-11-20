import { run, json, smsOk, smsWhy } from '../src/store.js';

export async function onRequestPost({ request, env }) {
  let body;
  try { body = await request.json(); } catch (e) { return json({ ok: false, msg: '잘못된 요청입니다.' }, 400); }
  const { fn, args, token } = body || {};
  try {
    const { result, calls } = await run(env, gas => gas.api(fn, args || [], token || ''));
    /* 시험 발송은 알리고의 실제 답을 그대로 보여줍니다 */
    if (fn === 'api_testSms' && result && result.ok && calls && calls.length) {
      const c = calls[calls.length - 1];
      const ok = smsOk(c);
      return json({ ok, msg: ok ? '보냈어요. 휴대폰에서 받았는지 확인해 주세요.' : ((/solapi/.test(c.url) ? '솔라피' : '알리고') + ' 답: ' + smsWhy(c)) });
    }
    return json(result === undefined ? null : result);
  } catch (e) {
    return json({ ok: false, msg: '서버 오류: ' + e.message }, 500);
  }
}

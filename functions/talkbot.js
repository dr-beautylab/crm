import { run, json } from '../src/store.js';
import { SHOP_NAMES, getCfg, putCfg, listLogs, askAI, decide } from '../src/talk.js';

/* 앱 화면에서 쓰는 톡톡 자동 답변 설정·내역 입구 (대표만) */
const tail = s => (s ? String(s).slice(-4) : '');
function view(origin, shop, c, g, env) {
  return { ok: true, shop, name: SHOP_NAMES[shop], mode: c.mode, hookUrl: origin + '/talk/' + shop + '?k=' + c.hook,
    keyTail: tail(c.sendKey), talkId: c.talkId, info: c.info, handoff: c.handoff, holdMsg: c.holdMsg, bookUrl: c.bookUrl,
    notifyOn: !!c.notifyOn, notifyTo: c.notifyTo, testers: c.testers || [],
    aiReady: !!(env.ANTHROPIC_API_KEY || g.aiKey), aiTail: env.ANTHROPIC_API_KEY ? '(서버 비밀값)' : tail(g.aiKey) };
}

export async function onRequestPost({ request, env }) {
  let b;
  try { b = await request.json(); } catch (e) { return json({ ok: false, msg: '잘못된 요청입니다.' }, 400); }
  const { result: who } = await run(env, g => g.authOf(b.token || ''));
  if (!who || !who.ok) return json({ ok: false, authFail: true, msg: '다시 로그인해 주세요.' });
  if (who.role === 'staff') return json({ ok: false, msg: '대표만 쓸 수 있어요.' });

  const origin = new URL(request.url).origin;
  const shop = Number(b.shop);
  if (shop !== 1 && shop !== 2) return json({ ok: false, msg: '매장을 골라 주세요.' });
  const db = env.DB;
  try {
    const c = await getCfg(db, shop);
    const g = await getCfg(db, 0);
    const p = b.patch || {};

    if (b.fn === 'get') return json(view(origin, shop, c, g, env));

    if (b.fn === 'save') {
      if (p.mode !== undefined) {
        if (['off', 'test', 'on'].indexOf(p.mode) < 0) return json({ ok: false, msg: '상태 값이 이상합니다.' });
        if (p.mode !== 'off' && !c.sendKey && !String(p.sendKey || '').trim()) return json({ ok: false, msg: '톡톡 보내기 API 열쇠를 먼저 넣어 주세요.' });
        if (p.mode !== 'off' && !(env.ANTHROPIC_API_KEY || g.aiKey || String(p.aiKey || '').trim())) return json({ ok: false, msg: '인공지능 열쇠를 먼저 넣어 주세요.' });
        c.mode = p.mode;
      }
      ['info', 'handoff', 'holdMsg', 'bookUrl', 'notifyTo', 'talkId'].forEach(k => { if (p[k] !== undefined) c[k] = String(p[k]); });
      if (p.notifyOn !== undefined) c.notifyOn = !!p.notifyOn;
      if (p.sendKey !== undefined && String(p.sendKey).trim()) c.sendKey = String(p.sendKey).trim();
      if (p.bookUrl && !/^https:\/\//.test(String(p.bookUrl).trim())) return json({ ok: false, msg: '예약 링크는 https:// 로 시작해야 합니다.' });
      c.bookUrl = String(c.bookUrl || '').trim();
      if (!String(c.holdMsg || '').trim()) c.holdMsg = '확인 후 답변드리겠습니다.';
      if (p.newHook) c.hook = crypto.randomUUID().replace(/-/g, '').slice(0, 24);
      await putCfg(db, shop, c);
      if (p.aiKey !== undefined && String(p.aiKey).trim()) { g.aiKey = String(p.aiKey).trim(); await putCfg(db, 0, g); }
      return json(view(origin, shop, c, g, env));
    }

    if (b.fn === 'tester') {
      const u = String(b.user || '');
      if (!u) return json({ ok: false, msg: '대화를 골라 주세요.' });
      const set = new Set(c.testers || []);
      if (b.on) set.add(u); else set.delete(u);
      c.testers = [...set];
      await putCfg(db, shop, c);
      return json(view(origin, shop, c, g, env));
    }

    if (b.fn === 'logs') {
      const rows = await listLogs(db, shop, Math.min(Number(b.limit) || 50, 200), Number(b.before) || 0);
      return json({ ok: true, rows, testers: c.testers || [] });
    }

    /* 미리 물어보기: 실제로 보내지 않고, 손님에게 나갈 답만 보여줍니다 (인공지능 요금은 듭니다) */
    if (b.fn === 'try') {
      const text = String(b.text || '').trim();
      if (!text) return json({ ok: false, msg: '물어볼 말을 적어 주세요.' });
      if (p.info !== undefined) c.info = String(p.info);
      if (p.handoff !== undefined) c.handoff = String(p.handoff);
      const r = await askAI(env, shop, c, [], text);
      const d = decide(c, r);
      return json({ ok: true, status: d.status, category: d.category || '', text: d.text, booking: !!d.booking && !!c.bookUrl, note: d.note || '' });
    }
    return json({ ok: false, msg: '알 수 없는 요청입니다.' });
  } catch (e) {
    return json({ ok: false, msg: '서버 오류: ' + e.message }, 500);
  }
}

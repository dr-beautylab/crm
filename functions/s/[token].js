import { run, json } from '../../src/store.js';
import { I18N } from './i18n.js';

/* 손님용 상담지·동의서 작성 페이지 (로그인 없이 토큰 링크로만) */
export async function onRequestGet({ params, env }) {
  const { result } = await run(env, g => g.publicForm(params.token));
  return new Response(page(result, trFor(result)), {
    headers: { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store', 'x-robots-tag': 'noindex' }
  });
}

export async function onRequestPost({ params, env, request }) {
  let body;
  try { body = await request.json(); } catch (e) { return json({ ok: false, msg: '잘못된 요청입니다.' }, 400); }
  const meta = { ua: request.headers.get('user-agent') || '', ip: request.headers.get('cf-connecting-ip') || '' };
  meta.lang = ['en', 'ja'].includes(body && body.lang) ? body.lang : 'ko';
  const { result } = await run(env, g => g.publicSubmit(params.token, body, meta));
  return json(result);
}

/* 이 양식에 필요한 번역만 골라 페이지에 넣습니다. 번역이 거의 없는 양식(예: 사전 상담지)은 언어 버튼을 숨깁니다 */
const UI_KEYS = ['본인 서명', '손가락으로 서명해 주세요', '지우기', '제출하기', '보내는 중…', '필수 항목을 작성해 주세요',
  '보호자 서명을 해주세요', '서명을 해주세요', '인터넷 연결을 확인해 주세요', '다시 시도해 주세요', '작성이 완료되었습니다',
  '감사합니다. 방문하실 때 뵙겠습니다.', '만 {n}세', ' · 보호자 동의가 함께 필요해요', '위 내용을 확인했으며 동의합니다.'];
function trFor(f) {
  const out = {};
  if (!f || !f.ok) return { tr: out, multi: false };
  let need = 0, have = 0;
  const put = k => { if (k && I18N[k]) out[k] = I18N[k]; };
  f.questions.forEach(q => {
    if (/^lang\s*=/.test(q.cond || '')) return;
    need++; if (I18N[q.q]) have++;
    put(q.q); put(q.help); (q.opts || []).forEach(put);
  });
  put(f.form); put(f.shop); UI_KEYS.forEach(put);
  return { tr: out, multi: need > 0 && have / need >= 0.8 };
}

const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

function page(f, T) {
  const data = JSON.stringify(f || {}).replace(/</g, '\\u003c');
  const tdata = JSON.stringify(T || { tr: {}, multi: false }).replace(/</g, '\\u003c');
  return `<!doctype html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="robots" content="noindex, nofollow">
<title>${esc(f && f.ok ? f.form : '상담지')}</title>
<style>
:root{--c:#7B3F61;--bg:#f7f3f5;--line:#e5dbe0;--t:#2b2227}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--t);font:15px/1.6 -apple-system,BlinkMacSystemFont,"Apple SD Gothic Neo","Malgun Gothic","Noto Sans KR",sans-serif}
header{background:var(--c);color:#fff;padding:22px 18px 18px}header h1{margin:0;font-size:19px}header p{margin:4px 0 0;opacity:.85;font-size:13px}
main{max-width:560px;margin:0 auto;padding:14px 14px 40px}
.q{background:#fff;border:1px solid var(--line);border-radius:14px;padding:14px 14px 12px;margin:10px 0}
.q h3{margin:0 0 8px;font-size:15px;font-weight:600}.q h3 em{color:#c0392b;font-style:normal;margin-left:3px}
.q.info{background:#fbf8fa}.q.info h3{font-weight:500;white-space:pre-wrap}
label.opt{display:flex;gap:9px;align-items:flex-start;padding:7px 2px;border-top:1px solid #f1eaee}label.opt:first-of-type{border-top:0}
label.opt input{margin-top:4px;width:18px;height:18px;accent-color:var(--c);flex:none}
input[type=text],textarea{width:100%;border:1px solid var(--line);border-radius:10px;padding:10px 11px;font:inherit;background:#fff}
textarea{min-height:84px;resize:vertical}
.sign{background:#fff;border:1px solid var(--line);border-radius:14px;padding:14px;margin:16px 0}
canvas{width:100%;height:190px;border:1.5px dashed #c9b6c0;border-radius:10px;touch-action:none;background:#fff;display:block}
.row{display:flex;justify-content:space-between;align-items:center;margin-top:8px;font-size:13px;color:#7a6b73}
button{font:inherit}.clear{border:1px solid var(--line);background:#fff;border-radius:8px;padding:5px 12px}
.go{width:100%;border:0;background:var(--c);color:#fff;font-size:16px;font-weight:600;border-radius:12px;padding:15px;margin-top:6px}
.go:disabled{opacity:.6}.msg{text-align:center;padding:60px 20px}.msg h2{color:var(--c)}
.err{color:#c0392b;font-size:14px;text-align:center;min-height:20px;margin-top:8px}
.q.hide{display:none}
input[type=date],input[type=tel]{width:100%;border:1px solid var(--line);border-radius:10px;padding:10px 11px;font:inherit;background:#fff}
.q .sub{font-size:12.5px;color:#7a6b73;margin:-4px 0 8px;white-space:pre-wrap}
.age{font-size:13px;color:var(--c);margin-top:6px}
canvas.pad2{height:150px}
.langs{display:flex;gap:6px;margin-top:12px}.langs button{border:1px solid rgba(255,255,255,.55);background:transparent;color:#fff;border-radius:20px;padding:5px 12px;font-size:13px}
.langs button.on{background:#fff;color:var(--c);font-weight:600}
.ko{display:none;font-size:12px;color:#9a8b93;font-weight:400;margin-top:2px;white-space:pre-wrap}body.xl .ko{display:block}
label.opt .ko{display:none}body.xl label.opt .ko{display:inline;margin-left:6px}
</style></head><body>
<div id="app"></div>
<script>
var F = ${data};
var TT = ${tdata};
var app = document.getElementById('app');
var LANG = 'ko';
function e(s){return String(s==null?'':s).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];});}
/* 번역: 한국어 원문을 키로 찾고, 없으면 한국어 그대로 */
function T(k){ if(LANG==='ko'||!k) return k; var t=TT.tr[k]; return (t&&t[LANG])||k; }
function has(k){ return LANG!=='ko' && TT.tr[k] && TT.tr[k][LANG] && TT.tr[k][LANG]!==k; }
/* 화면 글자: 고른 언어로 쓰고, 외국어일 때는 아래에 한국어를 작게 같이 */
function X(k, long, mid){ return '<span data-k="'+e(k)+'"'+(long?' data-long="1"':'')+'>'+e(k)+'</span>'+(mid||'')+'<span class="ko" data-ko="'+e(k)+'"></span>'; }
function applyLang(){
  document.body.classList.toggle('xl', LANG!=='ko');
  document.documentElement.lang = LANG;
  [].forEach.call(document.querySelectorAll('[data-k]'), function(el){ el.textContent = T(el.getAttribute('data-k')); });
  [].forEach.call(document.querySelectorAll('[data-ko]'), function(el){
    var k=el.getAttribute('data-ko'), prev=el.parentNode && el.parentNode.querySelector('[data-long]');
    var show = has(k) && !prev;
    el.textContent = show ? k : ''; el.style.display = show ? '' : 'none';
  });
  [].forEach.call(document.querySelectorAll('.langs button'), function(b){ b.classList.toggle('on', b.getAttribute('data-l')===LANG); });
  var nm=document.getElementById('nm'); if(nm) nm.textContent = nameLine();
  document.title = T(F.form||'');
  if(window.refresh) refresh();
}
function setLang(l){ LANG=l; try{ localStorage.setItem('formLang', l); }catch(x){} applyLang(); }
function nameLine(){ var n=String(F.name||'').replace(/님$/,''); return T(F.shop)+(n?' · '+n+(LANG==='ko'?'님':''):''); }
function done(t,s){app.innerHTML='<div class="msg"><h2>'+e(T(t))+'</h2><p>'+e(T(s||''))+'</p></div>';}
function ageOf(v){ var m=/^(\\d{4})-(\\d{2})-(\\d{2})$/.exec(v||''); if(!m) return null; var n=new Date(), a=n.getFullYear()-m[1]; if(n.getMonth()+1<+m[2]||(n.getMonth()+1==+m[2]&&n.getDate()<+m[3])) a--; return a; }
var pads={};
function makePad(cv){ var o={cv:cv,drawn:false,last:null,cx:null};
  o.fit=function(){ var r=cv.getBoundingClientRect(), d=window.devicePixelRatio||1; if(!r.width) return; cv.width=r.width*d; cv.height=r.height*d; o.cx=cv.getContext('2d'); o.cx.scale(d,d); o.cx.lineWidth=2.4; o.cx.lineCap='round'; o.cx.lineJoin='round'; o.cx.strokeStyle='#1d1a1c'; o.drawn=false; };
  function pt(ev){ var r=cv.getBoundingClientRect(); return {x:ev.clientX-r.left,y:ev.clientY-r.top}; }
  cv.addEventListener('pointerdown',function(ev){ o.last=pt(ev); cv.setPointerCapture(ev.pointerId); });
  cv.addEventListener('pointermove',function(ev){ if(!o.last||!o.cx) return; var p=pt(ev); o.cx.beginPath(); o.cx.moveTo(o.last.x,o.last.y); o.cx.lineTo(p.x,p.y); o.cx.stroke(); o.last=p; o.drawn=true; });
  cv.addEventListener('pointerup',function(){ o.last=null; });
  o.fit(); return o; }
if(!F.ok){ app.innerHTML='<div class="msg"><h2>'+e(F.done?'작성 완료':'안내')+'</h2><p>'+e(F.msg||'')+'</p></div>'; }
else {
  var Q=F.questions;
  /* 손님 기기 언어로 먼저 고르고, 손님이 위 버튼으로 바꿀 수 있어요 */
  if(TT.multi){
    var saved=''; try{ saved=localStorage.getItem('formLang')||''; }catch(x){}
    var nav=(navigator.language||'').toLowerCase();
    LANG = /^(ko|en|ja)$/.test(saved) ? saved : (/^ko/.test(nav) ? 'ko' : (/^ja/.test(nav) ? 'ja' : (nav ? 'en' : 'ko')));
  }
  var h='<header><h1>'+X(F.form)+'</h1><p id="nm"></p>'+
    (TT.multi?'<div class="langs"><button type="button" data-l="ko" onclick="setLang(\\'ko\\')">한국어</button><button type="button" data-l="en" onclick="setLang(\\'en\\')">English</button><button type="button" data-l="ja" onclick="setLang(\\'ja\\')">日本語</button></div>':'')+
    '</header><main>';
  Q.forEach(function(q,i){
    var cls='q'+(q.type==='안내'?' info':'')+(q.cond?' hide':'');
    if(q.type==='안내'){ h+='<div class="'+cls+'" data-i="'+i+'"><h3>'+X(q.q, q.q.length>120)+'</h3></div>'; return; }
    h+='<div class="'+cls+'" data-i="'+i+'"><h3>'+X(q.q, false, q.need?'<em>*</em>':'')+'</h3>';
    if(q.help) h+='<div class="sub">'+X(q.help)+'</div>';
    if(q.type==='단일선택'||q.type==='복수선택'){
      var t=q.type==='단일선택'?'radio':'checkbox';
      q.opts.forEach(function(o){ h+='<label class="opt"><input type="'+t+'" name="q'+i+'" value="'+e(o)+'"><span>'+X(o)+'</span></label>'; });
    } else if(q.type==='동의'){
      h+='<label class="opt"><input type="checkbox" name="q'+i+'" value="동의"><span>'+X(q.opts[0]||'위 내용을 확인했으며 동의합니다.')+'</span></label>';
    } else if(q.type==='서술'){ h+='<textarea name="q'+i+'"></textarea>'; }
    else if(q.type==='생년월일'){ h+='<input type="date" name="q'+i+'" max="'+new Date().toISOString().slice(0,10)+'"><div class="age" id="age'+i+'"></div>'; }
    else if(q.type==='전화'){ h+='<input type="tel" inputmode="tel" name="q'+i+'" placeholder="010-0000-0000">'; }
    else if(q.type==='서명'){ h+='<canvas class="pad2" data-pad="'+i+'"></canvas><div class="row"><span>'+X('손가락으로 서명해 주세요')+'</span><button class="clear" type="button" data-clr="'+i+'" data-k="지우기">지우기</button></div>'; }
    else { h+='<input type="text" name="q'+i+'">'; }
    h+='</div>';
  });
  h+='<div class="sign"><h3 style="margin:0 0 8px;font-size:15px">'+X('본인 서명', false, '<em style="color:#c0392b;font-style:normal"> *</em>')+'</h3>'+
     '<canvas id="pad"></canvas><div class="row"><span>'+X('손가락으로 서명해 주세요')+'</span><button class="clear" id="clr" type="button" data-k="지우기">지우기</button></div></div>'+
     '<button class="go" id="go" data-k="제출하기">제출하기</button><div class="err" id="err"></div></main>';
  app.innerHTML=h;
  var main=makePad(document.getElementById('pad'));
  document.getElementById('clr').onclick=main.fit;
  function valOf(i){ var q=Q[i]; if(q.type==='안내') return '';
    if(q.type==='서명'){ var p=pads[i]; return p&&p.drawn? p.cv.toDataURL('image/png') : ''; }
    var els=document.getElementsByName('q'+i);
    if(q.type==='복수선택'){ var a=[]; els.forEach(function(x){ if(x.checked) a.push(x.value); }); return a; }
    if(q.type==='단일선택'){ var s=''; els.forEach(function(x){ if(x.checked) s=x.value; }); return s; }
    if(q.type==='동의'){ return !!(els[0]&&els[0].checked); }
    return els[0]? els[0].value.trim() : ''; }
  function byText(t){ for(var k=0;k<Q.length;k++) if(Q[k].q===t) return k; return -1; }
  function age(){ for(var k=0;k<Q.length;k++) if(Q[k].type==='생년월일') return ageOf(valOf(k)); return null; }
  function shown(i){ var c=Q[i].cond; if(!c) return true;
    var lg=/^lang\\s*=\\s*(\\w+)$/.exec(c); if(lg) return LANG===lg[1];
    var m=/^나이\\s*<\\s*(\\d+)$/.exec(c); if(m){ var a=age(); return a!==null && a<+m[1]; }
    var p=c.indexOf('='); if(p<0) return true;
    var k=byText(c.slice(0,p).trim()); if(k<0 || !shown(k)) return false;
    var want=c.slice(p+1).split('|').map(function(x){return x.trim();}), v=valOf(k);
    if(Array.isArray(v)) return v.some(function(x){return want.indexOf(x)>=0;});
    return want.indexOf(v)>=0; }
  window.refresh=function(){
    Q.forEach(function(q,i){ var box=document.querySelector('.q[data-i="'+i+'"]'); if(!box) return;
      var on=shown(i); box.classList.toggle('hide',!on);
      if(on && q.type==='서명' && !pads[i]){ pads[i]=makePad(box.querySelector('canvas')); box.querySelector('[data-clr]').onclick=pads[i].fit; }
      if(q.type==='생년월일'){ var a=ageOf(valOf(i)), el=document.getElementById('age'+i); if(el) el.textContent = a===null?'':(T('만 {n}세').replace('{n}',a)+(a<19?T(' · 보호자 동의가 함께 필요해요'):'')); }
    }); };
  app.addEventListener('change',refresh); app.addEventListener('input',refresh);
  applyLang();
  document.getElementById('go').onclick=function(){
    var err=document.getElementById('err'); err.textContent='';
    var answers=Q.map(function(q,i){ return {v: shown(i)? valOf(i) : ''}; });
    for(var i=0;i<Q.length;i++){ var q=Q[i], a=answers[i].v;
      if(q.need && q.type!=='안내' && shown(i) && (a===''||a===false||(Array.isArray(a)&&!a.length))){
        err.textContent= T(q.type==='서명' ? '보호자 서명을 해주세요' : '필수 항목을 작성해 주세요');
        var box=document.querySelector('.q[data-i="'+i+'"]'); if(box) box.scrollIntoView({behavior:'smooth',block:'center'}); return; } }
    if(!main.drawn){ err.textContent=T('서명을 해주세요'); return; }
    var b=this; b.disabled=true; b.textContent=T('보내는 중…');
    fetch(location.pathname,{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({answers:answers,sign:main.cv.toDataURL('image/png'),lang:LANG})})
      .then(function(r){return r.json();}).then(function(r){
        if(r.ok){ window.scrollTo(0,0); done('작성이 완료되었습니다','감사합니다. 방문하실 때 뵙겠습니다.'); }
        else { b.disabled=false; b.textContent=T('제출하기'); err.textContent= LANG==='ko' ? (r.msg||'다시 시도해 주세요') : T('다시 시도해 주세요')+' ('+(r.msg||'')+')'; }
      }).catch(function(){ b.disabled=false; b.textContent=T('제출하기'); err.textContent=T('인터넷 연결을 확인해 주세요'); });
  };
}
</script></body></html>`;
}

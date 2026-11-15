import { run, json } from '../../src/store.js';

/* 손님용 상담지·동의서 작성 페이지 (로그인 없이 토큰 링크로만) */
export async function onRequestGet({ params, env }) {
  const { result } = await run(env, g => g.publicForm(params.token));
  return new Response(page(result), {
    headers: { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store', 'x-robots-tag': 'noindex' }
  });
}

export async function onRequestPost({ params, env, request }) {
  let body;
  try { body = await request.json(); } catch (e) { return json({ ok: false, msg: '잘못된 요청입니다.' }, 400); }
  const meta = { ua: request.headers.get('user-agent') || '', ip: request.headers.get('cf-connecting-ip') || '' };
  const { result } = await run(env, g => g.publicSubmit(params.token, body, meta));
  return json(result);
}

const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

function page(f) {
  const data = JSON.stringify(f || {}).replace(/</g, '\\u003c');
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
</style></head><body>
<div id="app"></div>
<script>
var F = ${data};
var app = document.getElementById('app');
function e(s){return String(s==null?'':s).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];});}
function done(t,s){app.innerHTML='<div class="msg"><h2>'+e(t)+'</h2><p>'+e(s||'')+'</p></div>';}
if(!F.ok){ done(F.done?'작성 완료':'안내', F.msg); }
else {
  var h='<header><h1>'+e(F.form)+'</h1><p>'+e(F.shop)+(F.name?' · '+e(F.name)+'님':'')+'</p></header><main>';
  F.questions.forEach(function(q,i){
    if(q.type==='안내'){ h+='<div class="q info"><h3>'+e(q.q)+'</h3></div>'; return; }
    h+='<div class="q" data-i="'+i+'"><h3>'+e(q.q)+(q.need?'<em>*</em>':'')+'</h3>';
    if(q.type==='단일선택'||q.type==='복수선택'){
      var t=q.type==='단일선택'?'radio':'checkbox';
      q.opts.forEach(function(o){ h+='<label class="opt"><input type="'+t+'" name="q'+i+'" value="'+e(o)+'"><span>'+e(o)+'</span></label>'; });
    } else if(q.type==='동의'){
      h+='<label class="opt"><input type="checkbox" name="q'+i+'" value="동의"><span>위 내용을 확인했으며 동의합니다.</span></label>';
    } else if(q.type==='서술'){ h+='<textarea name="q'+i+'"></textarea>'; }
    else { h+='<input type="text" name="q'+i+'">'; }
    h+='</div>';
  });
  h+='<div class="sign"><h3 style="margin:0 0 8px;font-size:15px">서명<em style="color:#c0392b;font-style:normal"> *</em></h3>'+
     '<canvas id="pad"></canvas><div class="row"><span>손가락으로 서명해 주세요</span><button class="clear" id="clr" type="button">지우기</button></div></div>'+
     '<button class="go" id="go">제출하기</button><div class="err" id="err"></div></main>';
  app.innerHTML=h;
  var cv=document.getElementById('pad'), cx, drawn=false, last=null;
  function fit(){ var r=cv.getBoundingClientRect(), d=window.devicePixelRatio||1; cv.width=r.width*d; cv.height=r.height*d; cx=cv.getContext('2d'); cx.scale(d,d); cx.lineWidth=2.4; cx.lineCap='round'; cx.lineJoin='round'; cx.strokeStyle='#1d1a1c'; drawn=false; }
  fit();
  function pt(ev){ var r=cv.getBoundingClientRect(); return {x:ev.clientX-r.left,y:ev.clientY-r.top}; }
  cv.addEventListener('pointerdown',function(ev){ last=pt(ev); cv.setPointerCapture(ev.pointerId); });
  cv.addEventListener('pointermove',function(ev){ if(!last) return; var p=pt(ev); cx.beginPath(); cx.moveTo(last.x,last.y); cx.lineTo(p.x,p.y); cx.stroke(); last=p; drawn=true; });
  cv.addEventListener('pointerup',function(){ last=null; });
  document.getElementById('clr').onclick=fit;
  document.getElementById('go').onclick=function(){
    var err=document.getElementById('err'); err.textContent='';
    var answers=F.questions.map(function(q,i){
      if(q.type==='안내') return {v:''};
      var els=document.getElementsByName('q'+i);
      if(q.type==='복수선택'){ var a=[]; els.forEach(function(x){ if(x.checked) a.push(x.value); }); return {v:a}; }
      if(q.type==='단일선택'){ var s=''; els.forEach(function(x){ if(x.checked) s=x.value; }); return {v:s}; }
      if(q.type==='동의'){ return {v: !!(els[0]&&els[0].checked)}; }
      return {v: els[0]? els[0].value.trim() : ''};
    });
    for(var i=0;i<F.questions.length;i++){ var q=F.questions[i], a=answers[i].v;
      if(q.need && q.type!=='안내' && (a===''||a===false||(Array.isArray(a)&&!a.length))){
        err.textContent='필수 항목을 작성해 주세요';
        var box=document.querySelector('.q[data-i="'+i+'"]'); if(box) box.scrollIntoView({behavior:'smooth',block:'center'}); return; } }
    if(!drawn){ err.textContent='서명을 해주세요'; return; }
    var b=this; b.disabled=true; b.textContent='보내는 중…';
    fetch(location.pathname,{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({answers:answers,sign:cv.toDataURL('image/png')})})
      .then(function(r){return r.json();}).then(function(r){
        if(r.ok){ window.scrollTo(0,0); done('작성이 완료되었습니다','감사합니다. 방문하실 때 뵙겠습니다.'); }
        else { b.disabled=false; b.textContent='제출하기'; err.textContent=r.msg||'다시 시도해 주세요'; }
      }).catch(function(){ b.disabled=false; b.textContent='제출하기'; err.textContent='인터넷 연결을 확인해 주세요'; });
  };
}
</script></body></html>`;
}

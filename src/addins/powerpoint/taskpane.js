// --- guard: don't run twice ---
if (window.__NIS_TASKPANE_JS__) {
  console.debug("NIS: taskpane.js already loaded; skipping re-eval");
} else {
  window.__NIS_TASKPANE_JS__ = true;

// ====== (كل كود الملف الحالي يبدأ من هنا كما هو) ======
/* ======== Nano Interactive Slides - taskpane.js (per-slide + linked sequence + nano progress/cancel + caching) ======== */

/* ---------- Keys ---------- */
const NIS_KEY_PREFIX = 'NIS:scene:'; 
function nisKey(k){ return NIS_KEY_PREFIX + k; }

const NIS_STYLE_KEY_PREFIX='NIS:style:'; 
function nmStyleKey(slideKey){ return NIS_STYLE_KEY_PREFIX + (slideKey||'default'); }

const NIS_LINK_KEY_PREFIX='NIS:link:'; // {enabled,next,auto,autoMs,inherit,inheritNm}

const NIS_IMG_CACHE_PREFIX='NIS:img:'; 
function nmCacheKey(s){ return NIS_IMG_CACHE_PREFIX+[s.theme||'',s.prompt||'',String(s.seed||0),s.aspect||'16:9'].join('|'); }

/* ---------- Defaults (Simulation Controls) ---------- */
const NIS_DEFAULT_PARAMS = Object.freeze({
  speed: 50,
  capacity: 100,
  delay: 1,
  projectToSlide: false,
  projectMs: 1000,
  autoStart: false,
  stopOnChange: false
});

/* ---------- Fast in-memory cache ---------- */
const __NIS_STATE_CACHE = new Map();   // slideKey -> params
const __NIS_LINK_CACHE  = new Map();   // slideKey -> link cfg
let   __NIS_ACTIVE_SLIDE_KEY=null;
let   __NIS_GEN_ABORT = null;          // AbortController أثناء توليد الصورة

/* transient flag: when advancing via Linked Sequence, we may inherit on first visit */
let __NIS_INHERIT_NEXT = null;

/* ---------- Debounced settings save (150ms) ---------- */
let __nisSaveTimer=null, __nisSavePending=false;
function nisScheduleSave(delayMs=150){
  if(__nisSaveTimer){ __nisSavePending=true; return; }
  __nisSaveTimer=setTimeout(()=>{
    try{ Office.context.document.settings.saveAsync(()=>{}); }catch(e){}
    __nisSaveTimer=null;
    if(__nisSavePending){ __nisSavePending=false; nisScheduleSave(delayMs); }
  }, delayMs);
}

/* ---------- Persist helpers ---------- */
const NISPersist = {
  saveScene(k,d){ try{ Office.context.document.settings.set(nisKey(k), JSON.stringify(d)); nisScheduleSave(); }catch(e){} },
  loadScene(k){  try{ const r=Office.context.document.settings.get(nisKey(k)); return r?JSON.parse(r):null; }catch(e){ return null; } },
  saveLink(k,d){  try{ Office.context.document.settings.set(NIS_LINK_KEY_PREFIX+k, JSON.stringify(d)); nisScheduleSave(); }catch(e){} },
  loadLink(k){   try{ const r=Office.context.document.settings.get(NIS_LINK_KEY_PREFIX+k); return r?JSON.parse(r):null; }catch(e){ return null; } },
  cacheGet(key){ try{ return Office.context.document.settings.get(key)||null; }catch(e){ return null; } },
  cacheSet(key,v){ try{ Office.context.document.settings.set(key,v); nisScheduleSave(); }catch(e){} }
};

/* ---------- Engine Registry ---------- */
const EngineRegistry=(()=>{ const m=new Map(); 
  return {
    get(k,f){ if(!m.has(k)) m.set(k,f(k)); return m.get(k); },
    dispose(k,d){ if(m.has(k)){ try{ d(m.get(k)); }catch(e){} m.delete(k); } }
  };
})();

/* ---------- Host helpers ---------- */
function q(id){ return document.getElementById(id); }
function hostIsOffice(){ try{ return !!(Office&&Office.context&&Office.context.host); }catch(e){ return false; } }
function hostIsPowerPoint(){ try{ return Office.context.host==='PowerPoint'; }catch(e){ return false; } }

/* Show host hint (restored helper) */
function setHostHint(){
  const h=q('hostHint'); if(!h) return;
  try{
    if(hostIsOffice()) h.textContent='Host: '+Office.context.host+(hostIsPowerPoint()?' (projection enabled)':'');
    else h.textContent='Host: Browser preview';
  }catch(e){ h.textContent='Host: Browser'; }
}

/* ---------- UI params (Simulation Controls) ---------- */
function getUIParams(){
  const s=q('speed'), c=q('capacity'), d=q('delay');
  const pt=q('projectToggle'), pm=q('projectMs');
  const as=q('autoStartToggle'), st=q('stopOnChangeToggle');
  return {
    speed: s?Number(s.value):null,
    capacity: c?Number(c.value):null,
    delay: d?Number(d.value):null,
    projectToSlide: pt?!!pt.checked:false,
    projectMs: pm?Number(pm.value||1000):1000,
    autoStart: as?!!as.checked:false,
    stopOnChange: st?!!st.checked:false
  };
}
function setUIParams(p){
  const s=q('speed'), sv=q('speedVal');
  const c=q('capacity'), cv=q('capacityVal');
  const d=q('delay'), dv=q('delayVal');
  const pt=q('projectToggle'), pm=q('projectMs');
  const as=q('autoStartToggle'), st=q('stopOnChangeToggle');

  if(s && typeof p.speed==='number'){ s.value=String(p.speed); if(sv) sv.textContent=String(p.speed); s.dispatchEvent(new Event('input',{bubbles:true})); }
  if(c && typeof p.capacity==='number'){ c.value=String(p.capacity); if(cv) cv.textContent=String(p.capacity); c.dispatchEvent(new Event('input',{bubbles:true})); }
  if(d && typeof p.delay==='number'){ d.value=String(p.delay); if(dv) dv.textContent=String(p.delay); d.dispatchEvent(new Event('input',{bubbles:true})); }
  if(pt && typeof p.projectToSlide==='boolean'){ pt.checked=p.projectToSlide; }
  if(pm && typeof p.projectMs==='number'){ pm.value=String(p.projectMs); }
  if(as && typeof p.autoStart==='boolean'){ as.checked=p.autoStart; }
  if(st && typeof p.stopOnChange==='boolean'){ st.checked=p.stopOnChange; }

  const e=getActiveEngine();
  if(e){
    e.setSpeed?.(p.speed);
    e.setCapacity?.(p.capacity);
    e.setDelay?.(p.delay);
    e.setProjectToSlide?.(!!p.projectToSlide);
    if(typeof p.projectMs==='number') e.setProjectMs?.(p.projectMs);
  }
}
nmAbortInFlight('slide-change');

/* ---------- Scene persist / restore ---------- */
function persistCurrentSlide(){
  if(!__NIS_ACTIVE_SLIDE_KEY) return;
  const p = getUIParams();
  __NIS_STATE_CACHE.set(__NIS_ACTIVE_SLIDE_KEY, {...p});
  NISPersist.saveScene(__NIS_ACTIVE_SLIDE_KEY, p);
}
function restoreCurrentSlide(){
  if(!__NIS_ACTIVE_SLIDE_KEY) return;

  const cached = __NIS_STATE_CACHE.get(__NIS_ACTIVE_SLIDE_KEY);
  if(cached) setUIParams(cached);

  const persisted = NISPersist.loadScene(__NIS_ACTIVE_SLIDE_KEY);
  if(persisted){
    setUIParams(persisted);
    __NIS_STATE_CACHE.set(__NIS_ACTIVE_SLIDE_KEY, {...persisted});
    return;
  }

  const defaults={...NIS_DEFAULT_PARAMS};
  setUIParams(defaults);
  __NIS_STATE_CACHE.set(__NIS_ACTIVE_SLIDE_KEY, {...defaults});
  NISPersist.saveScene(__NIS_ACTIVE_SLIDE_KEY, defaults);
}

/* ---------- Nano style helpers (per slide) ---------- */
function nmLoadStyleForSlideKey(k){
  try{ const raw=Office.context.document.settings.get(nmStyleKey(k)); return raw?JSON.parse(raw):null; }catch(e){ return null; }
}
function nmSaveStyleForSlideKey(k,s){
  try{ Office.context.document.settings.set(nmStyleKey(k), JSON.stringify(s)); nisScheduleSave(); }catch(e){}
}

/* ---------- Fast slide id (race) ---------- */
function captureSlideKeyFast(){
  return new Promise((resolve)=>{
    let done=false; const once=(id)=>{ if(!done){ done=true; resolve(id||'default-slide'); } };

    try{
      if(window.PowerPoint && PowerPoint.run){
        PowerPoint.run(async (ctx)=>{
          const s = ctx.presentation.getSelectedSlides().getItemAt(0);
          s.load("id"); await ctx.sync();
          once(String(s.id||'default-slide'));
        }).catch(()=>{});
      }
    }catch(e){}

    try{
      Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        r=>{
          if(r.status===Office.AsyncResultStatus.Succeeded &&
             r.value && r.value.slides && r.value.slides[0] && r.value.slides[0].id){
            once(String(r.value.slides[0].id));
          }else{
            once('default-slide');
          }
        }
      );
    }catch(e){}

    setTimeout(()=>once('default-slide'), 250);
  });
}

/* ---------- Slide change wiring ---------- */
let __NIS_LINK_AUTO_TIMER = null;
function linkAutoClear(){
  if(__NIS_LINK_AUTO_TIMER){ clearTimeout(__NIS_LINK_AUTO_TIMER); __NIS_LINK_AUTO_TIMER=null; }
}

function wireSlideChange(){
  try{
    if(!(Office && Office.context && Office.context.document && Office.EventType)) return;

    let lastSlideKey = null;

    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      async ()=>{
        // persist old slide + clear timers
        persistCurrentSlide();
        linkAutoClear();

        // stop engine if requested
        const prev = getUIParams();
        if(prev.stopOnChange){
          const e = getActiveEngine();
          e?.stop?.();
        }

        const k = await captureSlideKeyFast();
        if(!k) return;
        if(k === lastSlideKey) return;

        lastSlideKey = k;
        __NIS_ACTIVE_SLIDE_KEY = k;

        // --- INHERIT (first visit) if requested ---
        if(__NIS_INHERIT_NEXT && __NIS_INHERIT_NEXT.from){
          const fromKey = __NIS_INHERIT_NEXT.from;
          __NIS_INHERIT_NEXT = null; // consume flag

          const hasScene = NISPersist.loadScene(k);
          const linkCfg  = __NIS_LINK_CACHE.get(fromKey) || linkLoad(fromKey);

          if(!hasScene && linkCfg?.inherit){
            const src = __NIS_STATE_CACHE.get(fromKey) || NISPersist.loadScene(fromKey) || {...NIS_DEFAULT_PARAMS};
            NISPersist.saveScene(k, src);
            __NIS_STATE_CACHE.set(k, {...src});
          }

          if(linkCfg?.inheritNm){
            const tgtStyle = nmLoadStyleForSlideKey(k);
            if(!tgtStyle){
              const srcStyle = nmLoadStyleForSlideKey(fromKey);
              if(srcStyle) nmSaveStyleForSlideKey(k, srcStyle);
            }
          }
        }
        // -------------------------------------------

        getActiveEngine();
        restoreCurrentSlide();
        nmInit();
        linkRestoreForSlide();   // refresh Linked Sequence UI

        const cur = getUIParams();
        if(cur.autoStart){
          const e = getActiveEngine();
          e?.start?.();
          linkAutoArmForActive(); // arm auto-advance if enabled
        }
      }
    );
  }catch(e){}
}

/* ---------- Simple internal preview engine ---------- */
function drawPreview(ctx,state){
  const w=ctx.canvas.width,h=ctx.canvas.height;
  ctx.clearRect(0,0,w,h);
  ctx.fillStyle='#f9fafb'; ctx.fillRect(0,0,w,h);
  ctx.fillStyle='#e5e7eb'; for(let i=0;i<10;i++){ ctx.fillRect(i*(w/10),0,1,h); }
  ctx.fillStyle='#111827'; ctx.fillRect(32,h-40,Math.max(20,Math.min(w-64,state.capacity)),12);
  const y=h/2; ctx.beginPath(); ctx.arc(state.x,y,12,0,Math.PI*2); ctx.fillStyle='#2563eb'; ctx.fill();
  ctx.font='14px system-ui,Segoe UI,Arial'; ctx.fillStyle='#374151';
  ctx.fillText('spd '+state.speed+'  cap '+state.capacity+'  dly '+state.delay,32,28);
}

function createInternalEngine(slideKey){
  // ensure we have a <canvas id="preview">
  let host = q('preview'); 
  let canvas = host;
  if(!canvas || typeof canvas.getContext !== 'function'){
    const c = document.createElement('canvas');
    c.id='preview'; c.width=480; c.height=220;
    if(host){ host.innerHTML=''; host.appendChild(c); }
    canvas = c;
  }
  const ctx = canvas.getContext('2d');

  let running=false, tm=null, lastProject=0;
  let state={ ...NIS_DEFAULT_PARAMS, x:40 };

  const step=()=>{
    if(!running) return;
    const v=Math.max(1,Math.floor((state.speed||50)/3));
    state.x+=v; if(state.x>canvas.width-40) state.x=40;

    if(__NIS_ACTIVE_SLIDE_KEY===slideKey){
      drawPreview(ctx,state);
      const now=Date.now();
      if(state.projectToSlide && hostIsPowerPoint() && now-lastProject>=state.projectMs){
        projectCanvas(canvas); lastProject=now;
      }
    }
    const tickMs=Math.max(5,200-(state.speed||50)*1.5)+(state.delay||0)*100;
    tm=setTimeout(step,tickMs);
  };

  return {
    start(){ if(running) return; running=true; step(); },
    stop(){ running=false; if(tm){ clearTimeout(tm); tm=null; } },
    setSpeed(v){ state.speed=v; },
    setCapacity(v){ state.capacity=v; },
    setDelay(v){ state.delay=v; },
    setProjectToSlide(v){ state.projectToSlide=!!v; },
    setProjectMs(v){ state.projectMs=Number(v)||1000; },
    reset(){ state={ ...NIS_DEFAULT_PARAMS, x:40 }; 
             if(__NIS_ACTIVE_SLIDE_KEY===slideKey) drawPreview(ctx,state); },
    snapshot(){ if(__NIS_ACTIVE_SLIDE_KEY===slideKey){ drawPreview(ctx,state); if(hostIsPowerPoint()) projectCanvas(canvas); } },
    download(){ downloadPNG(); },
    export(){ exportJSON(); },
    import(file){ importJSON(file); }
  };
}

function createEngineForSlide(slideKey){
  if(typeof window.NIS_createEngine==='function') return window.NIS_createEngine(slideKey);
  if(typeof window.createEngine==='function') return window.createEngine(slideKey);
  return createInternalEngine(slideKey);
}
function getActiveEngine(){
  if(!__NIS_ACTIVE_SLIDE_KEY) return null;
  return EngineRegistry.get(__NIS_ACTIVE_SLIDE_KEY, createEngineForSlide);
}

/* ---------- Simulation UI bindings ---------- */
function bindSimUI(){
  const startBtn=q('start'), stopBtn=q('stop'), resetBtn=q('reset');
  const snapBtn=q('snapshot'), expBtn=q('exportJson');
  const impBtn=q('importJson'), impFile=q('jsonFile'), pngBtn=q('downloadPng');
  const s=q('speed'), sv=q('speedVal');
  const c=q('capacity'), cv=q('capacityVal');
  const d=q('delay'), dv=q('delayVal');
  const pt=q('projectToggle'), pm=q('projectMs');

  if(startBtn){ startBtn.addEventListener('click',()=>{ const e=getActiveEngine(); e?.start?.(); linkAutoArmForActive(); }); }
  if(stopBtn){  stopBtn .addEventListener('click',()=>{ const e=getActiveEngine(); e?.stop?.();  linkAutoClear(); }); }
  if(resetBtn){ resetBtn.addEventListener('click',()=>{ setUIParams({...NIS_DEFAULT_PARAMS}); const e=getActiveEngine(); e?.reset?.(); persistCurrentSlide(); }); }

  if(s){ s.addEventListener('input',()=>{ const v=+s.value; if(sv) sv.textContent=String(v); const e=getActiveEngine(); e?.setSpeed?.(v); });
        s.addEventListener('change',()=>{ persistCurrentSlide(); }); }
  if(c){ c.addEventListener('input',()=>{ const v=+c.value; if(cv) cv.textContent=String(v); const e=getActiveEngine(); e?.setCapacity?.(v); });
        c.addEventListener('change',()=>{ persistCurrentSlide(); }); }
  if(d){ d.addEventListener('input',()=>{ const v=+d.value; if(dv) dv.textContent=String(v); const e=getActiveEngine(); e?.setDelay?.(v); });
        d.addEventListener('change',()=>{ persistCurrentSlide(); }); }

  if(pt){ pt.addEventListener('change',()=>{ const v=!!pt.checked; const e=getActiveEngine(); e?.setProjectToSlide?.(v); persistCurrentSlide(); }); }
  if(pm){ pm.addEventListener('input',()=>{ const v=Number(pm.value)||1000; const e=getActiveEngine(); e?.setProjectMs?.(v); });
        pm.addEventListener('change',()=>{ persistCurrentSlide(); }); }

  if(snapBtn){ snapBtn.addEventListener('click',()=>{ const e=getActiveEngine(); e?.snapshot?.(); }); }
  if(expBtn){  expBtn .addEventListener('click',()=>{ const e=getActiveEngine(); e?.export?.();  }); }
  if(impBtn){  impBtn .addEventListener('click',()=>{ if(impFile) impFile.click(); }); }
  if(impFile){ impFile.addEventListener('change',ev=>{ const f=ev.target.files&&ev.target.files[0]; if(f){ const e=getActiveEngine(); e?.import?.(f); impFile.value=''; } }); }
  if(pngBtn){  pngBtn .addEventListener('click',()=>{ const e=getActiveEngine(); e?.download?.(); }); }
}

/* ---------- Project/export helpers ---------- */
function projectCanvas(canvas){
  try{
    const dataUrl=canvas.toDataURL('image/png');
    const base64=dataUrl.split(',')[1];
    Office.context.document.setSelectedDataAsync(base64,{coercionType:Office.CoercionType.Image},()=>{});
  }catch(e){}
}
function exportJSON(){
  if(!__NIS_ACTIVE_SLIDE_KEY) return;
  const payload={ slideKey:__NIS_ACTIVE_SLIDE_KEY, params:getUIParams(), ts:Date.now() };
  const blob=new Blob([JSON.stringify(payload,null,2)],{type:'application/json'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download='nis-scene-'+__NIS_ACTIVE_SLIDE_KEY+'.json'; a.click(); URL.revokeObjectURL(a.href);
}
function importJSON(file){
  const r=new FileReader();
  r.onload=()=>{ try{ const data=JSON.parse(r.result); const p=data&&data.params?data.params:data; setUIParams(p); persistCurrentSlide(); const e=getActiveEngine(); e?.reset?.(); }catch(e){} };
  r.readAsText(file);
}
function downloadPNG(){
  let canvas=q('preview');
  if(canvas && typeof canvas.toDataURL==='function'){
    const a=document.createElement('a'); a.href=canvas.toDataURL('image/png');
    a.download='nis-slide-'+(__NIS_ACTIVE_SLIDE_KEY||'default')+'.png'; a.click();
  }
}

/* ---------- Nano Mode (with Progress/Cancel/Cache) ---------- */
function nmGetStyle(){
  try{
    const key=nmStyleKey(__NIS_ACTIVE_SLIDE_KEY);
    const raw=Office.context.document.settings.get(key);
    return raw ? JSON.parse(raw) : {theme:'',seed:42,prompt:'',aspect:'16:9',caption:false,autoInc:true};
  }catch(e){ return {theme:'',seed:42,prompt:'',aspect:'16:9',caption:false,autoInc:true}; }
}
function nmSaveStyle(s){
  try{
    const key=nmStyleKey(__NIS_ACTIVE_SLIDE_KEY);
    Office.context.document.settings.set(key, JSON.stringify(s));
    nisScheduleSave();
  }catch(e){}
}
function nmReadInputs(){
  return {
    theme: q('nmTheme')?.value||'',
    seed:  q('nmSeed') ? Number(q('nmSeed').value||0) : 0,
    prompt:q('nmPrompt')?.value||'',
    aspect:q('nmAspect')?.value||'16:9',
    caption:q('nmCaption')?q('nmCaption').checked:false,
    autoInc:q('nmAutoInc')?q('nmAutoInc').checked:true
  };
}
function nmWriteInputs(s){
  if(q('nmTheme')) q('nmTheme').value = s.theme||'';
  if(q('nmSeed'))  q('nmSeed').value  = String(typeof s.seed==='number'?s.seed:42);
  if(q('nmPrompt'))q('nmPrompt').value= s.prompt||'';
  if(q('nmAspect'))q('nmAspect').value= s.aspect||'16:9';
  if(q('nmCaption'))q('nmCaption').checked=!!s.caption;
  if(q('nmAutoInc'))q('nmAutoInc').checked=(s.autoInc!==false);
}
function nmHash(str){ let h=2166136261>>>0; for(let i=0;i<str.length;i++){ h^=str.charCodeAt(i); h=Math.imul(h,16777619); } return h>>>0; }
function nmSize(aspect){ if(aspect==='4:3')return{w:1024,h:768}; if(aspect==='1:1')return{w:1024,h:1024}; return{w:1920,h:1080}; }

// ---- Abort helper: cancel any in-flight generation safely ----
function nmAbortInFlight(reason='user-cancel'){
  try{
    if(__NIS_GEN_ABORT){ __NIS_GEN_ABORT.abort(reason); __NIS_GEN_ABORT=null; }
  }catch(e){}
  nmShowBusy(false);
  const h=q('nmHint');
  if(h){
    if(reason==='slide-change') h.textContent='Canceled (slide changed).';
    else if(reason==='preempt') h.textContent='Canceled (new request started).';
    else h.textContent='Canceled.';
  }
}

// ---- Throttled progress to reduce UI jank ----
let __nmProgTs = 0;
function nmSetProgressThrottled(pct,msg){
  const now = (typeof performance!=='undefined' && performance.now) ? performance.now() : Date.now();
  if(now - __nmProgTs < 50) return;  // ~20fps
  __nmProgTs = now;
  nmSetProgress(pct,msg);
}

// ---- Cache helpers ----
function nmFindCachedForStyle(style){
  try{
    const key = nmCacheKey(style);
    const b64 = NISPersist.cacheGet(key);
    return b64 || null;
  }catch(e){ return null; }
}
function nmShowCachedButton(style){
  const btn = q('nmApplyCached'), hint=q('nmHint');
  if(!btn) return;
  const has = !!nmFindCachedForStyle(style);
  btn.style.display = has ? 'inline-block' : 'none';
  if(hint && has) hint.textContent = 'Cached image available for this style.';
}

/* Placeholder generator (fallback) */
function nmGeneratePNG(style){
  const sz=nmSize(style.aspect||'16:9'); const w=sz.w,h=sz.h;
  const c=document.createElement('canvas'); c.width=w; c.height=h; const ctx=c.getContext('2d');
  const seed=nmHash((style.theme||'')+'|'+(style.prompt||'')+'|'+String(style.seed||0));
  const a1=(seed%360), a2=((seed>>3)%360);
  const g=ctx.createLinearGradient(0,0,w,h); g.addColorStop(0,'hsl('+a1+' 70% 60%)'); g.addColorStop(1,'hsl('+a2+' 70% 40%)');
  ctx.fillStyle=g; ctx.fillRect(0,0,w,h);
  const motif=(seed%3); ctx.save(); ctx.globalAlpha=0.18;
  if(motif===0){ for(let i=0;i<10;i++){ const r=((seed>>i)&255)/255; const x=r*w,y=(1-r)*h; ctx.beginPath(); ctx.arc(x,y,80*(0.3+r),0,Math.PI*2); ctx.fillStyle='#fff'; ctx.fill(); } }
  else if(motif===1){ for(let i=0;i<14;i++){ const r=((seed>>i)&255)/255; ctx.fillStyle='#fff'; ctx.fillRect(r*w,0,6,h); } }
  else{ ctx.translate(w/2,h/2); for(let i=0;i<8;i++){ ctx.rotate(((seed>>i)&7)*0.15); ctx.fillStyle='#fff'; ctx.fillRect(0,0,w*0.35,3); } }
  ctx.restore();
  if(style.caption){ ctx.fillStyle='rgba(0,0,0,0.6)'; ctx.fillRect(0,h-64,w,64);
    ctx.fillStyle='#fff'; ctx.font='bold 22px system-ui,Segoe UI,Arial';
    const text=(style.theme||'')+' | '+(style.prompt||'')+' | #'+String(style.seed||0);
    ctx.fillText(text,24,h-24);
  }
  return c.toDataURL('image/png').split(',')[1];
}

/* UI helpers (busy/progress/cancel) */
function nmShowBusy(on){
  const busy=q('nmBusy'), prog=q('nmProg'), btnCancel=q('nmCancel');
  ['nmStyleSelected','nmRegenSelected','nmSave'].forEach(id=>{ const b=q(id); if(b) b.disabled=on; });
  if(busy) busy.style.display=on?'block':'none';
  if(prog){ prog.style.display=on?'inline-block':'none'; if(!on){ prog.value=0; } }
  if(btnCancel) btnCancel.style.display=on?'inline-block':'none';
}
function nmSetProgress(pct,msg){
  const prog=q('nmProg'), hint=q('nmHint'), busy=q('nmBusy');
  if(prog && typeof pct==='number'){ prog.value=Math.max(0,Math.min(100,Math.floor(pct))); }
  if(busy) busy.textContent = msg ? ('Generating… '+msg) : 'Generating…';
  if(hint && msg) hint.textContent = msg;
}
function nmApplyToSelection(base64){
  try{ Office.context.document.setSelectedDataAsync(base64,{coercionType:Office.CoercionType.Image},()=>{}); }catch(e){}
}

/* Core generate with provider/timeout/cancel/cache (+ catch for canceled) */
async function nmGenerate(style){
  nmAbortInFlight('preempt');
  const key=nmCacheKey(style);
  nmShowBusy(true); nmSetProgress(0,'Starting');
  try{
    const cached = NISPersist.cacheGet(key);
    if(cached){ nmSetProgress(100,'Cached'); return cached; }

    // Provider available?
    const provider = typeof window.NIS_generateImage==='function' ? window.NIS_generateImage : null;

    // Abort support
    const ctrl = new AbortController();
    __NIS_GEN_ABORT = ctrl;

    // Progress callback (provider may ignore it)
    const onProgress = (p,msg)=>{ try{ nmSetProgressThrottled(p,msg||''); }catch(e){} };

    // Timeout race (45s)
    const timeoutMs = 45000;
    const timeout = new Promise((_,rej)=>setTimeout(()=>rej(new Error('timeout')), timeoutMs));

    let resBase64 = null;

    if(provider){
      const result = await Promise.race([
        provider(style, onProgress, ctrl.signal),
        timeout
      ]);
      if (ctrl.signal.aborted) throw new Error("canceled");
      if(typeof result === 'string' && result.startsWith('data:image/')){
        resBase64 = result.split(',')[1];
      }else if(result && typeof result.base64 === 'string'){
        resBase64 = result.base64;
      }
    }

    // Fallback if provider missing أو رجّع فورغ
    if(!resBase64){
      onProgress(25,'Fallback generator');
      resBase64 = nmGeneratePNG(style);
    }

    onProgress(90,'Caching');
    NISPersist.cacheSet(key, resBase64);

    onProgress(100,'Done');
    try{ nmShowCachedButton(style); }catch(e){}
    return resBase64;
  } catch(err){
  const h=q('nmHint'); 
  if(h) h.textContent=(err && err.message==="canceled") ? "Canceled." : ("Error: "+(err?.message||"failed"));
  return null;
} finally {
  __NIS_GEN_ABORT = null;
  nmShowBusy(false);
}
}

/* Bind Nano UI */
function bindNanoUI(){
  const save=q('nmSave'), styleSel=q('nmStyleSelected'), regen=q('nmRegenSelected'), hint=q('nmHint');
  const btnCancel=q('nmCancel');
  const applyCached = q('nmApplyCached');

  if(save){ save.addEventListener('click',()=>{ const s=nmReadInputs(); nmSaveStyle(s); if(hint) hint.textContent='Style saved for this slide'; }); }

  if(styleSel){ styleSel.addEventListener('click',async()=>{
    const s=nmReadInputs(); nmSaveStyle(s); nmShowCachedButton(s);
    try{ const b64=await nmGenerate(s); if(b64) nmApplyToSelection(b64); if(hint && b64) hint.textContent='Applied to selection'; }
    catch(err){ if(hint) hint.textContent=(err&&err.message==='timeout')?'Generation timeout.':'Generation failed.'; }
  }); }

  if(regen){ regen.addEventListener('click',async()=>{
    let s=nmReadInputs();
    if(s.autoInc){ s.seed=(s.seed||0)+1; nmWriteInputs(s); }
    nmSaveStyle(s); nmShowCachedButton(s);
    try{ const b64=await nmGenerate(s); if(b64) nmApplyToSelection(b64); if(hint && b64) hint.textContent=s.autoInc?'Next seed applied':'Re-generated with same seed'; }
    catch(err){ if(hint) hint.textContent=(err&&err.message==='timeout')?'Generation timeout.':'Generation failed.'; }
  }); }

  if(applyCached){
  applyCached.addEventListener('click', ()=>{
    const s = nmReadInputs();
    const b64 = nmFindCachedForStyle(s);
    const hint = q('nmHint');
    if(b64){
      nmApplyToSelection(b64);
      if(hint) hint.textContent = 'Cached image applied.';
    }else{
      if(hint) hint.textContent = 'No cached image for current style.';
      applyCached.style.display = 'none';
    }
  });
}

  if(btnCancel){
  btnCancel.addEventListener('click',()=>{
    nmAbortInFlight('user-cancel');
  });
}
}
function nmInit(){
  const s = nmGetStyle();
  nmWriteInputs(s);
  // لو فيه صورة متخزّنة للستايل الحالي، نعرض زرار Apply cached
  nmShowCachedButton(s);
}

/* ---------- Linked Sequence (MVP + Auto-advance + Inherit) ---------- */
/* Storage */
function linkLoad(k){
  const persisted = NISPersist.loadLink(k);
  if(persisted && typeof persisted==='object'){
    return {
      enabled: !!persisted.enabled,
      next: persisted.next || null,
      auto: !!persisted.auto,
      autoMs: Number(persisted.autoMs||3000),
      inherit: !!persisted.inherit,
      inheritNm: !!persisted.inheritNm
    };
  }
  return {enabled:false, next:null, auto:false, autoMs:3000, inherit:false, inheritNm:false};
}
function linkSave(k, data){
  const clean={
    enabled:!!data.enabled, next:data.next||null,
    auto:!!data.auto, autoMs:Number(data.autoMs||3000),
    inherit:!!data.inherit, inheritNm:!!data.inheritNm
  };
  __NIS_LINK_CACHE.set(k, clean);
  NISPersist.saveLink(k, clean);
}

/* UI helpers */
async function linkPopulateDropdown(){
  const sel=q('linkNext'); if(!sel) return;
  sel.innerHTML='<option value="">— None —</option>';
  if(!hostIsPowerPoint() || !(window.PowerPoint&&PowerPoint.run)){
    const hint=q('linkHint'); if(hint) hint.textContent='(PowerPoint API not available)';
    return;
  }
  try{
    await PowerPoint.run(async (ctx)=>{
      const coll = ctx.presentation.slides;
      coll.load("items"); await ctx.sync();
      coll.items.forEach(s=>s.load("id,index")); 
      await ctx.sync();
      coll.items.forEach(sl=>{
        const opt=document.createElement('option');
        opt.value=String(sl.id);
        opt.textContent='Slide '+(Number(sl.index)+1);
        sel.appendChild(opt);
      });
    });
  }catch(e){
    const hint=q('linkHint'); if(hint) hint.textContent='(Cannot enumerate slides)';
  }
}
function linkWriteToUI(k){
  const state = __NIS_LINK_CACHE.get(k) || linkLoad(k);
  __NIS_LINK_CACHE.set(k,state);
  const en=q('linkEnable'), nextSel=q('linkNext'), au=q('linkAuto'), auMs=q('linkAutoMs');
  const inh=q('linkInherit'), inhNm=q('linkInheritNm');
  if(en) en.checked=!!state.enabled;
  if(nextSel){ nextSel.value = state.next || ""; }
  if(au) au.checked=!!state.auto;
  if(auMs) auMs.value=String(Number(state.autoMs||3000));
  if(inh) inh.checked=!!state.inherit;
  if(inhNm) inhNm.checked=!!state.inheritNm;
}
function linkReadFromUI(){
  const en=q('linkEnable'), nextSel=q('linkNext'), au=q('linkAuto'), auMs=q('linkAutoMs');
  const inh=q('linkInherit'), inhNm=q('linkInheritNm');
  return { 
    enabled: !!(en && en.checked), 
    next: (nextSel && nextSel.value) ? nextSel.value : null,
    auto: !!(au && au.checked),
    autoMs: Number(auMs && auMs.value ? auMs.value : 3000),
    inherit: !!(inh && inh.checked),
    inheritNm: !!(inhNm && inhNm.checked)
  };
}

/* Auto-advance timer (per active slide) */
function linkAutoArmForActive(){
  linkAutoClear();
  const k=__NIS_ACTIVE_SLIDE_KEY; if(!k) return;
  const st = __NIS_LINK_CACHE.get(k) || linkLoad(k);
  if(!st.enabled || !st.next || !st.auto) return;
  const ms = Math.max(200, Number(st.autoMs||3000));
  __NIS_LINK_AUTO_TIMER = setTimeout(async ()=>{
    __NIS_LINK_AUTO_TIMER=null;
    await linkAdvanceFrom(k);
  }, ms);
}
async function linkAdvanceFrom(k){
  const st = __NIS_LINK_CACHE.get(k) || linkLoad(k);
  if(!st.enabled || !st.next) return false;
  if(!(window.PowerPoint&&PowerPoint.run)) return false;

  // mark to inherit on first visit if flags enabled
  __NIS_INHERIT_NEXT = st.inherit || st.inheritNm ? {from:k} : null;

  try{
    await PowerPoint.run(async (ctx)=>{
      ctx.presentation.setSelectedSlides([st.next]); // API set 1.5
      await ctx.sync();
    });
    return true;
  }catch(e){
    const hint=q('linkHint'); if(hint) hint.textContent='(Advance failed)';
    __NIS_INHERIT_NEXT = null;
    return false;
  }
}

/* Bind UI */
function bindLinkUI(){
  const en=q('linkEnable'), nextSel=q('linkNext');
  const play=q('linkPlay'), adv=q('linkAdvance'), hint=q('linkHint');
  const au=q('linkAuto'), auMs=q('linkAutoMs'), inh=q('linkInherit'), inhNm=q('linkInheritNm');

  const onSave=()=>{ if(!__NIS_ACTIVE_SLIDE_KEY) return; const cur=linkReadFromUI(); linkSave(__NIS_ACTIVE_SLIDE_KEY,cur); if(hint) hint.textContent='Saved.'; };

  en?.addEventListener('change', onSave);
  nextSel?.addEventListener('change', onSave);
  au?.addEventListener('change', onSave);
  auMs?.addEventListener('change', onSave);
  inh?.addEventListener('change', onSave);
  inhNm?.addEventListener('change', onSave);

  play?.addEventListener('click',()=>{
    const cur=getUIParams();
    if(cur.autoStart){ const e=getActiveEngine(); e?.start?.(); linkAutoArmForActive(); }
    if(hint) hint.textContent='Sequence ready — auto/advance as set.';
  });
  adv?.addEventListener('click', async ()=>{
    linkAutoClear();
    const ok = await linkAdvanceFrom(__NIS_ACTIVE_SLIDE_KEY);
    if(!ok){ if(hint) hint.textContent='No next slide set for this slide.'; }
    else{ if(hint) hint.textContent='Advanced.'; }
  });
}

/* Restore for current slide */
function linkRestoreForSlide(){
  if(!__NIS_ACTIVE_SLIDE_KEY) return;
  linkPopulateDropdown().then(()=>{ linkWriteToUI(__NIS_ACTIVE_SLIDE_KEY); });
}

/* ---------- Boot ---------- */
function initBoot(){
  bindSimUI();
  bindNanoUI();
  bindLinkUI();

  captureSlideKeyFast().then(k=>{
    __NIS_ACTIVE_SLIDE_KEY=k;
    getActiveEngine();
    restoreCurrentSlide();
    nmInit();
    linkRestoreForSlide();

    const cur=getUIParams();
    if(cur.autoStart){ const e=getActiveEngine(); e?.start?.(); linkAutoArmForActive(); }
  });

  wireSlideChange();
  setHostHint();

  const e=getActiveEngine(); e?.reset?.();

  // Shortcuts: Ctrl+Alt+S toggle, Ctrl+Alt+Right advance
  document.addEventListener('keydown', (ev)=>{
    if(!(ev.ctrlKey && ev.altKey)) return;
    if(ev.code==='KeyS'){
      ev.preventDefault();
      linkAutoClear();
      const e=getActiveEngine(); e?.stop?.(); setTimeout(()=>{ e?.start?.(); linkAutoArmForActive(); },0);
    }else if(ev.code==='ArrowRight'){
      ev.preventDefault();
      linkAutoClear();
      linkAdvanceFrom(__NIS_ACTIVE_SLIDE_KEY);
    }
  });
}

if(window.Office && Office.onReady){
  Office.onReady(()=>{ 
    if(document.readyState==='loading'){ document.addEventListener('DOMContentLoaded', initBoot); }
    else { initBoot(); }
  });
}else{
  if(document.readyState==='loading'){ document.addEventListener('DOMContentLoaded', initBoot); }
  else { initBoot(); }
}
// ====== end of taskpane.js ======
} // <--- close guard

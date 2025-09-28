/* ======== Nano Interactive Slides - taskpane.js (per-slide, fast, linked sequence + auto-advance) ======== */

/* ---------- Keys ---------- */
const NIS_KEY_PREFIX = 'NIS:scene:'; 
function nisKey(k){ return NIS_KEY_PREFIX + k; }

const NIS_STYLE_KEY_PREFIX='NIS:style:'; 
function nmStyleKey(slideKey){ return NIS_STYLE_KEY_PREFIX + (slideKey||'default'); }

const NIS_LINK_KEY_PREFIX='NIS:link:'; // per-slide link settings {enabled,next,auto,autoMs}

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
const __NIS_LINK_CACHE  = new Map();   // slideKey -> {enabled,next,auto,autoMs}

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
  loadLink(k){   try{ const r=Office.context.document.settings.get(NIS_LINK_KEY_PREFIX+k); return r?JSON.parse(r):null; }catch(e){ return null; } }
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

/* ---------- Active slide ---------- */
let __NIS_ACTIVE_SLIDE_KEY=null;

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
    if(typeof p.speed==='number' && e.setSpeed) e.setSpeed(p.speed);
    if(typeof p.capacity==='number' && e.setCapacity) e.setCapacity(p.capacity);
    if(typeof p.delay==='number' && e.setDelay) e.setDelay(p.delay);
    if(e.setProjectToSlide) e.setProjectToSlide(!!p.projectToSlide);
    if(e.setProjectMs && typeof p.projectMs==='number') e.setProjectMs(p.projectMs);
  }
}

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
function wireSlideChange(){
  try{
    if(!(Office && Office.context && Office.context.document && Office.EventType)) return;

    let lastSlideKey = null;

    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      async ()=>{
        // persist old slide + clear auto-advance timer
        persistCurrentSlide();
        linkAutoClear();

        // stop engine if requested
        const prev = getUIParams();
        if(prev.stopOnChange){
          const e = getActiveEngine();
          if(e && e.stop) e.stop();
        }

        const k = await captureSlideKeyFast();
        if(!k) return;
        if(k === lastSlideKey) return;

        lastSlideKey = k;
        __NIS_ACTIVE_SLIDE_KEY = k;

        getActiveEngine();
        restoreCurrentSlide();
        nmInit();
        linkRestoreForSlide();   // refresh Linked Sequence UI

        const cur = getUIParams();
        if(cur.autoStart){
          const e = getActiveEngine();
          if(e && e.start) e.start();
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

/* ---------- Simulation UI bindings (sliders: input=live, change=save) ---------- */
function applyPreset(name){
  if(name==='slow'){ setUIParams({speed:20,capacity:120,delay:2}); }
  else if(name==='normal'){ setUIParams({speed:50,capacity:200,delay:1}); }
  else if(name==='fast'){ setUIParams({speed:85,capacity:320,delay:0.3}); }
  persistCurrentSlide();
  const e=getActiveEngine(), p=getUIParams();
  if(e&&e.setSpeed)e.setSpeed(p.speed);
  if(e&&e.setCapacity)e.setCapacity(p.capacity);
  if(e&&e.setDelay)e.setDelay(p.delay);
}

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

/* ---------- Host hint ---------- */
function setHostHint(){
  const h=q('hostHint'); if(!h) return;
  try{
    if(hostIsOffice()) h.textContent='Host: '+Office.context.host+(hostIsPowerPoint()?' (projection enabled)':'');
    else h.textContent='Host: Browser preview';
  }catch(e){ h.textContent='Host: Browser'; }
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

/* ---------- Nano Mode (unchanged) ---------- */
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
function nmApplyToSelection(base64){
  try{ Office.context.document.setSelectedDataAsync(base64,{coercionType:Office.CoercionType.Image},()=>{}); }catch(e){}
}
const NIS_IMG_CACHE_PREFIX='NIS:img:'; 
function nmCacheKey(s){ return NIS_IMG_CACHE_PREFIX+[s.theme||'',s.prompt||'',String(s.seed||0),s.aspect||'16:9'].join('|'); }
function nmCacheGet(key){ try{ const raw=Office.context.document.settings.get(key); return raw||null; }catch(e){ return null; } }
function nmCacheSet(key,b64){ try{ Office.context.document.settings.set(key,b64); nisScheduleSave(); }catch(e){} }
function showBusy(on){ const el=q('nmBusy'); if(el) el.style.display=on?'block':'none'; ['nmStyleSelected','nmRegenSelected'].forEach(id=>{const b=q(id); if(b) b.disabled=on;}); }
async function nmGenerate(style){
  showBusy(true); const key=nmCacheKey(style);
  try{
    const cached=nmCacheGet(key); if(cached) return cached;
    if(typeof window.NIS_generateImage==='function'){
      const res=await Promise.race([ window.NIS_generateImage(style), new Promise((_,rej)=>setTimeout(()=>rej(new Error('timeout')),45000)) ]);
      if(typeof res==='string' && res.startsWith('data:image/')){ const b64=res.split(',')[1]; nmCacheSet(key,b64); return b64; }
      if(res && typeof res.base64==='string'){ nmCacheSet(key,res.base64); return res.base64; }
    }
    const b64=nmGeneratePNG(style); nmCacheSet(key,b64); return b64;
  }finally{ showBusy(false); }
}
function bindNanoUI(){
  const save=q('nmSave'), styleSel=q('nmStyleSelected'), regen=q('nmRegenSelected'), hint=q('nmHint');
  if(save){ save.addEventListener('click',()=>{ const s=nmReadInputs(); nmSaveStyle(s); if(hint) hint.textContent='Style saved for this slide'; }); }
  if(styleSel){ styleSel.addEventListener('click',async()=>{ const s=nmReadInputs(); nmSaveStyle(s); const b64=await nmGenerate(s); nmApplyToSelection(b64); if(hint) hint.textContent='Applied to selection'; }); }
  if(regen){ regen.addEventListener('click',async()=>{ let s=nmReadInputs(); if(s.autoInc){ s.seed=(s.seed||0)+1; nmWriteInputs(s); } nmSaveStyle(s); const b64=await nmGenerate(s); nmApplyToSelection(b64); if(hint) hint.textContent=s.autoInc?'Next seed applied':'Re-generated with same seed'; }); }
}
function nmInit(){ const s=nmGetStyle(); nmWriteInputs(s); }

/* ---------- Linked Sequence (MVP + Auto-advance) ---------- */
/* Storage */
function linkLoad(k){
  const persisted = NISPersist.loadLink(k);
  if(persisted && typeof persisted==='object'){
    return {
      enabled: !!persisted.enabled,
      next: persisted.next || null,
      auto: !!persisted.auto,
      autoMs: Number(persisted.autoMs||3000)
    };
  }
  return {enabled:false, next:null, auto:false, autoMs:3000};
}
function linkSave(k, data){
  const clean={enabled:!!data.enabled, next:data.next||null, auto:!!data.auto, autoMs:Number(data.autoMs||3000)};
  __NIS_LINK_CACHE.set(k, clean);
  NISPersist.saveLink(k, clean);
}

/* UI helpers */
async function linkPopulateDropdown(){
  const sel=q('linkNext'); if(!sel) return;
  // reset
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
  if(en) en.checked=!!state.enabled;
  if(nextSel){ nextSel.value = state.next || ""; }
  if(au) au.checked=!!state.auto;
  if(auMs) auMs.value=String(Number(state.autoMs||3000));
}
function linkReadFromUI(){
  const en=q('linkEnable'), nextSel=q('linkNext'), au=q('linkAuto'), auMs=q('linkAutoMs');
  return { 
    enabled: !!(en && en.checked), 
    next: (nextSel && nextSel.value) ? nextSel.value : null,
    auto: !!(au && au.checked),
    autoMs: Number(auMs && auMs.value ? auMs.value : 3000)
  };
}

/* Auto-advance timer (per active slide) */
let __NIS_LINK_AUTO_TIMER = null;
function linkAutoClear(){
  if(__NIS_LINK_AUTO_TIMER){ clearTimeout(__NIS_LINK_AUTO_TIMER); __NIS_LINK_AUTO_TIMER=null; }
}
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

  try{
    await PowerPoint.run(async (ctx)=>{
      ctx.presentation.setSelectedSlides([st.next]); // API set 1.5
      await ctx.sync();
    });
    return true;
  }catch(e){
    const hint=q('linkHint'); if(hint) hint.textContent='(Advance failed)';
    return false;
  }
}

/* Bind UI */
function bindLinkUI(){
  const en=q('linkEnable'), nextSel=q('linkNext');
  const play=q('linkPlay'), adv=q('linkAdvance'), hint=q('linkHint');
  const au=q('linkAuto'), auMs=q('linkAutoMs');

  if(en){ en.addEventListener('change',()=>{ if(!__NIS_ACTIVE_SLIDE_KEY) return; const cur=linkReadFromUI(); linkSave(__NIS_ACTIVE_SLIDE_KEY,cur); if(hint) hint.textContent='Saved.'; }); }
  if(nextSel){ nextSel.addEventListener('change',()=>{ if(!__NIS_ACTIVE_SLIDE_KEY) return; const cur=linkReadFromUI(); linkSave(__NIS_ACTIVE_SLIDE_KEY,cur); if(hint) hint.textContent='Saved.'; }); }
  if(au){ au.addEventListener('change',()=>{ if(!__NIS_ACTIVE_SLIDE_KEY) return; const cur=linkReadFromUI(); linkSave(__NIS_ACTIVE_SLIDE_KEY,cur); if(hint) hint.textContent='Saved.'; }); }
  if(auMs){ auMs.addEventListener('change',()=>{ if(!__NIS_ACTIVE_SLIDE_KEY) return; const cur=linkReadFromUI(); linkSave(__NIS_ACTIVE_SLIDE_KEY,cur); if(hint) hint.textContent='Saved.'; }); }

  if(play){
    play.addEventListener('click',()=>{
      const cur=getUIParams();
      if(cur.autoStart){ const e=getActiveEngine(); e?.start?.(); linkAutoArmForActive(); }
      if(hint) hint.textContent='Sequence ready — use auto-advance or "Advance now".';
    });
  }
  if(adv){
    adv.addEventListener('click', async ()=>{
      linkAutoClear();
      const ok = await linkAdvanceFrom(__NIS_ACTIVE_SLIDE_KEY);
      if(!ok){ if(hint) hint.textContent='No next slide set for this slide.'; }
      else{ if(hint) hint.textContent='Advanced.'; }
    });
  }
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

  /* Simple keyboard shortcuts (optional):
     Ctrl+Alt+S => Start/Stop toggle,  Ctrl+Alt+Right => Advance now */
  document.addEventListener('keydown', (ev)=>{
    if(!(ev.ctrlKey && ev.altKey)) return;
    if(ev.code==='KeyS'){
      ev.preventDefault();
      const e=getActiveEngine();
      // naive toggle: try stop then start
      linkAutoClear();
      e?.stop?.(); setTimeout(()=>{ e?.start?.(); linkAutoArmForActive(); },0);
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

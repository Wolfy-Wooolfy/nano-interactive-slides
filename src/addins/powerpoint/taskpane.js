// [NIS] taskpane.js (powerpoint root) - Step 3: persistence + reset + bindings
(() => {
  const inOffice = typeof Office !== "undefined" && Office?.context?.host;
  let scene = null, running = false, nano = false, lastTick = 0, acc = 0;
  let currentSlideId = "unknown", saveTimer = null, lastSerialized = "";

  console.log("[NIS] LOADED src/addins/powerpoint/taskpane.js ✅");

  // ---- default scene ----
  const defaultScene = {
    version: "0.1",
    nodes: [
      { id: "factory",   type: "producer", rate: 5,  stock: 0 },
      { id: "warehouse", type: "buffer",   capacity: 50, stock: 10 },
      { id: "store",     type: "consumer", rate: 4,  stock: 0 }
    ],
    edges: [
      { from: "factory",   to: "warehouse", delay: 2 },
      { from: "warehouse", to: "store",     delay: 1 }
    ],
    params: { tickMs: 100, initialStock: { factory: 0, warehouse: 10, store: 0 } }
  };

  const keyFor = (slideId) => `scene:${slideId || "default"}`;
  const log = (...a) => console.log("[NIS]", ...a);

  async function getSlideId() {
    if (!inOffice) return "default";
    try {
      return await PowerPoint.run(async (ctx) => {
        const sel = ctx.presentation.getSelectedSlides();
        sel.load("items/id");
        await ctx.sync();
        return sel.items[0]?.id || "default";
      });
    } catch { return "default"; }
  }

  function saveScene(slideId, scn) {
    return new Promise((resolve, reject) => {
      try {
        Office.context.document.settings.set(keyFor(slideId), JSON.stringify(scn));
        Office.context.document.settings.saveAsync((res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(res.error?.message || "saveAsync failed"));
        });
      } catch (e) { reject(e); }
    });
  }
  function loadSceneFromSettings(slideId) {
    try {
      const v = Office.context.document.settings.get(keyFor(slideId));
      if (typeof v === "string" && v.length) return JSON.parse(v);
      return null;
    } catch { return null; }
  }

  // ---------- HUD ----------
  function renderStatus(text) {
    const hud = document.getElementById("time") || document.getElementById("status");
    if (!hud) return;
    const f = scene?.nodes?.find(n=>n.id==="factory");
    const w = scene?.nodes?.find(n=>n.id==="warehouse");
    const st = scene?.nodes?.find(n=>n.id==="store");
    hud.textContent = `t=${Date.now()%100000} | factory=${f?.stock??0} | warehouse=${w?.stock??0} | store=${st?.stock??0} | nano=${nano?"ON":"OFF"} | ${text}`;
  }

  // ---------- رسم ومحاكاة ----------
  function draw() {
    if (!scene) return;
    const canvas = document.getElementById("sim");
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    ctx.clearRect(0,0,canvas.width,canvas.height);

    const nodes = scene.nodes, pos = {}, nx = Math.max(1, nodes.length);
    nodes.forEach((n, i) => pos[n.id] = { x: 120 + i * ((canvas.width - 240) / (nx - 1 || 1)), y: canvas.height/2 });

    ctx.lineWidth = 2; ctx.strokeStyle = nano ? "#0078d4" : "#888";
    scene.edges.forEach(e => { const a=pos[e.from], b=pos[e.to]; ctx.beginPath(); ctx.moveTo(a.x,a.y); ctx.lineTo(b.x,b.y); ctx.stroke(); });

    nodes.forEach(n => {
      const p = pos[n.id];
      let color = n.type==="producer" ? "#34c759" : (n.type==="buffer" ? "#ffcc00" : "#ff375f");
      if (nano) color="#0078d4";
      ctx.fillStyle = color; ctx.beginPath(); ctx.arc(p.x,p.y,26,0,Math.PI*2); ctx.fill();
      ctx.fillStyle = "#000"; ctx.font = "12px Segoe UI"; ctx.textAlign = "center";
      ctx.fillText(`${n.id}`, p.x, p.y-36); ctx.fillText(`${n.stock??0}`, p.x, p.y+4);
    });
  }

  function tick(dtMs) {
    if (!scene) return;
    const speed = nano ? 3 : 1;
    const steps = Math.max(1, Math.floor((dtMs / scene.params.tickMs) * speed));
    for (let s = 0; s < steps; s++) {
      scene.nodes.forEach(n => { if (n.type==="producer" && n.rate) n.stock=(n.stock||0)+1; });
      scene.edges.forEach(e => {
        const from = scene.nodes.find(n=>n.id===e.from), to = scene.nodes.find(n=>n.id===e.to);
        if ((from.stock||0)>0 && Math.random() < 1/Math.max(1,e.delay)) { from.stock=(from.stock||0)-1; to.stock=(to.stock||0)+1; }
      });
      scene.nodes.forEach(n => {
        if (n.type==="consumer" && n.rate && (n.stock||0)>0) n.stock=(n.stock||0)-1;
        if (n.type==="buffer" && Number.isFinite(n.capacity) && (n.stock||0)>n.capacity) n.stock=n.capacity;
      });
    }
    draw();
  }
  function loop(ts){ if(!running) return; if(!lastTick) lastTick=ts; const dt=ts-lastTick; lastTick=ts; acc+=dt; if(acc>=30){ tick(acc); acc=0; } requestAnimationFrame(loop); }
  function start(){ if(!scene) return; running=true; nano=false; lastTick=0; acc=0; renderStatus("Started"); requestAnimationFrame(loop); }
  function stop(){ running=false; nano=false; lastTick=0; acc=0; if(scene) scene.nodes.forEach(n=>n.stock=(scene.params.initialStock[n.id]??0)); draw(); renderStatus("Stopped"); }
  function nanoMode(){ if(!scene) return; running=true; nano=true; lastTick=0; acc=0; renderStatus(inOffice?"Nano Mode (Office)":"Nano Mode (Browser)"); draw(); requestAnimationFrame(loop); }

  // ---------- Persistence ----------
  const serialize = (s)=>{ try{return JSON.stringify(s);}catch{return"";} };
  function scheduleSave(){
    if (!inOffice || currentSlideId==="unknown") return;
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(async () => {
      try { const now = serialize(scene); if (now !== lastSerialized) { await saveScene(currentSlideId, scene); lastSerialized = now; } }
      catch {}
    }, 1000);
  }

  function bindInputs(){
    const f = document.getElementById("input-factory");
    const w = document.getElementById("input-warehouse");
    const s = document.getElementById("input-store");
    const factory   = () => scene.nodes.find(n=>n.id==="factory");
    const warehouse = () => scene.nodes.find(n=>n.id==="warehouse");
    const store     = () => scene.nodes.find(n=>n.id==="store");

    function syncFromScene(){
      if(!scene) return;
      if(f) f.value = String(factory()?.rate ?? 5);
      if(w) w.value = String(warehouse()?.capacity ?? 50);
      if(s) s.value = String(store()?.rate ?? 4);
    }
    function updateFromInputs(){
      if(!scene) return;
      if(f){ const v=Number(f.value||0); const n=factory(); if(n) n.rate=v; }
      if(w){ const v=Number(w.value||0); const n=warehouse(); if(n) n.capacity=v; }
      if(s){ const v=Number(s.value||0); const n=store(); if(n) n.rate=v; }
      draw(); scheduleSave(); renderStatus("Changed");
    }
    [f,w,s].forEach(inp => inp?.addEventListener("input", updateFromInputs));
    syncFromScene();
  }

  async function loadForCurrentSlide(){
    if (inOffice) currentSlideId = await getSlideId();
    let data = inOffice ? loadSceneFromSettings(currentSlideId) : null;
    if (!data) {
      try {
        const res = await fetch(`${location.origin}/examples/supply-chain/scene.json`, { cache: "no-store" });
        if (res.ok) data = await res.json();
      } catch {}
    }
    if (!data) data = JSON.parse(JSON.stringify(defaultScene));
    data.nodes.forEach(n => (n.stock = data.params.initialStock[n.id] ?? 0));
    scene = data; lastSerialized = serialize(scene);
    draw(); bindInputs(); renderStatus("Ready");
  }

  // ---------- ربط الأزرار (مرن مع IDs متعددة) ----------
  function wireButtons(){
    const byIds = (...ids) => ids.map(id => document.getElementById(id)).find(Boolean);
    const startBtn = byIds("start","btn-start");
    const stopBtn  = byIds("stop","btn-stop");
    const nanoBtn  = byIds("nano-btn","btn-nano","nano","nano-mode","nanoMode","NanoMode");
    const resetBtn = byIds("reset","btn-reset");

    if (startBtn) { startBtn.addEventListener("click", start); log("bound start"); }
    if (stopBtn)  { stopBtn.addEventListener("click", stop);  log("bound stop"); }
    if (nanoBtn)  { nanoBtn.addEventListener("click", nanoMode); log("bound nano"); }
    if (resetBtn) { resetBtn.addEventListener("click", async () => {
      scene = JSON.parse(JSON.stringify(defaultScene));
      scene.nodes.forEach(n => (n.stock = scene.params.initialStock[n.id] ?? 0));
      draw(); bindInputs();
      if (inOffice && currentSlideId !== "unknown") { await saveScene(currentSlideId, scene); lastSerialized = serialize(scene); }
      renderStatus("Reset");
    }); log("bound reset"); }
  }

  function watchSelectionChange(){
    if (!inOffice) return;
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      async () => { await loadForCurrentSlide(); }
    );
  }

  function boot(){ wireButtons(); bindInputs(); watchSelectionChange(); loadForCurrentSlide(); }
  if (inOffice) { Office.onReady(() => { log("Office ready"); boot(); }); }
  else { window.addEventListener("DOMContentLoaded", () => { log("DOM ready (browser)"); boot(); }); }
})();

(() => {
  const inOffice = typeof window.Office !== "undefined" && Office?.context?.host;
  let scene = null, running = false, nano = false, lastTick = 0, acc = 0;
  let currentSlideId = "unknown", saveTimer = null, lastSerialized = "";

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

  const keyFor = (slideId) => `scene:${slideId}`;

  async function getSlideId() {
    try {
      return await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.getSelectedSlides();
        slides.load("items/id,items/index");
        await ctx.sync();
        if (slides.items.length > 0 && slides.items[0].id) return slides.items[0].id;
        return "unknown";
      });
    } catch {
      return "unknown";
    }
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

  function ensureUI() {
    // زر Reset لو مش موجود
    if (!document.getElementById("reset")) {
      const row = document.querySelector(".row") || document.body;
      const btn = document.createElement("button");
      btn.id = "reset";
      btn.textContent = "Reset";
      btn.style.marginLeft = "8px";
      row.appendChild(btn);
    }
    // كانفاس العرض
    let view = document.getElementById("view");
    if (view && !document.getElementById("sim")) {
      const c = document.createElement("canvas");
      c.id = "sim"; c.width = 720; c.height = 360;
      c.style.border = "1px solid #ddd"; c.style.marginTop = "8px";
      view.appendChild(c);
    } else if (!view) {
      view = document.createElement("div");
      view.id = "view";
      const c = document.createElement("canvas");
      c.id = "sim"; c.width = 720; c.height = 360;
      c.style.border = "1px solid #ddd"; c.style.marginTop = "8px";
      view.appendChild(c);
      document.body.appendChild(view);
    }
    return { canvas: document.getElementById("sim"), ctx: document.getElementById("sim").getContext("2d") };
  }

  function renderHud(text) {
    const s = document.getElementById("time");
    if (!s) return;
    const f = scene?.nodes?.find(n=>n.id==="factory");
    const w = scene?.nodes?.find(n=>n.id==="warehouse");
    const st = scene?.nodes?.find(n=>n.id==="store");
    s.textContent = `t=${Date.now()%100000} | factory=${f?.stock??0} | warehouse=${w?.stock??0} | store=${st?.stock??0} | nano=${nano?"ON":"OFF"} | ${text}`;
  }

  function draw() {
    if (!scene) return;
    const { canvas, ctx } = ensureUI();
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

  function loop(ts) {
    if (!running) return;
    if (!lastTick) lastTick = ts;
    const dt = ts - lastTick; lastTick = ts; acc += dt;
    if (acc >= 30) { tick(acc); acc = 0; }
    requestAnimationFrame(loop);
  }

  function start(){ if(!scene) return; running=true; nano=false; lastTick=0; acc=0; renderHud("Started"); requestAnimationFrame(loop); }
  function stop(){ running=false; nano=false; lastTick=0; acc=0; if(scene) scene.nodes.forEach(n=>n.stock=(scene.params.initialStock[n.id]??0)); draw(); renderHud("Stopped"); }
  function nanoMode(){
    if(!scene) return; running=true; nano=true; lastTick=0; acc=0; renderHud(inOffice?"Nano Mode (Office)":"Nano Mode (Browser)"); draw();
    if(inOffice){ try{ Office.context.document.setSelectedDataAsync("[NANO MODE]",()=>requestAnimationFrame(loop)); } catch{ requestAnimationFrame(loop); } }
    else requestAnimationFrame(loop);
  }

  function serializeScene(s){ try{ return JSON.stringify(s); } catch{ return ""; } }
  function scheduleSave(){
    if (!inOffice || currentSlideId==="unknown") return;
    if (saveTimer) clearTimeout(saveTimer);
    saveTimer = setTimeout(async () => {
      try {
        const now = serializeScene(scene);
        if (now !== lastSerialized) { await saveScene(currentSlideId, scene); lastSerialized = now; }
      } catch {}
    }, 1000);
  }

  function getInput(idCandidates){
    for(const id of idCandidates){ const el = document.getElementById(id); if(el) return el; }
    return null;
  }

  function bindInputs(){
    const f = getInput(["input-factory","nis-factory"]);
    const w = getInput(["input-warehouse","nis-warehouse"]);
    const s = getInput(["input-store","nis-store"]);
    const factory = () => scene.nodes.find(n=>n.id==="factory");
    const warehouse = () => scene.nodes.find(n=>n.id==="warehouse");
    const store = () => scene.nodes.find(n=>n.id==="store");

    function syncInputsFromScene(){
      if(!scene) return;
      if(f) f.value = String(factory()?.rate ?? 5);
      if(w) w.value = String(warehouse()?.capacity ?? 50);
      if(s) s.value = String(store()?.rate ?? 4);
    }
    function updateFromInputs(){
      if(!scene) return;
      if(f){ const v = Number(f.value||0); const n=factory(); if(n) n.rate = v; }
      if(w){ const v = Number(w.value||0); const n=warehouse(); if(n) n.capacity = v; }
      if(s){ const v = Number(s.value||0); const n=store(); if(n) n.rate = v; }
      draw(); scheduleSave(); renderHud("Changed");
    }
    [f,w,s].forEach(inp => inp?.addEventListener("input", updateFromInputs));
    syncInputsFromScene();
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
    scene = data; lastSerialized = serializeScene(scene);
    draw(); bindInputs(); renderHud("Ready");
  }

  function wireButtons(){
    const $ = (id)=>document.getElementById(id);
    ($("#start"))?.addEventListener("click", start);
    ($("#stop"))?.addEventListener("click", stop);
    ($("#nano-btn") || $("#nano") || $("#nano-mode") || $("#nanoMode") || $("#NanoMode"))?.addEventListener("click", nanoMode);
    ($("#reset"))?.addEventListener("click", async ()=> {
      scene = JSON.parse(JSON.stringify(defaultScene));
      scene.nodes.forEach(n => (n.stock = scene.params.initialStock[n.id] ?? 0));
      draw(); bindInputs();
      if (inOffice && currentSlideId !== "unknown") { await saveScene(currentSlideId, scene); lastSerialized = serializeScene(scene); }
      renderHud("Reset");
    });
  }

  function watchSelectionChange(){
    if (!inOffice) return;
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      async () => { await loadForCurrentSlide(); }
    );
  }

  function boot(){ ensureUI(); wireButtons(); watchSelectionChange(); loadForCurrentSlide(); }
  if (inOffice) { Office.onReady(() => boot()); } else { window.addEventListener("DOMContentLoaded", boot); }
})();

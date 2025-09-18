(() => {
  const inOffice = typeof window.Office !== "undefined" && Office?.context?.host;
  let scene = null, running = false, nano = false, lastTick = 0, acc = 0;

  // Fallback scene لو ملف JSON مش موجود
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

  function ensureUI() {
    let hud = document.getElementById("hud");
    if (!hud) {
      hud = document.createElement("div");
      hud.id = "hud";
      hud.style.marginTop = "12px";
      hud.style.fontFamily = "Segoe UI, Arial, sans-serif";
      document.body.appendChild(hud);
    }
    let canvas = document.getElementById("sim");
    if (!canvas) {
      canvas = document.createElement("canvas");
      canvas.id = "sim";
      canvas.width = 720;
      canvas.height = 360;
      canvas.style.border = "1px solid #ddd";
      canvas.style.marginTop = "8px";
      document.body.appendChild(canvas);
    }
    return { hud, canvas, ctx: canvas.getContext("2d") };
  }

  async function loadScene() {
    const url = `${location.origin}/examples/supply-chain/scene.json`;
    let data;
    try {
      const res = await fetch(url, { cache: "no-store" });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      data = await res.json();
    } catch {
      data = JSON.parse(JSON.stringify(defaultScene));
    }
    data.nodes.forEach(n => (n.stock = data.params.initialStock[n.id] ?? 0));
    scene = data;
    draw();
    renderHud("Ready");
  }

  function renderHud(text) {
    const hud = document.getElementById("hud");
    if (hud) hud.textContent = `State: ${running ? (nano ? "Running (Nano)" : "Running") : "Stopped"} | ${text}`;
  }

  function draw() {
    if (!scene) return;
    const { canvas, ctx } = ensureUI();
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    const pos = {};
    const nodes = scene.nodes;
    const nx = Math.max(1, nodes.length);
    nodes.forEach((n, i) => {
      pos[n.id] = { x: 120 + i * ((canvas.width - 240) / (nx - 1 || 1)), y: canvas.height / 2 };
    });

    // edges
    ctx.lineWidth = 2;
    ctx.strokeStyle = nano ? "#0078d4" : "#888";
    scene.edges.forEach(e => {
      const a = pos[e.from], b = pos[e.to];
      ctx.beginPath(); ctx.moveTo(a.x, a.y); ctx.lineTo(b.x, b.y); ctx.stroke();
    });

    // nodes
    nodes.forEach(n => {
      const p = pos[n.id];
      let color = n.type === "producer" ? "#34c759" : (n.type === "buffer" ? "#ffcc00" : "#ff375f");
      if (nano) color = "#0078d4"; // لون nano
      ctx.fillStyle = color;
      ctx.beginPath(); ctx.arc(p.x, p.y, 26, 0, Math.PI * 2); ctx.fill();
      ctx.fillStyle = "#000"; ctx.font = "12px Segoe UI"; ctx.textAlign = "center";
      ctx.fillText(`${n.id}`, p.x, p.y - 36);
      ctx.fillText(`${n.stock ?? 0}`, p.x, p.y + 4);
    });
  }

  function tick(dtMs) {
    if (!scene) return;
    const speed = nano ? 3 : 1;
    const steps = Math.max(1, Math.floor((dtMs / scene.params.tickMs) * speed));
    for (let s = 0; s < steps; s++) {
      // production
      scene.nodes.forEach(n => {
        if (n.type === "producer" && n.rate) n.stock = (n.stock || 0) + 1;
      });
      // movement
      scene.edges.forEach(e => {
        const from = scene.nodes.find(n => n.id === e.from);
        const to   = scene.nodes.find(n => n.id === e.to);
        if ((from.stock || 0) > 0 && Math.random() < 1 / Math.max(1, e.delay)) {
          from.stock = (from.stock || 0) - 1;
          to.stock   = (to.stock   || 0) + 1;
        }
      });
      // consumption & caps
      scene.nodes.forEach(n => {
        if (n.type === "consumer" && n.rate && (n.stock || 0) > 0) n.stock = (n.stock || 0) - 1;
        if (n.type === "buffer" && Number.isFinite(n.capacity) && (n.stock || 0) > n.capacity) n.stock = n.capacity;
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

  function start() { if (!scene) return; running = true; nano = false; lastTick = 0; acc = 0; renderHud("Started"); requestAnimationFrame(loop); }
  function stop()  { running = false; nano = false; lastTick = 0; acc = 0; if (scene) scene.nodes.forEach(n => n.stock = (scene.params.initialStock[n.id] ?? 0)); draw(); renderHud("Stopped"); }

  function nanoMode() {
    if (!scene) return;
    console.debug("[NIS] Nano Mode clicked");
    running = true; nano = true; lastTick = 0; acc = 0;
    renderHud(inOffice ? "Nano Mode (Office)" : "Nano Mode (Browser)");
    draw(); // غيّر اللون فورًا

    if (inOffice) {
      try { Office.context.document.setSelectedDataAsync("[NANO MODE]", () => requestAnimationFrame(loop)); }
      catch { requestAnimationFrame(loop); }
    } else {
      requestAnimationFrame(loop);
    }
  }

  // يربط الزرار حتى لو الـ ID مختلف
  function wireButtons() {
    const byId = (id) => document.getElementById(id);
    const nanoCandidates = [
      byId("nano"), byId("nano-mode"), byId("nanoMode"), byId("NanoMode"),
      // fallback: دور على زرار نصه "Nano Mode"
      Array.from(document.querySelectorAll("button")).find(b => (b.textContent||"").trim().toLowerCase() === "nano mode")
    ].filter(Boolean);

    byId("start")?.addEventListener("click", start);
    byId("stop")?.addEventListener("click", stop);
    nanoCandidates[0]?.addEventListener("click", nanoMode);
  }

  function boot() { wireButtons(); ensureUI(); loadScene(); }
  if (inOffice) { Office.onReady(() => boot()); } else { window.addEventListener("DOMContentLoaded", boot); }
})();

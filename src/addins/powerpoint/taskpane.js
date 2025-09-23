let refreshingUI = false;
const engines = new Map();
let lastKey = null;
let polling = false;

class Engine {
  constructor(state) {
    this.state = Object.assign(
      { speed: 1, capacity: 50, delay: 1, progress: 0, running: false },
      state || {}
    );
  }

  start(onTick) {
    if (this.state.running) return;
    this.state.running = true;

    const step = () => {
      if (!this.state.running) return;

      const inc = Number(this.state.speed) || 0;
      this.state.progress = Math.max(
        0,
        Math.min(100, this.state.progress + inc)
      );

      console.log("Tick:", this.state); // للتأكد من التشغيل

      if (typeof onTick === "function") onTick(this.state);

      const d = Math.max(50, (Number(this.state.delay) || 1) * 100);
      setTimeout(step, d);
    };

    step();
  }

  stop() {
    this.state.running = false;
    console.log("Engine stopped");
  }

  setParam(k, v) {
    this.state[k] = v;
    console.log("Param set:", k, v);
  }
}

function setBadge(text) {
  const el = document.getElementById("slideId");
  if (el) el.textContent = text || "SlideKey: -";
}

function keyFromSlide(slideObj) {
  if (!slideObj) return null;
  const id = slideObj.id ? String(slideObj.id) : "";
  const idx = (typeof slideObj.index !== "undefined") ? String(slideObj.index) : "";
  const title = slideObj.title ? String(slideObj.title) : "";
  const key = id || (idx ? ("idx:" + idx) : "") || (title ? ("title:" + title) : "");
  return key || null;
}

function getCurrentSlideKey() {
  return new Promise((resolve) => {
    try {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          const v = res.value;
          if (v && v.slides && v.slides.length > 0) {
            resolve(keyFromSlide(v.slides[0]));
            return;
          }
        }
        resolve(null);
      });
    } catch (e) { resolve(null); }
  });
}

async function loadState(key) {
  try {
    const s = await OfficeRuntime.storage.getItem("nis:slide:" + key);
    return s ? JSON.parse(s) : null;
  } catch (e) { return null; }
}

async function saveState(key, state) {
  try { await OfficeRuntime.storage.setItem("nis:slide:" + key, JSON.stringify(state)); } catch (e) {}
}

async function getEngineFor(key) {
  let eng = engines.get(key);
  if (!eng) {
    const st = await loadState(key);
    eng = new Engine(st);
    engines.set(key, eng);
  }
  return eng;
}

function updateUiFromState(state) {
  const p = document.getElementById("progress");
  if (p) p.textContent = String(Math.round(state.progress));

  const bar = document.getElementById("progressBar");
  if (bar) {
    const w = Math.max(0, Math.min(100, Number(state.progress) || 0));
    bar.style.width = w + "%";
    bar.style.background =
      w < 33 ? "#ef4444" : w < 66 ? "#f59e0b" : "#10b981";
  }

  const s = document.getElementById("speed");
  const sn = document.getElementById("speedNum");
  if (s) s.value = state.speed;
  if (sn) sn.value = state.speed;

  const c = document.getElementById("capacity");
  if (c) c.value = state.capacity;

  const d = document.getElementById("delay");
  if (d) d.value = state.delay;
}

async function onSlideActivated(key) {
  if (!key) return;
  lastKey = key;
  setBadge("SlideKey: " + key);
  const eng = await getEngineFor(key);
  updateUiFromState(eng.state);
}

function startPollingSelection() {
  if (polling) return;
  polling = true;
  const tick = async () => {
    try {
      Office.context.document.getActiveViewAsync(async function (viewRes) {
        const view = viewRes && viewRes.value;
        if (view === "edit") {
          const key = await getCurrentSlideKey();
          if (key && key !== lastKey) {
            await onSlideActivated(key);
          }
        }
      });
    } catch (e) {}
    setTimeout(tick, 500);
  };
  tick();
}

async function ensureActiveKey() {
  if (!hasOffice()) {
    setStatus("Running outside PowerPoint: no slide detected");
    return null;
  }
  if (lastKey) {
    setStatus("");
    return lastKey;
  }
  const key = await getCurrentSlideKey();
  if (key) {
    await onSlideActivated(key);
    setStatus("");
    return key;
  }
  setStatus("No slide selected. Click a slide thumbnail, then press Sync.");
  return null;
}

async function NISStart() {
  const key = await ensureActiveKey();
  if (!key) return;
  const eng = await getEngineFor(key);

  if (eng.state.progress >= 100) {
    eng.state.progress = 0;
    updateUiFromState(eng.state);
    await saveState(key, eng.state);
  }

  eng.start(async (st) => {
    updateUiFromState(st);
    await saveState(key, st);
  });
  await saveState(key, eng.state);
}

async function NISStop() {
  const key = await ensureActiveKey();
  if (!key) return;
  const eng = await getEngineFor(key);
  eng.stop();
  await saveState(key, eng.state);
}

async function NISReset() {
  const key = await ensureActiveKey();
  if (!key) return;
  const eng = await getEngineFor(key);
  eng.stop();
  eng.state.progress = 0;
  updateUiFromState(eng.state);
  await saveState(key, eng.state);
}
window.NISReset = NISReset;

async function NISSetParam(name, value) {
  const key = await ensureActiveKey();
  if (!key) return;
  const eng = await getEngineFor(key);
  eng.setParam(name, typeof value === "string" ? (value.trim() === "" ? value : Number.isNaN(Number(value)) ? value : Number(value)) : value);
  updateUiFromState(eng.state);
  await saveState(key, eng.state);
}

async function NISSync() {
  const key = await getCurrentSlideKey();
  await onSlideActivated(key);
}

function autoWire() {
  const btnStart = document.getElementById("btnStart") || document.getElementById("start") || document.querySelector("[data-nis-start]");
  const btnStop  = document.getElementById("btnStop")  || document.getElementById("stop")  || document.querySelector("[data-nis-stop]");
  if (btnStart) btnStart.addEventListener("click", NISStart);
  if (btnStop)  btnStop.addEventListener("click", NISStop);

  // تزامن speed slider <-> number
  const speed = document.getElementById("speed");
  const speedNum = document.getElementById("speedNum");
  if (speed && speedNum) {
    const syncFromSlider = () => {
      if (refreshingUI) return;
      speedNum.value = speed.value;
      NISSetParam("speed", speed.value);
    };
    const syncFromNumber = () => {
      if (refreshingUI) return;
      speed.value = speedNum.value;
      NISSetParam("speed", speedNum.value);
    };
    speed.addEventListener("input",  syncFromSlider);
    speed.addEventListener("change", syncFromSlider);
    speedNum.addEventListener("input",  syncFromNumber);
    speedNum.addEventListener("change", syncFromNumber);
  }

  // باقي الباراميترات (capacity, delay)
  document.querySelectorAll("[data-nis-param]").forEach(function (el) {
    // اتأكدنا فوق إن speed ليه هاندلرز خاصة، فلو ده speedNum أو speed هنتجاهله هنا
    if (el.id === "speed" || el.id === "speedNum") return;
    el.addEventListener("input", function () { if (!refreshingUI) NISSetParam(el.getAttribute("data-nis-param"), el.value); });
    el.addEventListener("change", function () { if (!refreshingUI) NISSetParam(el.getAttribute("data-nis-param"), el.value); });
  });
}

function setStatus(msg) {
  const el = document.getElementById("status");
  if (el) el.textContent = msg || "";
}
function hasOffice() {
  return typeof Office !== "undefined" && Office.context && Office.context.document;
}

Office.initialize = function () {
  autoWire();
  startPollingSelection();
  setTimeout(NISSync, 300);
};

window.NISStart = NISStart;
window.NISStop = NISStop;
window.NISSetParam = NISSetParam;
window.NISSync = NISSync;

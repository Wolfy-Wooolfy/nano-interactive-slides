import { Engine, Scene } from "./engine";
import defaultScene from "./scene-default.json";

const statusEl = document.getElementById("status") as HTMLElement | null;
const btnStart = document.getElementById("btn-start");
const btnStop = document.getElementById("btn-stop");
const btnNano = document.getElementById("btn-nano");

let engine: Engine | null = null;
let uiTimer: any = null;

function setStatus(t: string) {
  if (statusEl) statusEl.textContent = t;
}

async function loadDefaultScene() {
  const scene = defaultScene as unknown as Scene;
  engine = new Engine(scene);
  setStatus("Scene loaded");
  if (uiTimer) clearInterval(uiTimer);
  uiTimer = setInterval(() => {
    const s = engine?.getSnapshot();
    if (s) setStatus(`t=${s.t} | factory=${s.stock["factory"] ?? 0} | warehouse=${s.stock["warehouse"] ?? 0} | store=${s.stock["store"] ?? 0} | nano=${s.nano ? "ON" : "OFF"}`);
  }, 300);
}

btnStart?.addEventListener("click", () => {
  engine?.start();
  setStatus("Engine: START");
});

btnStop?.addEventListener("click", () => {
  engine?.stop();
  setStatus("Engine: STOP");
});

btnNano?.addEventListener("click", () => {
  if (!engine) return;
  const v = engine.toggleNano();
  setStatus(`Nano Mode: ${v ? "ON" : "OFF"}`);
});

loadDefaultScene();

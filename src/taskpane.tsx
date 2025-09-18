const statusEl = document.getElementById("status") as HTMLElement | null;
const btnStart = document.getElementById("btn-start");
const btnStop = document.getElementById("btn-stop");
const btnNano = document.getElementById("btn-nano");

let nano = false;

function setStatus(t: string) {
  if (statusEl) statusEl.textContent = t;
}

btnStart?.addEventListener("click", () => setStatus("Engine: START"));
btnStop?.addEventListener("click", () => setStatus("Engine: STOP"));
btnNano?.addEventListener("click", () => {
  nano = !nano;
  setStatus(`Nano Mode: ${nano ? "ON" : "OFF"}`);
});

setStatus("Ready");

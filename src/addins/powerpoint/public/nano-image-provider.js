// src/addins/powerpoint/nano-image-provider.js
window.NIS_generateImage = async (style, onProgress, signal) => {
  const inc = (p, m) => { try { onProgress?.(p, m || ""); } catch(_) {} };
  const size = (a) => a === "4:3" ? { w: 1024, h: 768 } : a === "1:1" ? { w: 1024, h: 1024 } : { w: 1920, h: 1080 };
  const s = size(style?.aspect || "16:9");
  const body = {
    prompt: `${style?.theme || ""} ${style?.prompt || ""}`.trim(),
    seed: Number(style?.seed || 0) || 0,
    width: s.w,
    height: s.h
  };

  let timer = null, p = 5;
  inc(5, "Preparing");
  timer = setInterval(() => { p = Math.min(85, p + 2); inc(p, "Generating"); }, 250);

  try {
    const res = await fetch("/nano/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
      signal
    });
    if (!res.ok) throw new Error(`provider-${res.status}`);
    const data = await res.json();
    if (data?.base64) return { base64: data.base64 };
    throw new Error("provider-no-base64");
  } finally {
    if (timer) clearInterval(timer);
  }
};

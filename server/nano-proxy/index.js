import "dotenv/config";
import express from "express";
import cors from "cors";
import sharp from "sharp";

const app = express();
app.use(cors({ origin: [/^https?:\/\/localhost:3000$/] }));
app.use(express.json({ limit: "8mb" }));

const PORT   = Number(process.env.NANO_PROXY_PORT || 8787);
const PROV   = (process.env.NANO_PROVIDER || "").toLowerCase(); // "mock" | "replicate" | "gemini"
const MODEL  = process.env.NANO_MODEL || "black-forest-labs/flux-schnell";
const TOKEN  = process.env.REPLICATE_API_TOKEN || "";
const GKEY   = process.env.GOOGLE_API_KEY || "";
const PXBKEY = process.env.PIXABAY_API_KEY || ""; // <<< جديد

/* ---------------- helpers ---------------- */
function log(...a){ console.log("[NANO]", ...a); }
function clamp(n,min,max){ return Math.max(min, Math.min(max, n)); }
function hashInt(s){ let x=0; for (let i=0;i<s.length;i++) x=((x*31)+s.charCodeAt(i))>>>0; return x%1_000_000; }

async function fetchWithTimeout(url, opts={}, ms=9000){
  const c = new AbortController();
  const id = setTimeout(()=>c.abort(), ms);
  try {
    const r = await fetch(url, {
      ...opts,
      signal: c.signal,
      headers: {
        "Accept": "image/*,application/json;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.8",
        "User-Agent": "NIS-Mock/1.3",
        "Referer": "https://localhost:3000/",
        ...(opts.headers||{})
      },
      redirect: "follow",
      cache: "no-store"
    });
    return r;
  } finally { clearTimeout(id); }
}

async function fetchImageAsBase64(url, timeoutMs=9000){
  const r = await fetchWithTimeout(url, {}, timeoutMs);
  if (!r.ok) throw new Error(`status=${r.status}`);
  const ct = (r.headers.get("content-type") || "").toLowerCase();
  if (!ct.startsWith("image/")) throw new Error(`not-image:${ct}`);
  const buf = Buffer.from(await r.arrayBuffer());
  return buf.toString("base64");
}

/* ---------------- مصادر مجانية ---------------- */
// A) Pixabay (needs free API key)

// A) Pixabay (with resize)
async function getFromPixabay(topics, w, h, seed) {
  if (!PXBKEY) throw new Error("pxb-no-key");
  const q = encodeURIComponent(topics.replace(/,/g," "));
  const api = `https://pixabay.com/api/?key=${PXBKEY}&q=${q}&image_type=photo&per_page=50&safesearch=true`;
  const r = await fetchWithTimeout(api, { headers:{ "Accept":"application/json" } }, 9000);
  if (!r.ok) throw new Error(`pxb-status=${r.status}`);
  const j = await r.json();
  const hits = Array.isArray(j?.hits) ? j.hits : [];
  if (!hits.length) throw new Error("pxb-empty");
  const idx = seed % hits.length;
  const src = hits[idx].largeImageURL || hits[idx].webformatURL;
  if (!src) throw new Error("pxb-no-url");

  // fetch original image
  const rawResp = await fetchWithTimeout(src, {}, 9000);
  if (!rawResp.ok) throw new Error(`pxb-fetch-fail:${rawResp.status}`);
  const buf = Buffer.from(await rawResp.arrayBuffer());

  // resize with sharp
  const resized = await sharp(buf).resize(w, h, { fit: "cover" }).png().toBuffer();
  return { base64: resized.toString("base64"), via: "pixabay-resized" };
}

// B) Wikimedia Commons (strong search)
async function getFromCommons(topics, w, h, seed){
  const q = encodeURIComponent(topics.replace(/,/g," "));
  // محاولة 1: generator=search على namespace=6 (File) + نوع الصورة
  let api = `https://commons.wikimedia.org/w/api.php?action=query&format=json&origin=*&prop=imageinfo&generator=search&gsrsearch=intitle:${q}%20filetype:bitmap|drawing&gsrnamespace=6&gsrlimit=50&iiprop=url|mime|extmetadata&iiurlwidth=${w}`;
  let r = await fetchWithTimeout(api, { headers:{ "Accept":"application/json" } }, 9000);
  if (!r.ok) throw new Error(`commons-status=${r.status}`);
  let j = await r.json();
  let pages = j?.query?.pages ? Object.values(j.query.pages) : [];
  let imgs = pages.filter(p => p?.imageinfo?.[0]?.thumburl || p?.imageinfo?.[0]?.url);
  if (!imgs.length) {
    // محاولة 2: list=search للحصول على عناوين ثم جلب imageinfo
    const api2 = `https://commons.wikimedia.org/w/api.php?action=query&format=json&origin=*&list=search&srsearch=intitle:${q}%20filetype:bitmap|drawing&srnamespace=6&srinfo=totalhits|suggestion&srlimit=50&formatversion=2`;
    const r2 = await fetchWithTimeout(api2, { headers:{ "Accept":"application/json" } }, 9000);
    if (!r2.ok) throw new Error(`commons-status2=${r2.status}`);
    const j2 = await r2.json();
    const titles = (j2?.query?.search || []).map(s => s.title).filter(Boolean);
    if (!titles.length) throw new Error("commons-empty");
    const ti = titles.slice(0, 50).map(encodeURIComponent).join("|");
    const api3 = `https://commons.wikimedia.org/w/api.php?action=query&format=json&origin=*&prop=imageinfo&titles=${ti}&iiprop=url|mime|extmetadata&iiurlwidth=${w}`;
    const r3 = await fetchWithTimeout(api3, { headers:{ "Accept":"application/json" } }, 9000);
    if (!r3.ok) throw new Error(`commons-status3=${r3.status}`);
    const j3 = await r3.json();
    pages = j3?.query?.pages ? Object.values(j3.query.pages) : [];
    imgs = pages.filter(p => p?.imageinfo?.[0]?.thumburl || p?.imageinfo?.[0]?.url);
    if (!imgs.length) throw new Error("commons-empty");
  }
  const idx = seed % imgs.length;
  const info = imgs[idx].imageinfo[0];
  const src = info.thumburl || info.url;
  const base64 = await fetchImageAsBase64(src, 9000);
  return { base64, via: "mock-commons" };
}

// C) Unsplash / LoremFlickr / Picsum
async function getFromUnsplash(topics, w, h, seed){
  const u = `https://source.unsplash.com/${w}x${h}/?${encodeURIComponent(topics)}&sig=${seed}`;
  const base64 = await fetchImageAsBase64(u, 9000);
  return { base64, via: "mock-unsplash" };
}
async function getFromLoremFlickr(topics, w, h, seed){
  const u = `https://loremflickr.com/${w}/${h}/${topics}?lock=${seed}`;
  const base64 = await fetchImageAsBase64(u, 9000);
  return { base64, via: "mock-loremflickr" };
}
async function getFromPicsum(seed, w, h){
  const u = `https://picsum.photos/seed/${seed}/${w}/${h}`;
  const base64 = await fetchImageAsBase64(u, 9000);
  return { base64, via: "mock-picsum" };
}

/* ---------------- مولد MOCK الرئيسي ---------------- */
async function generateMock({ prompt = "", seed = 0, width = 1024, height = 1024 }) {
  const w = clamp(Number(width)  || 1024, 64, 4096);
  const h = clamp(Number(height) || 1024, 64, 4096);
  const topics = (String(prompt)||"nature").toLowerCase()
    .replace(/[^a-z0-9]+/g, ",").replace(/^,|,$/g,"") || "photos";
  const lock = hashInt(String(seed) + "|" + topics);

  // لو فيه Pixabay key → جرّبه أولاً
  if (PXBKEY) {
    try { return await getFromPixabay(topics, w, h, lock); }
    catch(e){ log("pixabay failed:", e.message); }
  }

  // وإلا: Commons → Unsplash → LoremFlickr → Picsum → Placeholder
  try { return await getFromCommons(topics, w, h, lock); }
  catch(e){ log("commons failed:", e.message); }

  try { return await getFromUnsplash(topics, w, h, lock); }
  catch(e){ log("unsplash failed:", e.message); }

  try { return await getFromLoremFlickr(topics, w, h, lock); }
  catch(e){ log("loremflickr failed:", e.message); }

  try { return await getFromPicsum(lock, w, h); }
  catch(e){ log("picsum failed:", e.message); }

  const text = encodeURIComponent((prompt || "NIS mock").slice(0,40));
  const ph = `https://placehold.co/${w}x${h}.jpg?text=${text}`;
  const base64 = await fetchImageAsBase64(ph, 6000);
  return { base64, via: "mock-placeholder" };
}

/* ---------------- Replicate (جاهز لاحقًا) ---------------- */
async function rfetch(url, opts = {}) {
  return fetch(url, {
    ...opts,
    headers: { ...(opts.headers||{}), Authorization:`Token ${TOKEN}`, "Content-Type":"application/json" }
  });
}
const isFlux = MODEL.startsWith("black-forest-labs/flux");
const isSDXL = MODEL.includes("sdxl") || MODEL.includes("stability-ai/sdxl");
function buildInput(body){
  const prompt = String(body?.prompt || "");
  const seed   = typeof body?.seed === "number" ? body.seed : 0;
  const width  = typeof body?.width === "number" ? body.width : 1024;
  const height = typeof body?.height === "number" ? body.height : 1024;
  if (isFlux) return { prompt };
  return { prompt, seed, width, height, num_inference_steps:28, guidance_scale:7 };
}
async function createViaModel(model, input){
  const res = await rfetch(`https://api.replicate.com/v1/models/${model}/predictions`, {
    method:"POST", body: JSON.stringify({ input })
  });
  if (!res.ok) {
    const txt = await res.text();
    const err = new Error(`create-via-model-failed: ${res.status} ${txt}`);
    err.status = res.status;
    throw err;
  }
  return res.json();
}
async function getLatestVersion(model){
  const res = await rfetch(`https://api.replicate.com/v1/models/${model}/versions`);
  if (!res.ok) throw new Error(`versions-failed: ${res.status} ${await res.text()}`);
  const data = await res.json();
  const ver = data?.results?.[0]?.id || data?.versions?.[0]?.id;
  if (!ver) throw new Error("no-version-found");
  return ver;
}
async function createViaVersion(version, input){
  const res = await rfetch("https://api.replicate.com/v1/predictions", {
    method:"POST", body: JSON.stringify({ version, input })
  });
  if (!res.ok) throw new Error(`create-via-version-failed: ${res.status} ${await res.text()}`);
  return res.json();
}
async function pollPred(id, timeoutMs=90_000){
  const t0 = Date.now();
  for(;;){
    const r = await rfetch(`https://api.replicate.com/v1/predictions/${id}`);
    if (!r.ok) throw new Error(`poll-failed: ${r.status} ${await r.text()}`);
    const data = await r.json();
    if (data?.status === "succeeded") return data;
    if (["failed","canceled"].includes(data?.status)) throw new Error(`prediction-${data.status}`);
    if (Date.now()-t0 > timeoutMs) throw new Error("timeout");
    await new Promise(r=>setTimeout(r,1200));
  }
}

/* ---------------- Route ---------------- */
app.post("/nano/generate", async (req, res) => {
  try {
    const body = req.body || {};

    if (PROV === "mock" || !PROV) {
      const out = await generateMock(body);
      res.setHeader("x-nis-provider", out.via || "mock");
      return res.json({ base64: out.base64 });
    }

    if (PROV === "replicate") {
      if (!TOKEN) return res.status(503).json({ error: "no-replicate-token" });
      const input = buildInput(body);
      let pred;
      try { pred = await createViaModel(MODEL, input); }
      catch (e1) { if (isSDXL) { const ver = await getLatestVersion(MODEL); pred = await createViaVersion(ver, input); } else { throw e1; } }
      const done = await pollPred(pred?.id);
      const url = Array.isArray(done?.output) ? done.output[0] : done?.output;
      if (!url || typeof url !== "string") throw new Error("no-output");
      const img = await fetch(url);
      const buf = Buffer.from(await img.arrayBuffer());
      res.setHeader("x-nis-provider", "replicate");
      return res.json({ base64: buf.toString("base64") });
    }

    if (PROV === "gemini") {
      return res.status(503).json({ error: "gemini-disabled-for-free" });
    }

    return res.status(400).json({ error: "unknown-provider" });
  } catch (err) {
    const msg = String(err?.message || err);
    const status = err?.status === 402 ? 402 : 502;
    log("ERROR:", msg);
    res.status(status).json({ error: msg });
  }
});

app.listen(PORT, () => {
  log(`listening on http://localhost:${PORT}`);
  log(`provider=${PROV || "mock"}  model=${MODEL}`);
});

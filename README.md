# Nano Interactive Slides (NIS)

**Nano Interactive Slides (NIS)** is a PowerPoint add-in that adds a per-slide control panel for:
- **Simulation Controls** (speed / capacity / delay + projection to the slide),
- **Nano Mode** (style inputs and image generation hooks with caching),
- **Linked Sequence** (per-slide “Next slide” mapping, auto-advance, and one-time inheritance).

All settings are **per slide** and persist inside the PowerPoint document.

---

## Table of Contents
- [Prerequisites](#prerequisites)
- [Getting Started](#getting-started)
- [Current Features](#current-features)
- [How to Use](#how-to-use)
  - [Simulation Controls](#simulation-controls)
  - [Linked Sequence](#linked-sequence)
  - [Nano Mode](#nano-mode)
- [Persistence & Caching](#persistence--caching)
- [Active Source Files](#active-source-files)
- [Keyboard Shortcuts](#keyboard-shortcuts)
- [Developer Notes](#developer-notes)
- [Troubleshooting](#troubleshooting)
- [Roadmap](#roadmap)
- [License](#license)

---

## Prerequisites
- **Node.js 20+**
- **pnpm** (you can use npm, but scripts here assume pnpm)
- **Microsoft PowerPoint** (Windows desktop with Office.js support)

> If HTTPS dev certs are not trusted on your machine, install Microsoft’s dev certs:
> ```bash
> npx office-addin-dev-certs install
> ```

---

## Getting Started

### 1) Install dependencies
pnpm install

### 2) Run from repository root (recommended)

pnpm run dev-addin     # starts the dev server at https://localhost:3000
pnpm run start-addin   # sideloads the manifest and launches PowerPoint


### Alternative (single command per directory)

pnpm --dir src/addins/powerpoint dev


> **Windows note:** Some shells mis-handle `-C`. Use `--dir` or the root scripts above.

---

## Current Features

* **Per-Slide State**

  * All Simulation Controls and Nano Mode inputs persist per slide and are restored after save/reopen.
* **Linked Sequence (MVP)**

  * Enable per slide and choose a **Next slide**. Includes **Play from this slide** and **Advance now** actions.
* **Auto-advance While Running**

  * A per-slide timer that arms on **Start/Auto-Start** and cancels on **Stop** or on slide change.
* **Inherit on First Visit**

  * Optionally copy Simulation settings (and optionally Nano Style) **once** when visiting the mapped next slide for the first time, only if it has no saved state yet.
* **Live Preview & Projection**

  * Built-in canvas preview, **Snapshot to slide**, **Export/Import JSON**, and **Download PNG**.
* **Image Caching (Nano)**

  * Generated images are cached by a stable style hash.
* **Quality-of-Life**

  * Faster slide switching (hybrid detection of the active slide).
  * Graceful fallbacks when APIs are unavailable.

---

## How to Use

### Simulation Controls

* **Speed / Capacity / Delay** — live sliders; values are saved on `change`.
* **Project to slide** + **Project ms** — periodically snapshot the preview to the slide.
* **Auto-Start on slide** — engine starts automatically when you visit the slide.
* **Stop on slide change** — automatically stop when leaving the slide.
* Actions: **Start**, **Stop**, **Reset**, **Snapshot Now**, **Export JSON**, **Import JSON**, **Download PNG**.

### Linked Sequence

* **Enable linked sequence on this slide** — turns on sequencing for the current slide.
* **Next slide** — pick the target slide; navigation uses Office.js.
* **Play from this slide** / **Advance now** — manual controls.
* **Auto-advance while running** (+ ms) — per-slide timer.
* **Copy Simulation…** / **Also copy Nano style** — one-time inheritance on the **first visit** to the next slide if it has no state yet.

> Programmatic navigation requires an Office.js build that supports
> `presentation.setSelectedSlides` (typically available in recent desktop builds).

### Nano Mode

* Inputs: **Style Theme**, **Seed**, **Prompt Add-on**, **Aspect**, **Caption**, **Auto-increment seed on Re-generate**.
* Actions: **Save Style**, **Style Selected**, **Re-generate Selected**.
* Image generation can be plugged in via:

  // Optional: supplied by host environment
  // Returns a Promise<string | { base64: string }>
  window.NIS_generateImage = async (style) => { /* generate and return base64 */ };
  
* A simple fallback gradient generator is provided for development; results are cached.


## Persistence & Caching

All state is stored in `Office.context.document.settings`:

* `NIS:scene:<slideId>` → Simulation Controls per slide
* `NIS:style:<slideId>` → Nano Mode per slide
* `NIS:link:<slideId>` → Linked Sequence config (next/auto/inherit…)
* `NIS:img:<hash>` → Nano image cache

An in-memory cache also accelerates slide switching.

---

## Active Source Files

```
src/addins/powerpoint/
├─ manifest.xml
├─ taskpane.html   # Task pane UI
└─ taskpane.js     # Task pane logic (all features above)
```

> These are the only files you need to modify for UI/logic changes in the current setup.

---

## Keyboard Shortcuts

* **Ctrl + Alt + S** — Toggle Start/Stop (re-arms auto-advance if applicable)
* **Ctrl + Alt + Right** — Advance now

---

## Developer Notes

* Recommended dev flow:

  pnpm run dev-addin
  pnpm run start-addin

* Opening `taskpane.html` directly in a browser will log *“Office.js is loaded outside of Office”* — that’s okay for previewing the UI.
* If navigation APIs are unavailable, the add-in degrades gracefully: UI shows hints and other features continue to work.

---

## Troubleshooting

* **Advance/Auto-advance does nothing:** Ensure your PowerPoint build supports `presentation.setSelectedSlides`.
* **HTTPS errors during development:**
  `npx office-addin-dev-certs install`
* **Projection does not appear on the slide:** Enable **Project to slide** and make sure there is a valid insertion target/selection on the slide.

---

## Roadmap

* **Nano Generator Integration:** progress/cancel UI around `NIS_generateImage`, resilient caching, and placeholder-aware insert.
* **Sequence Enhancements:** loops/branches, run stats, lightweight status indicator.
* **Tooling:** automated tests and linting.

---

## License

Add your license here (MIT/Apache-2.0/etc.), or keep internal if this is private.

```
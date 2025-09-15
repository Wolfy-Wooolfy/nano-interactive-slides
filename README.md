# Nano Interactive Slides (NIS)

> Turn slides into *living simulations* with live, on‑slide editing and quasi‑3D reality layers.

## Why NIS?
- **Slides as Living Simulations:** Convert static visuals into interactive, parameterized simulations.
- **Dynamic Real‑Time Slide Editor:** Edit elements on slide *during* the presentation.
- **Nano Reality Layers:** Explore slides as quasi‑3D layered scenes – no VR hardware needed.

## MVP Scope
- PowerPoint Add‑in (Task Pane): single showcase **Supply Chain Simulation**.
- Basic element inspector + live parameter knobs (speed, capacity, delay).

## Repo Layout
```
docs/                         # Specs, roadmap, and reference docs
src/
  addins/
    powerpoint/               # Office Add-in (manifest + taskpane app)
      taskpane/
    google-slides/            # (Phase 4) Slides add-on placeholder
  webapp/                     # Shared UI components (React/Vite placeholder)
packages/
  simulation-engine/          # Core state/update/render loop (headless)
  reality-layers/             # Quasi-3D layering helpers
examples/                     # Sample decks, JSON scenes, demos
.github/
  ISSUE_TEMPLATE/             # Bug/feature templates
.vscode/                      # Recommended editor settings
```

## Getting Started (Dev)
1. Install Node.js 20+ and pnpm or npm.
2. Install deps:
   ```bash
   pnpm install
   # or: npm install
   ```
3. PowerPoint add-in (dev server):
   ```bash
   pnpm -C src/addins/powerpoint dev
   # opens https://localhost:3000 and sideloads manifest.xml
   ```

## Contributing
Please read **CONTRIBUTING.md** and follow our **CODE_OF_CONDUCT.md**.

## License
MIT

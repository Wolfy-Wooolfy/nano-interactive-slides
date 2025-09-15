# MVP Spec — PowerPoint Add-in

## User Story
As a presenter, I want to turn a static supply chain diagram into an interactive simulation to explain bottlenecks live.

## Features
- Start/Stop simulation
- Adjust node/edge parameters (throughput, delay)
- Simple charts for queue length / throughput
- Reset button

## Tech
- Task Pane (Office.js)
- React/Vite UI served on https://localhost:3000
- Headless simulation-engine package

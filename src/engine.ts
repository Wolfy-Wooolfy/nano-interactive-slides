export type NodeType = "producer" | "buffer" | "consumer";

export type Scene = {
  nodes: { id: string; type: NodeType; rate?: number; capacity?: number }[];
  edges: { from: string; to: string; delay?: number }[];
  params: { tickMs: number; initialStock: Record<string, number> };
};

export type Snapshot = { t: number; stock: Record<string, number>; flow: Record<string, number>; nano: boolean };

export class Engine {
  private scene: Scene;
  private time = 0;
  private stock: Record<string, number>;
  private flow: Record<string, number> = {};
  private nano = false;
  private running = false;
  private interval: any = null;

  constructor(scene: Scene) {
    this.scene = scene;
    this.stock = { ...scene.params.initialStock };
  }

  load(scene: Scene) {
    this.scene = scene;
    this.stock = { ...scene.params.initialStock };
    this.flow = {};
    this.time = 0;
  }

  start() {
    if (this.running) return;
    this.running = true;
    this.interval = setInterval(() => this.tick(), Math.max(16, this.scene.params.tickMs));
  }

  stop() {
    this.running = false;
    if (this.interval) clearInterval(this.interval);
    this.interval = null;
  }

  toggleNano() {
    this.nano = !this.nano;
    return this.nano;
  }

  setParam(key: string, value: number) {
    if (key === "tickMs") this.scene.params.tickMs = Math.max(16, value);
  }

  getSnapshot(): Snapshot {
    return { t: this.time, stock: { ...this.stock }, flow: { ...this.flow }, nano: this.nano };
  }

  private tick() {
    this.time += this.scene.params.tickMs;

    for (const n of this.scene.nodes) {
      if (n.type === "producer" && n.rate) {
        this.stock[n.id] = (this.stock[n.id] ?? 0) + n.rate;
        this.flow[n.id] = n.rate;
      } else if (n.type === "consumer" && n.rate) {
        const have = this.stock[n.id] ?? 0;
        const take = Math.min(have, n.rate);
        this.stock[n.id] = have - take;
        this.flow[n.id] = -take;
      } else {
        this.flow[n.id] = 0;
      }
    }

    for (const e of this.scene.edges) {
      const d = e.delay ?? 0;
      if (d <= 0) {
        const fromStock = this.stock[e.from] ?? 0;
        if (fromStock > 0) {
          const move = Math.min(fromStock, 1);
          this.stock[e.from] = fromStock - move;
          this.stock[e.to] = (this.stock[e.to] ?? 0) + move;
        }
      }
    }
  }
}

// Stage 2 public API: gradients of any output cell with respect to every input cell, and a report
// of the kinks (IF / MIN / MAX / ...) that sit near the current inputs.

import { Workbook, Cell } from './workbook';
import { Recorder, Kink, InputLeaf, LoopInfo, plain } from './ad';
import { Dual } from './tape';
import { fullAddr, parseA1 } from './refs';
import { XLError } from './values';

export interface KinkReport {
  cell: string;
  kind: Kink['kind'];
  detail: string;
  /** Distance to the kink in the margin's own units (0 = sitting exactly on it). */
  margin: number;
  onKink: boolean;
  /** False when no input can move the margin (e.g. a balance check that is 0 by construction). */
  movable: boolean;
  /**
   * To first order, the smallest proportional move of *all* inputs at once (each by the same
   * percentage, in the worst direction) that reaches the kink. Infinity if no input can move it
   * proportionally (e.g. it only depends on inputs that are 0, or on none).
   */
  uniformMovePct: number;
  /** The single input that reaches the kink with the smallest move, and how big that move is. */
  nearestInput?: { addr: string; value: number; absChange: number; pctChange: number };
  near: boolean;
}

export interface GradientResult {
  output: string;
  value: number;
  /** d(output)/d(input), aligned with `Gradients.inputs`. */
  d: Float64Array;
}

export interface MismatchRow { cell: string; engine: unknown; tape: unknown }

const show = (v: unknown) => (v instanceof XLError ? v.code : v);

export class Gradients {
  readonly rec: Recorder;
  readonly inputs: InputLeaf[];
  readonly kinks: Kink[];
  readonly loops: LoopInfo[];
  readonly warnings: string[];
  /** Cells where the tape's value differs from the engine's (should always be empty). */
  readonly forwardMismatches: MismatchRow[] = [];
  readonly recordMs: number;

  /**
   * Records the tape at the workbook's current values. Unless `resolve` is false, first re-solves
   * circular blocks to machine precision (warm-started from the current values) so the implicit
   * derivative is taken at a true fixed point.
   */
  constructor(public readonly wb: Workbook, opts: { resolve?: boolean; allInputs?: boolean } = {}) {
    if (opts.resolve !== false) wb.recalc({ tight: true, cycleStart: 'current' });
    const t0 = performance.now();
    this.rec = new Recorder(wb);
    if (opts.allInputs) this.rec.addAllInputs();
    this.rec.record();
    this.recordMs = performance.now() - t0;
    this.inputs = this.rec.inputs;
    this.kinks = this.rec.kinks;
    this.loops = this.rec.loops;
    this.warnings = this.rec.warnings;
    for (const c of wb.allFormulaCells()) {
      const tv = plain(this.rec.nodeOf.get(c) ?? null);
      const ev = c.value;
      const same = tv === ev || (tv instanceof XLError && ev instanceof XLError && tv.code === ev.code) || (typeof tv === 'number' && typeof ev === 'number' && Object.is(tv + 0, ev + 0));
      if (!same) this.forwardMismatches.push({ cell: fullAddr(c.sheet, c.row, c.col), engine: show(ev), tape: show(tv) });
    }
  }

  cellAt(ref: string): Cell {
    const i = ref.lastIndexOf('!');
    const sheet = ref.slice(0, i).replace(/^'|'$/g, '').replace(/''/g, "'"), addr = ref.slice(i + 1).replace(/\$/g, '');
    const c = this.wb.getCell(sheet, addr);
    if (!c || !parseA1(addr)) throw new Error(`no cell ${ref}`);
    return c;
  }

  /** The output's tape node, or null if no input can change it. */
  nodeFor(cell: Cell): Dual | number | null {
    const v = this.rec.nodeOf.get(cell);
    if (v instanceof Dual) return v;
    return typeof v === 'number' ? v : null;
  }

  /** One backward pass: d(output)/d(every input). */
  gradient(output: string | Cell): GradientResult {
    const cell = typeof output === 'string' ? this.cellAt(output) : output;
    const node = this.nodeFor(cell);
    const d = new Float64Array(this.inputs.length);
    const out = { output: fullAddr(cell.sheet, cell.row, cell.col), value: NaN, d };
    if (node === null) return out;
    if (!(node instanceof Dual)) return { ...out, value: node };
    const adj = this.rec.tape.backward([[node.i, 1]]);
    this.inputs.forEach((x, k) => (d[k] = adj[x.node]));
    return { ...out, value: node.v };
  }

  /** Gradient as { "Sheet!A1": derivative }. */
  gradientByAddr(output: string | Cell): Record<string, number> {
    const g = this.gradient(output);
    return Object.fromEntries(this.inputs.map((x, k) => [x.addr, g.d[k]]));
  }

  /** Formula cells the output reads, directly or indirectly (including itself). */
  upstreamCells(output: string | Cell): Set<string> {
    const cell = typeof output === 'string' ? this.cellAt(output) : output;
    const seen = new Set<Cell>([cell]);
    const stack = [cell];
    while (stack.length) for (const d of this.wb.deps.get(stack.pop()!) ?? []) if (!seen.has(d)) { seen.add(d); stack.push(d); }
    return new Set([...seen].map((c) => fullAddr(c.sheet, c.row, c.col)));
  }

  /**
   * For every kink, how far the current inputs are from flipping it.
   * `nearPct`: a kink counts as "near" if moving all inputs by at most this percentage could reach it.
   */
  kinkReport(nearPct = 5, only?: Set<string>): KinkReport[] {
    const out: KinkReport[] = [];
    for (const k of this.kinks) {
      if (only && !only.has(k.cell)) continue;
      const adj = this.rec.tape.backward([[k.margin.i, 1]]);
      const m = k.margin.v;
      let denom = 0;
      let best: KinkReport['nearestInput'];
      for (const x of this.inputs) {
        const g = adj[x.node];
        if (g === 0) continue;
        denom += Math.abs(g * x.value);
        const absChange = Math.abs(m / g);
        const pctChange = x.value !== 0 ? (absChange / Math.abs(x.value)) * 100 : Infinity;
        if (!best || pctChange < best.pctChange || (pctChange === best.pctChange && absChange < best.absChange)) best = { addr: x.addr, value: x.value, absChange, pctChange };
      }
      const movable = best !== undefined;
      const uniformMovePct = !movable ? Infinity : m === 0 ? 0 : denom > 0 ? (Math.abs(m) / denom) * 100 : Infinity;
      out.push({ cell: k.cell, kind: k.kind, detail: k.detail, margin: m, onKink: m === 0, movable, uniformMovePct, nearestInput: best, near: movable && (m === 0 || uniformMovePct <= nearPct) });
    }
    return out.sort((a, b) => a.uniformMovePct - b.uniformMovePct);
  }
}

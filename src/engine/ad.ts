// Recording the tape: a second evaluator that mirrors the engine's Excel rules exactly, but carries
// derivatives along with the numbers.
//
// Any number that depends on an input cell is a `Dual` (its value plus its node on the tape). Text,
// TRUE/FALSE, blanks, errors and numbers that no input can change stay plain. After recording we
// check that every cell's value from this evaluator is bit-for-bit the value the engine computed,
// so the tape describes the same function the engine (and Excel) evaluates.
//
// Kinks: IF, comparisons, MIN, MAX, ABS, ROUND and XNPV's date truncation make the model piecewise.
// The derivative we report is the one for the branch Excel actually took. Each time a branch
// decision depends on an input, we record a "margin" node: a number that is positive on one side
// of the kink and crosses zero exactly where the branch would flip. See grad.ts for the report.

import { Node } from './parser';
import { Matrix, Scalar, XLError, isError } from './values';
import { compare, toBool, toNumber, toText } from './coerce';
import { MAX_COL, MAX_ROW, cellKey, fullAddr } from './refs';
import { solveIrr } from './functions';
import { Dual, Tape } from './tape';
import type { Cell, Workbook } from './workbook';

export type DS = Scalar | Dual;

export class DMatrix {
  constructor(public readonly rows: DS[][]) {}
  get height() { return this.rows.length; }
  get width() { return this.rows[0]?.length ?? 0; }
  get(r: number, c: number): DS { return this.rows[r][c]; }
  *values(): Generator<DS> { for (const row of this.rows) for (const v of row) yield v; }
}

export class DRange {
  constructor(
    public readonly sheet: string,
    public readonly r1: number, public readonly c1: number,
    public readonly r2: number, public readonly c2: number,
    private readonly rec: Recorder,
  ) {}
  get isCell() { return this.r1 === this.r2 && this.c1 === this.c2; }
  get height() { return this.r2 - this.r1 + 1; }
  get width() { return this.c2 - this.c1 + 1; }
  cell(r: number, c: number): DS { return this.rec.read(this.sheet, r, c); }
  toMatrix(): DMatrix {
    const b = this.rec.wb.getSheet(this.sheet)!;
    const r2 = Math.min(this.r2, Math.max(b.maxRow, this.r1));
    const c2 = Math.min(this.c2, Math.max(b.maxCol, this.c1));
    const rows: DS[][] = [];
    for (let r = this.r1; r <= r2; r++) {
      const row: DS[] = [];
      for (let c = this.c1; c <= c2; c++) row.push(this.cell(r, c));
      rows.push(row);
    }
    return new DMatrix(rows);
  }
}

export type DV = DS | DMatrix | DRange;

interface Ctx { sheet: string; row: number; col: number; array: boolean; cell: string }

export type KinkKind = 'compare' | 'IF' | 'AND/OR' | 'MIN' | 'MAX' | 'ABS' | 'ROUND' | 'XNPV date';

/** One branch decision that depends on the inputs. `margin` crosses zero where the branch flips. */
export interface Kink { cell: string; kind: KinkKind; detail: string; margin: Dual }

export interface InputLeaf { cell: Cell; addr: string; node: number; value: number }

/** A circular block as recorded: `residual` = how far the solved values are from a true fixed point (scaled). */
export interface LoopInfo { cells: string[]; residual: number }

const plain = (x: DS): Scalar => (x instanceof Dual ? x.v : x);
const isNum = (x: DS): x is number | Dual => typeof x === 'number' || x instanceof Dual;
const numVal = (x: number | Dual) => (x instanceof Dual ? x.v : x);

function dToNumber(x: DS): number | Dual | XLError {
  return x instanceof Dual ? x : toNumber(x);
}

function toDMatrix(v: DV): DMatrix {
  if (v instanceof DMatrix) return v;
  if (v instanceof DRange) return v.toMatrix();
  return new DMatrix([[v]]);
}

/** Same rule as the engine's normalise(): non-finite numbers become #NUM!, -0 becomes 0. */
function normalise(v: DS): DS {
  const x = v instanceof Dual ? v.v : v;
  if (typeof x === 'number' && !Number.isFinite(x)) return new XLError('#NUM!');
  if (typeof v === 'number') return v === 0 ? 0 : v;
  return v;
}

const err = (code: XLError['code'], detail?: string) => new XLError(code, detail);

export class Recorder {
  tape = new Tape();
  /** Each formula cell's value as recorded on the tape (a Dual if it depends on an input). */
  nodeOf = new Map<Cell, DS>();
  inputs: InputLeaf[] = [];
  private inputOf = new Map<Cell, Dual>();
  kinks: Kink[] = [];
  loops: LoopInfo[] = [];
  warnings: string[] = [];
  /** Loop cells currently being linearised: reads return their "loop leaf" instead. */
  private loopLeaf = new Map<Cell, DS>();

  constructor(public readonly wb: Workbook) {}

  // ---------- recording ----------

  record() {
    const wb = this.wb;
    for (const comp of wb.order) {
      const circular = comp.length > 1 || (wb.deps.get(comp[0]) ?? []).includes(comp[0]);
      if (!circular) { this.evalCell(comp[0]); continue; }
      this.recordLoop(comp);
    }
  }

  private recordLoop(comp: Cell[]) {
    const t = this.tape;
    const lo = t.size;
    // 1. One leaf per numeric loop cell, holding its solved value x*.
    const leaves: (Dual | null)[] = comp.map((c) => {
      if (c.array) throw new Error(`array formula inside a circular block is not supported (${fullAddr(c.sheet, c.row, c.col)})`);
      const leaf = typeof c.value === 'number' ? t.leaf(c.value) : null;
      this.loopLeaf.set(c, leaf ?? c.value);
      return leaf;
    });
    // 2. Evaluate every loop formula once at x*: y = F(x*, inputs).
    const kinksBefore = this.kinks.length;
    const ys: DS[] = comp.map((c) => this.evalFormula(c));
    let residual = 0;
    comp.forEach((c, k) => {
      const y = plain(ys[k]);
      if (typeof y === 'number' && typeof c.value === 'number') residual = Math.max(residual, Math.abs(y - c.value) / Math.max(1, Math.abs(c.value)));
      else if (y !== c.value && !(y instanceof XLError && c.value instanceof XLError && y.code === c.value.code)) residual = Infinity;
    });
    // Numeric cells whose formula ended up not depending on anything still need a node.
    const yNodes = ys.map((y, k) => (leaves[k] === null ? null : y instanceof Dual ? y : t.leaf(numVal(y as number))));
    const hi = t.size;
    // 3. J[k][j] = ∂y_k/∂x_j, one backward sweep over just this block per loop cell.
    const idx = comp.map((_, k) => k).filter((k) => leaves[k] !== null);
    const n = idx.length;
    const m = new Float64Array(n * n);
    for (let a = 0; a < n; a++) {
      const k = idx[a];
      const adj = t.backward([[yNodes[k]!.i, 1]], lo, hi, false);
      for (let b = 0; b < n; b++) {
        const jkb = adj[leaves[idx[b]]!.i - lo]; // ∂y_k/∂x_b
        // M = I − Jᵀ, so M[b][a] = δ − J[a][b]
        m[b * n + a] = (a === b ? 1 : 0) - jkb;
      }
    }
    // 4. X_k: the solved loop value, wired to y_k with slope 1. The block's linear solve happens
    //    when the backward pass reaches the last X.
    const xs: number[] = [];
    comp.forEach((c, k) => {
      this.loopLeaf.delete(c);
      if (leaves[k] === null) { this.nodeOf.set(c, c.value); return; }
      const X = t.node(c.value as number, [yNodes[k]!.i], [1]);
      xs.push(X.i);
      this.nodeOf.set(c, X);
    });
    if (n) t.blocks.push({ xs, trigger: xs[xs.length - 1], m, n });
    // 5. Kinks found inside the loop: their margin was recorded against the loop leaves, which are
    //    dead ends on the tape. Re-wire each margin through the X nodes so its gradient includes
    //    how the whole loop moves (margin' = margin + Σ ∂margin/∂x_j · (X_j − x_j), same value).
    for (let q = kinksBefore; q < this.kinks.length; q++) {
      const mg = this.kinks[q].margin;
      const adj = t.backward([[mg.i, 1]], lo, hi, false);
      const slopes = idx.map((k) => adj[leaves[k]!.i - lo]);
      this.kinks[q].margin = t.node(mg.v, [mg.i, ...xs], [1, ...slopes]);
    }
    this.loops.push({ cells: comp.map((c) => fullAddr(c.sheet, c.row, c.col)), residual });
  }

  read(sheet: string, r: number, c: number): DS {
    const cell = this.wb.getSheet(sheet)?.cells.get(cellKey(r, c));
    if (!cell) return null;
    if (this.loopLeaf.has(cell)) return this.loopLeaf.get(cell)!;
    if (cell.node || cell.arrayMaster || cell.parseError) return this.nodeOf.has(cell) ? this.nodeOf.get(cell)! : cell.value;
    if (typeof cell.value !== 'number') return cell.value;
    let d = this.inputOf.get(cell);
    if (!d) {
      d = this.tape.leaf(cell.value);
      this.inputOf.set(cell, d);
      this.inputs.push({ cell, addr: fullAddr(cell.sheet, cell.row, cell.col), node: d.i, value: cell.value });
    }
    return d;
  }

  /** Make sure every numeric constant cell has a leaf, even ones no formula reads (their gradient is 0). */
  addAllInputs() {
    for (const name of this.wb.sheetOrder) {
      const sh = this.wb.getSheet(name)!;
      for (const c of sh.cells.values()) if (!c.node && !c.arrayMaster && typeof c.value === 'number') this.read(c.sheet, c.row, c.col);
    }
  }

  private evalFormula(c: Cell): DS {
    if (!c.node) return new XLError('#NAME?', c.parseError);
    const ctx: Ctx = { sheet: c.sheet, row: c.row, col: c.col, array: !!c.array, cell: fullAddr(c.sheet, c.row, c.col) };
    const v = this.eval(c.node, ctx);
    return normalise(v instanceof DRange ? this.intersect(v, ctx) : v instanceof DMatrix ? v.get(0, 0) : v);
  }

  private evalCell(c: Cell) {
    if (!c.array) { this.nodeOf.set(c, this.evalFormula(c)); return; }
    const ctx: Ctx = { sheet: c.sheet, row: c.row, col: c.col, array: true, cell: fullAddr(c.sheet, c.row, c.col) };
    const m = toDMatrix(this.eval(c.node!, ctx));
    const sh = this.wb.getSheet(c.sheet)!;
    for (let r = c.array.r1; r <= c.array.r2; r++) for (let col = c.array.c1; col <= c.array.c2; col++) {
      const i = r - c.array.r1, j = col - c.array.c1;
      const x = m.height === 1 && m.width === 1 ? m.get(0, 0)
        : i < m.height && j < m.width ? m.get(i, j) : m.height === 1 && j < m.width ? m.get(0, j) : m.width === 1 && i < m.height ? m.get(i, 0) : new XLError('#N/A');
      this.nodeOf.set(sh.cells.get(cellKey(r, col))!, normalise(x));
    }
  }

  // ---------- tape helpers ----------

  /** Record value v with local slopes to its Dual parents. Returns a plain number if none are Duals. */
  mk(v: number, deps: [DS, number][]): number | Dual {
    const ps: number[] = [], ws: number[] = [];
    for (const [x, s] of deps) if (x instanceof Dual) { ps.push(x.i); ws.push(s); }
    return ps.length ? this.tape.node(v, ps, ws) : v;
  }

  kink(ctx: Ctx, kind: KinkKind, detail: string, margin: number | Dual) {
    if (margin instanceof Dual) this.kinks.push({ cell: ctx.cell, kind, detail, margin });
  }

  // ---------- evaluation (mirrors Workbook.eval) ----------

  eval(n: Node, ctx: Ctx): DV {
    switch (n.t) {
      case 'num': return n.v;
      case 'str': return n.v;
      case 'bool': return n.v;
      case 'err': return new XLError(n.v);
      case 'missing': return null;
      case 'ref': {
        const sheet = n.sheet ?? ctx.sheet;
        const sh = this.wb.getSheet(sheet);
        if (!sh) return err('#REF!', `no sheet "${sheet}"`);
        if (n.r1 < 1 || n.c1 < 1 || n.r2 > MAX_ROW || n.c2 > MAX_COL) return err('#REF!');
        return new DRange(sh.name, n.r1, n.c1, n.r2, n.c2, this);
      }
      case 'name': {
        const def = this.wb.lookupName(n.name, n.sheet ?? ctx.sheet);
        if (!def?.node) return err('#NAME?', `unknown name ${n.name}`);
        return this.eval(def.node, ctx);
      }
      case 'fn': {
        const f = AD_FUNCTIONS[n.name];
        if (!f) return err('#NAME?', `unsupported function ${n.name}`);
        return f(n.args, ctx, this);
      }
      case 'neg': return this.unary(n.e, ctx, (x) => this.mk(-numVal(x), [[x, -1]]));
      case 'pos': return this.eval(n.e, ctx);
      case 'pct': return this.unary(n.e, ctx, (x) => this.mk(numVal(x) / 100, [[x, 1 / 100]]));
      case 'arr': return new DMatrix(n.rows.map((r) => r.map((x) => this.eval(x, ctx) as DS)));
      case 'bin': return this.binary(n.op, this.operand(this.eval(n.l, ctx), ctx), this.operand(this.eval(n.r, ctx), ctx), ctx);
    }
  }

  scalar(n: Node, ctx: Ctx): DS { return this.intersect(this.eval(n, ctx), ctx); }

  intersect(v: DV, ctx: Ctx): DS {
    if (v instanceof DMatrix) return v.get(0, 0);
    if (!(v instanceof DRange)) return v;
    if (v.isCell) return v.cell(v.r1, v.c1);
    if (v.width === 1 && ctx.row >= v.r1 && ctx.row <= v.r2) return v.cell(ctx.row, v.c1);
    if (v.height === 1 && ctx.col >= v.c1 && ctx.col <= v.c2) return v.cell(v.r1, ctx.col);
    return err('#VALUE!', 'range does not intersect the formula cell');
  }

  private operand(v: DV, ctx: Ctx): DS | DMatrix {
    if (v instanceof DRange) return v.isCell ? v.cell(v.r1, v.c1) : ctx.array ? v.toMatrix() : this.intersect(v, ctx);
    return v;
  }

  private unary(e: Node, ctx: Ctx, f: (x: number | Dual) => number | Dual): DS | DMatrix {
    const v = this.operand(this.eval(e, ctx), ctx);
    const one = (s: DS): DS => { const x = dToNumber(s); return isError(x) ? x : f(x); };
    return v instanceof DMatrix ? new DMatrix(v.rows.map((r) => r.map(one))) : one(v);
  }

  private binary(op: string, a: DS | DMatrix, b: DS | DMatrix, ctx: Ctx): DS | DMatrix {
    if (a instanceof DMatrix || b instanceof DMatrix) {
      const A = toDMatrix(a), B = toDMatrix(b);
      const h = Math.max(A.height, B.height), w = Math.max(A.width, B.width);
      const pick = (M: DMatrix, r: number, c: number): DS => {
        const rr = M.height === 1 ? 0 : r, cc = M.width === 1 ? 0 : c;
        return rr < M.height && cc < M.width ? M.get(rr, cc) : err('#N/A');
      };
      const rows: DS[][] = [];
      for (let r = 0; r < h; r++) { const row: DS[] = []; for (let c = 0; c < w; c++) row.push(this.scalarOp(op, pick(A, r, c), pick(B, r, c), ctx)); rows.push(row); }
      return new DMatrix(rows);
    }
    return this.scalarOp(op, a, b, ctx);
  }

  scalarOp(op: string, a: DS, b: DS, ctx: Ctx): DS {
    if (op === '&') {
      const x = toText(plain(a)), y = toText(plain(b));
      if (isError(x)) return x; if (isError(y)) return y;
      return x + y;
    }
    if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
      const pa = plain(a), pb = plain(b);
      if (isError(pa)) return pa; if (isError(pb)) return pb;
      const c = compare(pa, pb);
      // A comparison between numbers (blank counts as 0) that an input can move is a kink.
      const na = a === null ? 0 : a, nb = b === null ? 0 : b;
      if ((a instanceof Dual || b instanceof Dual) && isNum(na) && isNum(nb)) {
        this.kink(ctx, 'compare', `${op} (left − right = ${numVal(na) - numVal(nb)})`, this.mk(numVal(na) - numVal(nb), [[na, 1], [nb, -1]]));
      }
      switch (op) {
        case '=': return c === 0; case '<>': return c !== 0; case '<': return c < 0;
        case '>': return c > 0; case '<=': return c <= 0; default: return c >= 0;
      }
    }
    const x = dToNumber(a); if (isError(x)) return x;
    const y = dToNumber(b); if (isError(y)) return y;
    const xv = numVal(x), yv = numVal(y);
    switch (op) {
      case '+': return this.mk(xv + yv, [[x, 1], [y, 1]]);
      case '-': return this.mk(xv - yv, [[x, 1], [y, -1]]);
      case '*': return this.mk(xv * yv, [[x, yv], [y, xv]]);
      case '/': return yv === 0 ? err('#DIV/0!') : this.mk(xv / yv, [[x, 1 / yv], [y, -xv / (yv * yv)]]);
      case '^': {
        if (xv === 0 && yv === 0) return err('#NUM!', '0^0');
        if (xv === 0 && yv < 0) return err('#DIV/0!');
        const r = Math.pow(xv, yv);
        if (!Number.isFinite(r)) return err('#NUM!');
        let dy = 0;
        if (y instanceof Dual) {
          if (xv > 0) dy = r * Math.log(xv);
          else this.warnings.push(`${ctx.cell}: x^y with x ≤ 0 and y depending on an input; ∂/∂y set to 0`);
        }
        const dx = yv === 0 ? 0 : yv * Math.pow(xv, yv - 1);
        if (x instanceof Dual && !Number.isFinite(dx)) this.warnings.push(`${ctx.cell}: x^y has an infinite slope at x = 0`);
        return this.mk(r, [[x, dx], [y, dy]]);
      }
    }
    return err('#VALUE!', `unknown operator ${op}`);
  }
}

// ---------- functions (mirror functions.ts) ----------

type AdFn = (args: Node[], ctx: Ctx, ev: Recorder) => DV;
const fromRefOrArray = (v: DV) => v instanceof DRange || v instanceof DMatrix;

function collectNumbers(args: Node[], ctx: Ctx, ev: Recorder): (number | Dual)[] | XLError {
  const out: (number | Dual)[] = [];
  for (const a of args) {
    if (a.t === 'missing') { out.push(0); continue; }
    const v = ev.eval(a, ctx);
    if (fromRefOrArray(v)) {
      for (const x of toDMatrix(v).values()) {
        if (isNum(x)) out.push(x);
        else if (isError(x)) return x;
      }
    } else {
      const n = dToNumber(v as DS);
      if (isError(n)) return n;
      out.push(n);
    }
  }
  return out;
}

function num(n: Node | undefined, ctx: Ctx, ev: Recorder, dflt?: number): number | Dual | XLError {
  if (!n || n.t === 'missing') return dflt ?? 0;
  return dToNumber(ev.scalar(n, ctx));
}

const checkNum = (x: number | Dual): number | Dual | XLError => (Number.isFinite(numVal(x)) ? x : err('#NUM!', 'result is not a finite number'));

/** toBool, recording a kink when an input-dependent number is tested against zero. */
function dBool(x: DS, ctx: Ctx, ev: Recorder, kind: KinkKind): boolean | XLError {
  if (x instanceof Dual) ev.kink(ctx, kind, `number tested as TRUE/FALSE (value ${x.v})`, x);
  return toBool(plain(x));
}

function collectBools(args: Node[], ctx: Ctx, ev: Recorder): boolean[] | XLError {
  const out: boolean[] = [];
  for (const a of args) {
    const v = ev.eval(a, ctx);
    if (fromRefOrArray(v)) {
      for (const x of toDMatrix(v).values()) {
        if (typeof x === 'boolean') out.push(x);
        else if (isNum(x)) out.push(dBool(x, ctx, ev, 'AND/OR') as boolean);
        else if (isError(x)) return x;
      }
    } else {
      const b = dBool(v as DS, ctx, ev, 'AND/OR');
      if (isError(b)) return b;
      out.push(b);
    }
  }
  if (out.length === 0) return err('#VALUE!', 'AND/OR got no logical values');
  return out;
}

/** NPV-style discounting with exact slopes for every flow and for the rate. */
function npvNode(ev: Recorder, rate: number | Dual, flows: (number | Dual)[], times: number[]): number | Dual {
  const r = numVal(rate);
  let s = 0, dr = 0;
  const deps: [DS, number][] = [];
  for (let i = 0; i < flows.length; i++) {
    const f = numVal(flows[i]);
    const d = Math.pow(1 + r, times[i]);
    s += f / d;
    deps.push([flows[i], 1 / d]);
    dr -= (times[i] * f) / (d * (1 + r));
  }
  deps.push([rate, dr]);
  return ev.mk(s, deps);
}

function minMax(ev: Recorder, xs: (number | Dual)[], ctx: Ctx, kind: 'MIN' | 'MAX'): number | Dual {
  if (xs.length === 0) return 0;
  const better = kind === 'MIN' ? (b: number, a: number) => b < a : (b: number, a: number) => b > a;
  let k = 0;
  for (let i = 1; i < xs.length; i++) if (better(numVal(xs[i]), numVal(xs[k]))) k = i;
  if (xs.length > 1) {
    // Runner-up: the best of the others. The branch flips when the gap between them hits zero.
    let r = k === 0 ? 1 : 0;
    for (let i = 0; i < xs.length; i++) if (i !== k && better(numVal(xs[i]), numVal(xs[r]))) r = i;
    const sign = kind === 'MIN' ? 1 : -1;
    const gap = sign * (numVal(xs[r]) - numVal(xs[k]));
    ev.kink(ctx, kind, `picks argument ${k + 1} of ${xs.length}; runner-up is argument ${r + 1}, gap ${gap}`, ev.mk(gap, [[xs[r], sign], [xs[k], -sign]]));
  }
  return xs[k];
}

export const AD_FUNCTIONS: Record<string, AdFn> = {
  SUM(args, ctx, ev) {
    const xs = collectNumbers(args, ctx, ev);
    if (isError(xs)) return xs;
    let s = 0; for (const x of xs) s += numVal(x);
    return checkNum(ev.mk(s, xs.map((x) => [x, 1])));
  },
  AVERAGE(args, ctx, ev) {
    const xs = collectNumbers(args.filter((a) => a.t !== 'missing'), ctx, ev);
    if (isError(xs)) return xs;
    if (xs.length === 0) return err('#DIV/0!', 'AVERAGE of no numbers');
    let s = 0; for (const x of xs) s += numVal(x);
    return ev.mk(s / xs.length, xs.map((x) => [x, 1 / xs.length]));
  },
  MIN(args, ctx, ev) { const xs = collectNumbers(args, ctx, ev); return isError(xs) ? xs : minMax(ev, xs, ctx, 'MIN'); },
  MAX(args, ctx, ev) { const xs = collectNumbers(args, ctx, ev); return isError(xs) ? xs : minMax(ev, xs, ctx, 'MAX'); },
  SUMPRODUCT(args, ctx, ev) {
    const actx = { ...ctx, array: true };
    const ms = args.map((a) => toDMatrix(ev.eval(a, actx)));
    const h = ms[0].height, w = ms[0].width;
    if (ms.some((m) => m.height !== h || m.width !== w)) return err('#VALUE!', 'SUMPRODUCT arrays differ in size');
    let s = 0;
    const deps: [DS, number][] = [];
    for (let r = 0; r < h; r++) for (let c = 0; c < w; c++) {
      let p = 1;
      const xs: DS[] = [], vs: number[] = [];
      for (const m of ms) {
        const x = m.get(r, c);
        if (isError(x)) return x;
        const v = isNum(x) ? numVal(x) : 0; // text, booleans and blanks count as 0
        p *= v; xs.push(x); vs.push(v);
      }
      s += p;
      for (let k = 0; k < xs.length; k++) if (xs[k] instanceof Dual) {
        let others = 1; for (let j = 0; j < vs.length; j++) if (j !== k) others *= vs[j];
        deps.push([xs[k], others]);
      }
    }
    return ev.mk(s, deps);
  },
  IF(args, ctx, ev) {
    if (args.length < 1 || args.length > 3) return err('#VALUE!', 'IF takes 1 to 3 arguments');
    const cv = ev.eval(args[0], ctx); // evaluated once (the engine evaluates it twice; the result is the same)
    const branch = (b: boolean): DV => {
      const n = args[b ? 1 : 2];
      if (n === undefined) return b;
      if (n.t === 'missing') return 0;
      return ev.eval(n, ctx);
    };
    if (ctx.array && (cv instanceof DMatrix || (cv instanceof DRange && !cv.isCell))) {
      const m = toDMatrix(cv);
      const t = toDMatrix(branch(true)), f = toDMatrix(branch(false));
      const pick = (x: DMatrix, r: number, c: number) => x.get(Math.min(r, x.height - 1), Math.min(c, x.width - 1));
      return new DMatrix(m.rows.map((row, r) => row.map((x, c) => {
        const b = dBool(x, ctx, ev, 'IF');
        return isError(b) ? b : pick(b ? t : f, r, c);
      })));
    }
    const b = dBool(ev.intersect(cv, ctx), ctx, ev, 'IF');
    if (isError(b)) return b;
    return branch(b);
  },
  AND(args, ctx, ev) { const bs = collectBools(args, ctx, ev); return isError(bs) ? bs : bs.every((x) => x); },
  OR(args, ctx, ev) { const bs = collectBools(args, ctx, ev); return isError(bs) ? bs : bs.some((x) => x); },
  NOT(args, ctx, ev) { const b = dBool(ev.scalar(args[0], ctx), ctx, ev, 'AND/OR'); return isError(b) ? b : !b; },
  NPV(args, ctx, ev) {
    if (args.length < 2) return err('#VALUE!', 'NPV needs a rate and at least one value');
    const rate = num(args[0], ctx, ev);
    if (isError(rate)) return rate;
    const xs = collectNumbers(args.slice(1), ctx, ev);
    if (isError(xs)) return xs;
    if (numVal(rate) === -1) return err('#DIV/0!', 'NPV rate of -100%');
    return checkNum(npvNode(ev, rate, xs, xs.map((_, i) => i + 1)));
  },
  XNPV(args, ctx, ev) {
    const rate = num(args[0], ctx, ev);
    if (isError(rate)) return rate;
    const vals = [...toDMatrix(ev.eval(args[1], ctx)).values()];
    const dates = [...toDMatrix(ev.eval(args[2], ctx)).values()];
    for (const x of [...vals, ...dates]) { if (isError(x)) return x; if (!isNum(x)) return err('#VALUE!', 'XNPV values/dates must be numbers'); }
    if (vals.length !== dates.length) return err('#NUM!', 'XNPV values and dates differ in length');
    if (numVal(rate) <= -1) return err('#NUM!', 'XNPV rate must be > -100%');
    // Excel truncates dates to whole days, so the result is a step function of each date: slope 0,
    // with a kink wherever a date crosses midnight.
    for (const x of dates) if (x instanceof Dual) {
      const frac = x.v - Math.trunc(x.v);
      const down = frac, up = 1 - frac;
      ev.kink(ctx, 'XNPV date', `date ${x.v} is truncated to a whole day`, ev.mk(down <= up ? down : up, [[x, down <= up ? 1 : -1]]));
    }
    const d = (dates as (number | Dual)[]).map((x) => Math.trunc(numVal(x)));
    if (d.some((x) => x < d[0])) return err('#NUM!', 'XNPV date precedes the first date');
    return checkNum(npvNode(ev, rate, vals as (number | Dual)[], d.map((x) => (x - d[0]) / 365)));
  },
  IRR(args, ctx, ev) {
    const flows: (number | Dual)[] = [];
    for (const x of toDMatrix(ev.eval(args[0], ctx)).values()) {
      if (isError(x)) return x;
      if (isNum(x)) flows.push(x);
    }
    const guess = num(args[1], ctx, ev, 0.1);
    if (isError(guess)) return guess;
    const r = solveIrr(flows.map(numVal), numVal(guess));
    if (isError(r)) return r;
    // Implicit function theorem: r solves g(r, f) = Σ f_i (1+r)^-i = 0, so
    // ∂r/∂f_i = −(∂g/∂f_i)/(∂g/∂r) = −(1+r)^-i / g'(r).
    let gr = 0;
    for (let i = 0; i < flows.length; i++) gr -= (i * numVal(flows[i])) / Math.pow(1 + r, i + 1);
    if (gr === 0) { ev.warnings.push(`${ctx.cell}: IRR has a flat NPV curve at the root; slope undefined`); return r; }
    return ev.mk(r, flows.map((f, i) => [f, -1 / Math.pow(1 + r, i) / gr]));
  },
  ABS(args, ctx, ev) {
    const x = num(args[0], ctx, ev);
    if (isError(x)) return x;
    if (x instanceof Dual) ev.kink(ctx, 'ABS', `argument ${x.v}`, x);
    const v = numVal(x);
    return ev.mk(Math.abs(v), [[x, v > 0 ? 1 : v < 0 ? -1 : 0]]);
  },
  ROUND(args, ctx, ev) {
    const x = num(args[0], ctx, ev), d = num(args[1], ctx, ev);
    if (isError(x)) return x; if (isError(d)) return d;
    const xv = numVal(x);
    const f = Math.pow(10, Math.trunc(numVal(d)));
    const y = Math.abs(xv) * f;
    const r = Math.round(y + y * Number.EPSILON);
    // ROUND is a staircase: slope 0, with a kink at every half-step boundary.
    if (x instanceof Dual) {
      const dist = Math.abs(y - Math.floor(y) - 0.5) / f;
      const s = Math.sign(y - Math.floor(y) - 0.5) * (xv < 0 ? -1 : 1);
      ev.kink(ctx, 'ROUND', `rounds ${xv} to ${Math.trunc(numVal(d))} digits`, ev.mk(dist, [[x, s]]));
    }
    return (Math.sign(xv) * r) / f;
  },
  IFERROR(args, ctx, ev) {
    const v = ev.eval(args[0], ctx);
    const s = v instanceof DRange || v instanceof DMatrix ? ev.intersect(v, ctx) : v;
    if (isError(s)) return args[1] && args[1].t !== 'missing' ? ev.eval(args[1], ctx) : 0;
    return v;
  },
  COUNT(args, ctx, ev) {
    let n = 0;
    for (const a of args) {
      const v = ev.eval(a, ctx);
      if (fromRefOrArray(v)) { for (const x of toDMatrix(v).values()) if (isNum(x)) n++; }
      else { const p = plain(v as DS); if (typeof p === 'number' || typeof p === 'boolean' || (typeof p === 'string' && !isError(toNumber(p)))) n++; }
    }
    return n;
  },
  TRUE: () => true,
  FALSE: () => false,
  SINGLE(args, ctx, ev) { return ev.scalar(args[0], ctx); },
};

export { plain };
export type { Matrix };

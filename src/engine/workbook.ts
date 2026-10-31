// The workbook model, its dependency graph, and recalculation.
//
// How recalculation works, in plain English:
//  1. Every formula cell is parsed once into a tree (AST).
//  2. We list which cells each formula reads ("precedents"). That gives a directed graph.
//  3. We split the graph into "strongly connected components" (Tarjan's algorithm). A component
//     with one cell is a normal formula; a component with several cells (or a cell that reads
//     itself) is a circular reference.
//  4. Tarjan's algorithm hands the components back in dependency order, so we evaluate each one
//     after everything it reads. Circular components are solved by repeated sweeps, like Excel's
//     "Enable iterative calculation" setting, until no cell changes by more than Max Change.

import { Node, ParseError, RefNode, functionsUsed, parseFormula, shiftNode } from './parser';
import { MAX_COL, MAX_ROW, cellKey, fullAddr, keyCol, keyRow, parseA1 } from './refs';
import { Matrix, RangeRef, Scalar, Value, XLError, isError } from './values';
import { compare, toMatrix, toNumber, toText } from './coerce';
import { Ctx, Evaluator, FUNCTIONS } from './functions';

export interface CellInput {
  /** Formula text without '=' (as stored in the file), if this is a formula cell. */
  formula?: string;
  /** Excel's cached value for formula cells, or the constant for input cells. */
  value: Scalar;
  /** For an array formula: the block it fills (only on its top-left "master" cell). */
  arrayRef?: string;
  /** Pre-parsed AST (used for shared formulas, which are stored once and shifted). */
  node?: Node;
}

export interface SheetInput { name: string; cells: Map<string, CellInput> }
export interface NameInput { name: string; formula: string; scope?: string }
export interface CalcSettings { iterate: boolean; maxIterations: number; maxChange: number }
export interface WorkbookInput { sheets: SheetInput[]; names: NameInput[]; calc: CalcSettings }

export interface Cell {
  sheet: string; row: number; col: number;
  formula?: string; node?: Node; parseError?: string;
  cached: Scalar; value: Scalar;
  array?: { r1: number; c1: number; r2: number; c2: number };
  arrayMaster?: Cell; // set on cells filled by another cell's array formula
}

export interface CycleReport { cells: string[]; iterations: number; converged: boolean; lastChange: number }

export interface RecalcOptions {
  /** Where circular cells start before iterating. 'zero' (default) proves we really solve the loop. */
  cycleStart?: 'zero' | 'cached';
}

class Sheet {
  cells = new Map<number, Cell>();
  formulaCells: Cell[] = [];
  maxRow = 0; maxCol = 0;
  constructor(public name: string) {}
}

const upper = (s: string) => s.toUpperCase();

export class Workbook implements Evaluator {
  sheets = new Map<string, Sheet>(); // keyed by upper-case name
  sheetOrder: string[] = [];
  names = new Map<string, { node?: Node; error?: string; text: string }>(); // key: "SCOPE!NAME" or "!NAME"
  calc: CalcSettings;
  order: Cell[][] = []; // strongly connected components in evaluation order
  cycles: CycleReport[] = [];
  unsupported = new Map<string, string[]>(); // function or feature -> cells that use it

  constructor(input: WorkbookInput) {
    this.calc = input.calc;
    for (const s of input.sheets) {
      const sh = new Sheet(s.name);
      this.sheets.set(upper(s.name), sh);
      this.sheetOrder.push(s.name);
      for (const [addr, ci] of s.cells) {
        const a = parseA1(addr)!;
        const cell: Cell = { sheet: s.name, row: a.row, col: a.col, cached: ci.value, value: ci.value };
        if (ci.formula !== undefined || ci.node) {
          cell.formula = ci.formula;
          try { cell.node = ci.node ?? parseFormula(ci.formula!); }
          catch (e) { cell.parseError = (e as Error).message; this.flag(`PARSE: ${(e as Error).message}`, cell); }
          if (ci.arrayRef) {
            const [p, q] = ci.arrayRef.split(':'); const a1 = parseA1(p)!, a2 = parseA1(q ?? p)!;
            cell.array = { r1: a1.row, c1: a1.col, r2: a2.row, c2: a2.col };
          }
          sh.formulaCells.push(cell);
        }
        sh.cells.set(cellKey(a.row, a.col), cell);
        sh.maxRow = Math.max(sh.maxRow, a.row); sh.maxCol = Math.max(sh.maxCol, a.col);
      }
    }
    // Cells inside an array formula's block point back at the master cell.
    for (const sh of this.sheets.values()) for (const m of sh.formulaCells) if (m.array) {
      for (let r = m.array.r1; r <= m.array.r2; r++) for (let c = m.array.c1; c <= m.array.c2; c++) {
        if (r === m.row && c === m.col) continue;
        let cell = sh.cells.get(cellKey(r, c));
        if (!cell) { cell = { sheet: sh.name, row: r, col: c, cached: null, value: null }; sh.cells.set(cellKey(r, c), cell); }
        cell.arrayMaster = m;
      }
    }
    for (const n of input.names) {
      const key = `${n.scope ? upper(n.scope) : ''}!${upper(n.name)}`;
      try { this.names.set(key, { node: parseFormula(n.formula), text: n.formula }); }
      catch (e) { this.names.set(key, { error: (e as Error).message, text: n.formula }); }
    }
    for (const sh of this.sheets.values()) for (const c of sh.formulaCells) if (c.node) {
      for (const f of functionsUsed(c.node)) if (!FUNCTIONS[f]) this.flag(f, c);
      this.scanNames(c.node, c);
    }
    this.buildGraph();
  }

  private flag(what: string, cell: Cell) {
    const list = this.unsupported.get(what) ?? [];
    list.push(fullAddr(cell.sheet, cell.row, cell.col));
    this.unsupported.set(what, list);
  }

  private scanNames(n: Node, cell: Cell, seen = new Set<string>()) {
    const walk = (x: Node) => {
      if (x.t === 'name') {
        const def = this.lookupName(x.name, x.sheet ?? cell.sheet);
        if (!def) { this.flag(`UNKNOWN NAME: ${x.name}`, cell); return; }
        if (def.error) { this.flag(`NAME PARSE: ${x.name}: ${def.error}`, cell); return; }
        if (def.node && !seen.has(x.name)) { seen.add(x.name); for (const f of functionsUsed(def.node)) if (!FUNCTIONS[f]) this.flag(f, cell); walk(def.node); }
      } else if (x.t === 'fn') x.args.forEach(walk);
      else if (x.t === 'bin') { walk(x.l); walk(x.r); }
      else if (x.t === 'neg' || x.t === 'pos' || x.t === 'pct') walk(x.e);
    };
    walk(n);
  }

  lookupName(name: string, sheet: string) {
    return this.names.get(`${upper(sheet)}!${upper(name)}`) ?? this.names.get(`!${upper(name)}`);
  }

  getSheet(name: string) { return this.sheets.get(upper(name)); }

  getCell(sheet: string, addr: string): Cell | undefined {
    const a = parseA1(addr); if (!a) return undefined;
    return this.getSheet(sheet)?.cells.get(cellKey(a.row, a.col));
  }

  /** Change an input cell's value (call recalc() afterwards). */
  setValue(sheet: string, addr: string, v: Scalar) {
    const sh = this.getSheet(sheet)!; const a = parseA1(addr)!;
    let c = sh.cells.get(cellKey(a.row, a.col));
    if (!c) { c = { sheet: sh.name, row: a.row, col: a.col, cached: null, value: null }; sh.cells.set(cellKey(a.row, a.col), c); }
    if (c.node) throw new Error(`${sheet}!${addr} is a formula cell`);
    c.value = v;
  }

  // ---------- dependency graph ----------

  /** Every block of cells a formula can read, with names expanded. */
  precedentRanges(n: Node, sheet: string, out: { sheet: string; r1: number; c1: number; r2: number; c2: number }[] = [], seen = new Set<string>()) {
    switch (n.t) {
      case 'ref': out.push({ sheet: n.sheet ?? sheet, r1: n.r1, c1: n.c1, r2: n.r2, c2: n.c2 }); break;
      case 'name': {
        const def = this.lookupName(n.name, n.sheet ?? sheet);
        const k = `${n.sheet ?? sheet}!${n.name}`;
        if (def?.node && !seen.has(k)) { seen.add(k); this.precedentRanges(def.node, sheet, out, seen); }
        break;
      }
      case 'fn': n.args.forEach((a) => this.precedentRanges(a, sheet, out, seen)); break;
      case 'bin': this.precedentRanges(n.l, sheet, out, seen); this.precedentRanges(n.r, sheet, out, seen); break;
      case 'neg': case 'pos': case 'pct': this.precedentRanges(n.e, sheet, out, seen); break;
    }
    return out;
  }

  /** Formula cells (or array masters) inside a block. */
  private formulaCellsIn(r: { sheet: string; r1: number; c1: number; r2: number; c2: number }): Cell[] {
    const sh = this.getSheet(r.sheet); if (!sh) return [];
    const out = new Set<Cell>();
    const add = (c: Cell | undefined) => { if (!c) return; if (c.node) out.add(c); else if (c.arrayMaster) out.add(c.arrayMaster); };
    const r2 = Math.min(r.r2, sh.maxRow), c2 = Math.min(r.c2, sh.maxCol);
    const area = (r2 - r.r1 + 1) * (c2 - r.c1 + 1);
    if (area <= 0) return [];
    if (area <= 4 * sh.cells.size) {
      for (let row = r.r1; row <= r2; row++) for (let col = r.c1; col <= c2; col++) add(sh.cells.get(cellKey(row, col)));
    } else {
      for (const [k, c] of sh.cells) { const row = keyRow(k), col = keyCol(k); if (row >= r.r1 && row <= r2 && col >= r.c1 && col <= c2) add(c); }
    }
    return [...out];
  }

  deps = new Map<Cell, Cell[]>();

  private buildGraph() {
    const all: Cell[] = [];
    for (const name of this.sheetOrder) all.push(...this.getSheet(name)!.formulaCells);
    for (const c of all) {
      const ds = new Set<Cell>();
      if (c.node) for (const r of this.precedentRanges(c.node, c.sheet)) for (const d of this.formulaCellsIn(r)) ds.add(d);
      this.deps.set(c, [...ds]);
    }
    this.order = tarjan(all, (c) => this.deps.get(c) ?? []);
  }

  // ---------- evaluation ----------

  private readCell = (sheet: string, r: number, c: number): Scalar => {
    const cell = this.getSheet(sheet)?.cells.get(cellKey(r, c));
    return cell ? cell.value : null;
  };
  private bounds = (sheet: string) => { const s = this.getSheet(sheet)!; return { maxRow: s.maxRow, maxCol: s.maxCol }; };

  eval(n: Node, ctx: Ctx): Value {
    switch (n.t) {
      case 'num': return n.v;
      case 'str': return n.v;
      case 'bool': return n.v;
      case 'err': return new XLError(n.v);
      case 'missing': return null;
      case 'ref': {
        const sheet = n.sheet ?? ctx.sheet;
        const sh = this.getSheet(sheet);
        if (!sh) return new XLError('#REF!', `no sheet "${sheet}"`);
        if (n.r1 < 1 || n.c1 < 1 || n.r2 > MAX_ROW || n.c2 > MAX_COL) return new XLError('#REF!');
        return new RangeRef(sh.name, n.r1, n.c1, n.r2, n.c2, this.readCell, this.bounds);
      }
      case 'name': {
        const def = this.lookupName(n.name, n.sheet ?? ctx.sheet);
        if (!def?.node) return new XLError('#NAME?', `unknown name ${n.name}`);
        return this.eval(def.node, ctx);
      }
      case 'fn': {
        const f = FUNCTIONS[n.name];
        if (!f) return new XLError('#NAME?', `unsupported function ${n.name}`);
        return f(n.args, ctx, this);
      }
      case 'neg': return this.unary(n.e, ctx, (x) => -x);
      case 'pos': return this.eval(n.e, ctx);
      case 'pct': return this.unary(n.e, ctx, (x) => x / 100);
      case 'arr': return new Matrix(n.rows.map((r) => r.map((x) => this.eval(x, ctx) as Scalar)));
      case 'bin': return this.binary(n.op, this.operand(this.eval(n.l, ctx), ctx), this.operand(this.eval(n.r, ctx), ctx));
    }
  }

  scalar(n: Node, ctx: Ctx): Scalar {
    const v = this.eval(n, ctx);
    return this.intersect(v, ctx);
  }

  /** Excel's implicit intersection: a range used where one value is expected picks the cell in the formula's row/column. */
  intersect(v: Value, ctx: Ctx): Scalar {
    if (v instanceof Matrix) return v.get(0, 0);
    if (!(v instanceof RangeRef)) return v;
    if (v.isCell) return v.cell(v.r1, v.c1);
    if (v.width === 1 && ctx.row >= v.r1 && ctx.row <= v.r2) return v.cell(ctx.row, v.c1);
    if (v.height === 1 && ctx.col >= v.c1 && ctx.col <= v.c2) return v.cell(v.r1, ctx.col);
    return new XLError('#VALUE!', 'range does not intersect the formula cell');
  }

  private operand(v: Value, ctx: Ctx): Scalar | Matrix {
    if (v instanceof RangeRef) return v.isCell ? v.cell(v.r1, v.c1) : ctx.array ? v.toMatrix() : this.intersect(v, ctx);
    return v;
  }

  private unary(e: Node, ctx: Ctx, f: (x: number) => number): Scalar | Matrix {
    const v = this.operand(this.eval(e, ctx), ctx);
    const one = (s: Scalar): Scalar => { const x = toNumber(s); return isError(x) ? x : f(x); };
    return v instanceof Matrix ? new Matrix(v.rows.map((r) => r.map(one))) : one(v);
  }

  private binary(op: string, a: Scalar | Matrix, b: Scalar | Matrix): Scalar | Matrix {
    if (a instanceof Matrix || b instanceof Matrix) {
      const A = toMatrix(a), B = toMatrix(b);
      const h = Math.max(A.height, B.height), w = Math.max(A.width, B.width);
      const pick = (M: Matrix, r: number, c: number): Scalar => {
        const rr = M.height === 1 ? 0 : r, cc = M.width === 1 ? 0 : c;
        return rr < M.height && cc < M.width ? M.get(rr, cc) : new XLError('#N/A');
      };
      const rows: Scalar[][] = [];
      for (let r = 0; r < h; r++) { const row: Scalar[] = []; for (let c = 0; c < w; c++) row.push(scalarOp(op, pick(A, r, c), pick(B, r, c))); rows.push(row); }
      return new Matrix(rows);
    }
    return scalarOp(op, a, b);
  }

  private evalCell(c: Cell) {
    if (!c.node) { c.value = new XLError('#NAME?', c.parseError); return; }
    const ctx: Ctx = { sheet: c.sheet, row: c.row, col: c.col, array: !!c.array };
    const v = this.eval(c.node, ctx);
    if (c.array) {
      const m = toMatrix(v);
      const sh = this.getSheet(c.sheet)!;
      for (let r = c.array.r1; r <= c.array.r2; r++) for (let col = c.array.c1; col <= c.array.c2; col++) {
        const i = r - c.array.r1, j = col - c.array.c1;
        const x = m.height === 1 && m.width === 1 ? m.get(0, 0)
          : i < m.height && j < m.width ? m.get(i, j) : m.height === 1 && j < m.width ? m.get(0, j) : m.width === 1 && i < m.height ? m.get(i, 0) : new XLError('#N/A');
        const target = sh.cells.get(cellKey(r, col))!;
        target.value = normalise(x);
      }
    } else {
      c.value = normalise(v instanceof RangeRef ? this.intersect(v, ctx) : v instanceof Matrix ? v.get(0, 0) : v);
    }
  }

  recalc(opts: RecalcOptions = {}) {
    this.cycles = [];
    for (const comp of this.order) {
      const circular = comp.length > 1 || (this.deps.get(comp[0]) ?? []).includes(comp[0]);
      if (!circular) { this.evalCell(comp[0]); continue; }
      const cells = comp.map((c) => fullAddr(c.sheet, c.row, c.col));
      if (!this.calc.iterate) {
        for (const c of comp) c.value = new XLError('#CIRC!', 'circular reference with iterative calculation off');
        this.cycles.push({ cells, iterations: 0, converged: false, lastChange: NaN });
        continue;
      }
      if ((opts.cycleStart ?? 'zero') === 'zero') for (const c of comp) c.value = 0;
      else for (const c of comp) c.value = c.cached;
      let it = 0, change = Infinity;
      while (it < this.calc.maxIterations) {
        it++; change = 0;
        for (const c of comp) {
          const before = c.value;
          this.evalCell(c);
          const d = typeof before === 'number' && typeof c.value === 'number' ? Math.abs(c.value - before) : before === c.value ? 0 : Infinity;
          if (d > change) change = d;
        }
        if (change < this.calc.maxChange) break;
      }
      this.cycles.push({ cells, iterations: it, converged: change < this.calc.maxChange, lastChange: change });
    }
  }

  /** Circular blocks (strongly connected components that contain a loop). */
  circularBlocks(): Cell[][] {
    return this.order.filter((comp) => comp.length > 1 || (this.deps.get(comp[0]) ?? []).includes(comp[0]));
  }

  /**
   * How far each circular block is from being a true fixed point: plug values in, evaluate every
   * loop cell once, and measure the biggest change. A solved loop gives ~0. With source 'cached'
   * this tests the values stored in the file instead of ours (e.g. to audit another program's answer).
   */
  fixedPointResiduals(source: 'current' | 'cached'): { cells: string[]; residual: number; worstCell: string }[] {
    const all = [...this.allFormulaCells()];
    const saved = all.map((c) => c.value);
    if (source === 'cached') for (const c of all) c.value = c.cached;
    const out = [];
    for (const comp of this.circularBlocks()) {
      let residual = 0, worst = '';
      for (const c of comp) {
        const before = c.value;
        this.evalCell(c);
        const d = typeof before === 'number' && typeof c.value === 'number' ? Math.abs(c.value - before) : before === c.value ? 0 : Infinity;
        if (d >= residual) { residual = d; worst = fullAddr(c.sheet, c.row, c.col); }
        c.value = before;
      }
      out.push({ cells: comp.map((c) => fullAddr(c.sheet, c.row, c.col)), residual, worstCell: worst });
    }
    all.forEach((c, i) => (c.value = saved[i]));
    return out;
  }

  /** Formula cells that are in a circular block or read one (directly or indirectly). */
  cyclicOrDownstream(): Set<Cell> {
    const tainted = new Set<Cell>(this.circularBlocks().flat());
    for (const comp of this.order) for (const c of comp) if (!tainted.has(c) && (this.deps.get(c) ?? []).some((d) => tainted.has(d))) tainted.add(c);
    return tainted;
  }

  *allFormulaCells(): Generator<Cell> {
    for (const name of this.sheetOrder) {
      const sh = this.getSheet(name)!;
      for (const c of sh.cells.values()) if (c.node || c.arrayMaster || c.parseError) yield c;
    }
  }
}

function normalise(v: Scalar): Scalar {
  if (typeof v === 'number') {
    if (Number.isNaN(v) || !Number.isFinite(v)) return new XLError('#NUM!');
    return v === 0 ? 0 : v; // drop -0
  }
  return v;
}

export function scalarOp(op: string, a: Scalar, b: Scalar): Scalar {
  if (op === '&') {
    const x = toText(a), y = toText(b);
    if (isError(x)) return x; if (isError(y)) return y;
    return x + y;
  }
  if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
    if (isError(a)) return a; if (isError(b)) return b;
    const c = compare(a, b);
    switch (op) {
      case '=': return c === 0; case '<>': return c !== 0; case '<': return c < 0;
      case '>': return c > 0; case '<=': return c <= 0; default: return c >= 0;
    }
  }
  const x = toNumber(a); if (isError(x)) return x;
  const y = toNumber(b); if (isError(y)) return y;
  switch (op) {
    case '+': return x + y;
    case '-': return x - y;
    case '*': return x * y;
    case '/': return y === 0 ? new XLError('#DIV/0!') : x / y;
    case '^': {
      if (x === 0 && y === 0) return new XLError('#NUM!', '0^0');
      if (x === 0 && y < 0) return new XLError('#DIV/0!');
      const r = Math.pow(x, y);
      return Number.isFinite(r) ? r : new XLError('#NUM!');
    }
  }
  return new XLError('#VALUE!', `unknown operator ${op}`);
}

/** Iterative Tarjan SCC. Returns components with dependencies first. */
export function tarjan<T>(nodes: T[], succ: (n: T) => T[]): T[][] {
  const index = new Map<T, number>(), low = new Map<T, number>(), onStack = new Set<T>();
  const stack: T[] = []; const out: T[][] = []; let idx = 0;
  for (const root of nodes) {
    if (index.has(root)) continue;
    const work: { n: T; i: number; s: T[] }[] = [{ n: root, i: 0, s: succ(root) }];
    index.set(root, idx); low.set(root, idx); idx++; stack.push(root); onStack.add(root);
    while (work.length) {
      const top = work[work.length - 1];
      if (top.i < top.s.length) {
        const m = top.s[top.i++];
        if (!index.has(m)) {
          index.set(m, idx); low.set(m, idx); idx++; stack.push(m); onStack.add(m);
          work.push({ n: m, i: 0, s: succ(m) });
        } else if (onStack.has(m)) low.set(top.n, Math.min(low.get(top.n)!, index.get(m)!));
      } else {
        work.pop();
        if (work.length) { const p = work[work.length - 1].n; low.set(p, Math.min(low.get(p)!, low.get(top.n)!)); }
        if (low.get(top.n) === index.get(top.n)) {
          const comp: T[] = []; let m: T;
          do { m = stack.pop()!; onStack.delete(m); comp.push(m); } while (m !== top.n);
          out.push(comp.reverse());
        }
      }
    }
  }
  return out;
}

export { shiftNode };
export type { RefNode, ParseError };

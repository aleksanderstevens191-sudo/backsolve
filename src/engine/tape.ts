// A reverse-mode automatic differentiation "tape".
//
// In plain English: while we recalculate the workbook, every arithmetic step that depends on an
// input cell is written down as one line on a tape: "node 57 = node 12 × node 40", together with
// the local slopes (∂57/∂12 = value of node 40, ∂57/∂40 = value of node 12). To get the derivative
// of one output with respect to *every* input, we walk the tape once, backwards, passing each
// node's sensitivity ("adjoint") to its parents times the local slope. That is the chain rule,
// applied in the cheap direction: one backward pass costs about as much as one recalculation, no
// matter how many inputs there are.
//
// Circular references get special treatment (an "implicit block"): see `ImplicitBlock` below.

/** A number that depends on the inputs: its value and its node number on the tape. */
export class Dual {
  constructor(public readonly v: number, public readonly i: number) {}
}

/**
 * A circular block at its fixed point x* = F(x*, p).
 *
 * Plain English: the loop cells don't have a formula "x = something of the inputs" we can just
 * record; they are whatever values make the loop agree with itself. The implicit function theorem
 * says how that agreement point moves when an input moves:
 *   dx/dp = (I − J)⁻¹ · ∂F/∂p,   where J = ∂F/∂x (how each loop cell's formula reacts to the others).
 * In reverse mode that becomes: take the sensitivities ā arriving at the loop cells, solve
 * (I − Jᵀ) λ = ā, and send λ (instead of ā) into one evaluation of the loop formulas.
 */
export interface ImplicitBlock {
  /** Tape nodes standing for the loop cells' solved values (X_k = y_k with slope 1). */
  xs: number[];
  /** Highest X index: when the backward pass reaches it, all of the X adjoints are final. */
  trigger: number;
  /** The matrix I − Jᵀ (row-major, n×n). */
  m: Float64Array;
  n: number;
}

export class Tape {
  values: number[] = [];
  /** Node i's parents are par[off[i]] .. par[off[i+1]-1], with local slopes w[...]. */
  off: number[] = [0];
  par: number[] = [];
  w: number[] = [];
  blocks: ImplicitBlock[] = [];

  get size() { return this.values.length; }

  leaf(v: number): Dual {
    this.values.push(v);
    this.off.push(this.par.length);
    return new Dual(v, this.values.length - 1);
  }

  node(v: number, parents: number[], slopes: number[]): Dual {
    for (let k = 0; k < parents.length; k++) { this.par.push(parents[k]); this.w.push(slopes[k]); }
    this.values.push(v);
    this.off.push(this.par.length);
    return new Dual(v, this.values.length - 1);
  }

  /**
   * One backward pass. `seeds` gives the starting adjoint for some nodes (usually: the output = 1).
   * Returns the adjoint of every node; read the input leaves' entries to get the gradient.
   * Nodes below `lo` are not visited (used to differentiate a single loop block on its own).
   */
  backward(seeds: [number, number][], lo = 0, hi = this.size, useBlocks = true): Float64Array {
    const adj = new Float64Array(hi - lo);
    for (const [i, a] of seeds) adj[i - lo] += a;
    const { off, par, w, blocks } = this;
    let b = blocks.length - 1;
    while (b >= 0 && blocks[b].trigger >= hi) b--;
    for (let i = hi - 1; i >= lo; i--) {
      if (useBlocks) while (b >= 0 && blocks[b].trigger === i) { applyBlock(blocks[b], adj, lo); b--; }
      const a = adj[i - lo];
      if (a === 0) continue;
      for (let k = off[i], e = off[i + 1]; k < e; k++) {
        const p = par[k];
        if (p >= lo) adj[p - lo] += a * w[k];
      }
    }
    return adj;
  }
}

function applyBlock(bl: ImplicitBlock, adj: Float64Array, lo: number) {
  const a = bl.xs.map((x) => adj[x - lo]);
  if (a.every((x) => x === 0)) return;
  const lambda = solveLinear(bl.m, bl.n, a);
  bl.xs.forEach((x, k) => (adj[x - lo] = lambda[k]));
}

/** Solve M λ = a by Gaussian elimination with partial pivoting (M is small: one row per loop cell). */
export function solveLinear(m: Float64Array, n: number, a: number[]): number[] {
  const A = Float64Array.from(m), b = a.slice();
  for (let col = 0; col < n; col++) {
    let piv = col;
    for (let r = col + 1; r < n; r++) if (Math.abs(A[r * n + col]) > Math.abs(A[piv * n + col])) piv = r;
    if (A[piv * n + col] === 0) throw new Error('circular block: I − J is singular (the loop has no unique local solution)');
    if (piv !== col) {
      for (let c = 0; c < n; c++) { const t = A[col * n + c]; A[col * n + c] = A[piv * n + c]; A[piv * n + c] = t; }
      const t = b[col]; b[col] = b[piv]; b[piv] = t;
    }
    for (let r = col + 1; r < n; r++) {
      const f = A[r * n + col] / A[col * n + col];
      if (f === 0) continue;
      for (let c = col; c < n; c++) A[r * n + c] -= f * A[col * n + c];
      b[r] -= f * b[col];
    }
  }
  const x = new Array<number>(n).fill(0);
  for (let r = n - 1; r >= 0; r--) {
    let s = b[r];
    for (let c = r + 1; c < n; c++) s -= A[r * n + c] * x[c];
    x[r] = s / A[r * n + r];
  }
  return x;
}

// Stage 3 unit tests: the solver vs answers worked out by hand (algebra and geometry, not our
// engine), plus the real test workbooks.

import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { Workbook, WorkbookInput, loadWorkbook } from '../src/index';
import { Backsolver, goalSeek, tangencyOf, jacobiEigenvalues, newtonBracketed } from '../src/engine/solve';
import type { Scalar } from '../src/engine/values';
import { fullAddr } from '../src/engine/refs';

type Cells = Record<string, Scalar | `=${string}`>;
function book(cells: Cells, iterate = false) {
  const input: WorkbookInput = {
    sheets: [{ name: 'S', cells: new Map(Object.entries(cells).map(([a, v]) => [a, typeof v === 'string' && v.startsWith('=') ? { formula: v.slice(1), value: null } : { value: v as Scalar }])) }],
    names: [],
    calc: { iterate, maxIterations: 1000, maxChange: 1e-12 },
  };
  const wb = new Workbook(input);
  wb.recalc();
  return wb;
}
const rel = (a: number, b: number) => Math.abs(a - b) / Math.max(1, Math.abs(b));

describe('one free input (Goal Seek)', () => {
  it('Gordon growth: price = D/(r − g); the g for a target price is r − D/price', () => {
    const wb = book({ A1: 5, A2: 0.09, A3: 0.02, Z1: '=A1/(A2-A3)' });
    const { answer, result } = goalSeek(wb, 'S!Z1', 100, 'S!A3');
    expect(rel(answer, 0.09 - 5 / 100)).toBeLessThan(1e-14);
    expect(Math.abs(result.missBy)).toBeLessThanOrEqual(1e-12 * 100);
    expect(result.converged).toBe(true);
  });
  it('damped Newton across a pole: 10/(1 − x) = 1000 from x = 0 (the full Newton step lands far past the pole)', () => {
    const wb = book({ A1: 0, Z1: '=10/(1-A1)' });
    const { answer } = goalSeek(wb, 'S!Z1', 1000, 'S!A1');
    expect(rel(answer, 0.99)).toBeLessThan(1e-13);
  });
  it('through a circular reference: x = p + 0.5x, output 2x = 10 needs p = 2.5', () => {
    const wb = book({ A1: 1, A2: '=A1+0.5*A2', Z1: '=2*A2' }, true);
    expect(wb.cycles[0].converged).toBe(true);
    const { answer } = goalSeek(wb, 'S!Z1', 10, 'S!A1');
    expect(rel(answer, 2.5)).toBeLessThan(1e-13);
  });
  it('through IRR: equal flows p for 3 years on 100 invested; IRR = 10% needs p = 100 / annuity factor', () => {
    const wb = book({ A1: -100, A2: '=B1', A3: '=B1', A4: '=B1', B1: 40, Z1: '=IRR(A1:A4)' });
    const { answer } = goalSeek(wb, 'S!Z1', 0.1, 'S!B1');
    const annuity = (1 - 1.1 ** -3) / 0.1;
    expect(rel(answer, 100 / annuity)).toBeLessThan(1e-13);
  });
  it('the workbook is left exactly as it was', () => {
    const wb = book({ A1: 5, A2: 0.09, A3: 0.02, Z1: '=A1/(A2-A3)' });
    goalSeek(wb, 'S!Z1', 100, 'S!A3');
    expect(wb.getCell('S', 'A3')!.value).toBe(0.02);
    expect(wb.getCell('S', 'Z1')!.value).toBe(5 / (0.09 - 0.02));
  });
  it('flat start (zero slope at the base): probes outward until it finds the root', () => {
    // (x − 1)^2 has zero slope at x = 1; output 4 at x = 3 or x = −1, both 2 away.
    const r = newtonBracketed((x) => ({ f: (x - 1) ** 2 - 4, df: 2 * (x - 1) }), 1, 1e-12, 1);
    expect(r.converged).toBe(true);
    expect(Math.abs(Math.abs(r.x - 1) - 2)).toBeLessThan(1e-12);
  });
  it('implicit-function slope of the answer: g* = r − D/P, so dg*/dr = 1 and dg*/dD = −1/P', () => {
    const wb = book({ A1: 5, A2: 0.09, A3: 0.02, Z1: '=A1/(A2-A3)' });
    const s = new Backsolver(wb, 'S!Z1', [{ name: 'g', cells: 'S!A3' }]);
    const r = s.solve(100);
    expect(s.sensitivity1D(r, 'S!A2').dAnswer).toBeCloseTo(1, 13);
    expect(s.sensitivity1D(r, 'S!A1').dAnswer).toBeCloseTo(-1 / 100, 14);
  });
});

describe('several free inputs: closest point to the base case that hits the target', () => {
  it('linear model: matches the closed form d = σ²·g·(T − P0) / Σ σ²g²', () => {
    const wb = book({ A1: 1, A2: 1, A3: 1, Z1: '=2*A1+3*A2+5*A3+7' });
    const sig = [1, 2, 0.5], g = [2, 3, 5];
    const s = new Backsolver(wb, 'S!Z1', ['A1', 'A2', 'A3'].map((a, k) => ({ name: a, cells: `S!${a}`, sigma: sig[k] })));
    const r = s.solve(40);
    const denom = g.reduce((t, x, k) => t + sig[k] ** 2 * x * x, 0);
    g.forEach((x, k) => expect(rel(r.d[k], (sig[k] ** 2 * x * (40 - 17)) / denom)).toBeLessThan(1e-13));
    expect(Math.abs(r.missBy)).toBeLessThanOrEqual(40e-12);
    expect(r.tangency).toBeLessThan(1e-12);
    expect(r.reducedHessianMinEig).toBeCloseTo(1, 6); // a flat surface: curvature of the distance is exactly 1
  });
  it('weighting: a driver with twice the sigma moves 4 times as much (equal sensitivity)', () => {
    const wb = book({ A1: 0, A2: 0, Z1: '=A1+A2' });
    const s = new Backsolver(wb, 'S!Z1', [{ name: 'a', cells: 'S!A1', sigma: 2 }, { name: 'b', cells: 'S!A2', sigma: 1 }]);
    const r = s.solve(10);
    expect(r.d[0] / r.d[1]).toBeCloseTo(4, 12);
    // u = d/σ = (4, 2): driver a carries 16/(16 + 4) = 80% of the squared distance.
    expect(r.drivers[0].shareOfDistance).toBeCloseTo(0.8, 12);
  });
  it('circle: from (3, 4) to the circle of radius 10, the closest point is (6, 8)', () => {
    const wb = book({ A1: 3, A2: 4, Z1: '=A1^2+A2^2' });
    const s = new Backsolver(wb, 'S!Z1', [{ name: 'a', cells: 'S!A1', sigma: 1 }, { name: 'b', cells: 'S!A2', sigma: 1 }]);
    const r = s.solve(100);
    expect(rel(r.d[0], 3)).toBeLessThan(1e-12);
    expect(rel(r.d[1], 4)).toBeLessThan(1e-12);
    expect(r.distance).toBeCloseTo(5, 12);
    expect(r.reducedHessianMinEig!).toBeGreaterThan(0);
  });
  it('hyperbola: from (1, 1) to a·b = 4, the closest point is (2, 2); the curvature check sees a minimum', () => {
    const wb = book({ A1: 1, A2: 1, Z1: '=A1*A2' });
    const s = new Backsolver(wb, 'S!Z1', [{ name: 'a', cells: 'S!A1', sigma: 1 }, { name: 'b', cells: 'S!A2', sigma: 1 }]);
    const r = s.solve(4);
    expect(rel(r.d[0], 1)).toBeLessThan(1e-12);
    expect(rel(r.d[1], 1)).toBeLessThan(1e-12);
    // Along a·b = 4 at (2, 2), the squared distance to (1, 1) curves upward: a true minimum.
    expect(r.reducedHessianMinEig!).toBeGreaterThan(0.1);
  });
  it('scale drivers: SUM of three cells scaled by (1 + d) to hit 9 from 6 needs d = 0.5', () => {
    const wb = book({ A1: 1, A2: 2, A3: 3, Z1: '=SUM(A1:A3)' });
    const s = new Backsolver(wb, 'S!Z1', [{ name: 'all', cells: 'S!A1:A3', mode: 'scale' }]);
    const r = s.solve(9);
    expect(rel(r.d[0], 0.5)).toBeLessThan(1e-14);
    expect(r.drivers[0].cells.map((c) => c.implied)).toEqual([1.5, 3, 4.5]);
  });
  it('driver groups: a shift driver moves every cell by the same amount', () => {
    const wb = book({ A1: 0.1, A2: 0.2, B1: 0, Z1: '=A1+A2+B1' });
    const s = new Backsolver(wb, 'S!Z1', [{ name: 'g', cells: 'S!A1:A2' }, { name: 'b', cells: 'S!B1' }]);
    const r = s.solve(0.6);
    // Sensitivities: 2 for the pair (both cells move), 1 for B1; equal sigma (1 pp) → d ∝ (2, 1).
    expect(r.d[0] / r.d[1]).toBeCloseTo(2, 12);
    expect(rel(2 * r.d[0] + r.d[1], 0.3)).toBeLessThan(1e-13);
  });
  it('refuses formula cells, text cells and cells claimed by two drivers', () => {
    const wb = book({ A1: 1, A2: 'x', Z1: '=A1*2' });
    expect(() => new Backsolver(wb, 'S!Z1', [{ name: 'f', cells: 'S!Z1' }])).toThrow(/not an input cell/);
    expect(() => new Backsolver(wb, 'S!Z1', [{ name: 't', cells: 'S!A2' }])).toThrow(/not a number/);
    expect(() => new Backsolver(wb, 'S!Z1', [{ name: 'a', cells: 'S!A1' }, { name: 'b', cells: 'S!A1' }])).toThrow(/another driver/);
  });
});

describe('the curve of equal price', () => {
  it('a·b = 4: every point is on the curve and it ends exactly on the box', () => {
    const wb = book({ A1: 1, A2: 1, Z1: '=A1*A2' });
    const s = new Backsolver(wb, 'S!Z1', [{ name: 'a', cells: 'S!A1', sigma: 1 }, { name: 'b', cells: 'S!A2', sigma: 1 }]);
    const c = s.traceCurve(4, 0, 1, [[-0.5, 6], [-0.5, 6]]); // a, b ∈ [0.5, 7]
    expect(c.branches.length).toBe(1);
    expect(c.grid.covered).toBe(c.grid.crossings);
    for (const p of c.branches[0].points) expect(Math.abs(p.values[0] * p.values[1] - 4)).toBeLessThan(4e-12);
    const ends = [c.branches[0].points[0], c.branches[0].points.at(-1)!].map((p) => p.values).sort((x, y) => x[0] - y[0]);
    expect(ends[0][0]).toBeCloseTo(4 / 7, 9); expect(ends[0][1]).toBeCloseTo(7, 9);
    expect(ends[1][0]).toBeCloseTo(7, 9); expect(ends[1][1]).toBeCloseTo(4 / 7, 9);
  });
  it('a² − b² = 1 has two separate branches: both are found', () => {
    const wb = book({ A1: 2, A2: 0, Z1: '=A1^2-A2^2' });
    const s = new Backsolver(wb, 'S!Z1', [{ name: 'a', cells: 'S!A1', sigma: 1 }, { name: 'b', cells: 'S!A2', sigma: 1 }]);
    const c = s.traceCurve(1, 0, 1, [[-5, 1], [-3, 3]]); // a, b ∈ [−3, 3]
    expect(c.branches.length).toBe(2);
    expect(c.grid.covered).toBe(c.grid.crossings);
    const sides = c.branches.map((b) => Math.sign(b.points[0].values[0])).sort();
    expect(sides).toEqual([-1, 1]);
    for (const b of c.branches) for (const p of b.points) expect(Math.abs(p.values[0] ** 2 - p.values[1] ** 2 - 1)).toBeLessThan(1e-11);
  });
  it('a² + b² = 4 is a closed loop', () => {
    const wb = book({ A1: 0, A2: 0, Z1: '=A1^2+A2^2' });
    const s = new Backsolver(wb, 'S!Z1', [{ name: 'a', cells: 'S!A1', sigma: 1 }, { name: 'b', cells: 'S!A2', sigma: 1 }]);
    const c = s.traceCurve(4, 0, 1, [[-3, 3], [-3, 3]]);
    expect(c.branches.length).toBe(1);
    expect(c.branches[0].closed).toBe(true);
    expect(c.grid.covered).toBe(c.grid.crossings);
  });
  it('a corner (MAX kink) in the curve is walked through, not stopped at', () => {
    // MAX(a, b) = 2: an L-shaped curve with a corner at (2, 2).
    const wb = book({ A1: 0, A2: 0, Z1: '=MAX(A1,A2)' });
    const s = new Backsolver(wb, 'S!Z1', [{ name: 'a', cells: 'S!A1', sigma: 1 }, { name: 'b', cells: 'S!A2', sigma: 1 }]);
    const c = s.traceCurve(2, 0, 1, [[-1, 3], [-1, 3]]);
    expect(c.branches.length).toBe(1);
    expect(c.branches[0].ends).toEqual(['edge of the box', 'edge of the box']);
    expect(c.grid.covered).toBe(c.grid.crossings);
  });
});

describe('helpers', () => {
  it('tangency resolves tiny angles (projection, not sqrt(1 − cos²))', () => {
    expect(tangencyOf([1, 1e-12], [1, 0])).toBeCloseTo(1e-12, 24);
    expect(tangencyOf([0, 1], [1, 0])).toBe(1);
  });
  it('symmetric eigenvalues', () => {
    expect(jacobiEigenvalues(Float64Array.from([2, 1, 0, 1, 2, 0, 0, 0, 5]), 3).sort((a, b) => a - b).map((x) => +x.toFixed(12))).toEqual([1, 3, 5]);
  });
});

describe('real test workbooks', () => {
  const load = (f: string) => { const { wb } = loadWorkbook(new Uint8Array(readFileSync(join(__dirname, '..', 'workbooks', f)))); wb.recalc({ cycleStart: 'zero' }); return wb; };
  it('three-statement: closest growth + gross margin for $30 hits it exactly, at a true minimum, and restores the workbook', () => {
    const wb = load('three_statement.xlsx');
    const before = new Map([...wb.allFormulaCells()].map((c) => [fullAddr(c.sheet, c.row, c.col), c.value]));
    const s = new Backsolver(wb, 'Valuation!B13', [{ name: 'growth', cells: 'Assumptions!C5:G5' }, { name: 'gross margin', cells: 'Assumptions!C6:G6' }]);
    const r = s.solve(30);
    expect(r.converged).toBe(true);
    expect(Math.abs(r.missBy)).toBeLessThanOrEqual(30e-12);
    expect(r.tangency).toBeLessThanOrEqual(1e-9);
    expect(r.reducedHessianMinEig!).toBeGreaterThan(0);
    for (const c of wb.allFormulaCells()) {
      const a = before.get(fullAddr(c.sheet, c.row, c.col)), b = c.value;
      if (typeof a === 'number' && typeof b === 'number') expect(Math.abs(a - b)).toBeLessThanOrEqual(1e-12 * Math.max(1, Math.abs(a)));
      else expect(b).toEqual(a);
    }
  });
  it('LBO: the margin improvement for a 20% IRR (5 loops + IRR) matches plain bisection on the engine', () => {
    const wb = load('lbo_circular.xlsx');
    const { answer } = goalSeek(wb, 'Returns!B8', 0.2, 'Inputs!B10');
    // Reference without derivatives: bisection on the recalculated IRR.
    const cell = wb.getCell('Inputs', 'B10')!, out = wb.getCell('Returns', 'B8')!;
    const irr = (x: number) => { cell.value = x; wb.recalc({ tight: true, cycleStart: 'current' }); return (out.value as number) - 0.2; };
    let a = 0, b = 0.02;
    for (let k = 0; k < 200 && b - a > 1e-17; k++) { const m = (a + b) / 2; if (irr(m) < 0) a = m; else b = m; }
    cell.value = 0.005;
    expect(Math.abs(answer - (a + b) / 2)).toBeLessThan(1e-12);
  });
});

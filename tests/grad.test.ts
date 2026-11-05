// Stage 2 unit tests: derivatives from the tape vs derivatives worked out by hand (calculus), and
// vs finite differences on the real test workbooks.

import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { Workbook, WorkbookInput, Gradients, loadWorkbook } from '../src/index';
import type { Scalar } from '../src/engine/values';

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
const grad = (cells: Cells, out = 'S!Z1', iterate = false) => {
  const g = new Gradients(book(cells, iterate));
  expect(g.forwardMismatches).toEqual([]);
  return { g, d: g.gradientByAddr(out) };
};

describe('tape: hand-derived derivatives', () => {
  it('a*b + c', () => {
    const { d } = grad({ A1: 3, A2: 5, A3: 7, Z1: '=A1*A2+A3' });
    expect(d).toEqual({ 'S!A1': 5, 'S!A2': 3, 'S!A3': 1 });
  });
  it('division, power (both arguments), unary minus, percent', () => {
    const { d } = grad({ A1: 2, A2: 3, Z1: '=-(A1^A2)/A2+A1%' });
    // f = -a^b/b + a/100: ∂f/∂a = -b·a^(b-1)/b + 0.01 = -a^(b-1) + 0.01; ∂f/∂b = -(a^b ln a)/b + a^b/b²
    expect(d['S!A1']).toBeCloseTo(-(2 ** 2) + 0.01, 14);
    expect(d['S!A2']).toBeCloseTo(-(8 * Math.log(2)) / 3 + 8 / 9, 14);
  });
  it('negative base with an input-driven exponent: slope in the exponent is undefined, so warn', () => {
    // Excel's unary minus binds tighter than ^, so -A1^A2 is (-2)^3; ln(-2) does not exist.
    const { g } = grad({ A1: 2, A2: 3, Z1: '=-A1^A2' });
    expect(g.warnings.join()).toMatch(/x\^y with x ≤ 0/);
  });
  it('SUM, AVERAGE and SUMPRODUCT over ranges', () => {
    const { d } = grad({ A1: 1, A2: 2, B1: 3, B2: 4, Z1: '=SUM(A1:A2)+AVERAGE(A1:B2)+SUMPRODUCT(A1:A2,B1:B2)' });
    expect(d).toEqual({ 'S!A1': 1 + 0.25 + 3, 'S!A2': 1 + 0.25 + 4, 'S!B1': 0.25 + 1, 'S!B2': 0.25 + 2 });
  });
  it('NPV: slope in the rate and in each cash flow', () => {
    const { d } = grad({ A1: 0.1, B1: 100, B2: 200, Z1: '=NPV(A1,B1:B2)' });
    expect(d['S!B1']).toBeCloseTo(1 / 1.1, 15);
    expect(d['S!B2']).toBeCloseTo(1 / 1.21, 15);
    expect(d['S!A1']).toBeCloseTo(-100 / 1.1 ** 2 - (2 * 200) / 1.1 ** 3, 12);
  });
  it('IRR: implicit function theorem, ∂r/∂f_i = −(1+r)^−i / g′(r)', () => {
    const { g, d } = grad({ A1: -100, A2: 30, A3: 40, A4: 50, Z1: '=IRR(A1:A4)' });
    const r = g.gradient('S!Z1').value;
    const gp = -(30 / (1 + r) ** 2) - (2 * 40) / (1 + r) ** 3 - (3 * 50) / (1 + r) ** 4;
    [-100, 30, 40, 50].forEach((_, i) => expect(d[`S!A${i + 1}`]).toBeCloseTo(-1 / (1 + r) ** i / gp, 13));
  });
  it('IF follows the branch Excel took, and records the comparison as a kink', () => {
    const { g, d } = grad({ A1: 5, B1: 4, Z1: '=IF(A1>B1,A1*2,B1*3)' });
    expect(d).toEqual({ 'S!A1': 2, 'S!B1': 0 });
    const k = g.kinkReport();
    expect(k).toHaveLength(1);
    expect(k[0]).toMatchObject({ cell: 'S!Z1', kind: 'compare', margin: 1, onKink: false, movable: true });
    // a 1-unit move in A1 or B1 flips it: A1 = 5 needs 20%, B1 = 4 needs 25%
    expect(k[0].nearestInput).toMatchObject({ addr: 'S!A1', absChange: 1 });
  });
  it('MIN / MAX: slope goes to the argument picked; a tie sits on the kink', () => {
    const { g, d } = grad({ A1: 2, A2: 2, A3: 9, Z1: '=MIN(A1:A3)+MAX(A1,A3)' });
    expect(d).toEqual({ 'S!A1': 1, 'S!A2': 0, 'S!A3': 1 });
    const mins = g.kinkReport().filter((k) => k.kind === 'MIN');
    expect(mins[0]).toMatchObject({ onKink: true, near: true });
  });
  it('ABS and ROUND: ABS has slope ±1, ROUND is a staircase (slope 0)', () => {
    const { d } = grad({ A1: -3.3, Z1: '=ABS(A1)+ROUND(A1,0)' });
    expect(d).toEqual({ 'S!A1': -1 });
  });
  it('a kink that no input can move is reported as not movable', () => {
    const { g } = grad({ A1: 4, Z1: '=IF(ABS(A1-A1)<0.001,A1,0)' });
    expect(g.kinkReport().every((k) => !k.near)).toBe(true);
  });
});

describe('circular references: implicit function theorem at the fixed point', () => {
  it('one-cell loop x = p + 0.5x gives x = 2p, dx/dp = 2', () => {
    const { d } = grad({ A1: 3, Z1: '=A1+0.5*Z1' }, 'S!Z1', true);
    expect(d['S!A1']).toBeCloseTo(2, 13);
  });
  it('two-cell loop and a cell downstream of it', () => {
    // x = p + 0.25y, y = q + 0.5x  →  x = (p + 0.25q)/0.875,  z = 2x
    const { d } = grad({ A1: 1, A2: 2, X1: '=A1+0.25*Y1', Y1: '=A2+0.5*X1', Z1: '=2*X1' }, 'S!Z1', true);
    expect(d['S!A1']).toBeCloseTo(2 / 0.875, 13);
    expect(d['S!A2']).toBeCloseTo((2 * 0.25) / 0.875, 13);
  });
  it('a kink inside a loop: its distance accounts for the loop moving too', () => {
    // x = MAX(0, p − 0.5x). For p > 0, x = p/1.5 and the MAX margin is p − 0.5x = x, which
    // reaches 0 only when p does: the report must say "p would have to fall by all of p".
    const { g, d } = grad({ A1: 3, Z1: '=MAX(0,A1-0.5*Z1)' }, 'S!Z1', true);
    expect(d['S!A1']).toBeCloseTo(1 / 1.5, 13);
    const k = g.kinkReport()[0];
    expect(k.margin).toBeCloseTo(2, 12);
    expect(k.nearestInput!.absChange).toBeCloseTo(3, 12);
  });
  it('tight recalc solves loops to rounding level', () => {
    const wb = book({ A1: 3, X1: '=A1+0.9*Y1', Y1: '=0.5*X1' }, true);
    wb.recalc({ tight: true, cycleStart: 'current' });
    expect(wb.cycles[0].converged).toBe(true);
    expect(wb.getCell('S', 'X1')!.value as number).toBeCloseTo(3 / 0.55, 13);
  });
});

describe('test workbooks: tape matches the engine and finite differences', () => {
  const dir = join(__dirname, '..', 'workbooks');
  it.each([
    ['three_statement.xlsx', 'Valuation!B13', []],
    // Model!B4 (the "Year 0" counter) sits exactly on IF(year=5) kinks: a finite difference there jumps.
    ['lbo_circular.xlsx', 'Returns!B8', ['Model!B4']],
  ])('%s: d %s / d every input within 1e-6 of finite differences', (file, out, skip) => {
    const { wb } = loadWorkbook(new Uint8Array(readFileSync(join(dir, file))));
    wb.recalc();
    const g = new Gradients(wb);
    expect(g.forwardMismatches).toEqual([]);
    const cell = g.cellAt(out);
    const gr = g.gradient(cell);
    for (let k = 0; k < g.inputs.length; k++) {
      const x = g.inputs[k];
      if (skip.includes(x.addr)) continue;
      const h = 1e-4 * Math.max(1, Math.abs(x.value));
      const at = (v: number) => { x.cell.value = v; wb.recalc({ tight: true, cycleStart: 'current' }); return cell.value as number; };
      const D = (s: number) => (at(x.value + s) - at(x.value - s)) / (2 * s);
      const fd = (4 * D(h / 2) - D(h)) / 3;
      at(x.value);
      if (Math.abs(gr.d[k]) <= 1e-12 && Math.abs(fd) <= 1e-12) continue; // both zero up to rounding
      expect(Math.abs(gr.d[k] - fd) / Math.max(Math.abs(gr.d[k]), Math.abs(fd)), `${x.addr}: AD ${gr.d[k]} FD ${fd}`).toBeLessThanOrEqual(1e-6);
    }
  });
});

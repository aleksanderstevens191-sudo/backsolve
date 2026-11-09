// The Stage 2 gradient check as a reusable function: every gradient vs central finite differences,
// the kink report, and a benchmark. Used by scripts/stage2.ts (test workbooks) and
// scripts/stage4.ts (SEC workbooks). See scripts/stage2.ts for what is checked and why.

import os from 'node:os';
import { Gradients, Workbook } from '../src/index';
import type { Cell } from '../src/engine/workbook';
import { fullAddr } from '../src/engine/refs';

const REL_TOL = 1e-6;
const NEAR_PCT = 5;

export const machine = { os: `${os.type()} ${os.release()}`, cpu: os.cpus()[0]?.model, cores: os.cpus().length, node: process.version };

export function median(xs: number[]) { const s = [...xs].sort((a, b) => a - b); return s[Math.floor(s.length / 2)]; }
export function timeIt(f: () => void, minMs = 200, minRuns = 5): number {
  f(); // warm up
  const ts: number[] = [];
  const start = performance.now();
  while (ts.length < minRuns || performance.now() - start < minMs) { const t = performance.now(); f(); ts.push(performance.now() - t); }
  return median(ts);
}

function numericOutputs(wb: Workbook): Cell[] {
  return [...wb.allFormulaCells()].filter((c) => typeof c.value === 'number');
}

interface PairFail { output: string; input: string; ad: number; fd: number; rel: number }

export function checkWorkbook(name: string, wb: Workbook, headline: string, fullJacobian: boolean, quiet = false) {
  const g = new Gradients(wb);
  const outputs = fullJacobian ? numericOutputs(wb) : [g.cellAt(headline)];
  const N = g.inputs.length;

  // AD: one backward pass per output.
  const adRows = outputs.map((o) => g.gradient(o).d);

  // Kink margins' gradients (to recognise steps that straddle a kink).
  const kinkAdj = g.kinks.map((k) => { const adj = g.rec.tape.backward([[k.margin.i, 1]]); return { k, m: k.margin.v, dm: g.inputs.map((x) => adj[x.node]) }; });
  const upstream = outputs.map((o) => g.upstreamCells(o));
  // Largest |value| among the cells each output is computed from (for the FD rounding-noise bound).
  const valueOf = new Map([...wb.allFormulaCells()].map((c) => [fullAddr(c.sheet, c.row, c.col), c.value]));
  const magnitude = upstream.map((u) => { let m = 1; for (const a of u) { const v = valueOf.get(a); if (typeof v === 'number') m = Math.max(m, Math.abs(v)); } for (const x of g.inputs) m = Math.max(m, Math.abs(x.value)); return m; });

  // FD: perturb each input, re-solve, read every output.
  const base = outputs.map((o) => o.value as number);
  const fdCols: Float64Array[] = [];
  const steps: number[] = [];
  for (let j = 0; j < N; j++) {
    const x = g.inputs[j];
    const h = 1e-4 * Math.max(1, Math.abs(x.value));
    steps.push(h);
    const at = (v: number) => { x.cell.value = v; wb.recalc({ tight: true, cycleStart: 'current' }); return outputs.map((o) => (typeof o.value === 'number' ? o.value : NaN)); };
    const p1 = at(x.value + h), m1 = at(x.value - h), p2 = at(x.value + h / 2), m2 = at(x.value - h / 2);
    x.cell.value = x.value;
    const col = new Float64Array(outputs.length);
    for (let i = 0; i < outputs.length; i++) {
      const D1 = (p1[i] - m1[i]) / (2 * h), D2 = (p2[i] - m2[i]) / h;
      col[i] = (4 * D2 - D1) / 3;
    }
    fdCols.push(col);
  }
  wb.recalc({ tight: true, cycleStart: 'current' });
  // After all the perturbations, putting the inputs back must give the same outputs (loops to 1e-14).
  const restored = outputs.every((o, i) => typeof o.value === 'number' && Math.abs(o.value - base[i]) <= 1e-14 * Math.max(1, Math.abs(base[i])));

  let maxAdNoise = 0, pairs = 0, exactZero = 0, passRel = 0, nearZero = 0, kinkExcluded = 0, maxRel = 0;
  let worst: PairFail | undefined;
  const fails: PairFail[] = [];
  const noiseSamples: { output: string; input: string; ad: number; fd: number; noise: number }[] = [];
  const excluded: { output: string; input: string; ad: number; fd: number; kinks: string[] }[] = [];
  for (let i = 0; i < outputs.length; i++) {
    const out = fullAddr(outputs[i].sheet, outputs[i].row, outputs[i].col);
    // Every input that can reach this output (non-zero AD), plus any where FD sees movement.
    for (let j = 0; j < N; j++) {
      const ad = adRows[i][j], fd = fdCols[j][i];
      if (ad === 0 && fd === 0) { pairs++; exactZero++; continue; }
      pairs++;
      const rel = Math.abs(ad - fd) / Math.max(Math.abs(ad), Math.abs(fd));
      if (Number.isFinite(rel) && rel <= REL_TOL) {
        passRel++;
        if (rel > maxRel) { maxRel = rel; worst = { output: out, input: g.inputs[j].addr, ad, fd, rel }; }
        continue;
      }
      // A finite difference can't resolve slopes smaller than its own rounding noise: each output
      // is computed from intermediate values as big as M, carrying rounding errors of about ε·M,
      // and dividing by the step h magnifies them. If AD and FD differ by less than 10·ε·M/h, the
      // difference is noise (typical case: a balance check that is exactly 0 by construction).
      const noise = (10 * Number.EPSILON * magnitude[i]) / steps[j];
      if (Math.abs(ad - fd) <= noise) { nearZero++; maxAdNoise = Math.max(maxAdNoise, Math.abs(ad)); if (noiseSamples.length < 20) noiseSamples.push({ output: out, input: g.inputs[j].addr, ad, fd, noise }); continue; }
      const straddled = kinkAdj.filter(({ k, m, dm }) => upstream[i].has(k.cell) && Math.abs(dm[j]) * steps[j] >= Math.abs(m));
      if (straddled.length || !Number.isFinite(fd)) {
        kinkExcluded++;
        excluded.push({ output: out, input: g.inputs[j].addr, ad, fd: Number.isFinite(fd) ? fd : NaN, kinks: [...new Set(straddled.map(({ k }) => `${k.cell} ${k.kind}`))] });
        continue;
      }
      fails.push({ output: out, input: g.inputs[j].addr, ad, fd, rel });
    }
  }

  // Benchmark on the headline output.
  const head = headline ? g.cellAt(headline) : outputs[0];
  const recalcMs = timeIt(() => wb.recalc({ cycleStart: 'current' }));
  const recordMs = timeIt(() => new Gradients(wb, { resolve: false }));
  const node = g.nodeFor(head);
  const backwardMs = timeIt(() => { if (node && typeof node !== 'number') g.rec.tape.backward([[node.i, 1]]); });
  wb.recalc({ tight: true, cycleStart: 'current' });

  const kinks = g.kinkReport(NEAR_PCT, headline ? g.upstreamCells(headline) : undefined);
  const allKinks = g.kinkReport(NEAR_PCT);
  const result = {
    workbook: name,
    tape: { nodes: g.rec.tape.size, inputs: N, kinks: g.kinks.length, circularBlocks: g.loops.length, forwardMismatches: g.forwardMismatches, warnings: g.warnings,
      loops: g.loops.map((l) => ({ size: l.cells.length, fixedPointResidualScaled: l.residual, cells: l.cells })) },
    gradientCheck: {
      outputsChecked: outputs.length, inputs: N, pairs, tolerance: `|AD − FD| ≤ ${REL_TOL} × max(|AD|, |FD|)`,
      exactZero, passRelative: passRel, withinFdRoundingNoise: nearZero, maxAbsAdInNoiseCategory: maxAdNoise, noiseSamples, noiseRule: '|AD − FD| ≤ 10·ε·M/h, M = largest |value| among the cells the output is computed from', excludedStepCrossesKink: kinkExcluded, failed: fails.length,
      maxRelErrorAmongPassing: maxRel, worstPassing: worst, failures: fails.slice(0, 50), excluded: excluded.slice(0, 50), excludedTotal: excluded.length,
      fdMethod: 'central differences at h and h/2, Richardson-extrapolated; h = 1e-4 × max(1, |input|); loops re-solved to 1e-15 at each point',
      outputsRestoredExactly: restored,
    },
    headline: headline ? { output: headline, value: head.value, gradient: Object.fromEntries(Object.entries(g.gradientByAddr(head)).filter(([, v]) => v !== 0)) } : undefined,
    kinksUpstreamOfHeadline: headline ? { nearPct: NEAR_PCT, total: kinks.length, movable: kinks.filter((k) => k.movable).length, near: kinks.filter((k) => k.near) } : undefined,
    kinksWholeWorkbook: { nearPct: NEAR_PCT, total: allKinks.length, movable: allKinks.filter((k) => k.movable).length, near: allKinks.filter((k) => k.near) },
    benchmark: {
      inputs: N, recalcMs, recordTapeMs: recordMs, backwardMs,
      finiteDifferenceGradientMs: recalcMs * (N + 1),
      tapeGradientMs: recordMs + backwardMs,
      speedup: (recalcMs * (N + 1)) / (recordMs + backwardMs),
      note: 'medians; FD gradient = one baseline recalc + one recalc per input (forward differences, the cheapest FD); tape gradient = one recording pass + one backward pass',
    },
  };
  const log = quiet ? (..._: unknown[]) => {} : console.log;
  log(`\n${name}`);
  log(`  tape: ${result.tape.nodes} nodes, ${N} inputs, ${g.kinks.length} kinks, ${g.loops.length} circular blocks; forward mismatches vs engine: ${g.forwardMismatches.length}`);
  for (const l of result.tape.loops) log(`  loop of ${l.size} cells: fixed-point residual ${l.fixedPointResidualScaled.toExponential(1)}`);
  log(`  gradient check: ${outputs.length} outputs × ${N} inputs = ${pairs} pairs: ${exactZero} both exactly 0, ${passRel} within ${REL_TOL} relative (worst ${maxRel.toExponential(1)}), ${nearZero} within FD rounding noise, ${kinkExcluded} excluded (step crosses a kink), ${fails.length} FAILED`);
  for (const f of fails.slice(0, 10)) log(`    FAIL d${f.output}/d${f.input}: AD ${f.ad}, FD ${f.fd}, rel ${f.rel.toExponential(1)}`);
  for (const e of excluded.slice(0, 5)) log(`    excluded d${e.output}/d${e.input}: AD ${e.ad}, FD ${e.fd}; kinks: ${e.kinks.join(', ') || 'FD not finite'}`);
  if (headline) {
    log(`  kinks upstream of ${headline}: ${kinks.length} (${result.kinksUpstreamOfHeadline!.movable} movable by inputs), near (≤${NEAR_PCT}% uniform input move): ${result.kinksUpstreamOfHeadline!.near.length}`);
    for (const k of result.kinksUpstreamOfHeadline!.near.slice(0, 8)) log(`    ${k.cell} ${k.kind}: ${k.onKink ? 'ON the kink' : `${k.uniformMovePct.toFixed(2)}% uniform move`}; nearest input ${k.nearestInput?.addr} (${k.nearestInput?.value} → change ${k.nearestInput?.absChange.toPrecision(3)})  [${k.detail}]`);
  }
  log(`  kinks anywhere in the workbook: ${allKinks.length} (${result.kinksWholeWorkbook.movable} movable), near: ${result.kinksWholeWorkbook.near.length}`);
  for (const k of result.kinksWholeWorkbook.near.slice(0, 8)) log(`    ${k.cell} ${k.kind}: ${k.onKink ? 'ON the kink' : `${k.uniformMovePct.toFixed(2)}% uniform move`}; nearest input ${k.nearestInput?.addr} (${k.nearestInput?.value} → change ${k.nearestInput?.absChange.toPrecision(3)})  [${k.detail}]`);
  log(`  benchmark: recalc ${recalcMs.toFixed(3)} ms; FD gradient (${N + 1} recalcs) ${result.benchmark.finiteDifferenceGradientMs.toFixed(2)} ms vs tape ${result.benchmark.tapeGradientMs.toFixed(3)} ms (record ${recordMs.toFixed(3)} + backward ${backwardMs.toFixed(3)}) → ${result.benchmark.speedup.toFixed(1)}× faster`);
  const ok = fails.length === 0 && g.forwardMismatches.length === 0 && restored;
  if (!restored) log('  FAIL: outputs did not return to their base values after the finite-difference sweep');
  return { ok, result };
}


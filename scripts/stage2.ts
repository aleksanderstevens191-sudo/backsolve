// Stage 2 check: every gradient vs central finite differences, the kink report, and a benchmark.
// Writes results/stage2_gradients.json.
//
//   npx tsx scripts/stage2.ts
//
// What is checked, per workbook:
//  * The tape reproduces the engine's value for every formula cell, bit for bit.
//  * For EVERY numeric formula cell (not just the share price) and EVERY input cell it reads, the
//    tape's derivative matches a finite-difference derivative to 1e-6 relative.
//    Finite differences: central differences at steps h and h/2, combined by Richardson
//    extrapolation ((4·D(h/2) − D(h)) / 3), h = 1e-4 × max(1, |input|). Circular blocks are
//    re-solved to machine precision at each perturbed point.
//  * Pairs where the step crosses a kink (a branch flips inside ±h) are not comparable: a finite
//    difference across a jump measures the jump, not a slope. They are excluded and listed, never
//    silently dropped.

import { readFileSync, readdirSync, writeFileSync, mkdirSync } from 'node:fs';
import { basename, join } from 'node:path';
import { loadWorkbook } from '../src/index';
import { checkWorkbook, machine } from './gradcheck';
import { syntheticDcf } from './synthetic';

const root = new URL('..', import.meta.url).pathname;
/** The output each workbook is "about" (used for the kink report headline and the benchmark). */
const HEADLINE: Record<string, string> = {
  'three_statement.xlsx': 'Valuation!B13',
  'lbo_circular.xlsx': 'Returns!B8',
  'functions_coverage.xlsx': '',
};

const results: unknown[] = [];
let allOk = true;
const dir = join(root, 'workbooks');
for (const f of readdirSync(dir).filter((f) => f.endsWith('.xlsx'))) {
  const { wb } = loadWorkbook(new Uint8Array(readFileSync(join(dir, f))));
  wb.recalc({ cycleStart: 'zero' });
  const { ok, result } = checkWorkbook(basename(f), wb, HEADLINE[f] ?? '', true);
  allOk &&= ok; results.push(result);
}
for (const S of [10, 100]) {
  const { wb, output } = syntheticDcf(S);
  const { ok, result } = checkWorkbook(`synthetic DCF, ${S} segments × 10 years (in memory, no Excel file)`, wb, output, S <= 10);
  allOk &&= ok; results.push(result);
}

mkdirSync(join(root, 'results'), { recursive: true });
writeFileSync(join(root, 'results', 'stage2_gradients.json'), JSON.stringify({ generatedAt: new Date().toISOString(), machine, overall: allOk ? 'PASS' : 'FAIL', workbooks: results }, null, 2) + '\n');
console.log(`\nwrote results/stage2_gradients.json — overall ${allOk ? 'PASS' : 'FAIL'}`);
process.exit(allOk ? 0 : 1);

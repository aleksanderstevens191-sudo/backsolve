// Stage 1 check: recalculate every workbook from scratch and compare every formula cell with
// the value cached in the file. Writes results/stage1_forward.json.
//
//   npx tsx scripts/check.ts [extra.xlsx ...]

import { readFileSync, readdirSync, writeFileSync, mkdirSync } from 'node:fs';
import { basename, join } from 'node:path';
import os from 'node:os';
import { loadWorkbook, compareToCached, DEFAULT_TOL, XLError } from '../src/index';
import type { Scalar } from '../src/engine/values';
import { fullAddr } from '../src/engine/refs';

const root = new URL('..', import.meta.url).pathname;
const dir = join(root, 'workbooks');
const files = [...readdirSync(dir).filter((f) => f.endsWith('.xlsx')).map((f) => join(dir, f)), ...process.argv.slice(2)];

const known: Record<string, Record<string, { excel: unknown; rule: string }>> = JSON.parse(readFileSync(join(dir, 'known_differences.json'), 'utf8'));
const machine = { os: `${os.type()} ${os.release()}`, cpu: os.cpus()[0]?.model, cores: os.cpus().length, node: process.version };
const out: Record<string, unknown>[] = [];
let allOk = true;

for (const path of files) {
  const t0 = performance.now();
  const { wb, report } = loadWorkbook(new Uint8Array(readFileSync(path)));
  const t1 = performance.now();
  wb.recalc({ cycleStart: 'zero' });
  const t2 = performance.now();
  // LibreOffice-cached files only: cells where LibreOffice is known to differ from Excel are checked
  // against the value Excel's documented rule gives (listed in workbooks/known_differences.json).
  const isLibre = !!report.application?.startsWith('LibreOffice');
  const knownHere = isLibre ? known[basename(path)] ?? {} : {};
  const overrides = new Map(Object.entries(knownHere).map(([cell, e]) => [cell, e.excel as Scalar]));
  const cmp = compareToCached(wb, DEFAULT_TOL, overrides);
  const documented = Object.entries(knownHere).map(([cell, e]) => {
    const c = wb.getCell(cell.split('!')[0].replace(/^'|'$/g, ''), cell.split('!')[1])!;
    return { cell, ours: c.value, libreOfficeCached: c.cached instanceof XLError ? c.cached.code : c.cached, excelRule: e.excel, rule: e.rule };
  });
  // Circular blocks: is our answer a true fixed point? Is the file's cached answer one?
  const oursResid = wb.fixedPointResiduals('current');
  const cachedResid = wb.fixedPointResiduals('cached');
  const tainted = new Set([...wb.cyclicOrDownstream()].map((c) => fullAddr(c.sheet, c.row, c.col)));
  const mismatchesOutsideLoops = cmp.mismatches.filter((m) => !tainted.has(m.cell));
  // Dump every formula cell's value so the Python cross-check can compare engines cell by cell.
  const values: Record<string, unknown> = {};
  for (const c of wb.allFormulaCells()) { const v = c.value; values[fullAddr(c.sheet, c.row, c.col)] = v instanceof XLError ? { error: v.code } : v; }
  mkdirSync(join(root, 'results', 'engine_values'), { recursive: true });
  writeFileSync(join(root, 'results', 'engine_values', basename(path).replace(/\.xlsx$/, '.json')), JSON.stringify({ values, inOrDownstreamOfLoops: [...tainted] }, null, 1) + '\n');
  const unsupported = Object.fromEntries([...wb.unsupported].map(([k, v]) => [k, { count: v.length, cells: v.slice(0, 10) }]));
  const ok = cmp.mismatches.length === 0 && cmp.noCache === 0 && wb.unsupported.size === 0;
  allOk &&= ok;
  const r = {
    file: basename(path), cachedBy: report.application, ok,
    formulaCells: cmp.formulaCells, compared: cmp.compared, matched: cmp.matched, mismatched: cmp.mismatches.length, noCachedValue: cmp.noCache,
    tolerance: `|ours - cached| <= ${DEFAULT_TOL} * max(1, |cached|)`,
    maxAbsDiff: cmp.maxAbsDiff, maxScaledDiff: cmp.maxScaledDiff, worstCell: cmp.worstCell,
    circular: wb.cycles.map((c, i) => ({ size: c.cells.length, iterations: c.iterations, converged: c.converged, lastChange: c.lastChange,
      fixedPointResidualOurs: oursResid[i]?.residual, fixedPointResidualCached: cachedResid[i]?.residual, worstCachedCell: cachedResid[i]?.worstCell, cells: c.cells })),
    cellsInOrDownstreamOfLoops: tainted.size, mismatchesOutsideLoops: mismatchesOutsideLoops.length,
    calcSettings: wb.calc, unsupported, dataTables: report.dataTables,
    documentedLibreOfficeDifferences: documented,
    mismatches: cmp.mismatches.slice(0, 50), ms: { parse: +(t1 - t0).toFixed(1), recalc: +(t2 - t1).toFixed(1) },
  };
  out.push(r);
  console.log(`\n${r.file}  (cached values written by: ${r.cachedBy ?? 'unknown'})`);
  console.log(`  formula cells ${r.formulaCells}, compared ${r.compared}, matched ${r.matched}, mismatched ${r.mismatched}, no cached value ${r.noCachedValue}`);
  console.log(`  max |diff| ${r.maxAbsDiff.toExponential(2)}, max scaled diff ${r.maxScaledDiff.toExponential(2)} at ${r.worstCell ?? '-'}`);
  for (const c of r.circular) console.log(`  circular block of ${c.size} cells: ${c.iterations} iterations from zero, converged=${c.converged}; fixed-point residual ours ${c.fixedPointResidualOurs?.toExponential(2)}, file's cached values ${c.fixedPointResidualCached?.toExponential(2)} (${c.worstCachedCell})`);
  if (r.circular.length) console.log(`  ${tainted.size} formula cells are in or downstream of a loop; mismatches outside them: ${mismatchesOutsideLoops.length}`);
  if (wb.unsupported.size) for (const [k, v] of wb.unsupported) console.log(`  UNSUPPORTED ${k}: ${v.length} cells, e.g. ${v.slice(0, 3).join(', ')}`);
  for (const m of cmp.mismatches.slice(0, 10)) console.log(`  MISMATCH ${m.cell} =${m.formula}: ours ${m.ours}, cached ${m.cached}`);
  for (const d of documented) console.log(`  checked against Excel's rule, not LibreOffice's cache: ${d.cell} ours ${d.ours}, Excel rule ${d.excelRule}, LibreOffice cached ${d.libreOfficeCached}`);
  console.log(`  ${ok ? 'PASS' : 'FAIL'}  (parse ${r.ms.parse} ms, recalc ${r.ms.recalc} ms)`);
}

mkdirSync(join(root, 'results'), { recursive: true });
writeFileSync(join(root, 'results', 'stage1_forward.json'), JSON.stringify({ generatedAt: new Date().toISOString(), machine, workbooks: out }, null, 2) + '\n');
console.log(`\nwrote results/stage1_forward.json — overall ${allOk ? 'PASS' : 'FAIL'}`);
process.exit(allOk ? 0 : 1);

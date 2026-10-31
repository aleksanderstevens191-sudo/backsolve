// Unit tests for Excel's rules. Expected values come from Excel's documented behaviour
// (comments say which rule), not from our own engine.

import { describe, expect, it } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { Workbook, WorkbookInput, parseFormula, printNode, readXlsx, XLError, loadWorkbook } from '../src/index';
import { shiftNode } from '../src/engine/parser';
import { tarjan } from '../src/engine/workbook';
import type { Scalar } from '../src/engine/values';

type Cells = Record<string, Scalar | `=${string}`>;
function book(sheets: Record<string, Cells>, opts: Partial<WorkbookInput> = {}) {
  const input: WorkbookInput = {
    sheets: Object.entries(sheets).map(([name, cells]) => ({
      name,
      cells: new Map(Object.entries(cells).map(([a, v]) => [a, typeof v === 'string' && v.startsWith('=') ? { formula: v.slice(1), value: null } : { value: v as Scalar }])),
    })),
    names: opts.names ?? [],
    calc: opts.calc ?? { iterate: false, maxIterations: 100, maxChange: 0.001 },
  };
  const wb = new Workbook(input);
  wb.recalc();
  return wb;
}
const val = (wb: Workbook, ref: string) => { const [s, a] = ref.split('!'); return wb.getCell(s, a)!.value; };
const one = (f: string, cells: Cells = {}) => val(book({ S: { ...cells, Z99: `=${f}` } }), 'S!Z99');

describe('parser', () => {
  it('unary minus binds tighter than ^ (=-2^2 is 4)', () => expect(one('-2^2')).toBe(4));
  it('^ is left-associative (=2^3^2 is 64)', () => expect(one('2^3^2')).toBe(64));
  it('usual arithmetic precedence', () => expect(one('1+2*3-4/2^2')).toBe(6));
  it('percent', () => expect(one('50*10%')).toBe(5));
  it('string escapes', () => expect(one('"say ""hi"""')).toBe('say "hi"'));
  it('quoted sheet names, absolute refs and ranges', () => {
    expect(printNode(parseFormula("SUM('My Sheet'!$A$1:B$3)+Other!C4"))).toBe("(SUM('My Sheet'!$A$1:B$3)+Other!C4)");
  });
  it('whole-column and row ranges', () => {
    expect(printNode(parseFormula('SUM(A:B)+SUM($2:$3)'))).toBe('(SUM(A:B)+SUM($2:$3))');
  });
  it('_xlfn. prefixes are stripped from function names', () => {
    expect((parseFormula('_xlfn.STDEV.S(A1:A3)') as { name: string }).name).toBe('STDEV.S');
  });
  it('a word that looks like a cell past column XFD is a name', () => {
    expect(parseFormula('XFE1').t).toBe('name'); // column XFE is 16385, one past the last column
    expect(parseFormula('XFD1').t).toBe('ref');
    expect(parseFormula('LOG10(100)').t).toBe('fn');
  });
  it('shared-formula shift moves relative refs only', () => {
    const n = shiftNode(parseFormula('A1+$B$1+C$1+$D1'), 2, 1);
    expect(printNode(n)).toBe('(((B3+$B$1)+D$1)+$D3)');
  });
  it('rejects unsupported syntax loudly', () => {
    expect(() => parseFormula('Table1[Col]')).toThrow(/Structured/);
    expect(() => parseFormula("'S1:S3'!A1")).toThrow(/3-D/);
  });
});

describe('functions follow Excel rules', () => {
  const data: Cells = { A1: 1, A2: 2, A3: 'text', A4: true, A6: 5.5 };
  it('SUM ignores text and TRUE inside ranges', () => expect(one('SUM(A1:A6)', data)).toBe(8.5));
  it('SUM counts TRUE and number-text typed directly', () => expect(one('SUM(1,"2",TRUE)')).toBe(4));
  it('SUM errors on non-number text typed directly', () => expect((one('SUM(1,"x")') as XLError).code).toBe('#VALUE!'));
  it('SUM propagates errors inside ranges', () => expect((one('SUM(A1:A2)', { A1: 1, A2: '=1/0' }) as XLError).code).toBe('#DIV/0!'));
  it('AVERAGE of only text is #DIV/0!', () => expect((one('AVERAGE(A3)', data) as XLError).code).toBe('#DIV/0!'));
  it('MIN/MAX of no numbers is 0', () => { expect(one('MIN(A3:A5)', data)).toBe(0); expect(one('MAX(A3:A5)', data)).toBe(0); });
  it('SUMPRODUCT treats text as 0 and needs equal sizes', () => {
    expect(one('SUMPRODUCT(A1:A3,A1:A3)', data)).toBe(5);
    expect((one('SUMPRODUCT(A1:A3,A1:A2)', data) as XLError).code).toBe('#VALUE!');
  });
  it('SUMPRODUCT evaluates conditions as arrays', () => expect(one('SUMPRODUCT((A1:A2>1)*A1:A2)', data)).toBe(2));
  it('IF: missing else gives FALSE, empty arg gives 0', () => {
    expect(one('IF(FALSE,1)')).toBe(false);
    expect(one('IF(TRUE,,1)')).toBe(0);
  });
  it('IF only evaluates the branch it takes', () => expect(one('IF(TRUE,1,1/0)')).toBe(1));
  it('AND/OR ignore text in ranges; no logicals at all is #VALUE!', () => {
    expect(one('AND(A1:A4)', data)).toBe(true);
    expect(one('OR(A3)', data) instanceof XLError).toBe(true);
  });
  it('NPV discounts the first value one period', () => expect(one('NPV(0.1,110)')).toBeCloseTo(100, 12));
  it('NPV skips text/blank in ranges (they are not periods)', () => expect(one('NPV(0.1,A1:A6)', data)).toBeCloseTo(1 / 1.1 + 2 / 1.21 + 5.5 / 1.331, 14));
  it('XNPV uses actual/365 from the first date', () => {
    const v = one('XNPV(0.1,B1:B2,C1:C2)', { B1: -100, B2: 110, C1: 45292, C2: 45292 + 365 }) as number;
    expect(v).toBeCloseTo(0, 12);
  });
  it('IRR solves NPV = 0', () => {
    const r = one('IRR(B1:B4)', { B1: -100, B2: 30, B3: 40, B4: 50 }) as number;
    expect(-100 + 30 / (1 + r) + 40 / (1 + r) ** 2 + 50 / (1 + r) ** 3).toBeCloseTo(0, 10);
    expect(r).toBeCloseTo(0.0889633947, 9);
  });
  it('IRR with all positive flows is #NUM!', () => expect((one('IRR({1,2,3})') as XLError).code).toBe('#NUM!'));
  it('text comparison is case-insensitive; numbers < text < logicals', () => {
    expect(one('"abc"="ABC"')).toBe(true);
    expect(one('"a">99999')).toBe(true);
    expect(one('TRUE>"z"')).toBe(true);
  });
  it('blank cells are 0 in arithmetic and "" in text', () => { expect(one('A5+1', data)).toBe(1); expect(one('A5&"x"', data)).toBe('x'); });
  it('number-to-text uses General format', () => expect(one('1/3&""')).toBe('0.333333333333333'));
  it('division by zero', () => expect((one('1/0') as XLError).code).toBe('#DIV/0!'));
  it('unknown functions are #NAME? and reported', () => {
    const wb = book({ S: { A1: '=VLOOKUP(1,B1:C2,2,FALSE)' } });
    expect((val(wb, 'S!A1') as XLError).code).toBe('#NAME?');
    expect(wb.unsupported.get('VLOOKUP')).toEqual(['S!A1']);
  });
});

describe('references and names', () => {
  it('cross-sheet refs and quoted names', () => {
    const wb = book({ 'My Sheet': { A1: 2 }, S: { A1: "='My Sheet'!A1*3" } });
    expect(val(wb, 'S!A1')).toBe(6);
  });
  it('workbook and sheet-scoped names (sheet scope wins)', () => {
    const wb = book({ S: { A1: 1, A2: 2, B1: '=Rate' }, T: { B1: '=Rate' } }, {
      names: [{ name: 'Rate', formula: 'S!$A$1' }, { name: 'Rate', formula: 'S!$A$2', scope: 'T' }],
    });
    expect(val(wb, 'S!B1')).toBe(1);
    expect(val(wb, 'T!B1')).toBe(2);
  });
  it('a name can be a formula or a range', () => {
    const wb = book({ S: { A1: 1, A2: 2, B1: '=SUM(Block)*Twice' } }, { names: [{ name: 'Block', formula: 'S!$A$1:$A$2' }, { name: 'Twice', formula: '1+1' }] });
    expect(val(wb, 'S!B1')).toBe(6);
  });
  it('implicit intersection picks the cell on the same row', () => {
    const wb = book({ S: { A1: 10, A2: 20, B2: '=A1:A3*2' } });
    expect(val(wb, 'S!B2')).toBe(40);
  });
  it('whole-column SUM', () => expect(one('SUM(B:B)', { B1: 1, B500: 2 })).toBe(3));
  it('missing sheet is #REF!', () => expect((one('Nope!A1') as XLError).code).toBe('#REF!'));
});

describe('dependency graph and circular references', () => {
  it('Tarjan returns dependencies first and finds loops', () => {
    const g: Record<string, string[]> = { a: ['b'], b: ['c'], c: ['b'], d: [] };
    const comps = tarjan(Object.keys(g), (n) => g[n]).map((c) => c.sort().join(''));
    expect(comps).toEqual(['bc', 'a', 'd']);
  });
  it('evaluates in dependency order regardless of sheet layout', () => {
    const wb = book({ S: { A1: '=A2+1', A2: '=A3*2', A3: 5 } });
    expect(val(wb, 'S!A1')).toBe(11);
  });
  it('circular reference with iteration off is flagged, not silently computed', () => {
    const wb = book({ S: { A1: '=A2+1', A2: '=A1/2' } });
    expect((val(wb, 'S!A1') as XLError).code).toBe('#CIRC!');
    expect(wb.cycles).toHaveLength(1);
  });
  it('iterative calculation finds the fixed point (interest on average balance)', () => {
    // debt_end = 100 - (50 - interest);  interest = 0.1 * (100 + debt_end) / 2  =>  interest = 15/0.95
    const wb = book({ S: { A1: 100, A2: '=A1-(50-A3)', A3: '=0.1*AVERAGE(A1,A2)' } }, { calc: { iterate: true, maxIterations: 100, maxChange: 1e-14 } });
    expect(val(wb, 'S!A3')).toBeCloseTo(7.5 / 0.95, 13);
    expect(wb.cycles[0].converged).toBe(true);
  });
  it('stops at Max Iterations and reports non-convergence', () => {
    const wb = book({ S: { A1: '=A1*2+1' } }, { calc: { iterate: true, maxIterations: 5, maxChange: 0.001 } });
    expect(wb.cycles[0]).toMatchObject({ iterations: 5, converged: false });
  });
});

describe('xlsx reader', () => {
  const xml = (s: string) => strToU8(`<?xml version="1.0" encoding="UTF-8"?>${s}`);
  const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const file = (sheetXml: string, extraWb = '') => zipSync({
    'xl/workbook.xml': xml(`<workbook ${NS}><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>${extraWb}</workbook>`),
    'xl/_rels/workbook.xml.rels': xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    'xl/sharedStrings.xml': xml(`<sst ${NS}><si><t>hello</t></si><si><r><t>rich </t></r><r><t>text</t></r></si></sst>`),
    'xl/worksheets/sheet1.xml': xml(`<worksheet ${NS}><sheetData>${sheetXml}</sheetData></worksheet>`),
  });

  it('expands shared formulas (how Excel stores formulas filled down)', () => {
    const bytes = file(`
      <row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*10</f><v>10</v></c></row>
      <row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>20</v></c></row>
      <row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>30</v></c></row>`);
    const { wb } = loadWorkbook(bytes);
    wb.recalc();
    expect([1, 2, 3].map((r) => wb.getCell('Data', `B${r}`)!.value)).toEqual([10, 20, 30]);
  });
  it('reads shared strings, rich text, booleans, errors and calc settings', () => {
    const bytes = file(`<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="b"><v>1</v></c><c r="D1" t="e"><f>1/0</f><v>#DIV/0!</v></c></row>`,
      '<calcPr iterate="1" iterateCount="50" iterateDelta="1E-9"/>');
    const { input } = readXlsx(bytes);
    const cells = input.sheets[0].cells;
    expect(cells.get('A1')!.value).toBe('hello');
    expect(cells.get('B1')!.value).toBe('rich text');
    expect(cells.get('C1')!.value).toBe(true);
    expect((cells.get('D1')!.value as XLError).code).toBe('#DIV/0!');
    expect(input.calc).toEqual({ iterate: true, maxIterations: 50, maxChange: 1e-9 });
  });
  it('array formulas fill their block', () => {
    const bytes = file(`<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="array" ref="B1:B2">A1:A2*2</f><v>2</v></c></row><row r="2"><c r="A2"><v>4</v></c><c r="B2"><v>8</v></c></row>`);
    const { wb } = loadWorkbook(bytes);
    wb.recalc();
    expect(wb.getCell('Data', 'B2')!.value).toBe(8);
  });
});

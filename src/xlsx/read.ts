// Read an .xlsx file (a zip of XML files) into the engine's WorkbookInput.
// Works in the browser and in Node: it only needs the file's bytes, nothing is uploaded anywhere.

import { unzipSync, strFromU8 } from 'fflate';
import { XMLParser } from 'fast-xml-parser';
import { parseFormula, shiftNode, Node } from '../engine/parser';
import { parseA1 } from '../engine/refs';
import { Scalar, XLError, ErrorCode } from '../engine/values';
import type { CellInput, NameInput, SheetInput, WorkbookInput } from '../engine/workbook';

const ARRAYS = new Set(['row', 'c', 'sheet', 'definedName', 'si', 'r', 'Relationship']);
const parser = new XMLParser({
  ignoreAttributes: false, attributeNamePrefix: '@', textNodeName: '#text',
  parseTagValue: false, parseAttributeValue: false, trimValues: false,
  isArray: (name, _p, _leaf, isAttr) => !isAttr && ARRAYS.has(name),
});

/* eslint-disable @typescript-eslint/no-explicit-any */
const text = (x: any): string => (x === undefined || x === null ? '' : typeof x === 'object' ? (x['#text'] ?? '') : String(x));

export interface ReadReport {
  /** Formula cells whose file has no cached value (e.g. written by openpyxl, never recalculated). */
  missingCache: string[];
  dataTables: string[];
  application?: string;
}

export function readXlsx(bytes: Uint8Array): { input: WorkbookInput; report: ReadReport } {
  const files = unzipSync(bytes);
  const get = (p: string) => { const f = files[p.replace(/^\//, '')]; return f ? strFromU8(f) : undefined; };
  const report: ReadReport = { missingCache: [], dataTables: [] };

  const app = get('docProps/app.xml');
  if (app) { const m = /<Application>([^<]*)<\/Application>/.exec(app); const v = /<AppVersion>([^<]*)<\/AppVersion>/.exec(app); report.application = m ? `${m[1]}${v ? ' ' + v[1] : ''}` : undefined; }

  // Shared strings
  const sst: string[] = [];
  const sstXml = get('xl/sharedStrings.xml');
  if (sstXml) {
    const doc = parser.parse(sstXml);
    for (const si of doc.sst?.si ?? []) {
      if (si.t !== undefined) sst.push(text(si.t));
      else sst.push((si.r ?? []).map((r: any) => text(r.t)).join(''));
    }
  }

  const wbDoc = parser.parse(get('xl/workbook.xml')!);
  const wb = wbDoc.workbook;
  const rels = parser.parse(get('xl/_rels/workbook.xml.rels')!).Relationships.Relationship as any[];
  const target = (id: string) => {
    const t = rels.find((r) => r['@Id'] === id)['@Target'] as string;
    return t.startsWith('/') ? t.slice(1) : `xl/${t}`;
  };

  const sheetsMeta = (wb.sheets.sheet as any[]).map((s) => ({ name: s['@name'] as string, path: target(s['@r:id']) }));
  const calcPr = wb.calcPr ?? {};
  const truthy = (v: any) => v === '1' || v === 'true';
  const calc = {
    iterate: truthy(calcPr['@iterate']),
    maxIterations: calcPr['@iterateCount'] !== undefined ? Number(calcPr['@iterateCount']) : 100,
    maxChange: calcPr['@iterateDelta'] !== undefined ? Number(calcPr['@iterateDelta']) : 0.001,
  };

  const names: NameInput[] = [];
  for (const dn of wb.definedNames?.definedName ?? []) {
    const name = dn['@name'] as string;
    if (name.startsWith('_xlnm.')) continue; // print areas, filters: not part of the calculation
    const lsid = dn['@localSheetId'];
    names.push({ name, formula: text(dn), scope: lsid !== undefined ? sheetsMeta[Number(lsid)].name : undefined });
  }

  const sheets: SheetInput[] = [];
  for (const meta of sheetsMeta) {
    const xml = get(meta.path);
    const cells = new Map<string, CellInput>();
    sheets.push({ name: meta.name, cells });
    if (!xml) continue;
    const doc = parser.parse(xml);
    const shared = new Map<string, { node: Node; row: number; col: number }>();
    const pendingShared: { addr: string; si: string; cell: CellInput }[] = [];
    for (const row of doc.worksheet?.sheetData?.row ?? []) {
      for (const c of row.c ?? []) {
        const addr = c['@r'] as string;
        const t = c['@t'] as string | undefined;
        const vRaw = c.v !== undefined ? text(c.v) : undefined;
        let value: Scalar = null;
        if (t === 's') value = vRaw !== undefined ? sst[Number(vRaw)] : null;
        else if (t === 'b') value = vRaw === '1' || vRaw === 'TRUE';
        else if (t === 'e') value = new XLError((vRaw ?? '#VALUE!') as ErrorCode);
        else if (t === 'str') value = vRaw ?? '';
        else if (t === 'inlineStr') value = text(c.is?.t) || (c.is?.r ?? []).map((r: any) => text(r.t)).join('');
        else if (vRaw !== undefined && vRaw !== '') value = Number(vRaw);

        const f = c.f;
        if (f === undefined) { if (value !== null) cells.set(addr, { value }); continue; }
        const ftype = typeof f === 'object' ? f['@t'] : undefined;
        const ftext = text(f);
        const cell: CellInput = { value };
        if (vRaw === undefined && t !== 'inlineStr') report.missingCache.push(`${meta.name}!${addr}`);
        if (ftype === 'shared') {
          const si = f['@si'] as string;
          if (ftext.trim() !== '') {
            cell.formula = ftext;
            const a = parseA1(addr)!;
            try { shared.set(si, { node: parseFormula(ftext), row: a.row, col: a.col }); } catch { /* reported when the engine parses it */ }
          } else pendingShared.push({ addr, si, cell });
        } else if (ftype === 'dataTable') {
          report.dataTables.push(`${meta.name}!${addr}`);
          cell.formula = `TABLE()`; // What-If data tables are not supported: flagged as an unsupported function
        } else {
          cell.formula = ftext;
          if (ftype === 'array' && f['@ref']) cell.arrayRef = f['@ref'];
        }
        cells.set(addr, cell);
      }
    }
    for (const p of pendingShared) {
      const m = shared.get(p.si);
      const a = parseA1(p.addr)!;
      if (m) { p.cell.node = shiftNode(m.node, a.row - m.row, a.col - m.col); p.cell.formula = `(shared formula ${p.si})`; }
      else p.cell.formula = `#SHARED_FORMULA_${p.si}_MISSING`;
    }
  }
  return { input: { sheets, names, calc }, report };
}

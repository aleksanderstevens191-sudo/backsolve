export { Workbook } from './engine/workbook';
export type { Cell, CycleReport, WorkbookInput, CalcSettings } from './engine/workbook';
export { readXlsx } from './xlsx/read';
export { compareToCached, DEFAULT_TOL } from './engine/compare';
export { parseFormula, printNode } from './engine/parser';
export { SUPPORTED_FUNCTIONS } from './engine/functions';
export { XLError, Matrix } from './engine/values';

import { readXlsx } from './xlsx/read';
import { Workbook } from './engine/workbook';

/** Convenience: bytes of an .xlsx -> a Workbook ready to recalc(). */
export function loadWorkbook(bytes: Uint8Array) {
  const { input, report } = readXlsx(bytes);
  return { wb: new Workbook(input), report };
}

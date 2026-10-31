// A1-style address helpers. Rows and columns are 1-based like Excel.

export const MAX_ROW = 1048576;
export const MAX_COL = 16384;

export function colToNum(col: string): number {
  let n = 0;
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function numToCol(n: number): string {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

export function parseA1(addr: string): { row: number; col: number; absRow: boolean; absCol: boolean } | null {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(addr);
  if (!m) return null;
  const col = colToNum(m[2]); const row = Number(m[4]);
  if (col < 1 || col > MAX_COL || row < 1 || row > MAX_ROW) return null;
  return { row, col, absCol: m[1] === '$', absRow: m[3] === '$' };
}

export const toA1 = (row: number, col: number) => `${numToCol(col)}${row}`;

/** Numeric key for a cell within one sheet. */
export const cellKey = (row: number, col: number) => row * (MAX_COL + 1) + col;
export const keyRow = (k: number) => Math.floor(k / (MAX_COL + 1));
export const keyCol = (k: number) => k % (MAX_COL + 1);

/** Quote a sheet name the way Excel does in formulas when needed. */
export function quoteSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !parseA1(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

export const fullAddr = (sheet: string, row: number, col: number) => `${quoteSheet(sheet)}!${toA1(row, col)}`;

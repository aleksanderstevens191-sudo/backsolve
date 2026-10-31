// Value types that flow through the engine, mirroring Excel's cell value kinds.

export type ErrorCode =
  | '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A'
  | '#CIRC!'; // not a real Excel error: marks a circular reference with iteration off

export class XLError {
  constructor(public readonly code: ErrorCode, public readonly detail?: string) {}
  toString() { return this.code; }
}

/** A blank cell. Excel treats it as 0 in arithmetic and "" in text contexts. */
export const BLANK = null;
export type Blank = typeof BLANK;

export type Scalar = number | string | boolean | XLError | Blank;

/** A 2-D block of values: the result of evaluating a range or an array expression. */
export class Matrix {
  constructor(public readonly rows: Scalar[][]) {}
  get height() { return this.rows.length; }
  get width() { return this.rows[0]?.length ?? 0; }
  get(r: number, c: number): Scalar { return this.rows[r][c]; }
  *values(): Generator<Scalar> { for (const row of this.rows) for (const v of row) yield v; }
}

/**
 * A reference to a rectangular block of cells. Functions like SUM treat values
 * that come from references differently from values typed directly into the
 * formula (e.g. SUM ignores text in a range but errors on a direct text arg),
 * so we keep the reference around rather than flattening it immediately.
 */
export class RangeRef {
  constructor(
    public readonly sheet: string,
    public readonly r1: number, public readonly c1: number,
    public readonly r2: number, public readonly c2: number,
    private readonly read: (sheet: string, r: number, c: number) => Scalar,
    private readonly bounds: (sheet: string) => { maxRow: number; maxCol: number },
  ) {}
  get isCell() { return this.r1 === this.r2 && this.c1 === this.c2; }
  get height() { return this.r2 - this.r1 + 1; }
  get width() { return this.c2 - this.c1 + 1; }
  cell(r: number, c: number): Scalar { return this.read(this.sheet, r, c); }
  /** Materialise into a Matrix, clipping whole-column/row refs to the used area. */
  toMatrix(): Matrix {
    const b = this.bounds(this.sheet);
    const r2 = Math.min(this.r2, Math.max(b.maxRow, this.r1));
    const c2 = Math.min(this.c2, Math.max(b.maxCol, this.c1));
    const rows: Scalar[][] = [];
    for (let r = this.r1; r <= r2; r++) {
      const row: Scalar[] = [];
      for (let c = this.c1; c <= c2; c++) row.push(this.read(this.sheet, r, c));
      rows.push(row);
    }
    return new Matrix(rows);
  }
}

export type Value = Scalar | Matrix | RangeRef;

export const isError = (v: unknown): v is XLError => v instanceof XLError;

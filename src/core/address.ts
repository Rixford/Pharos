/**
 * A1-style address parsing, formatting and range arithmetic.
 */
import { CellRef, RangeRef } from './types';

export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;

export function colToLetter(col: number): string {
  if (col < 1 || !Number.isInteger(col)) throw new Error(`Column index must be a positive integer, got ${col}`);
  let s = '';
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function letterToCol(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    const v = ch.charCodeAt(0) - 64;
    if (v < 1 || v > 26) throw new Error(`Invalid column letters: "${letters}"`);
    n = n * 26 + v;
  }
  if (n < 1) throw new Error(`Invalid column letters: "${letters}"`);
  return n;
}

/** Quote a sheet name for use in an A1 reference, only when required. */
export function quoteSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;
const COL_RE = /^\$?([A-Za-z]{1,3})$/;
const ROW_RE = /^\$?(\d{1,7})$/;

/** Parse "B4" / "$B$4" (no sheet). Returns null if not a plain cell. */
export function parseLocalCell(text: string): { row: number; col: number } | null {
  const m = CELL_RE.exec(text);
  if (!m) return null;
  const col = letterToCol(m[1]);
  const row = parseInt(m[2], 10);
  if (col > MAX_COLS || row > MAX_ROWS || row < 1) return null;
  return { row, col };
}

/**
 * Split "Sheet!rest" into sheet and local part, handling quoted sheet names
 * ('My Sheet'!A1, with '' as an escaped quote). Sheet is undefined when absent.
 */
export function splitSheet(address: string): [string | undefined, string] {
  if (address.startsWith("'")) {
    let i = 1;
    let name = '';
    while (i < address.length) {
      if (address[i] === "'") {
        if (address[i + 1] === "'") {
          name += "'";
          i += 2;
          continue;
        }
        break;
      }
      name += address[i++];
    }
    if (address[i] !== "'" || address[i + 1] !== '!') {
      throw new Error(`Malformed quoted sheet reference: "${address}"`);
    }
    return [name, address.slice(i + 2)];
  }
  const bang = address.indexOf('!');
  if (bang === -1) return [undefined, address];
  return [address.slice(0, bang), address.slice(bang + 1)];
}

export function parseCellAddress(address: string, defaultSheet?: string): CellRef {
  const [sheet, rest] = splitSheet(address.trim());
  const cell = parseLocalCell(rest);
  if (!cell) throw new Error(`Invalid cell address: "${address}"`);
  const s = sheet ?? defaultSheet;
  if (!s) {
    throw new Error(`Address "${address}" has no sheet; qualify it like Sheet1!${rest.toUpperCase()}`);
  }
  return { sheet: s, row: cell.row, col: cell.col };
}

export function formatCell(ref: CellRef, includeSheet = true): string {
  const local = `${colToLetter(ref.col)}${ref.row}`;
  return includeSheet ? `${quoteSheet(ref.sheet)}!${local}` : local;
}

/** Parse "Sheet1!A1:B10", "A1:B10", "A1", "A:B" or "2:5" into a RangeRef. */
export function parseRange(text: string, defaultSheet?: string): RangeRef {
  const [sheet, rest] = splitSheet(text.trim());
  const s = sheet ?? defaultSheet;
  if (!s) throw new Error(`Range "${text}" has no sheet; qualify it like Sheet1!${rest}`);
  const parts = rest.split(':');
  if (parts.length === 1) {
    const c = parseLocalCell(parts[0]);
    if (!c) throw new Error(`Invalid range: "${text}"`);
    return { sheet: s, startRow: c.row, startCol: c.col, endRow: c.row, endCol: c.col };
  }
  if (parts.length !== 2) throw new Error(`Invalid range: "${text}"`);
  const [a, b] = parts;
  const ca = parseLocalCell(a);
  const cb = parseLocalCell(b);
  if (ca && cb) {
    return normalizeRange({
      sheet: s,
      startRow: ca.row,
      startCol: ca.col,
      endRow: cb.row,
      endCol: cb.col
    });
  }
  const colA = COL_RE.exec(a);
  const colB = COL_RE.exec(b);
  if (colA && colB) {
    return normalizeRange({
      sheet: s,
      startRow: 1,
      startCol: letterToCol(colA[1]),
      endRow: MAX_ROWS,
      endCol: letterToCol(colB[1]),
      open: 'columns'
    });
  }
  const rowA = ROW_RE.exec(a);
  const rowB = ROW_RE.exec(b);
  if (rowA && rowB) {
    return normalizeRange({
      sheet: s,
      startRow: parseInt(rowA[1], 10),
      startCol: 1,
      endRow: parseInt(rowB[1], 10),
      endCol: MAX_COLS,
      open: 'rows'
    });
  }
  throw new Error(`Invalid range: "${text}"`);
}

export function normalizeRange(r: RangeRef): RangeRef {
  const out = { ...r };
  if (out.startRow > out.endRow) [out.startRow, out.endRow] = [out.endRow, out.startRow];
  if (out.startCol > out.endCol) [out.startCol, out.endCol] = [out.endCol, out.startCol];
  return out;
}

export function formatRange(r: RangeRef, includeSheet = true): string {
  const a = `${colToLetter(r.startCol)}${r.startRow}`;
  const b = `${colToLetter(r.endCol)}${r.endRow}`;
  const local = a === b ? a : `${a}:${b}`;
  return includeSheet ? `${quoteSheet(r.sheet)}!${local}` : local;
}

export function sameSheet(a: string, b: string): boolean {
  return a.toLowerCase() === b.toLowerCase();
}

export function rangeContains(r: RangeRef, ref: CellRef): boolean {
  return (
    sameSheet(r.sheet, ref.sheet) &&
    ref.row >= r.startRow &&
    ref.row <= r.endRow &&
    ref.col >= r.startCol &&
    ref.col <= r.endCol
  );
}

export function rangeArea(r: RangeRef): number {
  return (r.endRow - r.startRow + 1) * (r.endCol - r.startCol + 1);
}

export function rangesOverlap(a: RangeRef, b: RangeRef): boolean {
  return (
    sameSheet(a.sheet, b.sheet) &&
    a.startRow <= b.endRow &&
    b.startRow <= a.endRow &&
    a.startCol <= b.endCol &&
    b.startCol <= a.endCol
  );
}

/** Clamp open-ended (whole row/column) ranges to a sheet's used extent. */
export function clampRange(r: RangeRef, maxRow: number, maxCol: number): RangeRef {
  return {
    ...r,
    endRow: Math.min(r.endRow, Math.max(maxRow, r.startRow)),
    endCol: Math.min(r.endCol, Math.max(maxCol, r.startCol))
  };
}

/** Iterate cells of a range, capped to avoid pathological expansion. */
export function* iterateRange(r: RangeRef, cap = Number.POSITIVE_INFINITY): Generator<CellRef> {
  let count = 0;
  for (let row = r.startRow; row <= r.endRow; row++) {
    for (let col = r.startCol; col <= r.endCol; col++) {
      if (count++ >= cap) return;
      yield { sheet: r.sheet, row, col };
    }
  }
}

/** Key used in per-sheet cell maps. */
export function cellKey(row: number, col: number): string {
  return `${row},${col}`;
}

/**
 * Helpers for building the benchmark workbooks with ExcelJS, plus a
 * placement manifest so the gold report can cite exact source ranges
 * without ever reading the generated files.
 */
import * as ExcelJS from 'exceljs';

export interface Placement {
  book: string;
  sheet: string;
  label: string;
  range: string;
  headerRow?: number;
  dataStart?: number;
  dataEnd?: number;
}

export class Manifest {
  readonly placements: Placement[] = [];
  constructor(readonly book: string) {}
  add(p: Omit<Placement, 'book'>): void {
    this.placements.push({ book: this.book, ...p });
  }
  find(sheet: string, label: string): Placement {
    const hit = this.placements.find((p) => p.sheet === sheet && p.label === label);
    if (!hit) throw new Error(`manifest: ${sheet}/${label} not found`);
    return hit;
  }
}

export const colLetter = (n: number): string => {
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
};

export const a1 = (row: number, col: number): string => `${colLetter(col)}${row}`;
export const rangeA1 = (r1: number, c1: number, r2: number, c2: number): string =>
  `${a1(r1, c1)}:${a1(r2, c2)}`;

/** Merged bold title spanning `span` columns at (row, col). */
export function title(ws: ExcelJS.Worksheet, row: number, col: number, span: number, text: string): void {
  ws.mergeCells(row, col, row, col + span - 1);
  const cell = ws.getCell(row, col);
  cell.value = text;
  cell.font = { bold: true, size: 13 };
}

export function bold(ws: ExcelJS.Worksheet, row: number, col: number, value: ExcelJS.CellValue): void {
  const cell = ws.getCell(row, col);
  cell.value = value;
  cell.font = { bold: true };
}

/** Single-row header. */
export function headerRow(ws: ExcelJS.Worksheet, row: number, col: number, headers: string[]): void {
  headers.forEach((h, i) => bold(ws, row, col + i, h));
}

/**
 * Two-row grouped header: merged group labels on top, sub-headers below.
 * Returns the row index of the sub-header row.
 */
export function groupedHeader(
  ws: ExcelJS.Worksheet,
  row: number,
  col: number,
  groups: { label: string; span: number }[],
  subs: string[]
): number {
  let c = col;
  for (const g of groups) {
    if (g.span > 1) ws.mergeCells(row, c, row, c + g.span - 1);
    const cell = ws.getCell(row, c);
    cell.value = g.label;
    cell.font = { bold: true, italic: true };
    c += g.span;
  }
  headerRow(ws, row + 1, col, subs);
  return row + 1;
}

/** Key/value block; returns next free row. */
export function kvBlock(
  ws: ExcelJS.Worksheet,
  row: number,
  col: number,
  pairs: [string, ExcelJS.CellValue][]
): number {
  pairs.forEach(([k, v], i) => {
    ws.getCell(row + i, col).value = k;
    ws.getCell(row + i, col + 1).value = v;
  });
  return row + pairs.length;
}

/** A small italic notes block (one string per row); returns next free row. */
export function notes(ws: ExcelJS.Worksheet, row: number, col: number, lines: string[]): number {
  lines.forEach((line, i) => {
    const cell = ws.getCell(row + i, col);
    cell.value = line;
    cell.font = { italic: true };
  });
  return row + lines.length;
}

export function f(formula: string, result: number | string): ExcelJS.CellValue {
  return { formula, result } as ExcelJS.CellValue;
}

export const r2 = (n: number): number => Math.round(n * 100) / 100;

// pharos:eof

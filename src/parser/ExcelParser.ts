/**
 * ExcelJS-backed implementation of WorkbookParser.
 *
 * Why ExcelJS and not SheetJS (`xlsx`)? The npm build of `xlsx` is frozen at
 * 0.18.5 with open CVEs and ships no cell styles in the community edition,
 * which Pharos needs for header/region heuristics. ExcelJS is maintained and
 * exposes formulas (incl. shared formulas), styles, merges, hidden sheets
 * and defined names. The WorkbookParser interface keeps this swappable.
 */
import * as ExcelJS from 'exceljs';
import { cellKey, parseLocalCell, parseRange } from '../core/address';
import { CellScalar, CellStyle, CellValueType, RangeRef } from '../core/types';
import { offsetFormula } from './FormulaParser';
import { ParsedCell, ParsedSheet, ParsedWorkbook, WorkbookParser } from './types';

interface Classified {
  value: CellScalar;
  type: CellValueType;
}

function classify(res: unknown): Classified {
  if (res === null || res === undefined) return { value: null, type: 'empty' };
  if (res instanceof Date) return { value: res, type: 'date' };
  if (typeof res === 'number') return { value: res, type: 'number' };
  if (typeof res === 'boolean') return { value: res, type: 'boolean' };
  if (typeof res === 'string') return { value: res, type: res === '' ? 'empty' : 'string' };
  if (typeof res === 'object') {
    const o = res as Record<string, unknown>;
    if ('error' in o) return { value: String(o.error), type: 'error' };
    if ('richText' in o && Array.isArray(o.richText)) {
      const text = (o.richText as { text?: string }[]).map((r) => r.text ?? '').join('');
      return { value: text, type: text === '' ? 'empty' : 'string' };
    }
    if ('text' in o) return classify(o.text);
  }
  return { value: String(res), type: 'string' };
}

function extractStyle(cell: ExcelJS.Cell): CellStyle | undefined {
  const style: CellStyle = {};
  const font = cell.font;
  if (font?.bold) style.bold = true;
  if (font?.italic) style.italic = true;
  if (cell.numFmt) style.numFmt = cell.numFmt;
  const fill = cell.fill as { fgColor?: { argb?: string } } | undefined;
  if (fill?.fgColor?.argb && fill.fgColor.argb !== 'FFFFFFFF') style.fillColor = fill.fgColor.argb;
  return Object.keys(style).length > 0 ? style : undefined;
}

function convertCell(cell: ExcelJS.Cell, row: number, col: number): ParsedCell | null {
  // Merged slave cells mirror their master's value in ExcelJS; skip them —
  // merge geometry is carried separately in ParsedSheet.merges.
  if (cell.type === ExcelJS.ValueType.Merge) return null;
  const master = (cell as { master?: { address?: string } }).master;
  if (master?.address && master.address !== cell.address) return null;

  const v = cell.value as unknown;
  let formula: string | undefined;
  let sharedFrom: string | undefined;
  let hyperlink: string | undefined;
  let classified: Classified;

  if (v !== null && v !== undefined && typeof v === 'object' && !(v instanceof Date)) {
    const o = v as Record<string, unknown>;
    if ('formula' in o || 'sharedFormula' in o) {
      if (typeof o.formula === 'string' && o.formula.length > 0) formula = o.formula;
      else if (typeof o.sharedFormula === 'string') sharedFrom = o.sharedFormula;
      classified = classify(o.result);
    } else if ('hyperlink' in o) {
      hyperlink = typeof o.hyperlink === 'string' ? o.hyperlink : undefined;
      classified = classify('text' in o ? o.text : null);
    } else {
      classified = classify(v);
    }
  } else {
    classified = classify(v);
  }

  const style = extractStyle(cell);
  if (classified.type === 'empty' && !formula && !sharedFrom) return null;

  const parsed: ParsedCell = { row, col, value: classified.value, type: classified.type };
  if (formula) parsed.formula = formula;
  if (sharedFrom) parsed.sharedFrom = sharedFrom;
  if (style) parsed.style = style;
  if (hyperlink) parsed.hyperlink = hyperlink;
  return parsed;
}

export class ExcelParser implements WorkbookParser {
  async parse(input: Buffer): Promise<ParsedWorkbook> {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(input as unknown as ExcelJS.Buffer);

    const warnings: string[] = [];
    const sheets: ParsedSheet[] = [];

    wb.worksheets.forEach((ws, index) => {
      const cells = new Map<string, ParsedCell>();
      const merges: RangeRef[] = [];
      const model = (ws as unknown as { model?: { merges?: string[] } }).model;
      for (const m of model?.merges ?? []) {
        try {
          merges.push(parseRange(m, ws.name));
        } catch {
          warnings.push(`Sheet "${ws.name}": could not parse merge range "${m}"`);
        }
      }

      let maxRow = 0;
      let maxCol = 0;
      ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          const parsed = convertCell(cell, rowNumber, colNumber);
          if (parsed) {
            cells.set(cellKey(rowNumber, colNumber), parsed);
            if (rowNumber > maxRow) maxRow = rowNumber;
            if (colNumber > maxCol) maxCol = colNumber;
          }
        });
      });

      // Translate shared formulas onto member cells by shifting the master.
      for (const cell of cells.values()) {
        if (cell.sharedFrom && !cell.formula) {
          const masterLoc = parseLocalCell(cell.sharedFrom);
          const master = masterLoc ? cells.get(cellKey(masterLoc.row, masterLoc.col)) : undefined;
          if (master?.formula && masterLoc) {
            cell.formula = offsetFormula(
              master.formula,
              ws.name,
              cell.row - masterLoc.row,
              cell.col - masterLoc.col
            );
          } else {
            warnings.push(
              `Sheet "${ws.name}": shared formula master ${cell.sharedFrom} not found for row ${cell.row}, col ${cell.col}`
            );
          }
        }
      }

      sheets.push({
        name: ws.name,
        index,
        hidden: ws.state !== 'visible',
        maxRow,
        maxCol,
        cells,
        merges
      });
    });

    let definedNames: { name: string; ranges: string[] }[] = [];
    try {
      const dn = (wb as unknown as { definedNames?: { model?: { name: string; ranges?: string[] }[] } })
        .definedNames;
      definedNames = (dn?.model ?? []).map((m) => ({ name: m.name, ranges: m.ranges ?? [] }));
    } catch (err) {
      warnings.push(`Failed to read defined names: ${err instanceof Error ? err.message : String(err)}`);
    }

    return { sheets, definedNames, warnings };
  }
}

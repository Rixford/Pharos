/**
 * FormulaParser — extracts cell/range/name references from Excel formulas.
 *
 * This is a character-level scanner, not a single regex pass: string
 * literals, quoted sheet names (with '' escapes), external workbook
 * prefixes, whole-row/column ranges and structured references are all
 * handled, so e.g. "A1" inside a string literal is never mistaken for a
 * reference and LOG10(...) is recognised as a function, not a cell.
 *
 * Limitations (documented in docs/DESIGN.md): R1C1 notation is not
 * resolved, structured references (Table1[Col]) are recorded but not
 * expanded to cells, and 3-D references (Sheet1:Sheet3!A1) traverse only
 * the first sheet (a warning is emitted).
 */
import { colToLetter, letterToCol, MAX_COLS, MAX_ROWS, normalizeRange } from '../core/address';
import { FormulaRef, RangeRef } from '../core/types';

export interface EndpointInfo {
  row?: number;
  col?: number;
  absRow?: boolean;
  absCol?: boolean;
}

export interface RefSpan {
  /** Start index (inclusive) of the reference text within the formula. */
  start: number;
  /** End index (exclusive). */
  end: number;
  ref: FormulaRef;
  /** Verbatim sheet prefix (including quotes and '!') or '' when local. */
  sheetPrefix: string;
  a?: EndpointInfo;
  b?: EndpointInfo;
  open?: 'columns' | 'rows';
}

export interface ExtractResult {
  refs: FormulaRef[];
  spans: RefSpan[];
  warnings: string[];
}

const CELL_TOKEN = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/;
const COL_TOKEN = /^(\$?)([A-Za-z]{1,3})$/;
const ROW_TOKEN = /^(\$?)(\d{1,7})$/;

const isIdentStart = (ch: string): boolean => /[A-Za-z_$\\]/.test(ch);
const isIdentChar = (ch: string): boolean => /[A-Za-z0-9_.$\\]/.test(ch);
const isDigit = (ch: string): boolean => ch >= '0' && ch <= '9';

function endpointFromCell(m: RegExpExecArray): EndpointInfo | null {
  const col = letterToCol(m[2]);
  const row = parseInt(m[4], 10);
  if (col > MAX_COLS || row < 1 || row > MAX_ROWS) return null;
  return { row, col, absRow: m[3] === '$', absCol: m[1] === '$' };
}

/**
 * Extract all references from a formula. `currentSheet` is used to resolve
 * unqualified references (e.g. `B4` inside Sheet1 means `Sheet1!B4`).
 */
export function extractRefs(formula: string, currentSheet: string): ExtractResult {
  const f = formula;
  const n = f.length;
  const spans: RefSpan[] = [];
  const warnings: string[] = [];
  let i = 0;

  const nextNonSpace = (pos: number): string => {
    while (pos < n && /\s/.test(f[pos])) pos++;
    return pos < n ? f[pos] : '';
  };

  const readIdent = (pos: number): number => {
    while (pos < n && isIdentChar(f[pos])) pos++;
    return pos;
  };

  const makeSpan = (
    start: number,
    end: number,
    sheetPrefix: string,
    sheet: string | undefined,
    external: string | undefined,
    a: EndpointInfo,
    b?: EndpointInfo,
    open?: 'columns' | 'rows'
  ): RefSpan => {
    const s = sheet ?? currentSheet;
    let range: RangeRef;
    if (open === 'columns') {
      range = normalizeRange({
        sheet: s,
        startRow: 1,
        endRow: MAX_ROWS,
        startCol: a.col!,
        endCol: b!.col!,
        open,
        external
      });
    } else if (open === 'rows') {
      range = normalizeRange({
        sheet: s,
        startRow: a.row!,
        endRow: b!.row!,
        startCol: 1,
        endCol: MAX_COLS,
        open,
        external
      });
    } else if (b) {
      range = normalizeRange({
        sheet: s,
        startRow: a.row!,
        endRow: b.row!,
        startCol: a.col!,
        endCol: b.col!,
        external
      });
    } else {
      range = {
        sheet: s,
        startRow: a.row!,
        endRow: a.row!,
        startCol: a.col!,
        endCol: a.col!,
        external
      };
    }
    const kind = b || open ? 'range' : 'cell';
    return { start, end, sheetPrefix, a, b, open, ref: { kind, raw: f.slice(start, end), range, external } };
  };

  /**
   * Parse the local part of a reference starting at `pos`. `tokenStart` is
   * where the whole span began (before any sheet prefix).
   */
  const parseLocal = (
    pos: number,
    tokenStart: number,
    sheetPrefix: string,
    sheet: string | undefined,
    external: string | undefined
  ): number => {
    const j = readIdent(pos);
    const token = f.slice(pos, j);
    if (!token) return Math.max(pos, tokenStart + 1);

    const cellM = CELL_TOKEN.exec(token);
    if (cellM) {
      const a = endpointFromCell(cellM);
      if (a) {
        if (f[j] === ':') {
          const k2 = readIdent(j + 1);
          const tok2 = f.slice(j + 1, k2);
          const cellM2 = CELL_TOKEN.exec(tok2);
          const b = cellM2 ? endpointFromCell(cellM2) : null;
          if (b) {
            spans.push(makeSpan(tokenStart, k2, sheetPrefix, sheet, external, a, b));
            return k2;
          }
        }
        // A function name can also look like a cell (LOG10, …): check call syntax.
        if (nextNonSpace(j) === '(') return j;
        spans.push(makeSpan(tokenStart, j, sheetPrefix, sheet, external, a));
        return j;
      }
    }

    const colM = COL_TOKEN.exec(token);
    if (colM && f[j] === ':') {
      const k2 = readIdent(j + 1);
      const tok2 = f.slice(j + 1, k2);
      const colM2 = COL_TOKEN.exec(tok2);
      if (colM2 && nextNonSpace(k2) !== '(') {
        const a: EndpointInfo = { col: letterToCol(colM[2]), absCol: colM[1] === '$' };
        const b: EndpointInfo = { col: letterToCol(colM2[2]), absCol: colM2[1] === '$' };
        spans.push(makeSpan(tokenStart, k2, sheetPrefix, sheet, external, a, b, 'columns'));
        return k2;
      }
    }

    const rowM = ROW_TOKEN.exec(token);
    if (rowM) {
      if (f[j] === ':') {
        const k2 = readIdent(j + 1);
        const tok2 = f.slice(j + 1, k2);
        const rowM2 = ROW_TOKEN.exec(tok2);
        if (rowM2) {
          const a: EndpointInfo = { row: parseInt(rowM[2], 10), absRow: rowM[1] === '$' };
          const b: EndpointInfo = { row: parseInt(rowM2[2], 10), absRow: rowM2[1] === '$' };
          spans.push(makeSpan(tokenStart, k2, sheetPrefix, sheet, external, a, b, 'rows'));
          return k2;
        }
      }
      return j; // bare numeric-ish token ($5) — not a reference
    }

    // Function call?
    if (nextNonSpace(j) === '(') return j;

    // Structured reference: Table1[...], with possible nested brackets.
    if (f[j] === '[') {
      let k = j;
      let depth = 0;
      do {
        if (f[k] === '[') depth++;
        else if (f[k] === ']') depth--;
        k++;
      } while (k < n && depth > 0);
      spans.push({
        start: tokenStart,
        end: k,
        sheetPrefix,
        ref: { kind: 'structured', raw: f.slice(tokenStart, k), name: token, external }
      });
      warnings.push(`Structured reference "${token}[…]" recorded but not resolved to cells`);
      return k;
    }

    if (/^(TRUE|FALSE)$/i.test(token)) return j;

    // Unquoted 3-D reference: Sheet1:Sheet2!A1
    if (f[j] === ':' && sheet === undefined) {
      const k2 = readIdent(j + 1);
      if (f[k2] === '!') {
        const sheet2 = f.slice(j + 1, k2);
        warnings.push(`3-D reference ${token}:${sheet2}!… — only first sheet "${token}" is traversed`);
        return parseLocal(k2 + 1, tokenStart, f.slice(tokenStart, k2 + 1), token, external);
      }
    }

    spans.push({
      start: tokenStart,
      end: j,
      sheetPrefix,
      ref: { kind: 'name', raw: f.slice(tokenStart, j), name: token, external }
    });
    return j;
  };

  while (i < n) {
    const ch = f[i];

    if (ch === '"') {
      i++;
      while (i < n) {
        if (f[i] === '"') {
          if (f[i + 1] === '"') {
            i += 2;
            continue;
          }
          i++;
          break;
        }
        i++;
      }
      continue;
    }

    if (ch === "'") {
      const start = i;
      i++;
      let name = '';
      let closed = false;
      while (i < n) {
        if (f[i] === "'") {
          if (f[i + 1] === "'") {
            name += "'";
            i += 2;
            continue;
          }
          closed = true;
          i++;
          break;
        }
        name += f[i++];
      }
      if (closed && f[i] === '!') {
        const prefix = f.slice(start, i + 1);
        let sheet = name;
        let external: string | undefined;
        const em = /^(.*)\[([^\]]+)\](.*)$/.exec(name);
        if (em) {
          external = (em[1] || '') + em[2];
          sheet = em[3];
        }
        if (sheet.includes(':')) {
          const [s1] = sheet.split(':');
          warnings.push(`3-D reference '${name}'!… — only first sheet "${s1}" is traversed`);
          sheet = s1;
        }
        i = parseLocal(i + 1, start, prefix, sheet, external);
      }
      continue;
    }

    if (ch === '[') {
      // Unquoted external prefix: [Budget.xlsx]FY26!B2 or [1]Sheet1!A1
      const m = /^\[([^\]]*)\]([A-Za-z_][A-Za-z0-9_.]*)!/.exec(f.slice(i));
      if (m) {
        const start = i;
        i = parseLocal(i + m[0].length, start, m[0], m[2], m[1]);
        continue;
      }
      i++;
      continue;
    }

    if (isIdentStart(ch)) {
      const j = readIdent(i);
      if (f[j] === '!') {
        const sheet = f.slice(i, j);
        i = parseLocal(j + 1, i, f.slice(i, j + 1), sheet, undefined);
        continue;
      }
      i = parseLocal(i, i, '', undefined, undefined);
      continue;
    }

    if (isDigit(ch)) {
      // Number literal, or a whole-row range like 2:5.
      let j = i;
      while (j < n && isDigit(f[j])) j++;
      if (f[j] === ':') {
        const dollar = f[j + 1] === '$' ? 1 : 0;
        let k2 = j + 1 + dollar;
        while (k2 < n && isDigit(f[k2])) k2++;
        if (k2 > j + 1 + dollar && !/[.A-Za-z(]/.test(f[k2] ?? '')) {
          const a: EndpointInfo = { row: parseInt(f.slice(i, j), 10), absRow: false };
          const b: EndpointInfo = { row: parseInt(f.slice(j + 1 + dollar, k2), 10), absRow: dollar === 1 };
          spans.push(makeSpan(i, k2, '', undefined, undefined, a, b, 'rows'));
          i = k2;
          continue;
        }
      }
      while (j < n && /[0-9.]/.test(f[j])) j++;
      if (f[j] === 'e' || f[j] === 'E') {
        let k = j + 1;
        if (f[k] === '+' || f[k] === '-') k++;
        if (isDigit(f[k] ?? '')) {
          while (k < n && isDigit(f[k])) k++;
          j = k;
        }
      }
      i = j;
      continue;
    }

    i++;
  }

  return { refs: spans.map((s) => s.ref), spans, warnings };
}

function shiftEndpoint(e: EndpointInfo, dRow: number, dCol: number): EndpointInfo {
  return {
    row: e.row !== undefined && !e.absRow ? Math.max(1, e.row + dRow) : e.row,
    col: e.col !== undefined && !e.absCol ? Math.max(1, e.col + dCol) : e.col,
    absRow: e.absRow,
    absCol: e.absCol
  };
}

function endpointText(e: EndpointInfo): string {
  let s = '';
  if (e.col !== undefined) s += (e.absCol ? '$' : '') + colToLetter(e.col);
  if (e.row !== undefined) s += (e.absRow ? '$' : '') + String(e.row);
  return s;
}

/**
 * Re-emit a formula with all *relative* references shifted by (dRow, dCol).
 * Used to translate shared formulas onto their member cells. Names,
 * structured and external references are left untouched.
 */
export function offsetFormula(formula: string, currentSheet: string, dRow: number, dCol: number): string {
  const { spans } = extractRefs(formula, currentSheet);
  let out = '';
  let pos = 0;
  for (const span of spans) {
    out += formula.slice(pos, span.start);
    if (span.a && (span.ref.kind === 'cell' || span.ref.kind === 'range')) {
      const a = shiftEndpoint(span.a, dRow, dCol);
      const text = span.b
        ? `${endpointText(a)}:${endpointText(shiftEndpoint(span.b, dRow, dCol))}`
        : endpointText(a);
      out += span.sheetPrefix + text;
    } else {
      out += formula.slice(span.start, span.end);
    }
    pos = span.end;
  }
  out += formula.slice(pos);
  return out;
}

function canonicalEndpoint(e: EndpointInfo, baseRow: number, baseCol: number): string {
  let s = '';
  if (e.row !== undefined) s += e.absRow ? `R${e.row}` : `R[${e.row - baseRow}]`;
  if (e.col !== undefined) s += e.absCol ? `C${e.col}` : `C[${e.col - baseCol}]`;
  return s;
}

/**
 * Canonical, position-independent form of a formula relative to the cell
 * that holds it: `=D4*E4` at F4 becomes `R[0]C[-2]*R[0]C[-1]`. Two cells in
 * a column computed "the same way" share a canonical form, which is how
 * Pharos detects computed columns. Non-reference text is uppercased so
 * `sum(...)` and `SUM(...)` compare equal.
 */
export function canonicalizeFormula(
  formula: string,
  currentSheet: string,
  baseRow: number,
  baseCol: number
): string {
  const { spans } = extractRefs(formula, currentSheet);
  let out = '';
  let pos = 0;
  for (const span of spans) {
    out += formula.slice(pos, span.start).toUpperCase();
    if (span.a && (span.ref.kind === 'cell' || span.ref.kind === 'range')) {
      const a = canonicalEndpoint(span.a, baseRow, baseCol);
      const text = span.b ? `${a}:${canonicalEndpoint(span.b, baseRow, baseCol)}` : a;
      out += span.sheetPrefix.toUpperCase() + text;
    } else {
      out += formula.slice(span.start, span.end).toUpperCase();
    }
    pos = span.end;
  }
  out += formula.slice(pos).toUpperCase();
  return out;
}

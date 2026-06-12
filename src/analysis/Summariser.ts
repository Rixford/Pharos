/**
 * Summariser — turns regions into human-readable, token-budgeted summaries
 * at six granularities:
 *
 *   summary  · one-paragraph English description
 *   compact  · summary + per-column statistics and samples
 *   evidence · compact + specific cell addresses backing each claim
 *   cells    · summary + the raw cell objects of the region
 *   formulas · summary + every formula with its references and templates
 *   audit    · everything, including styles and full region metadata
 *
 * Every summary carries `sourceCells` so any statement can be traced back
 * to concrete coordinates, and a `truncated` flag when a token budget
 * forced anything to be dropped.
 */
import { colToLetter, formatCell } from '../core/address';
import { CellLookup, GridCell } from '../core/grid';
import {
  ColumnProfile,
  GranularityMode,
  JsonScalar,
  RegionData,
  RegionKind,
  RegionSummary
} from '../core/types';
import { fmtDate, fmtNumber } from '../core/util';
import { extractRefs } from '../parser/FormulaParser';

/**
 * Crude but deterministic token estimate (≈4 characters per token). Good
 * enough for budgeting; swap in a real tokenizer if you need exact counts.
 */
export function estimateTokens(payload: string | unknown): number {
  if (payload === undefined || payload === null) return 0;
  const s = typeof payload === 'string' ? payload : JSON.stringify(payload);
  return Math.ceil((s?.length ?? 0) / 4);
}

const KIND_LABEL: Record<RegionKind, string> = {
  table: 'Table',
  matrix: 'Matrix',
  keyValue: 'Key/value block',
  list: 'List',
  block: 'Block',
  notes: 'Notes'
};

function jsonValue(cell: GridCell): JsonScalar {
  return cell.value instanceof Date ? fmtDate(cell.value) : cell.value;
}

function display(v: JsonScalar): string {
  if (v === null) return '';
  if (typeof v === 'number') return fmtNumber(v);
  const s = String(v);
  return s.length > 60 ? `${s.slice(0, 57)}…` : s;
}

function colName(c: ColumnProfile): string {
  return c.header ?? c.letter;
}

interface CellEntry {
  a: string;
  v: JsonScalar;
  t: string;
  f?: string;
  s?: unknown;
}

export function summariseRegion(
  region: RegionData,
  lookup: CellLookup,
  mode: GranularityMode = 'summary',
  tokenBudget?: number
): RegionSummary {
  const r = region;
  const cellAt = (row: number, col: number): GridCell | undefined => lookup(r.sheet, row, col);
  const sourceCells: string[] = [r.rangeA1];
  let truncated = false;

  // ── shared text fragments ─────────────────────────────────────────────
  const parts: string[] = [];
  parts.push(
    `${KIND_LABEL[r.kind]} ${r.rangeA1}` +
      (r.hiddenSheet ? ' (hidden sheet)' : '') +
      (r.title ? ` — “${r.title}”` : '')
  );
  let dims = `${r.dataRowCount} data row${r.dataRowCount === 1 ? '' : 's'} × ${r.colCount} column${r.colCount === 1 ? '' : 's'}`;
  if (r.headers) {
    const shown = r.headers.slice(0, 8);
    dims += ` (${shown.join(', ')}${r.headers.length > 8 ? ', …' : ''})`;
  }
  parts.push(dims);

  const keyCol = r.columns.find((c) => c.isKey);
  if (keyCol) parts.push(`key column: ${colName(keyCol)}`);

  for (const c of r.columns.filter((c) => c.formulaTemplate).slice(0, 3)) {
    parts.push(`${colName(c)} is computed (${c.formulaExample})`);
  }

  const numericCols = r.columns.filter((c) => c.stats);
  const headline = numericCols.sort(
    (a, b) => Math.abs(b.stats!.sum) - Math.abs(a.stats!.sum)
  )[0];
  if (headline?.stats) {
    parts.push(
      `${colName(headline)}: sum ${fmtNumber(headline.stats.sum)}, range ${fmtNumber(headline.stats.min)}–${fmtNumber(headline.stats.max)}`
    );
  }

  const dateCol = r.columns.find((c) => c.dateRange);
  if (dateCol?.dateRange) {
    parts.push(`${colName(dateCol)} spans ${dateCol.dateRange.min} → ${dateCol.dateRange.max}`);
  }

  if (r.purpose) parts.splice(1, 0, `purpose: ${r.purpose}`);
  if (r.headerRows && r.headerRows.length > 1) parts.push('2-row grouped header');
  const groupSections = (r.sections ?? []).filter((s) => s.kind === 'group');
  if (groupSections.length > 0) {
    parts.push(
      `${groupSections.length} grouped section(s)` +
        ((r.subtotalRows?.length ?? 0) > 0
          ? ` — subtotal rows ${r.subtotalRows!.slice(0, 8).join(', ')} excluded from stats`
          : '')
    );
  }
  if (r.notes && r.notes.length > 0) {
    parts.push(`note: ${r.notes[0].slice(0, 120)}${r.notes.length > 1 ? ` (+${r.notes.length - 1} more)` : ''}`);
  }

  const totalsPairs: string[] = [];
  if (r.totalsRow !== undefined) {
    for (let c = r.range.startCol; c <= r.range.endCol && totalsPairs.length < 4; c++) {
      const cell = cellAt(r.totalsRow, c);
      if (cell && (cell.formula || cell.type === 'number')) {
        totalsPairs.push(`${colToLetter(c)}${r.totalsRow}=${display(jsonValue(cell))}`);
      }
    }
    parts.push(`totals row at ${r.totalsRow}${totalsPairs.length ? ` (${totalsPairs.join(', ')})` : ''}`);
  }

  if (r.kind === 'keyValue') {
    const pairs: string[] = [];
    for (let row = r.dataStartRow; row <= r.dataEndRow && pairs.length < 6; row++) {
      const k = cellAt(row, r.range.startCol);
      const v = cellAt(row, r.range.startCol + 1);
      if (k && v) pairs.push(`${display(jsonValue(k))} = ${display(jsonValue(v))}`);
    }
    if (pairs.length > 0) parts.push(`pairs: ${pairs.join('; ')}`);
  }

  const summaryText = parts.join(' · ');

  const columnLines = (): string => {
    const lines: string[] = ['columns:'];
    for (const c of r.columns.slice(0, 12)) {
      if (c.nonEmpty === 0) continue;
      let line = `- ${colName(c)} (${c.letter}): ${c.type}, ${c.nonEmpty} values`;
      if (c.formulaExample) line += `, computed ${c.formulaExample}`;
      if (c.stats) {
        line += `, sum ${fmtNumber(c.stats.sum)}, min ${fmtNumber(c.stats.min)}, max ${fmtNumber(c.stats.max)}, mean ${fmtNumber(c.stats.mean)}`;
      }
      if (c.dateRange) line += `, ${c.dateRange.min} → ${c.dateRange.max}`;
      if (c.type === 'string') line += `, ${c.distinct} distinct`;
      if (c.samples.length > 0 && !c.stats) {
        line += `, e.g. ${c.samples.slice(0, 2).map(display).join(', ')}`;
      }
      lines.push(line);
    }
    if (r.columns.length > 12) lines.push(`- … ${r.columns.length - 12} more columns`);
    return lines.join('\n');
  };

  const evidenceLines = (): string => {
    const lines: string[] = ['evidence:'];
    if (r.headerRow !== undefined) {
      const span = `${r.sheet}!${colToLetter(r.range.startCol)}${r.headerRow}:${colToLetter(r.range.endCol)}${r.headerRow}`;
      lines.push(`- headers at ${span}`);
      sourceCells.push(span);
    }
    if (r.titleRange) {
      lines.push(`- title at ${r.titleRange}`);
      sourceCells.push(r.titleRange);
    }
    if (headline?.stats?.maxAt) {
      lines.push(`- max ${colName(headline)} ${fmtNumber(headline.stats.max)} at ${headline.stats.maxAt}`);
      sourceCells.push(headline.stats.maxAt);
    }
    if (headline?.stats?.minAt) {
      lines.push(`- min ${colName(headline)} ${fmtNumber(headline.stats.min)} at ${headline.stats.minAt}`);
      sourceCells.push(headline.stats.minAt);
    }
    if (r.totalsRow !== undefined) {
      let shown = 0;
      for (let c = r.range.startCol; c <= r.range.endCol && shown < 3; c++) {
        const cell = cellAt(r.totalsRow, c);
        if (cell?.formula) {
          const addr = formatCell({ sheet: r.sheet, row: r.totalsRow, col: c });
          lines.push(`- totals: ${addr} =${cell.formula} → ${display(jsonValue(cell))}`);
          sourceCells.push(addr);
          shown++;
        }
      }
    }
    // First data row as a concrete sample.
    if (r.dataRowCount > 0) {
      const sample: string[] = [];
      for (let c = r.range.startCol; c <= r.range.endCol && sample.length < 8; c++) {
        const cell = cellAt(r.dataStartRow, c);
        if (cell) {
          const local = `${colToLetter(c)}${r.dataStartRow}`;
          sample.push(`${local}=${cell.formula ? `=${cell.formula}→` : ''}${display(jsonValue(cell))}`);
        }
      }
      if (sample.length > 0) {
        lines.push(`- sample row ${r.dataStartRow}: ${sample.join(', ')}`);
        sourceCells.push(
          `${r.sheet}!${colToLetter(r.range.startCol)}${r.dataStartRow}:${colToLetter(r.range.endCol)}${r.dataStartRow}`
        );
      }
    }
    return lines.join('\n');
  };

  const collectCells = (withStyle: boolean): CellEntry[] => {
    const out: CellEntry[] = [];
    for (let row = r.range.startRow; row <= r.range.endRow; row++) {
      for (let col = r.range.startCol; col <= r.range.endCol; col++) {
        const cell = lookup(r.sheet, row, col);
        if (!cell) continue;
        const entry: CellEntry = {
          a: formatCell({ sheet: r.sheet, row, col }),
          v: jsonValue(cell),
          t: cell.type
        };
        if (cell.formula) entry.f = cell.formula;
        if (withStyle && cell.style) entry.s = cell.style;
        out.push(entry);
      }
    }
    return out;
  };

  // ── assemble by mode ──────────────────────────────────────────────────
  let text: string;
  let data: unknown;

  switch (mode) {
    case 'summary':
      text = summaryText;
      break;
    case 'compact':
      text = `${summaryText}\n${columnLines()}`;
      break;
    case 'evidence':
      text = `${summaryText}\n${columnLines()}\n${evidenceLines()}`;
      break;
    case 'cells': {
      text = summaryText;
      data = { cells: collectCells(false), omitted: 0 };
      break;
    }
    case 'formulas': {
      text = summaryText;
      const formulas = collectCells(false)
        .filter((c) => c.f)
        .map((c) => ({
          a: c.a,
          f: c.f!,
          refs: extractRefs(c.f!, r.sheet).refs.map((ref) => ref.raw)
        }));
      const templates = r.columns
        .filter((c) => c.formulaTemplate)
        .map((c) => ({ column: colName(c), template: c.formulaTemplate!, example: c.formulaExample! }));
      data = { templates, formulas, omitted: 0 };
      break;
    }
    case 'audit': {
      text = `${summaryText}\n${columnLines()}\n${evidenceLines()}`;
      data = { region: r, cells: collectCells(true), omitted: 0 };
      break;
    }
  }

  // ── token budgeting ───────────────────────────────────────────────────
  if (tokenBudget !== undefined && tokenBudget > 0) {
    const over = (): boolean => estimateTokens(text) + estimateTokens(data) > tokenBudget;

    // 1. shrink data arrays (cells / formulas modes)
    if (data !== undefined && over()) {
      const d = data as { cells?: CellEntry[]; formulas?: { a: string }[]; omitted: number };
      const shrink = (arrName: 'cells' | 'formulas'): void => {
        const arr = d[arrName] as unknown[] | undefined;
        if (!arr) return;
        const original = arr.length;
        let current = arr;
        while (current.length > 1 && over()) {
          current = current.slice(0, Math.floor(current.length / 2));
          (d as Record<string, unknown>)[arrName] = current;
          d.omitted = original - current.length;
          truncated = true;
        }
      };
      shrink('cells');
      shrink('formulas');
      if (over()) {
        data = { note: 'data dropped to fit token budget', omitted: -1 };
        truncated = true;
      }
    }

    // 2. degrade text granularity
    if (over()) {
      if (text !== summaryText) {
        text = `${summaryText}\n${columnLines()}`;
        truncated = true;
      }
      if (over()) {
        text = summaryText;
        truncated = true;
      }
      if (over()) {
        text = `${text.slice(0, Math.max(40, tokenBudget * 4 - 2))}…`;
        truncated = true;
      }
    }
  }

  return {
    regionId: r.id,
    sheet: r.sheet,
    rangeA1: r.rangeA1,
    kind: r.kind,
    mode,
    text,
    data,
    sourceCells: [...new Set(sourceCells)],
    tokens: estimateTokens(text) + estimateTokens(data),
    truncated
  };
}

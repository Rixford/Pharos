/**
 * Extractor — zoom level 6: transformation-ready table extraction.
 *
 * Returns typed rows keyed by header names with per-row provenance,
 * subtotal/grand-total rows excluded by default (they are summaries of
 * other rows — including them double-counts), Dates normalised to ISO,
 * and deterministic paging via offset/limit. This is the API an agent
 * uses when it has finished narrowing and needs the actual data.
 */
import { colToLetter } from '../core/address';
import { CellLookup } from '../core/grid';
import { ColumnRole, JsonScalar, RegionData } from '../core/types';
import { fmtDate } from '../core/util';
import { estimateTokens } from './Summariser';

export interface ExtractOptions {
  /** Data-row offset after subtotal exclusion (default 0). */
  offset?: number;
  /** Maximum rows returned (default: all). */
  limit?: number;
  /** Restrict to columns whose name contains one of these (case-insensitive). */
  columns?: string[];
  /** Include subtotal/grand-total rows (default false; flagged when true). */
  includeSubtotals?: boolean;
}

export interface ExtractedColumn {
  name: string;
  letter: string;
  role?: ColumnRole;
  type: string;
}

export interface ExtractedTable {
  regionId: string;
  workbook?: string;
  sheet: string;
  rangeA1: string;
  columns: ExtractedColumn[];
  /** Typed data rows keyed by column name (Dates → ISO strings). */
  rows: Record<string, JsonScalar>[];
  /** One A1 range per returned row — exact source provenance. */
  rowProvenance: string[];
  /** Grouped sections covering the returned span (labels from subtotals). */
  sections?: { label?: string; rows: [number, number] }[];
  /** Sheet rows excluded as subtotal/grand-total summaries. */
  excludedSubtotalRows?: number[];
  /** Echo of the excluded subtotal rows (label + values) for verification. */
  subtotals?: { row: number; label?: string; values: Record<string, JsonScalar> }[];
  totalDataRows: number;
  offset: number;
  returned: number;
  /** True when rows contains every data row of the region. */
  complete: boolean;
  warnings: string[];
  tokens: number;
}

export function extractTable(
  region: RegionData,
  lookup: CellLookup,
  opts: ExtractOptions = {}
): ExtractedTable {
  const warnings: string[] = [];
  const subtotalSet = new Set(opts.includeSubtotals ? [] : (region.subtotalRows ?? []));
  if (opts.includeSubtotals && (region.subtotalRows?.length ?? 0) > 0) {
    warnings.push('subtotal rows are INCLUDED — beware of double counting');
  }

  // Unique column names (collisions get the column letter appended).
  const seen = new Map<string, number>();
  const names = region.columns.map((c) => {
    let name = c.header ?? c.letter;
    const key = name.toLowerCase();
    const count = seen.get(key) ?? 0;
    seen.set(key, count + 1);
    if (count > 0) name = `${name} (${c.letter})`;
    return name;
  });
  if (region.headerRow === undefined) {
    warnings.push('no header row detected — columns are named by letter');
  }

  const columnsMeta: ExtractedColumn[] = region.columns.map((c, i) => ({
    name: names[i],
    letter: c.letter,
    role: c.role,
    type: c.type
  }));
  const wanted = opts.columns?.map((w) => w.toLowerCase());
  const keep = columnsMeta.map((c) => (wanted ? wanted.some((w) => c.name.toLowerCase().includes(w)) : true));
  if (wanted && keep.every((k) => !k)) warnings.push(`no columns matched filter [${opts.columns!.join(', ')}]`);

  // Candidate data rows: inside the data span, not subtotals, not empty.
  const startCol = region.range.startCol;
  const dataRows: number[] = [];
  for (let r = region.dataStartRow; r <= region.dataEndRow; r++) {
    if (subtotalSet.has(r)) continue;
    let any = false;
    for (let c = region.range.startCol; c <= region.range.endCol; c++) {
      if (lookup(region.sheet, r, c)) {
        any = true;
        break;
      }
    }
    if (any) dataRows.push(r);
  }

  const offset = Math.max(0, opts.offset ?? 0);
  const slice = dataRows.slice(offset, opts.limit !== undefined ? offset + opts.limit : undefined);

  const rows: Record<string, JsonScalar>[] = [];
  const rowProvenance: string[] = [];
  for (const r of slice) {
    const obj: Record<string, JsonScalar> = {};
    columnsMeta.forEach((meta, i) => {
      if (!keep[i]) return;
      const cell = lookup(region.sheet, r, startCol + i);
      let value: JsonScalar = null;
      if (cell) value = cell.value instanceof Date ? fmtDate(cell.value) : (cell.value as JsonScalar);
      obj[meta.name] = value;
    });
    rows.push(obj);
    rowProvenance.push(
      `${region.sheet}!${colToLetter(region.range.startCol)}${r}:${colToLetter(region.range.endCol)}${r}`
    );
  }

  const sections = region.sections
    ?.filter((s) => s.kind === 'group')
    .map((s) => ({ label: s.label, rows: [s.startRow, s.endRow] as [number, number] }));

  const subtotals = (region.subtotalRows ?? [])
    .filter(() => !opts.includeSubtotals)
    .map((r) => {
      const values: Record<string, JsonScalar> = {};
      let label: string | undefined;
      columnsMeta.forEach((meta, i) => {
        const cell = lookup(region.sheet, r, startCol + i);
        const value: JsonScalar = cell
          ? cell.value instanceof Date
            ? fmtDate(cell.value)
            : (cell.value as JsonScalar)
          : null;
        values[meta.name] = value;
        if (label === undefined && typeof value === 'string' && value.trim() !== '') label = value;
      });
      return { row: r, label, values };
    });

  const table: ExtractedTable = {
    regionId: region.id,
    sheet: region.sheet,
    rangeA1: region.rangeA1,
    columns: columnsMeta.filter((_, i) => keep[i]),
    rows,
    rowProvenance,
    sections,
    excludedSubtotalRows: opts.includeSubtotals ? undefined : region.subtotalRows,
    subtotals: subtotals.length > 0 ? subtotals : undefined,
    totalDataRows: dataRows.length,
    offset,
    returned: rows.length,
    complete: offset === 0 && rows.length === dataRows.length,
    warnings,
    tokens: 0
  };
  table.tokens = estimateTokens({ columns: table.columns, rows: table.rows });
  return table;
}

// pharos:eof

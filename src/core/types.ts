/**
 * Shared type definitions for the Pharos spreadsheet context engine.
 */

/** Plain JS values a cell can hold after parsing. */
export type CellScalar = string | number | boolean | Date | null;

/** JSON-safe scalar (Dates are serialised to ISO strings). */
export type JsonScalar = string | number | boolean | null;

/** The inferred type of a cell's (computed) value. */
export type CellValueType = 'number' | 'string' | 'boolean' | 'date' | 'error' | 'empty';

/** Subset of visual style Pharos extracts — used by region heuristics and audits. */
export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  numFmt?: string;
  fillColor?: string;
}

/** A single cell coordinate (1-based row/col). */
export interface CellRef {
  sheet: string;
  row: number;
  col: number;
}

/** A rectangular range. Open (whole-row/column) refs are clamped when resolved. */
export interface RangeRef {
  sheet: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  /** Set when the reference points at another workbook, e.g. `Budget.xlsx`. */
  external?: string;
  /** Set when the source was a whole-column (A:A) or whole-row (1:3) reference. */
  open?: 'columns' | 'rows';
}

export type RefKind = 'cell' | 'range' | 'name' | 'structured';

/** A reference extracted from a formula. */
export interface FormulaRef {
  kind: RefKind;
  /** Verbatim text as it appeared in the formula. */
  raw: string;
  range?: RangeRef;
  /** Defined-name or table name for kind 'name' / 'structured'. */
  name?: string;
  external?: string;
  /** Ranges a defined name resolves to (filled by WorkbookGraph). */
  resolved?: RangeRef[];
}

/** Granularity modes for summarisation, ordered roughly cheap → exhaustive. */
export type GranularityMode = 'summary' | 'compact' | 'evidence' | 'cells' | 'formulas' | 'audit';

export const GRANULARITY_MODES: readonly GranularityMode[] = [
  'summary',
  'compact',
  'evidence',
  'cells',
  'formulas',
  'audit'
] as const;

/** Edge families used by the diffusion engine. */
export type EdgeType = 'spatial' | 'structural' | 'formula' | 'semantic' | 'sheet';

export interface EdgeWeights {
  spatial: number;
  structural: number;
  formula: number;
  semantic: number;
  sheet: number;
}

export const DEFAULT_EDGE_WEIGHTS: EdgeWeights = {
  formula: 1.0,
  structural: 0.9,
  semantic: 0.6,
  spatial: 0.5,
  sheet: 0.3
};

export interface DiffusionOptions {
  /** Maximum number of hops from the seed (default 2). */
  depth?: number;
  /** Granularity used for the most relevant regions (default 'compact'). */
  mode?: GranularityMode;
  /** Approximate token budget for the whole packet (default 2000). */
  tokenBudget?: number;
  /** Hard cap on regions included (default 8). */
  maxRegions?: number;
  /** Frontier entries below this weight are dropped (default 0.15). */
  minWeight?: number;
  /** Multiplicative decay applied per hop (default 0.75). */
  decay?: number;
  weights?: Partial<EdgeWeights>;
  /** Include a precedent/dependent trace excerpt for formula seeds (default true). */
  includeTrace?: boolean;
}

export interface ResolvedDiffusionOptions {
  depth: number;
  mode: GranularityMode;
  tokenBudget: number;
  maxRegions: number;
  minWeight: number;
  decay: number;
  weights: EdgeWeights;
  includeTrace: boolean;
}

export interface NumericStats {
  sum: number;
  min: number;
  max: number;
  mean: number;
  /** Address of the cell holding the max value (evidence). */
  maxAt?: string;
  /** Address of the cell holding the min value (evidence). */
  minAt?: string;
}

export interface ColumnProfile {
  /** 1-based column index in the sheet. */
  col: number;
  letter: string;
  header?: string;
  type: CellValueType | 'mixed';
  nonEmpty: number;
  distinct: number;
  stats?: NumericStats;
  /** Canonical relative formula shared by most of the column's data cells. */
  formulaTemplate?: string;
  /** The template instantiated at the first data row, e.g. '=D4*E4'. */
  formulaExample?: string;
  isKey?: boolean;
  samples: JsonScalar[];
  /** Min/max for date-typed columns (ISO strings). */
  dateRange?: { min: string; max: string };
}

export type RegionKind = 'table' | 'matrix' | 'keyValue' | 'list' | 'block';

export interface RegionData {
  /** Stable identifier derived from sheet + range (same workbook ⇒ same id). */
  id: string;
  sheet: string;
  /** Range including header and totals rows (excludes a detached title). */
  range: RangeRef;
  rangeA1: string;
  kind: RegionKind;
  title?: string;
  titleRange?: string;
  headerRow?: number;
  headers?: string[];
  columns: ColumnProfile[];
  rowCount: number;
  colCount: number;
  /** Rows of actual data (excludes header and totals rows). */
  dataRowCount: number;
  /** First and last row holding data (excludes header/totals/gap rows). */
  dataStartRow: number;
  dataEndRow: number;
  totalsRow?: number;
  hiddenSheet: boolean;
  /** Non-empty cells / bounding-box area. */
  density: number;
  cellCount: number;
  formulaCellCount: number;
  /** Heuristic confidence 0..1 that this is a meaningful region. */
  confidence: number;
}

export interface RegionSummary {
  regionId: string;
  sheet: string;
  rangeA1: string;
  kind: RegionKind;
  mode: GranularityMode;
  /** Human-readable description (always present, even in data-heavy modes). */
  text: string;
  /** Mode-dependent structured payload (cells, formulas, audit, …). */
  data?: unknown;
  /** Cells/ranges this summary was derived from — drill down to verify. */
  sourceCells: string[];
  tokens: number;
  truncated: boolean;
}

export interface TraceNode {
  address: string;
  kind: 'cell' | 'range' | 'name' | 'external';
  value?: JsonScalar;
  formula?: string;
  /** For range nodes: number of cells the range covers. */
  cellCount?: number;
  children: TraceNode[];
  cycle?: boolean;
  truncated?: boolean;
  note?: string;
}

export interface Relation {
  from: string;
  to: string;
  type: EdgeType;
  weight: number;
  why: string;
}

export interface SeedSummary {
  address: string;
  value: JsonScalar;
  type: CellValueType;
  formula?: string;
  style?: CellStyle;
  regionId?: string;
  namedRanges: string[];
}

export interface ContextPacket {
  seed: string;
  seedCell: SeedSummary;
  options: ResolvedDiffusionOptions;
  regions: RegionSummary[];
  trace?: { precedents?: TraceNode; dependents?: TraceNode };
  relations: Relation[];
  nextActions: string[];
  warnings: string[];
  /** Union of all sourceCells across the packet. */
  sourceCells: string[];
  tokens: number;
  truncated: boolean;
}

export interface SheetInfo {
  name: string;
  index: number;
  hidden: boolean;
  rowCount: number;
  colCount: number;
  cellCount: number;
  formulaCellCount: number;
  usedRangeA1?: string;
}

export interface RegionBrief {
  id: string;
  rangeA1: string;
  kind: RegionKind;
  title?: string;
  rows: number;
  cols: number;
  confidence: number;
}

export interface WorkbookOverview {
  sheets: SheetInfo[];
  definedNames: { name: string; ranges: string[] }[];
  totalCells: number;
  totalFormulaCells: number;
  externalRefs: string[];
  regionsBySheet: Record<string, RegionBrief[]>;
  warnings: string[];
}

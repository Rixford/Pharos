/**
 * Minimal read-only view of a sheet shared by the analysis modules
 * (RegionDetector, Summariser, Diffuser) so they do not depend on
 * WorkbookGraph directly.
 */
import { CellScalar, CellStyle, CellValueType, RangeRef } from './types';

export interface GridCell {
  row: number;
  col: number;
  value: CellScalar;
  type: CellValueType;
  formula?: string;
  style?: CellStyle;
}

export interface SheetGrid {
  name: string;
  hidden: boolean;
  maxRow: number;
  maxCol: number;
  /** Keyed by cellKey(row, col). Only non-empty cells are present. */
  cells: ReadonlyMap<string, GridCell>;
  merges: RangeRef[];
}

/** Cell lookup used by the Summariser. */
export type CellLookup = (sheet: string, row: number, col: number) => GridCell | undefined;

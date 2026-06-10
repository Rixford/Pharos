/**
 * Parser abstraction. WorkbookGraph consumes a ParsedWorkbook, so any
 * format (xlsx via ExcelJS today; CSV or Google Sheets adapters later) can
 * feed the engine by implementing WorkbookParser.
 */
import { CellScalar, CellStyle, CellValueType, RangeRef } from '../core/types';

export interface ParsedCell {
  row: number;
  col: number;
  value: CellScalar;
  type: CellValueType;
  /** Formula text without the leading '='. */
  formula?: string;
  /** Local address ("E4") of the master cell when a shared formula was used. */
  sharedFrom?: string;
  style?: CellStyle;
  hyperlink?: string;
}

export interface ParsedSheet {
  name: string;
  index: number;
  hidden: boolean;
  maxRow: number;
  maxCol: number;
  /** Keyed by cellKey(row, col). Only non-empty cells are present. */
  cells: Map<string, ParsedCell>;
  merges: RangeRef[];
}

export interface ParsedWorkbook {
  sheets: ParsedSheet[];
  definedNames: { name: string; ranges: string[] }[];
  warnings: string[];
}

export interface WorkbookParser {
  parse(input: Buffer): Promise<ParsedWorkbook>;
}

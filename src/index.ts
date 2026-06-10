/**
 * Pharos — turn Excel workbooks into navigable graphs.
 * https://github.com/…/pharos
 */
export { WorkbookGraph } from './core/WorkbookGraph';
export type { CellInspection, DefinedName, FindHit, LoadOptions, TraceHooks } from './core/WorkbookGraph';
export { Collection } from './core/Collection';
export type {
  CollectionInput,
  CollectionOverview,
  CollectionContextPacket,
  CollectionWorkbookInfo,
  CrossLink,
  DataLink,
  DataLinkEnd,
  FormulaLinkGroup,
  QualifiedRef,
  SharedName,
  UnresolvedExternal
} from './core/Collection';
export { CellNode } from './core/CellNode';
export { Region } from './core/Region';
export * from './core/types';
export type { GridCell, SheetGrid, CellLookup } from './core/grid';
export {
  colToLetter,
  letterToCol,
  parseCellAddress,
  parseRange,
  formatCell,
  formatRange,
  rangeContains,
  rangeArea,
  quoteSheet
} from './core/address';
export { extractRefs, offsetFormula, canonicalizeFormula } from './parser/FormulaParser';
export type { ExtractResult, RefSpan } from './parser/FormulaParser';
export { ExcelParser } from './parser/ExcelParser';
export type { WorkbookParser, ParsedWorkbook, ParsedSheet, ParsedCell } from './parser/types';
export { detectRegions } from './analysis/RegionDetector';
export { summariseRegion, estimateTokens } from './analysis/Summariser';
export { expandContext } from './analysis/Diffuser';

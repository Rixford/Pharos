# Changelog

## 0.2.0 — 2026-06-10

Pharos Collections — multi-workbook graphs.

- New `Collection` class: load several workbooks as one graph; external
  references (`'[Book.xlsx]Sheet'!A1`) resolve to loaded workbooks, with
  qualified addressing (`[Book.xlsx]Sheet!A1`) throughout.
- Cross-workbook precedent traces (via new optional `TraceHooks` on
  `WorkbookGraph.tracePrecedents`) and dependent traces; cycle detection
  spans workbooks; unloaded externals remain explicit stub nodes.
- Cross-workbook context diffusion: `Collection.expandContext` extends the
  seed packet with regions from linked workbooks (formula edges, cross
  dependents, shared defined names, data links) under one token budget.
- Link analytics: `overview()`, `links()`, `crossDependentsOf()`,
  `sharedNames()`, `dataLinks()` (lookup-style key-overlap detection) and
  unresolved-external reporting with load suggestions.
- CLI: new `pharos collection <files...>` command (`--links`, `--inspect`,
  `--context`, `--precedents`, `--dependents`, `--find`, `--json`).
- `WorkbookGraph.cells()` iterator and `RegionSummary.workbook` field.
- Backward compatible: the entire v0.1.x API is unchanged; the new trace
  parameter is optional. New multi-workbook fixture set + 19 tests.

## 0.1.0 — 2026-06-10

Initial release.

- `WorkbookGraph`: load .xlsx (path/Buffer), cell index, precedent/dependent
  edges (range- and named-range-aware), `inspect`, `findValue`, `overview`.
- `FormulaParser`: character-level reference tokenizer (quoted sheets,
  externals, whole rows/columns, structured refs, 3-D warnings),
  `offsetFormula` (shared-formula translation), `canonicalizeFormula`
  (computed-column templates).
- `RegionDetector`: merge-aware flood fill, title/totals satellite
  attachment, header detection, column profiling, key columns, formula
  templates, kinds (table/keyValue/matrix/list/block), stable `rg_` ids,
  confidence scores.
- `Summariser`: six granularity modes with token budgeting, graceful
  degradation and `sourceCells` traceability.
- `Diffuser`: weighted multi-edge context diffusion with depth/weight/
  region/budget stopping rules, relations, warnings and next actions.
- CLI `pharos`: load, inspect, region, context, precedents, dependents,
  find — human and `--json` output.
- Jest suite (unit + CLI integration) over a generated fixture workbook
  with tables, formula chains, hidden sheet, merges, named range and an
  external reference. GitHub Actions CI (Node 18/20/22) + npm publish
  workflow.

# Changelog

## 0.3.0 — 2026-06-12

The zoomable context model — proven by a closed-loop liquidity benchmark
(`bench/`, see bench/REPORT.md for before/after numbers).

- `extractTable` (zoom L6): complete typed rows keyed by headers, per-row
  A1 provenance, offset/limit paging, column filters, subtotal rows
  excluded by default with a verification echo (`subtotals`).
- Region sections (zoom L3): subtotal/grand-total rows detected by label
  and by SUM-formula shape; grouped sections with cleaned labels; column
  stats and key columns now exclude subtotal rows.
- Two-row grouped headers (`headerRows`, headers like `Amounts · Billed`).
- Notes blocks: new `notes` region kind, attached to host tables.
- Column roles (zoom L4): key/id/category/measure/date/month/computed/text.
- Region purpose tags (`payments / cash receipts (by month)` …).
- `locate(question)` on WorkbookGraph and Collection: deterministic
  question→region ranking with a business-synonym lexicon (pluggable).
- `sheetMap(sheet)` (zoom L1): region inventory with purposes and notes.
- Question-aware diffusion: `expandContext(seed, { question })`.
- CLI: `pharos map`, `pharos locate`, `pharos extract`, `context --question`.
- Benchmark harness: deterministic two-workbook fixture generator, hidden
  gold liquidity report, blind solver (Pharos-only), tolerance scorer,
  multi-seed runner (`npm run bench`), Jest e2e; validated additionally by
  a clean-room LLM blind agent (100% score on an unseen seed).
- Keys: pure-number columns are no longer marked key columns.

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

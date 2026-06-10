# Changelog

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

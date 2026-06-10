# Pharos — Design

This document explains how the engine works, why it is built this way, and
where the sharp edges are. Read it before changing heuristics.

## Architecture

```
                 ┌────────────────────────────────────────────────┐
 .xlsx ──────▶  │ parser/                                         │
 (path/Buffer)  │   ExcelParser (ExcelJS)  → ParsedWorkbook       │
                │   FormulaParser (reference tokenizer)           │
                └───────────────┬────────────────────────────────┘
                                ▼
                ┌────────────────────────────────────────────────┐
                │ core/WorkbookGraph                              │
                │   cells indexed by sheet!row,col                │
                │   precedent edges (FormulaRef, names resolved)  │
                │   dependent index (per-cell + interval lists)   │
                │   region registry (stable ids, cached)          │
                └───────┬───────────────┬───────────────┬────────┘
                        ▼               ▼               ▼
                ┌──────────────┐ ┌──────────────┐ ┌──────────────┐
                │ analysis/    │ │ analysis/    │ │ analysis/    │
                │ RegionDetector│ │ Summariser  │ │ Diffuser     │
                └──────────────┘ └──────────────┘ └──────────────┘
                        ▲               ▲               ▲
                        └───────────────┴───────────────┘
                                        │
                              cli/ (commander)  ·  programmatic API
```

Separation rules: `parser/` knows nothing about graphs; `analysis/` modules
see only read-only views (`SheetGrid`, `CellLookup`, `RegionData`) plus a
type-only reference to `WorkbookGraph` in the Diffuser; `cli/` only calls
public `WorkbookGraph` methods. This keeps the engine portable (a Python
port maps module-for-module) and parsers swappable.

## Parser choice: ExcelJS, not SheetJS

The brief suggested the `xlsx` package. We deliberately chose **ExcelJS**:

* the npm build of `xlsx` is frozen at 0.18.5 (SheetJS moved distribution
  to their own CDN) and carries known CVEs (ReDoS, prototype pollution) —
  a poor foundation for a published library;
* the community edition exposes no usable cell styles, and bold-header
  detection materially improves region heuristics;
* ExcelJS reads formulas (including shared formulas), merges, hidden
  sheets and defined names through one maintained API.

Cost: ExcelJS is heavier and slower on very large files. The
`WorkbookParser` interface (`parser/types.ts`) isolates the choice — a
SheetJS, CSV or Google Sheets adapter can be added without touching the
engine. `WorkbookGraph.load(buffer, { parser })` accepts any implementation.

## Formula reference extraction

`FormulaParser.extractRefs` is a character-level scanner, not a regex pass.
It tracks string literals (`"…""…"`), quoted sheet names (`'O''Brien'!A1`),
external prefixes (`[Budget.xlsx]FY26!B2`, `'C:\…\[Book1.xlsx]Sheet1'!A1`),
whole-row/column ranges (`A:A`, `$2:$5`), structured references
(`Table1[[#All],[Revenue]]`, recorded but unexpanded) and 3-D references
(first sheet traversed, warning emitted). Function names that look like
cells (`LOG10(`) are disambiguated by call syntax.

The same span machinery powers two derived operations:

* `offsetFormula` — shifts relative endpoints; used to translate shared
  formulas onto member cells (`B2` with shared master `B1:'A1*2'` becomes
  `A2*2`), keeping `$`-anchored parts fixed.
* `canonicalizeFormula` — rewrites references relative to the holding cell
  (`=D4*E4` at F4 → `R[0]C[-2]*R[0]C[-1]`). Cells computed "the same way"
  share a canonical form; that is the definition of a computed column and
  of the "dominant pattern" notes in traces.

Known gaps (intentional, documented): R1C1 notation, structured-reference
expansion (needs table metadata), union/intersection operators treated as
plain scans, `INDEX(...):A1`-style dynamic range endpoints.

## Graph construction

Cells are stored sparsely per sheet (`Map<"row,col", CellNode>`). Formula
cells get `FormulaRef[]` with defined names resolved at load. Dependents
are indexed lazily on first use:

* ranges ≤ 512 cells (`RANGE_EXPAND_CAP`) expand into per-cell edges;
* larger ranges are kept as `(range, dependent)` interval entries per
  sheet, scanned on lookup — `dependentsOf` is O(direct + intervals on the
  sheet), avoiding the quadratic blow-up of `SUM(A:A)`-style references.

Traces (`tracePrecedents` / `traceDependents`) are recursive with a
path-stack for cycle detection (`cycle: true` nodes), depth limits
(`truncated: true`), and *range nodes*: a referenced range appears as one
node with cell/formula counts and the dominant canonical template plus one
representative child, instead of thousands of children.

## Region detection

Pipeline (see `RegionDetector.ts`):

1. **Occupancy** — non-empty cells, plus merge shadows (a merged range is
   occupied if its master has content) so a merged title spans its width.
2. **Components** — 8-connected flood fill; overlapping bounding boxes are
   merged until stable.
3. **Satellites** — single-row fragments grouped by row:
   * *titles*: ≤4 all-string cells, 1–2 blank rows above a host, within its
     column span → become region `title` metadata (excluded from the range);
   * *totals*: fragments exactly one blank row below a host containing a
     formula or a Total-ish label → extend the region, set `totalsRow`.
   Totals contiguous with the data (no gap) are detected separately by
   label or by `SUM`-over-own-column formulas.
4. **Analysis** — header row (≥60% coverage, ≥80% strings, distinct,
   type-discontinuity below or bold-only-on-top), column profiles (type
   tally, numeric stats with max/min addresses, distinct counts, samples,
   date ranges, formula templates), key column (leftmost fully-populated
   all-distinct), kind classification (`table`, `keyValue`, `matrix`,
   `list`, `block`) and a confidence score.

Confidence is a transparent additive score (header +0.2, density ≥0.5
+0.1, ≥3 data rows +0.1, type consistency up to +0.2, title +0.05,
classified kind +0.05, single-cell −0.35) clamped to [0.05, 0.99]. It is
deliberately interpretable rather than learned.

**Stable ids.** `rg_` + FNV-1a hash of `sheet|range`. Detection is
deterministic, so the same workbook yields the same ids across runs and
processes — callers can cache and cross-reference them.

## Diffusion

`expandContext` is a best-first frontier search over regions:

```
weight(entry) = edgeWeight(type) × decay^depth      (defaults in types.ts)
```

Seed wiring: structural edge to the seed's region (or spatial to the
nearest region if the seed is empty — with a warning); formula edges to
regions overlapping each precedent range and each dependent's region;
semantic edges via shared defined names and ≥2 shared column headers;
spatial/sheet edges to other regions on the seed sheet. Popped regions
expand further along formula edges only (the high-signal family).

Stopping rules, in order of bite: `depth`, `minWeight` (with `decay` this
bounds the frontier), `maxRegions`, and the token budget. The mode ladder
degrades with depth (depth 0 = requested mode, depth 1 ≤ compact, deeper =
summary), so far context is cheap context.

The packet always reports its own gaps: `truncated`, `warnings` (hidden
sheets, external workbooks, structured refs, empty seeds) and
`nextActions` — concrete follow-up calls (`tracePrecedents('…')`,
`summariseRegion('rg_…', 'cells')`, `findValue(…)`) chosen from what was
actually cut or left unexplored. The intent: an agent should never need to
guess its next tool call.

## Summarisation & token budgets

Six modes are projections of the same `RegionData` (see README table). All
text is deterministic templating over detected metadata — no model calls,
so summaries are stable across runs and safe to snapshot in tests.

`estimateTokens` is `ceil(len/4)` over the string/JSON payload: biased
slightly high for prose, slightly low for dense JSON — acceptable for
budgeting, documented so nobody mistakes it for a tokenizer. Budget
enforcement degrades gracefully: data arrays halve until they fit
(`omitted` records the cut), then granularity steps down
(evidence→compact→summary), then text truncates. Anything degraded sets
`truncated: true`.

`sourceCells` is non-negotiable: every summary names the range (and, in
evidence mode, the specific cells) it derives from.

## Performance notes

For a workbook with N stored cells, R regions, F formula cells: parsing is
O(N); component detection O(N) with small constants; column profiling
O(cells in region); dependent indexing O(F × avg refs) with the cap above;
diffusion touches at most `maxRegions` summaries. Region detection and the
dependent index are computed once and cached per graph. Pathological
inputs (e.g. a 10⁶-cell `A:A` reference) are clamped by used-range
clamping and `RANGE_EXPAND_CAP`.

Not yet addressed: streaming parses (ExcelJS holds the workbook in
memory), worker-thread parallelism. Both are roadmap items; medium
workbooks (10⁴–10⁵ cells) load in well under a second of analysis time.

## Future directions

* **Parsers**: CSV (trivial adapter), Google Sheets (API-backed), SheetJS.
* **Semantic search**: optional vector index over cell/region text for
  cross-workbook similarity queries.
* **Summariser plugins**: register kind-specific summarisers (financial
  statements, Gantt-ish schedules) keyed on header signatures.
* **Diff**: structural + value + formula diff of two workbook versions,
  reported region-by-region.
* **Streaming**: chunked region detection for files that exceed memory.
* **Header models**: multi-row headers, merged header groups, units rows.

# Pharos 🗼

**Turn Excel workbooks into navigable graphs.** Pharos parses a workbook into a graph of cells, formulas, regions and sheets, then gives you high-level operations to *understand* it: detect tables, trace formula chains, and diffuse outward from any cell into a token-budgeted context packet — built for AI agents, analysts and developers.

> Named for the lighthouse of Alexandria: a fixed point that makes a confusing coastline navigable.

[![CI](https://github.com/Rixford/pharos/actions/workflows/ci.yml/badge.svg)](https://github.com/Rixford/pharos/actions/workflows/ci.yml)
[![npm](https://img.shields.io/npm/v/pharos-sheets)](https://www.npmjs.com/package/pharos-sheets)
[![license: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

## Why

A spreadsheet is a program wearing a table costume. Cell `Summary!C2` isn't the number `15601.05` — it's `=Sales!F35`, which is `=SUM(F4:F33)`, which is thirty rows of `=Units × Unit Price`. Feeding raw cell dumps to an LLM wastes tokens and loses exactly this structure; feeding it one cell loses everything else.

Pharos sits in the middle. Ask it about a cell and it answers like an analyst would:

> Seed `Summary!C2` = 15,601.05, computed as `=Sales!F35`. That total sums the **Revenue** column of the “ACME Q1 Sales” table (`Sales!A3:F35`, 30 rows × 6 columns, Revenue computed as `=D4*E4`). Related: a Region/Target table at `Sales!H3:I7`. Want more? Trace `Sales!F35`'s precedents, or fetch the full cells of `rg_…`.

Everything it says carries `sourceCells` — the exact coordinates each claim came from — so nothing is an opaque inference.

## Install

```bash
npm install pharos-sheets        # library
npx pharos-sheets --help         # CLI without installing
```

Requires Node ≥ 18. The CLI binary is `pharos`.

## CLI quickstart

```bash
pharos load report.xlsx                 # sheets, regions, named ranges, warnings
pharos inspect report.xlsx 'Sales!F35'  # value, formula, style, region, refs
pharos region report.xlsx --list        # all detected regions with stable ids
pharos region report.xlsx 'Sales!B7' --mode evidence
pharos context report.xlsx 'Summary!C2' --depth 2 --budget 1500
pharos precedents report.xlsx 'Summary!C3' --depth 4
pharos dependents report.xlsx 'Sales!F35'
pharos find report.xlsx "North" --sheet Sales
```

Every command accepts `--json` for machine-readable output. A `context` run looks like:

```text
Seed Summary!C2 = 15601.05 [number]  =Sales!F35  region rg_l2zijg
Options: depth 2 · mode compact · budget 2000 (used ~572) · truncated: no

Regions (3):
1. [rg_l2zijg] keyValue Summary!B2:C5  (mode compact, ~95 tokens)
   Key/value block Summary!B2:C5 · 4 data rows × 2 columns · pairs: Total Revenue = 15,601.05;
   Commission (10%) = 1,560.11; North Rate = 0.05; …
   sources: Summary!B2:C5
2. [rg_1hqooki] table Sales!A3:F35  (mode compact, ~190 tokens)
   Table Sales!A3:F35 — “ACME Q1 Sales” · 30 data rows × 6 columns (Date, Region, Product, Units,
   Unit Price, Revenue) · key column: Date · Revenue is computed (=D4*E4) · Revenue: sum 15,601.05,
   range 49.95–1,349.46 · Date spans 2026-01-01 → 2026-01-30 · totals row at 35 (F35=15,601.05)
   sources: Sales!A3:F35
3. [rg_1ayprhb] table Rates!A1:B5  (mode compact, ~68 tokens)
   Table Rates!A1:B5 (hidden sheet) · 4 data rows × 2 columns (Region, Rate) · key column: Region
   sources: Rates!A1:B5

Precedent trace:
  Summary!C2 = 15601.05  =Sales!F35
  └─ Sales!F35 = 15601.05  =SUM(F4:F33)
     └─ Sales!F4:F33 (30 cells)  — 30 cells, 30 formula(s); dominant pattern =D4*E4 (30×)

Relations:
  Summary!C2 →(structural 0.9) rg_l2zijg — contains the seed cell
  Summary!C2 →(formula 0.75) rg_1hqooki — referenced by Summary!C2 via Sales!F35
  rg_l2zijg →(formula 0.75) rg_1ayprhb — formulas in rg_l2zijg reference Rates!A2:B5

Next actions:
  1. findValue("Total Revenue") to locate rows keyed by B across sheets
  2. summariseRegion('rg_l2zijg', 'cells') for full cell-level detail

Warnings:
  ! Region rg_1ayprhb lives on hidden sheet "Rates"
```

## Library quickstart

```ts
import { WorkbookGraph } from 'pharos-sheets';

const graph = await WorkbookGraph.load('report.xlsx'); // path or Buffer

// 1. Inspect a cell: value, formula, type, style, region, references
const cell = graph.inspect('Sales!F35');
// → { value: 15601.05, formula: 'SUM(F4:F33)', region: { id: 'rg_…', … }, dependents: ['Summary!C2'], … }

// 2. Detect and summarise the region containing a cell
const region = graph.regionAt('Sales!B7');
const summary = graph.summariseRegion(region!, 'evidence', 500);
// summary.text  → English description with cell addresses for every claim
// summary.sourceCells → ['Sales!A3:F35', 'Sales!F11', …]

// 3. Trace formulas, both directions
const upstream = graph.tracePrecedents('Summary!C3', 4);
const downstream = graph.traceDependents('Sales!F35');

// 4. Diffuse context from a seed cell under a token budget
const packet = graph.expandContext('Summary!C2', {
  depth: 2,
  mode: 'compact',
  tokenBudget: 1500
});
// packet.regions      → relevance-ordered region summaries
// packet.relations    → why each region was included (edge type + weight)
// packet.nextActions  → suggested follow-up calls
// packet.truncated    → true if the budget forced anything out
```

Full signatures: [docs/API.md](docs/API.md). TypeScript definitions ship with the package.

## Concepts

**Graph.** Cells, regions, sheets and the workbook are nodes. Edges come in five families: **spatial** (adjacency), **structural** (cell ∈ region ∈ sheet), **formula** (precedents/dependents, including through ranges and named ranges), **semantic** (shared headers, shared defined names) and **sheet** (same-sheet co-location).

**Regions.** Detected — not declared — via occupancy flood-fill plus heuristics: blank-row boundaries, header rows (string coverage, type discontinuity with the row below, bold), attached titles and totals rows, per-column type/statistics profiling, key-column inference, and repeated-formula templates (`=D4*E4` repeated 30× ⇒ “Revenue is computed”). Regions get **stable ids** (`rg_…`, a hash of sheet + range) so separate calls can refer to the same table, plus a confidence score.

**Diffusion.** `expandContext` runs a weighted frontier search from the seed: the seed's region first, then regions linked by formulas, names, headers, adjacency. Expansion stops on depth, weight decay, region cap, or token budget — whichever bites first. Deeper regions are summarised at coarser granularity.

**Granularity modes.**

| mode | returns | typical use |
|---|---|---|
| `summary` | one-paragraph English description | orientation, cheap context |
| `compact` | summary + per-column stats & samples | default for agents |
| `evidence` | compact + cell addresses backing each claim | verification, citation |
| `cells` | summary + raw cell objects | exact values needed |
| `formulas` | summary + every formula, refs, templates | dependency analysis |
| `audit` | everything incl. styles & region metadata | debugging, deep dives |

**Token budgets.** Estimates use a deterministic ≈4-chars/token heuristic (`estimateTokens`). Budgets degrade output gracefully — data arrays shrink, granularity steps down — and anything dropped sets `truncated: true` and usually a `nextActions` hint for retrieving it.

**Source traceability.** Every `RegionSummary` and `ContextPacket` carries `sourceCells`. If Pharos says “max Revenue 1,349.46”, it also says *at `Sales!F11`* — drill down and verify.

## Configuration

`expandContext(address, options)`:

| option | default | meaning |
|---|---|---|
| `depth` | `2` | max hops from the seed |
| `mode` | `'compact'` | granularity for the closest regions |
| `tokenBudget` | `2000` | approximate packet budget |
| `maxRegions` | `8` | hard cap on regions |
| `minWeight` | `0.15` | frontier entries below this are dropped |
| `decay` | `0.75` | per-hop weight multiplier |
| `weights` | formula 1.0 · structural 0.9 · semantic 0.6 · spatial 0.5 · sheet 0.3 | edge-family weights |
| `includeTrace` | `true` | include precedent/dependent excerpts |

## Performance

Designed for medium workbooks (10,000+ cells): single-pass parsing, lazy + cached region detection and dependent indexing, range references expanded per-cell only below a cap (512 cells) and kept as interval entries above it. No quadratic passes. See [docs/DESIGN.md](docs/DESIGN.md) for complexity notes and the full heuristic catalogue.

## Known limitations

R1C1 formulas are not resolved; structured references (`Table1[Col]`) and 3-D references (`Sheet1:Sheet3!A1`) are recorded with warnings rather than fully expanded; external workbook values are never loaded (referenced workbooks are surfaced in `overview()` and warnings). Multi-row headers are treated as single-row.

## Roadmap

CSV and Google Sheets parser adapters (the `WorkbookParser` interface is already pluggable) · vector-store indexing for semantic cell search · custom summariser plugins for domain tables (financial statements, schedules) · workbook `diff` · streaming mode for very large files.

## Contributing & license

PRs welcome — see [CONTRIBUTING.md](CONTRIBUTING.md). Tests: `npm test` (builds, regenerates the fixture workbook, runs Jest). MIT © Nicholas Rixford.

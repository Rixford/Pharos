# Liquidity Benchmark Report — Pharos zoom model (v0.2.0 → v0.3.0)

A closed-loop benchmark proving Pharos can give an AI agent enough
spreadsheet context to reconstruct a liquidity report from two complex
source workbooks **without ever seeing the answer**.

## Method: roles and knowledge separation

| Role | Implementation | Sees |
|---|---|---|
| Data Generation Agent | `bench/data.ts` + `build-billing.ts` + `build-costcenter.ts` (seeded PRNG) | the raw dataset |
| Gold Report Agent | `bench/gold.ts` — computes the answer **from the raw dataset**, never from the files | dataset + builder manifests |
| Blind Context Agent | (1) `bench/solve.ts`, a scripted reference agent; (2) a clean-room LLM agent | **only** the two .xlsx files + Pharos |
| Improvement Agent | this development loop | everything |

Separation mechanics: the orchestrator (`bench/run.ts`) runs the blind
solve **before** any gold artifact is written to disk; gold is then
derived from the dataset (deterministic, so order doesn't matter
mathematically — but nothing secret exists during the blind phase). The
LLM validation run went further: a fresh agent in `/tmp/blind-agent`
containing only the two workbooks and the packed `pharos-sheets` library,
on a seed (99) whose gold had never been generated, under explicit
rules: all reading through Pharos, no file access outside the directory.
Its working scripts were retained for audit.

The scorer (`bench/score.ts`) compares candidate output to gold with
0.05-absolute / 0.2%-relative currency tolerance across: required sheets,
required columns, month labels, monthly inflows/outflows/net, five
grouped summaries (customer / category / department / cost center /
vendor), risk flags (critical: negative-net months + customer
concentration), provenance coverage of 8 major metrics, and 5
reconciliation checks.

## The fixtures (same dataset, two business views)

`billing.xlsx` — 10 tabs: Executive Summary (KPI block + notes), Billing
Detail (**2-row grouped header + per-customer subtotal rows + grand
total**), Customer Contracts (two tables/sheet), Invoices, Payments
(register + by-month summary table on one sheet), Aging matrix,
Adjustments & Credits, Deferred Revenue (annual prepaids + recognition
schedule + cash-vs-recognition notes), Assumptions (the cash rulebook),
hidden CategoryMap (VLOOKUP target, named range).

`costcenter.xlsx` — 11 tabs: Executive Summary, Cost Center Detail
(**dept-grouped subtotal rows**), Department Spend matrix (SUMIFS over
detail), Payroll matrix + headcount table, Vendor Spend register +
vendor-totals table, Opex matrix, Capex, Allocations (formulas into the
hidden map), Forecast (decoy — excluded from actuals), Assumptions,
hidden AllocMap (dept shares + cost-center payroll shares).

Designed traps: subtotal rows that double-count if summed; a notes block
that *names* the data concept it annotates; capex that must appear in
vendor totals; SHARED vendors that must be allocated via a hidden tab;
late/partial payments crossing months; credits reducing inflows;
forecast that must be ignored; dept codes vs display names.

## Baseline (v0.2.0): behavior before changes

What worked: workbook orientation (`load`), region detection of the flat
registers, hidden-sheet visibility, formula tracing, provenance strings.

What failed, classified:

| Failure | Category | Evidence (seed 42) |
|---|---|---|
| `cells`-mode summaries truncate large registers under a realistic 1,500-token call budget; no paging, no typed rows | **Granularity / extraction (primary)** | 4 incomplete extractions; monthly inflows 16,783 vs gold 144,777; monthly accuracy 0%, group rows 2% |
| Subtotal rows are invisible metadata-wise; agents must regex-guess to avoid double counting | **Region metadata / sections** | Cost Center Detail & Billing Detail interleave subtotals with data |
| Grouped 2-row headers undetected → `headers: undefined` → name-based extraction impossible | **Multi-row header detection** | Billing Detail region became headerless |
| No question→region API; agents re-implement keyword scoring | **Semantic narrowing** | solver shipped its own concept lexicon |
| Notes blocks classified as data (`list`/`block`) | **Subregion awareness** | status-legend and capex notes |
| **Baseline score** | | **❌ FAIL — 3/8 sections** |

(Recorded in `bench/baseline-seed-42.score.json`.)

## Changes made (each mapped to a failure)

1. **`extractTable` — zoom level 6** (`src/analysis/Extractor.ts`): complete
   typed rows keyed by headers, Dates→ISO, per-row A1 provenance,
   offset/limit paging, column filters, **subtotal rows excluded by
   default with an echo of what was excluded** (`subtotals`).
2. **Sections & subtotal detection** (RegionDetector): label- and
   SUM-formula-based subtotal rows; grouped sections with cleaned labels
   (`Acme Industrial — Subtotal` → group “Acme Industrial”); column
   statistics and key-column inference now exclude subtotal rows.
3. **Two-row grouped headers**: merged group row + sub-header row →
   `headerRows: [r, r+1]`, headers like `Amounts (USD) · Billed`.
4. **Notes blocks**: new region kind `notes`, attached to the host table
   (`region.notes`), and given purpose `notes` so they never outrank the
   data they annotate.
5. **`locate(question)` — deterministic semantic narrowing**
   (`src/analysis/Locate.ts`): built-in business-synonym lexicon
   (inflows≈payments≈receipts≈collections…), scores sheet/title/purpose/
   headers/section-labels/samples/notes, hidden-sheet boost,
   data-kind-over-notes tie-breaking, extensible lexicon (pluggable; no
   model calls). Single-book and Collection-wide.
6. **`sheetMap` — zoom level 1**: per-sheet region inventory with
   purposes, section counts, notes.
7. **Column roles** (level 4): key/id/category/measure/date/month/
   computed/text on every `ColumnProfile`.
8. **Purpose tags** on regions (`payments / cash receipts (by month)`…).
9. **Question-aware diffusion**: `expandContext(..., { question })`
   boosts and injects semantically matching regions with explicit
   `matches the question (…)` relations.
10. **CLI**: `pharos map`, `pharos locate`, `pharos extract`,
    `context --question`.

Solver-side (agent-behavior, not engine): exact-before-substring header
matching, dept code→name resolution **from the hidden mapping table
itself**, concept queries routed through `locate` when available.
Dataset fix: cash months clamped to the reporting horizon (a genuine
test-dataset flaw the loop caught).

## Results

| Metric (seed 42) | v0.2.0 baseline | v0.3.0 |
|---|---|---|
| Required sheets / columns / month labels | 100% / 100% / 100% | 100% / 100% / 100% |
| Monthly totals accuracy (need ≥98%) | **0%** | **100%** |
| Grouped summary row match (need ≥95%) | **2%** | **100%** |
| Risk flags (critical complete) | 0–90% | **100%** |
| Provenance coverage (need ≥90%) | 100% | **100%** |
| Reconciliation checks | **0%** | **100%** |
| Overall | ❌ FAIL | ✅ PASS |

Generalization: seeds **42, 7, 1337 all pass every section** (`npm run
bench -- 42 7 1337`). The loop is locked in as a Jest e2e test
(`test/bench.e2e.test.ts`), so `npm test` regenerates the workbooks and
re-proves the blind reconstruction on every run.

**LLM blind-agent validation (seed 99, gold never on disk during the
run):** a clean-room agent given only the two workbooks + the packed
library scored **100% on all 8 sections**. It independently: found both
hidden mapping tabs and used them for dept/cost-center rollups; derived
the cash rules from the Assumptions tabs; included capex in vendor
totals; netted credits by month; excluded the forecast; produced 8/8
provenance mappings and 5/5 reconciliations (~82k agent tokens, 29 tool
calls). Its friction feedback was folded back into the engine: the
subtotal-label heuristic no longer fires on mid-sentence “…totals…”
wording (it had silently excluded an assumptions row), and
`extractTable` now echoes excluded subtotal rows for verification.

Efficiency: the scripted solver consumes ~12.3k Pharos-reported tokens
for the full reconstruction (~15 region extractions + orientation); the
baseline consumed ~8.5k and **failed** — completeness was unobtainable at
any prose-summary budget without paging or typed extraction.

## Limitations and honesty notes

- The scripted solver is a *reference* agent: generic (keyword concepts,
  no fixture coordinates — it survives seed changes that reshuffle row
  counts/anomalies) but not an LLM; the LLM run covers realistic usage
  once. Isolation of the LLM agent is enforced by instruction + artifact
  audit, not an OS sandbox.
- Layout variation across seeds is data-level (row counts, values,
  statuses, anomalies), not structural re-layout; structural variants
  (column order shuffles, alternative tab names) are the next hardening
  step for the harness.
- `locate` is deliberately lexicon-based and deterministic; an optional
  embedder can re-rank `LocateHit`s but the core never requires one.

## Reproduce

```bash
npm run bench -- 42 7 1337         # full closed loop, three seeds
npx ts-node bench/gen-sources.ts 123 /tmp/somewhere    # sources only
npx ts-node bench/score-blind.ts 123 /path/to/candidate.xlsx
```

// pharos:eof

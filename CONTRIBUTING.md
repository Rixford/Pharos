# Contributing to Pharos

Thanks for helping! This project aims to be the reference open-source
toolkit for spreadsheet understanding — clarity and verifiability beat
cleverness.

## Setup

```bash
git clone https://github.com/Rixford/pharos.git
cd pharos
npm install
npm test          # builds, regenerates the fixture workbook, runs Jest
```

Node ≥ 18 required.

## Scripts

| command | what it does |
|---|---|
| `npm run build` | compile `src/` → `dist/` (tsc, CommonJS + d.ts) |
| `npm run fixtures` | regenerate `test/fixtures/sample.xlsx` (deterministic) |
| `npm test` | pretest = build + fixtures, then Jest (unit + integration) |
| `npm run lint` | eslint over `src`, `test`, `scripts` |
| `npm run typecheck` | tsc --noEmit over everything |
| `npm run bench -- 42 7 1337` | closed-loop liquidity benchmark (bench/REPORT.md) |

The fixture workbook is generated, never committed — edit
`scripts/make-fixtures.ts` if a test needs new shapes (hidden sheets,
merges, shared formulas, …) and keep it deterministic (no randomness, no
timestamps).

## Layout

```
src/core      types, A1 addressing, CellNode/Region, WorkbookGraph
src/parser    WorkbookParser interface, ExcelJS adapter, formula tokenizer
src/analysis  RegionDetector, Summariser, Diffuser (read-only views only)
src/cli       commander wiring + human renderers
test/         one spec per module + CLI integration on the built dist
docs/         DESIGN.md (read before changing heuristics), API.md
```

Dependency rules: `parser/` must not import from `core/WorkbookGraph` or
`analysis/`; `analysis/` sees grids/lookups and (type-only) the graph;
heuristic changes belong in `analysis/` with DESIGN.md updated in the same
PR.

## Tests

Every behavioural change needs a test. Heuristic changes (region
detection, summaries, diffusion weights) must keep existing fixture
assertions passing or update them *with reasoning in the PR description*.
Summaries are deterministic — assert on substrings/regexes, not whole
blobs, so wording can evolve.

## Pull requests

1. Branch from `main`; keep PRs focused.
2. `npm run lint && npm run typecheck && npm test` must pass (CI runs the
   same on Node 18/20/22).
3. Public API changes → update `docs/API.md` + README.
4. New heuristics or trade-offs → a paragraph in `docs/DESIGN.md`.

## Releases

Maintainers: bump `version` in package.json, update CHANGELOG.md, tag and
publish a GitHub Release — the `publish.yml` workflow ships to npm
(requires the `NPM_TOKEN` repository secret).

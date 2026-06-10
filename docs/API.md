# Pharos — API Reference

All exports come from the package root:

```ts
import { WorkbookGraph, Region, CellNode, extractRefs, estimateTokens } from 'pharos-sheets';
```

Address strings accept `Sheet1!A1`, `'My Sheet'!B4` (quoted, `''` escapes a
quote), `$A$1` anchors, and bare `A1` (resolved against the first visible
sheet). Sheet names are case-insensitive.

---

## class `WorkbookGraph`

### `static load(input: string | Buffer, options?: LoadOptions): Promise<WorkbookGraph>`

Parse a workbook from a file path or in-memory buffer.
`options.parser` accepts any `WorkbookParser` implementation (default:
`ExcelParser`). Throws on unreadable input.

### `sheets(): SheetInfo[]`

Name, index, `hidden`, used extent, cell and formula counts per sheet.

### `getCell(target: string | CellRef): CellNode | undefined`

The cell node, or `undefined` for empty cells. `CellNode` exposes
`address`, `value`, `valueJson` (Dates → ISO), `type`
(`number|string|boolean|date|error|empty`), `formula` (without `=`),
`style`, `refs` (precedent `FormulaRef[]`, names resolved), `regionId`.

### `inspect(address: string): CellInspection`

Everything known about one cell: value/type/formula/style/hyperlink,
merge info (with master), containing region brief, precedent references
(raw + resolved target), dependents (first 20 + total), containing named
ranges, hidden-sheet flag. Works for empty cells (`exists: false`).

### `detectRegions(sheet?: string): Region[]`

Detected regions for one sheet (cached) or all sheets. See `Region`.

### `regionAt(target: string | CellRef): Region | undefined`

The region containing a cell, if any.

### `getRegion(id: string): Region | undefined`

Resolve a stable region id (`rg_…`).

### `summariseRegion(target: string | Region, mode?: GranularityMode, tokenBudget?: number): RegionSummary`

`target` may be a `Region`, a region id, or any address inside the region.
Modes: `summary | compact | evidence | cells | formulas | audit` (default
`summary`). With a budget, output degrades gracefully and sets
`truncated`. The result always includes `text`, `sourceCells`, `tokens`;
data-bearing modes add `data` (`cells`, `formulas`+`templates`, or the
full audit payload).

### `tracePrecedents(address: string, depth = 2): TraceNode`

Recursive upstream trace. Nodes: `kind` (`cell | range | name |
external`), `value`, `formula`, `children`, plus `cycle`, `truncated`,
`cellCount` and human `note`s. Ranges are summarised (counts + dominant
formula template + one representative child) rather than exploded.

### `traceDependents(address: string, depth = 2): TraceNode`

Recursive downstream trace (max 30 children per node, noted when capped).

### `dependentsOf(target: string | CellRef): string[]`

Direct dependents (addresses), including through ranges and named ranges.

### `expandContext(address: string, options?: DiffusionOptions): ContextPacket`

Weighted diffusion from a seed cell. See `DiffusionOptions` defaults in
the README table. The `ContextPacket` contains:

| field | meaning |
|---|---|
| `seed`, `seedCell` | normalised address + value/formula/type/region/names |
| `regions: RegionSummary[]` | relevance-ordered, granularity falls with depth |
| `trace` | precedent/dependent excerpts (when affordable in budget) |
| `relations` | `{from, to, type, weight, why}` for every inclusion |
| `nextActions` | concrete follow-up calls for what was cut/unexplored |
| `warnings` | hidden sheets, external workbooks, empty seed, … |
| `sourceCells` | union of all evidence coordinates |
| `tokens`, `truncated` | budget accounting |

### `findValue(query: string | number | RegExp, opts?: { sheet?: string; limit?: number }): FindHit[]`

Value search: strings match case-insensitive substrings, numbers match
exactly, RegExp tests the string form. Default limit 50.

### `overview(): WorkbookOverview`

Workbook-level structure: sheets, regions per sheet (briefs), defined
names, external workbook references, totals, warnings. This is what
`pharos load` prints.

### Properties

`definedNames` (with resolved ranges) · `warnings` (deduplicated, capped)
· `externalRefs` · `defaultSheet`.

---

## class `Collection`

Multi-workbook layer (v0.2). Addresses are qualified `[Book.xlsx]Sheet!A1`;
unqualified addresses resolve against the first loaded workbook. Workbook
names are case-insensitive and default to file basenames.

### `static load(inputs, options?): Promise<Collection>`

`inputs: Array<string | { name?: string; input: string | Buffer }>` —
Buffers need an explicit `name`. `options` are the same `LoadOptions` as
`WorkbookGraph.load`.

### `add(name, graph)` · `workbooks()` · `graph(book)` · `defaultWorkbook`

A Collection is pure composition over ordinary `WorkbookGraph`s; `add`
lets you build one from graphs you already loaded.

### `overview(): CollectionOverview`

Per-workbook stats, grouped formula links (resolved and unresolved),
shared defined names, data links and warnings. This is what
`pharos collection <files...>` prints.

### `links(): CrossLink[]` · `crossDependentsOf(address): string[]`

Every formula-level cross-workbook reference; and the cells in *other*
workbooks whose formulas read a given cell (range- and named-range-aware).

### `tracePrecedents(address, depth = 2)` · `traceDependents(address, depth = 2)`

Like the `WorkbookGraph` methods, but external references continue into
loaded workbooks and dependents cross workbook boundaries. All node
addresses come back qualified; cycle detection spans workbooks; externals
that are *not* loaded remain explicit stub nodes.

### `expandContext(address, options?): CollectionContextPacket`

Runs single-workbook diffusion for the seed's book (~65% of the token
budget), then adds regions from other workbooks reached via formula
edges, cross-dependents, shared defined names and data links. The packet
gains `workbooks`, `crossLinks`, a `workbook` field on every region
summary, and next actions such as `Load "budget-2026.xlsx" into the
collection to resolve …`.

### `dataLinks(opts?): DataLink[]`

Lookup-style relationships detected from data: key/category columns
compared pairwise across workbooks (or across sheets of one workbook);
≥50% overlap of the smaller value set (and ≥2 shared values) yields a
link with coverage in both directions and sample keys.

### `sharedNames(): SharedName[]` · `findValue(query, { workbook?, sheet?, limit? })` · `summariseRegion(book, target, mode?, budget?)` · `inspect(address)`

Names defined in 2+ workbooks; collection-wide value search (qualified
addresses); region summaries tagged with their workbook; cell inspection
with `workbook` and `crossDependents` added.

### `resolveAddress(address): { book, ref }` · `resolveExternalName(name): string | undefined`

Address plumbing, exposed for tooling. `resolveExternalName` matches
basenames case-insensitively, tolerating paths and missing extensions;
numeric link-table indices (`[1]Sheet1!A1`) return undefined.

Related: `WorkbookGraph.tracePrecedents` accepts an optional third
`TraceHooks` argument (address qualifier, shared cycle stack, external
resolver) — this is the seam Collections use, available to other tooling.

---

## class `Region`

`id` (stable `rg_…`), `sheet`, `rangeA1` (sheet-qualified), `kind`
(`table | matrix | keyValue | list | block`), `title`, `headers`,
`confidence`, `contains(ref)`, `brief()`, and `data: RegionData` with the
full detection payload: `headerRow`, `columns: ColumnProfile[]` (type,
stats with `maxAt`/`minAt`, distinct, samples, `formulaTemplate`/`Example`,
`isKey`, `dateRange`), `dataStartRow/dataEndRow/dataRowCount`, `totalsRow`,
`density`, `cellCount`, `formulaCellCount`, `hiddenSheet`.

---

## Functions

### `extractRefs(formula: string, currentSheet: string): ExtractResult`

Tokenize a formula into `refs: FormulaRef[]` (`kind: cell | range | name |
structured`, with `range`, `external`, `name`), plus positional `spans`
and `warnings`. Handles quoted sheets, externals, whole rows/columns,
string literals, function-name lookalikes, 3-D refs.

### `offsetFormula(formula, currentSheet, dRow, dCol): string`

Shift relative references (shared-formula translation). `$` anchors and
names are preserved.

### `canonicalizeFormula(formula, currentSheet, baseRow, baseCol): string`

Position-independent form (`=D4*E4` @F4 → `R[0]C[-2]*R[0]C[-1]`); equal
canonical forms ⇔ same computation pattern.

### `estimateTokens(payload: string | unknown): number`

Deterministic ≈4-chars/token estimate used by all budgeting.

### `detectRegions(grid: SheetGrid): RegionData[]` · `summariseRegion(region, lookup, mode?, budget?)` · `expandContext(graph, address, options?)`

The underlying analysis functions, exported for advanced use (custom
grids, custom pipelines). Most consumers should use the `WorkbookGraph`
methods.

---

## Implementing a parser

```ts
import { WorkbookParser, ParsedWorkbook, WorkbookGraph } from 'pharos-sheets';

class CsvParser implements WorkbookParser {
  async parse(input: Buffer): Promise<ParsedWorkbook> {
    // one sheet, cells keyed by `${row},${col}`, no formulas
    …
  }
}
const graph = await WorkbookGraph.load(buffer, { parser: new CsvParser() });
```

---

## CLI

```
pharos load <file> [--json]
pharos inspect <file> <address> [--json]
pharos region <file> [target] [--list] [-s sheet] [-m mode] [-b tokens] [--json]
pharos context <file> <address> [-d depth] [-m mode] [-b tokens] [--max-regions n] [--no-trace] [--json]
pharos precedents <file> <address> [-d depth] [--json]
pharos dependents <file> <address> [-d depth] [--json]
pharos find <file> <query> [-s sheet] [--regex] [--limit n] [--json]

pharos collection <files...> [--json]              # overview (default)
pharos collection <files...> --links [--json]
pharos collection <files...> --inspect    "[book.xlsx]Sheet!A1"
pharos collection <files...> --context    "[book.xlsx]Sheet!A1" [-d depth] [-m mode] [-b tokens]
pharos collection <files...> --precedents "[book.xlsx]Sheet!A1" [-d depth]
pharos collection <files...> --dependents "[book.xlsx]Sheet!A1" [-d depth]
pharos collection <files...> --find <query>
```

Exit code 0 on success, 1 on any error (message on stderr, prefixed
`pharos:`). `--json` always emits a single parseable JSON document on
stdout.

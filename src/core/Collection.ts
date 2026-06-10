/**
 * Pharos Collections — represent multiple workbooks as one navigable graph.
 *
 * A Collection wraps several WorkbookGraphs and adds the cross-workbook
 * layer the single-workbook engine deliberately stops at:
 *
 *   · external references in formulas ('[Budget.xlsx]FY'!B2) resolve to the
 *     loaded workbook of that name, so precedent traces and context
 *     diffusion continue across files instead of ending at a stub;
 *   · cross-workbook dependents ("which other books read this cell?");
 *   · shared defined names across books;
 *   · data links: regions in different books whose key/category columns
 *     share values (the classic lookup/fact↔dimension relationship);
 *   · a collection-level overview of the link graph, including external
 *     workbooks that are referenced but not loaded.
 *
 * Addresses are qualified with the workbook in brackets: `[Book.xlsx]Sheet!A1`.
 * Unqualified addresses resolve against the first loaded workbook.
 *
 * The single-workbook v0.1.x API is untouched; a Collection is pure
 * composition on top of public WorkbookGraph methods plus two optional
 * trace hooks.
 */
import * as path from 'path';
import {
  cellKey,
  clampRange,
  formatCell,
  formatRange,
  parseRange,
  rangeContains,
  rangesOverlap
} from './address';
import { Region } from './Region';
import { LoadOptions, TraceHooks, WorkbookGraph } from './WorkbookGraph';
import {
  CellRef,
  ContextPacket,
  DiffusionOptions,
  GranularityMode,
  JsonScalar,
  RangeRef,
  RegionSummary,
  Relation,
  TraceNode
} from './types';
import { uniq } from './util';
import { estimateTokens } from '../analysis/Summariser';

export interface CollectionInput {
  /** Workbook name used in qualified addresses (defaults to the file basename). */
  name?: string;
  input: string | Buffer;
}

/** One formula-level reference from a cell in one workbook to another workbook. */
export interface CrossLink {
  fromBook: string;
  /** Local address of the referencing cell, e.g. `Dash!C2`. */
  fromCell: string;
  /** Verbatim reference text from the formula. */
  raw: string;
  /** External workbook name exactly as written in the formula. */
  external: string;
  /** Loaded workbook this resolves to (undefined when not in the collection). */
  toBook?: string;
  /** Resolved target range, local to the target workbook. */
  toRange?: string;
  /** @internal resolved target range used for containment queries. */
  range?: RangeRef;
}

export interface FormulaLinkGroup {
  fromBook: string;
  external: string;
  toBook?: string;
  refCount: number;
  cells: string[];
  targets: string[];
}

export interface SharedName {
  name: string;
  books: { book: string; ranges: string[] }[];
}

export interface DataLinkEnd {
  book: string;
  regionId: string;
  rangeA1: string;
  column: string;
}

/** Two regions (different books, or different sheets) sharing key values. */
export interface DataLink {
  a: DataLinkEnd;
  b: DataLinkEnd;
  shared: number;
  coverageA: number;
  coverageB: number;
  sample: JsonScalar[];
}

export interface CollectionWorkbookInfo {
  key: string;
  sheets: number;
  hiddenSheets: number;
  cells: number;
  formulaCells: number;
  regions: number;
}

export interface UnresolvedExternal {
  external: string;
  refCount: number;
  fromBooks: string[];
  sampleCells: string[];
}

export interface CollectionOverview {
  workbooks: CollectionWorkbookInfo[];
  formulaLinks: FormulaLinkGroup[];
  sharedNames: SharedName[];
  dataLinks: DataLink[];
  unresolved: UnresolvedExternal[];
  warnings: string[];
}

export interface CollectionContextPacket extends ContextPacket {
  workbooks: string[];
  /** Cross-workbook links touching the seed cell or its region. */
  crossLinks: CrossLink[];
}

export interface QualifiedRef {
  book: string;
  ref: CellRef;
}

const BOOK_RE = /^\[([^\]]+)\](.*)$/;
const MODE_RANK: Record<GranularityMode, number> = {
  summary: 0,
  compact: 1,
  evidence: 2,
  cells: 3,
  formulas: 3,
  audit: 4
};
const round2 = (n: number): number => Math.round(n * 100) / 100;

function qualifyTree(node: TraceNode, prefix: string): TraceNode {
  return {
    ...node,
    address: node.address.startsWith('[') ? node.address : `${prefix}${node.address}`,
    children: node.children.map((c) => qualifyTree(c, prefix))
  };
}

export class Collection {
  private readonly byKey = new Map<string, WorkbookGraph>();
  private readonly display = new Map<string, string>();
  private readonly order: string[] = [];
  readonly warnings: string[] = [];
  private crossIndex?: {
    links: CrossLink[];
    byTarget: Map<string, { range: RangeRef; link: CrossLink }[]>;
  };
  private dataLinksCache?: DataLink[];

  /** Load several workbooks (paths, or `{ name, input }` for buffers). */
  static async load(
    inputs: Array<string | CollectionInput>,
    options?: LoadOptions
  ): Promise<Collection> {
    const collection = new Collection();
    for (const item of inputs) {
      const spec: CollectionInput = typeof item === 'string' ? { input: item } : item;
      let name = spec.name;
      if (!name) {
        if (typeof spec.input !== 'string') {
          throw new Error('Collection.load: a Buffer input needs an explicit { name }');
        }
        name = path.basename(spec.input);
      }
      collection.add(name, await WorkbookGraph.load(spec.input, options));
    }
    return collection;
  }

  /** Add an already-loaded workbook under a name (usually its file name). */
  add(name: string, graph: WorkbookGraph): void {
    const key = name.toLowerCase();
    if (this.byKey.has(key)) {
      throw new Error(`Collection already contains a workbook named "${name}"`);
    }
    this.byKey.set(key, graph);
    this.display.set(key, name);
    this.order.push(name);
    this.crossIndex = undefined;
    this.dataLinksCache = undefined;
  }

  workbooks(): string[] {
    return [...this.order];
  }

  /** First loaded workbook — the default for unqualified addresses. */
  get defaultWorkbook(): string {
    return this.order[0];
  }

  graph(book: string): WorkbookGraph {
    const g = this.byKey.get(book.toLowerCase());
    if (!g) {
      throw new Error(`Workbook "${book}" is not in the collection. Loaded: ${this.order.join(', ')}`);
    }
    return g;
  }

  private displayName(book: string): string {
    return this.display.get(book.toLowerCase()) ?? book;
  }

  private q(book: string, s: string): string {
    return `[${this.displayName(book)}]${s}`;
  }

  /** Parse `[Book.xlsx]Sheet!A1` (or `Sheet!A1` against the default workbook). */
  resolveAddress(address: string): QualifiedRef {
    const m = BOOK_RE.exec(address.trim());
    if (this.order.length === 0) throw new Error('Collection is empty');
    const book = m ? m[1] : this.defaultWorkbook;
    const local = m ? m[2] : address;
    const ref = this.graph(book).resolveAddress(local);
    return { book: this.displayName(book), ref };
  }

  /** Map an external workbook name from a formula to a loaded workbook. */
  resolveExternalName(external: string): string | undefined {
    const base = external.replace(/\\/g, '/').split('/').pop() ?? external;
    if (/^\d+$/.test(base)) return undefined; // numeric link-table index — see DESIGN.md
    const lower = base.toLowerCase();
    const candidates = [lower, `${lower}.xlsx`, lower.replace(/\.(xlsx|xlsm|xls)$/i, '')];
    for (const candidate of candidates) {
      const hit = this.display.get(candidate);
      if (hit) return hit;
    }
    return undefined;
  }

  // ── cross-link index ──────────────────────────────────────────────────────

  private index(): NonNullable<Collection['crossIndex']> {
    if (this.crossIndex) return this.crossIndex;
    const links: CrossLink[] = [];
    const byTarget = new Map<string, { range: RangeRef; link: CrossLink }[]>();
    for (const book of this.order) {
      const g = this.graph(book);
      for (const cell of g.cells()) {
        for (const fr of cell.refs) {
          const ext = fr.external ?? fr.range?.external;
          if (!ext) continue;
          const toBook = this.resolveExternalName(ext);
          const link: CrossLink = {
            fromBook: book,
            fromCell: cell.address,
            raw: fr.raw,
            external: ext,
            toBook
          };
          if (toBook && fr.range) {
            const target = this.graph(toBook);
            const grid = target.grid(fr.range.sheet);
            const clamped =
              grid && fr.range.open ? clampRange(fr.range, grid.maxRow, grid.maxCol) : fr.range;
            const clean: RangeRef = { ...clamped, external: undefined };
            link.range = clean;
            link.toRange = formatRange(clean);
            const key = toBook.toLowerCase();
            const list = byTarget.get(key) ?? [];
            list.push({ range: clean, link });
            byTarget.set(key, list);
          } else if (fr.range) {
            link.toRange = formatRange({ ...fr.range, external: undefined });
          }
          links.push(link);
        }
      }
    }
    this.crossIndex = { links, byTarget };
    return this.crossIndex;
  }

  /** Every formula-level cross-workbook reference in the collection. */
  links(): CrossLink[] {
    return [...this.index().links];
  }

  /** Cells in *other* workbooks whose formulas reference the given cell. */
  crossDependentsOf(target: string | QualifiedRef): string[] {
    const { book, ref } = typeof target === 'string' ? this.resolveAddress(target) : target;
    const entries = this.index().byTarget.get(book.toLowerCase()) ?? [];
    const out: string[] = [];
    for (const entry of entries) {
      if (rangeContains(entry.range, ref)) {
        out.push(this.q(entry.link.fromBook, entry.link.fromCell));
      }
    }
    return uniq(out);
  }

  // ── shared names & data links ─────────────────────────────────────────────

  /** Defined names that appear in two or more workbooks. */
  sharedNames(): SharedName[] {
    const byName = new Map<string, { display: string; books: { book: string; ranges: string[] }[] }>();
    for (const book of this.order) {
      for (const dn of this.graph(book).definedNames) {
        const key = dn.name.toLowerCase();
        const entry = byName.get(key) ?? { display: dn.name, books: [] };
        entry.books.push({ book, ranges: dn.ranges });
        byName.set(key, entry);
      }
    }
    return [...byName.values()]
      .filter((e) => e.books.length >= 2)
      .map((e) => ({ name: e.display, books: e.books }));
  }

  /**
   * Regions whose key/category columns share values across workbooks (or
   * across sheets of one workbook) — lookup-style relationships detected
   * from the data itself.
   */
  dataLinks(opts?: { minOverlap?: number; maxDistinct?: number }): DataLink[] {
    if (this.dataLinksCache && !opts) return this.dataLinksCache;
    const minOverlap = opts?.minOverlap ?? 0.5;
    const maxDistinct = opts?.maxDistinct ?? 1000;

    interface Candidate {
      book: string;
      region: Region;
      column: string;
      values: Set<string>;
    }
    const candidates: Candidate[] = [];
    for (const book of this.order) {
      const g = this.graph(book);
      for (const region of g.detectRegions()) {
        for (const col of region.data.columns) {
          const eligible =
            col.isKey ||
            (col.type === 'string' && col.distinct >= 2 && col.distinct <= maxDistinct && col.nonEmpty >= 2);
          if (!eligible) continue;
          const values = this.columnValues(g, region, col.col);
          if (values.size >= 2) {
            candidates.push({ book, region, column: col.header ?? col.letter, values });
          }
        }
      }
    }

    const out: DataLink[] = [];
    for (let i = 0; i < candidates.length; i++) {
      for (let j = i + 1; j < candidates.length; j++) {
        const A = candidates[i];
        const B = candidates[j];
        if (A.book === B.book && A.region.id === B.region.id) continue;
        if (A.book === B.book && A.region.sheet.toLowerCase() === B.region.sheet.toLowerCase()) continue;
        let shared = 0;
        const sample: JsonScalar[] = [];
        const [small, large] = A.values.size <= B.values.size ? [A.values, B.values] : [B.values, A.values];
        for (const v of small) {
          if (large.has(v)) {
            shared++;
            if (sample.length < 3) sample.push(v);
          }
        }
        if (shared < 2) continue;
        const coverageA = shared / A.values.size;
        const coverageB = shared / B.values.size;
        if (Math.max(coverageA, coverageB) < minOverlap) continue;
        out.push({
          a: { book: A.book, regionId: A.region.id, rangeA1: A.region.rangeA1, column: A.column },
          b: { book: B.book, regionId: B.region.id, rangeA1: B.region.rangeA1, column: B.column },
          shared,
          coverageA: round2(coverageA),
          coverageB: round2(coverageB),
          sample
        });
      }
    }
    out.sort((x, y) => y.shared - x.shared);
    if (!opts) this.dataLinksCache = out;
    return out;
  }

  private columnValues(g: WorkbookGraph, region: Region, col: number): Set<string> {
    const grid = g.grid(region.sheet);
    const out = new Set<string>();
    if (!grid) return out;
    for (let r = region.data.dataStartRow; r <= region.data.dataEndRow; r++) {
      const cell = grid.cells.get(cellKey(r, col));
      if (!cell || cell.value === null) continue;
      const v =
        cell.value instanceof Date ? cell.value.toISOString().slice(0, 10) : String(cell.value);
      out.add(v.trim().toLowerCase());
    }
    return out;
  }

  // ── overview ──────────────────────────────────────────────────────────────

  overview(): CollectionOverview {
    const workbooks: CollectionWorkbookInfo[] = this.order.map((book) => {
      const g = this.graph(book);
      const sheets = g.sheets();
      return {
        key: book,
        sheets: sheets.length,
        hiddenSheets: sheets.filter((s) => s.hidden).length,
        cells: sheets.reduce((a, s) => a + s.cellCount, 0),
        formulaCells: sheets.reduce((a, s) => a + s.formulaCellCount, 0),
        regions: g.detectRegions().length
      };
    });

    const { links } = this.index();
    const groups = new Map<string, FormulaLinkGroup>();
    for (const link of links) {
      const key = `${link.fromBook}→${link.external}`.toLowerCase();
      const group =
        groups.get(key) ??
        ({ fromBook: link.fromBook, external: link.external, toBook: link.toBook, refCount: 0, cells: [], targets: [] } as FormulaLinkGroup);
      group.refCount++;
      const fromQ = this.q(link.fromBook, link.fromCell);
      if (group.cells.length < 5 && !group.cells.includes(fromQ)) group.cells.push(fromQ);
      if (link.toRange && group.targets.length < 5 && !group.targets.includes(link.toRange)) {
        group.targets.push(link.toRange);
      }
      groups.set(key, group);
    }

    const unresolvedMap = new Map<string, UnresolvedExternal>();
    for (const link of links) {
      if (link.toBook) continue;
      const key = link.external.toLowerCase();
      const u =
        unresolvedMap.get(key) ??
        ({ external: link.external, refCount: 0, fromBooks: [], sampleCells: [] } as UnresolvedExternal);
      u.refCount++;
      if (!u.fromBooks.includes(link.fromBook)) u.fromBooks.push(link.fromBook);
      const fromQ = this.q(link.fromBook, link.fromCell);
      if (u.sampleCells.length < 3 && !u.sampleCells.includes(fromQ)) u.sampleCells.push(fromQ);
      unresolvedMap.set(key, u);
    }

    const warnings = [...this.warnings];
    for (const u of unresolvedMap.values()) {
      warnings.push(
        `External workbook "${u.external}" is referenced ${u.refCount}× but not loaded into the collection`
      );
    }

    return {
      workbooks,
      formulaLinks: [...groups.values()],
      sharedNames: this.sharedNames(),
      dataLinks: this.dataLinks(),
      unresolved: [...unresolvedMap.values()],
      warnings
    };
  }

  // ── cell-level operations ─────────────────────────────────────────────────

  inspect(address: string): ReturnType<WorkbookGraph['inspect']> & {
    workbook: string;
    crossDependents: string[];
  } {
    const { book, ref } = this.resolveAddress(address);
    const base = this.graph(book).inspect(formatCell(ref));
    return {
      ...base,
      address: this.q(book, base.address),
      workbook: book,
      crossDependents: this.crossDependentsOf({ book, ref })
    };
  }

  findValue(
    query: string | number | RegExp,
    opts?: { workbook?: string; sheet?: string; limit?: number }
  ): Array<{ workbook: string; address: string; value: JsonScalar; type: string; regionId?: string }> {
    const limit = opts?.limit ?? 50;
    const books = opts?.workbook ? [this.graph(opts.workbook) && this.displayName(opts.workbook)] : this.order;
    const out: Array<{ workbook: string; address: string; value: JsonScalar; type: string; regionId?: string }> = [];
    for (const book of books) {
      for (const hit of this.graph(book).findValue(query, { sheet: opts?.sheet, limit: limit - out.length })) {
        out.push({ workbook: book, ...hit, address: this.q(book, hit.address) });
        if (out.length >= limit) return out;
      }
    }
    return out;
  }

  /** Summarise a region in a specific workbook (sourceCells come back qualified). */
  summariseRegion(
    book: string,
    target: string | Region,
    mode: GranularityMode = 'summary',
    tokenBudget?: number
  ): RegionSummary {
    const display = this.displayName(book);
    const summary = this.graph(book).summariseRegion(target, mode, tokenBudget);
    return {
      ...summary,
      workbook: display,
      sourceCells: summary.sourceCells.map((s) => this.q(display, s))
    };
  }

  // ── cross-workbook tracing ────────────────────────────────────────────────

  private hooksFor(book: string, stack: Set<string>): TraceHooks {
    return {
      qualify: (s) => this.q(book, s),
      stackKeyPrefix: `${book.toLowerCase()}::`,
      resolveExternal: (external, range, _raw, depth) => {
        const toBook = this.resolveExternalName(external);
        if (!toBook || !range) return undefined;
        const clean: RangeRef = { ...range, external: undefined };
        return this.graph(toBook).traceRange(clean, depth, stack, this.hooksFor(toBook, stack));
      }
    };
  }

  /** Precedent trace that follows external references into loaded workbooks. */
  tracePrecedents(address: string, depth = 2): TraceNode {
    const { book, ref } = this.resolveAddress(address);
    const stack = new Set<string>([`${book.toLowerCase()}::${ref.sheet.toLowerCase()}|${ref.row},${ref.col}`]);
    return this.graph(book).traceFrom(ref, depth, stack, this.hooksFor(book, stack));
  }

  /** Dependent trace across workbooks (who reads this cell, anywhere). */
  traceDependents(address: string, depth = 2): TraceNode {
    const { book, ref } = this.resolveAddress(address);
    const stack = new Set<string>([`${book.toLowerCase()}::${ref.sheet.toLowerCase()}|${ref.row},${ref.col}`]);
    return this.dependentNode(book, ref, depth, stack);
  }

  private dependentNode(book: string, ref: CellRef, depth: number, stack: Set<string>): TraceNode {
    const g = this.graph(book);
    const cell = g.getCell(ref);
    const node: TraceNode = {
      address: this.q(book, formatCell(ref)),
      kind: 'cell',
      value: cell?.valueJson ?? null,
      children: []
    };
    if (cell?.formula) node.formula = cell.formula;

    const internal = g.dependentsOf(ref).map((addr) => ({ book, addr }));
    const cross = this.crossDependentsOf({ book, ref }).map((qualified) => {
      const { book: b, ref: r } = this.resolveAddress(qualified);
      return { book: b, addr: formatCell(r) };
    });
    const all = [...internal, ...cross];
    if (all.length === 0) return node;
    if (depth <= 0) {
      node.truncated = true;
      node.note = `${all.length} dependent(s) below depth limit`;
      return node;
    }
    for (const dep of all.slice(0, 30)) {
      const depGraph = this.graph(dep.book);
      const depRef = depGraph.resolveAddress(dep.addr);
      const key = `${dep.book.toLowerCase()}::${depRef.sheet.toLowerCase()}|${depRef.row},${depRef.col}`;
      if (stack.has(key)) {
        node.children.push({ address: this.q(dep.book, formatCell(depRef)), kind: 'cell', children: [], cycle: true });
        continue;
      }
      stack.add(key);
      node.children.push(this.dependentNode(dep.book, depRef, depth - 1, stack));
      stack.delete(key);
    }
    if (all.length > 30) {
      node.truncated = true;
      node.note = `${all.length - 30} more dependent(s) omitted`;
    }
    return node;
  }

  // ── cross-workbook context diffusion ─────────────────────────────────────

  expandContext(address: string, options?: DiffusionOptions): CollectionContextPacket {
    const { book, ref } = this.resolveAddress(address);
    const g = this.graph(book);
    const local = formatCell(ref);
    const budget = options?.tokenBudget ?? 2500;
    const maxRegions = options?.maxRegions ?? 8;
    const mode = options?.mode ?? 'compact';

    const base = g.expandContext(local, { ...options, tokenBudget: Math.floor(budget * 0.65) });
    const qualify = (s: string): string => (s.startsWith('[') ? s : this.q(book, s));

    const regions: RegionSummary[] = base.regions.map((r) => ({
      ...r,
      workbook: book,
      sourceCells: r.sourceCells.map(qualify)
    }));
    const relations: Relation[] = base.relations.map((rel) => ({
      ...rel,
      from: qualify(rel.from),
      to: qualify(rel.to)
    }));
    const warnings = [...base.warnings];
    const nextActions = [...base.nextActions];
    let tokens = base.tokens;
    let truncated = base.truncated;
    let remaining = budget - tokens;

    const seedRegion = g.regionAt(ref);
    const seen = new Set<string>();
    if (seedRegion) seen.add(`${book.toLowerCase()}::${seedRegion.id}`);
    for (const r of base.regions) seen.add(`${book.toLowerCase()}::${r.regionId}`);

    interface Candidate {
      book: string;
      region: Region;
      why: string;
      from: string;
      weight: number;
    }
    const candidates: Candidate[] = [];
    const addCandidate = (
      toBook: string,
      region: Region | undefined,
      why: string,
      from: string,
      weight: number
    ): void => {
      if (!region) return;
      const key = `${toBook.toLowerCase()}::${region.id}`;
      if (seen.has(key)) return;
      seen.add(key);
      candidates.push({ book: toBook, region, why, from, weight });
    };
    const regionsOverlappingIn = (toBook: string, range: RangeRef): Region[] => {
      try {
        return this.graph(toBook)
          .detectRegions(range.sheet)
          .filter((r) => rangesOverlap(r.data.range, range));
      } catch {
        return [];
      }
    };

    // 1. External references in the seed cell's own formula.
    const cell = g.getCell(ref);
    for (const fr of cell?.refs ?? []) {
      const ext = fr.external ?? fr.range?.external;
      if (!ext) continue;
      const toBook = this.resolveExternalName(ext);
      if (!toBook || !fr.range) continue; // unresolved externals are reported via `touching` below
      for (const region of regionsOverlappingIn(toBook, { ...fr.range, external: undefined })) {
        addCandidate(toBook, region, `referenced by ${this.q(book, local)} via ${fr.raw}`, this.q(book, local), 1.0);
      }
    }

    // 2. Cells in other books that depend on the seed.
    for (const qualified of this.crossDependentsOf({ book, ref }).slice(0, 10)) {
      const { book: b, ref: r } = this.resolveAddress(qualified);
      addCandidate(b, this.graph(b).regionAt(r), `${qualified} depends on the seed`, qualified, 0.9);
    }

    if (seedRegion) {
      // 3. External references made anywhere inside the seed's region.
      for (const link of this.index().links) {
        if (link.fromBook.toLowerCase() !== book.toLowerCase() || !link.toBook || !link.range) continue;
        let fromRef: CellRef;
        try {
          fromRef = g.resolveAddress(link.fromCell);
        } catch {
          continue;
        }
        if (!seedRegion.contains(fromRef)) continue;
        for (const region of regionsOverlappingIn(link.toBook, link.range)) {
          addCandidate(
            link.toBook,
            region,
            `formulas in the seed's region reference ${link.raw}`,
            this.q(book, seedRegion.id),
            0.85
          );
        }
      }

      // 4. Defined names shared with other workbooks, containing the seed.
      const seedNames = new Set(g.namesContaining(ref).map((n) => n.toLowerCase()));
      for (const sn of this.sharedNames()) {
        if (!seedNames.has(sn.name.toLowerCase())) continue;
        for (const other of sn.books) {
          if (other.book.toLowerCase() === book.toLowerCase()) continue;
          for (const rangeStr of other.ranges) {
            try {
              const range = parseRange(rangeStr);
              for (const region of regionsOverlappingIn(other.book, range)) {
                addCandidate(other.book, region, `shares defined name "${sn.name}" with the seed`, this.q(book, local), 0.7);
              }
            } catch {
              /* unparsable name target */
            }
          }
        }
      }

      // 5. Data links (shared key values) involving the seed's region.
      for (const dl of this.dataLinks()) {
        const ends = [dl.a, dl.b];
        const mine = ends.find(
          (e) => e.book.toLowerCase() === book.toLowerCase() && e.regionId === seedRegion.id
        );
        if (!mine) continue;
        const other = ends.find((e) => e !== mine)!;
        addCandidate(
          other.book,
          this.graph(other.book).getRegion(other.regionId),
          `shares ${dl.shared} key value(s) on ${mine.column}↔${other.column} (e.g. ${dl.sample.join(', ')})`,
          this.q(book, seedRegion.id),
          0.6
        );
      }
    }

    candidates.sort((x, y) => y.weight - x.weight);
    const crossModeCap: GranularityMode = MODE_RANK[mode] > 1 ? 'compact' : mode;
    const room = Math.max(2, maxRegions - base.regions.length);
    let added = 0;
    for (const candidate of candidates) {
      if (added >= room || remaining < 100) {
        truncated = true;
        nextActions.push(
          `summariseRegion('${candidate.region.id}') in ${candidate.book} — candidate skipped (${candidate.why})`
        );
        break;
      }
      const summary = this.summariseRegion(
        candidate.book,
        candidate.region,
        crossModeCap,
        Math.min(remaining, Math.max(120, Math.floor(budget / 4)))
      );
      if (summary.tokens > remaining) {
        truncated = true;
        continue;
      }
      remaining -= summary.tokens;
      tokens += summary.tokens;
      truncated = truncated || summary.truncated;
      added++;
      regions.push(summary);
      relations.push({
        from: candidate.from,
        to: this.q(candidate.book, candidate.region.id),
        type: candidate.weight >= 0.85 ? 'formula' : 'semantic',
        weight: candidate.weight,
        why: candidate.why
      });
    }

    // Trace: qualify the base trace; upgrade precedents to a cross-book
    // trace when the seed formula reaches another loaded workbook.
    let trace: ContextPacket['trace'] = base.trace
      ? {
          precedents: base.trace.precedents ? qualifyTree(base.trace.precedents, `[${book}]`) : undefined,
          dependents: base.trace.dependents ? qualifyTree(base.trace.dependents, `[${book}]`) : undefined
        }
      : undefined;
    const reachesLoadedBook = (cell?.refs ?? []).some((fr) => {
      const ext = fr.external ?? fr.range?.external;
      return ext !== undefined && this.resolveExternalName(ext) !== undefined;
    });
    if (reachesLoadedBook && (options?.includeTrace ?? true)) {
      const crossTrace = this.tracePrecedents(this.q(book, local), Math.min(2, options?.depth ?? 2));
      const cost = estimateTokens(crossTrace) - estimateTokens(trace?.precedents);
      if (cost <= remaining) {
        trace = { ...(trace ?? {}), precedents: crossTrace };
        remaining -= Math.max(0, cost);
        tokens += Math.max(0, cost);
      }
    }

    const crossDeps = this.crossDependentsOf({ book, ref });
    if (crossDeps.length > 0) {
      nextActions.push(
        `collection.traceDependents('${this.q(book, local)}') — ${crossDeps.length} cross-workbook dependent(s)`
      );
    }

    const touching = this.index().links.filter((link) => {
      if (link.fromBook.toLowerCase() === book.toLowerCase()) {
        if (link.fromCell === local) return true;
        if (!seedRegion) return false;
        try {
          return seedRegion.contains(g.resolveAddress(link.fromCell));
        } catch {
          return false;
        }
      }
      return link.toBook?.toLowerCase() === book.toLowerCase() && link.range !== undefined
        ? rangeContains(link.range, ref)
        : false;
    });
    for (const link of touching) {
      if (!link.toBook) {
        nextActions.push(
          `Load "${link.external}" into the collection to resolve ${link.raw} (from ${this.q(link.fromBook, link.fromCell)})`
        );
      }
    }

    return {
      ...base,
      seed: this.q(book, local),
      seedCell: { ...base.seedCell, address: this.q(book, base.seedCell.address) },
      regions,
      relations,
      trace,
      nextActions: uniq(nextActions).slice(0, 7),
      warnings: uniq(warnings),
      sourceCells: uniq([this.q(book, local), ...regions.flatMap((r) => r.sourceCells)]),
      tokens,
      truncated,
      workbooks: this.workbooks(),
      crossLinks: touching.slice(0, 12)
    };
  }
}

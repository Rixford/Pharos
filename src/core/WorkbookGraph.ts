/**
 * WorkbookGraph — the heart of Pharos.
 *
 * Parses a workbook (via a pluggable WorkbookParser), indexes every cell,
 * resolves formula references into precedent/dependent edges, registers
 * detected regions, and exposes the high-level operations: inspect,
 * detectRegions, summariseRegion, tracePrecedents/Dependents, findValue
 * and expandContext.
 */
import * as fs from 'fs';
import {
  cellKey,
  clampRange,
  colToLetter,
  formatCell,
  formatRange,
  iterateRange,
  parseCellAddress,
  parseRange,
  rangeArea,
  rangeContains
} from './address';
import { CellNode } from './CellNode';
import { Region } from './Region';
import { CellLookup, SheetGrid } from './grid';
import {
  CellRef,
  CellStyle,
  CellValueType,
  RegionKind,
  ContextPacket,
  DiffusionOptions,
  FormulaRef,
  GranularityMode,
  JsonScalar,
  RangeRef,
  RefKind,
  RegionBrief,
  RegionSummary,
  SheetInfo,
  TraceNode,
  WorkbookOverview
} from './types';
import { uniq } from './util';
import { canonicalizeFormula, extractRefs } from '../parser/FormulaParser';
import { ExcelParser } from '../parser/ExcelParser';
import { ParsedWorkbook, WorkbookParser } from '../parser/types';
import { detectRegions as detectRegionsInGrid } from '../analysis/RegionDetector';
import { ExtractOptions, ExtractedTable, extractTable as extractTableImpl } from '../analysis/Extractor';
import { LocateHit, LocateOptions, locateRegions } from '../analysis/Locate';
import { summariseRegion as summariseRegionImpl } from '../analysis/Summariser';
import { expandContext as expandContextImpl } from '../analysis/Diffuser';

/** Ranges up to this many cells are expanded into per-cell dependent edges. */
const RANGE_EXPAND_CAP = 512;
/** Cells scanned when characterising a range node in a trace. */
const TRACE_RANGE_SCAN_CAP = 2000;
const MAX_TRACE_CHILDREN = 30;

interface SheetData {
  name: string;
  index: number;
  hidden: boolean;
  maxRow: number;
  maxCol: number;
  cells: Map<string, CellNode>;
  merges: RangeRef[];
  formulaCellCount: number;
}

export interface DefinedName {
  name: string;
  ranges: string[];
  resolved: RangeRef[];
}

export interface CellInspection {
  address: string;
  exists: boolean;
  sheetHidden: boolean;
  value: JsonScalar;
  type: CellValueType;
  formula?: string;
  style?: CellStyle;
  hyperlink?: string;
  merged?: { range: string; master: string; masterValue: JsonScalar };
  region?: RegionBrief & { title?: string };
  precedents: { raw: string; kind: RefKind; target?: string; external?: string }[];
  dependents: string[];
  dependentsTotal: number;
  namedRanges: string[];
}

export interface FindHit {
  address: string;
  value: JsonScalar;
  type: CellValueType;
  regionId?: string;
}

export interface SheetMapRegion {
  id: string;
  rangeA1: string;
  kind: RegionKind;
  purpose?: string;
  title?: string;
  headers?: string[];
  rows: number;
  cols: number;
  sections?: number;
  hasNotes: boolean;
  confidence: number;
}

/** Zoom level 1 — a sheet's layout at a glance. */
export interface SheetMap {
  sheet: string;
  hidden: boolean;
  usedRangeA1?: string;
  purpose?: string;
  regions: SheetMapRegion[];
  notes: string[];
}

export interface LoadOptions {
  /** Override the parser (e.g. a CSV adapter). Defaults to ExcelParser. */
  parser?: WorkbookParser;
}

/**
 * Optional hooks for precedent traces. Used by Collections to qualify
 * addresses with a workbook name, share one cycle stack across workbooks,
 * and resolve external references into other loaded workbooks. With no
 * hooks, traces behave exactly as in v0.1.x.
 */
export interface TraceHooks {
  /** Rewrite addresses placed in trace nodes (e.g. prefix `[Book.xlsx]`). */
  qualify?: (address: string) => string;
  /** Prefix for cycle-stack keys; must be unique per workbook in a collection. */
  stackKeyPrefix?: string;
  /**
   * Resolve an external-workbook reference into a subtree. Return undefined
   * to fall back to the default stub node.
   */
  resolveExternal?: (
    external: string,
    range: RangeRef | undefined,
    raw: string,
    depth: number,
    stack: Set<string>
  ) => TraceNode | undefined;
}

export class WorkbookGraph {
  private readonly sheetStore = new Map<string, SheetData>();
  private readonly order: string[] = [];
  readonly definedNames: DefinedName[] = [];
  readonly warnings: string[] = [];
  readonly externalRefs: string[] = [];
  private readonly warned = new Set<string>();
  private readonly regionsCache = new Map<string, Region[]>();
  private readonly regionIndex = new Map<string, Region>();
  private deps?: {
    direct: Map<string, string[]>;
    ranged: Map<string, { range: RangeRef; dep: string }[]>;
  };

  /** Load a workbook from a file path or an in-memory buffer. */
  static async load(input: string | Buffer, options?: LoadOptions): Promise<WorkbookGraph> {
    const buffer = typeof input === 'string' ? await fs.promises.readFile(input) : input;
    const parser = options?.parser ?? new ExcelParser();
    const parsed = await parser.parse(buffer);
    return new WorkbookGraph(parsed);
  }

  private constructor(parsed: ParsedWorkbook) {
    for (const w of parsed.warnings) this.warn(w, w);

    for (const dn of parsed.definedNames) {
      const resolved: RangeRef[] = [];
      for (const r of dn.ranges) {
        try {
          resolved.push(parseRange(r));
        } catch {
          this.warn(`name:${dn.name}:${r}`, `Defined name "${dn.name}": unparsable target "${r}"`);
        }
      }
      this.definedNames.push({ name: dn.name, ranges: dn.ranges, resolved });
    }
    const nameLookup = new Map<string, DefinedName>();
    for (const dn of this.definedNames) nameLookup.set(dn.name.toLowerCase(), dn);

    const externals = new Set<string>();
    const sheetNames = new Set(parsed.sheets.map((s) => s.name.toLowerCase()));

    for (const ps of parsed.sheets) {
      const cells = new Map<string, CellNode>();
      let formulaCellCount = 0;
      for (const [key, pc] of ps.cells) {
        let refs: FormulaRef[] | undefined;
        if (pc.formula) {
          formulaCellCount++;
          const addr = formatCell({ sheet: ps.name, row: pc.row, col: pc.col });
          const result = extractRefs(pc.formula, ps.name);
          refs = result.refs;
          for (const w of result.warnings) this.warn(w, `${addr}: ${w}`);
          for (const ref of refs) {
            if (ref.kind === 'name' && ref.name) {
              const dn = nameLookup.get(ref.name.toLowerCase());
              if (dn) ref.resolved = dn.resolved;
              else {
                this.warn(
                  `undefined-name:${ref.name.toLowerCase()}`,
                  `Formula(s) reference undefined name "${ref.name}" (first seen at ${addr})`
                );
              }
            }
            const ext = ref.external ?? ref.range?.external;
            if (ext) externals.add(ext);
            else if (ref.range && !sheetNames.has(ref.range.sheet.toLowerCase())) {
              this.warn(
                `missing-sheet:${ref.range.sheet.toLowerCase()}`,
                `Formula(s) reference missing sheet "${ref.range.sheet}" (first seen at ${addr})`
              );
            }
          }
        }
        cells.set(
          key,
          new CellNode({
            sheet: ps.name,
            row: pc.row,
            col: pc.col,
            value: pc.value,
            type: pc.type,
            formula: pc.formula,
            style: pc.style,
            hyperlink: pc.hyperlink,
            refs
          })
        );
      }
      this.sheetStore.set(ps.name.toLowerCase(), {
        name: ps.name,
        index: ps.index,
        hidden: ps.hidden,
        maxRow: ps.maxRow,
        maxCol: ps.maxCol,
        cells,
        merges: ps.merges,
        formulaCellCount
      });
      this.order.push(ps.name);
    }
    this.externalRefs.push(...externals);
    if (externals.size > 0) {
      this.warn(
        'external-refs',
        `Workbook references ${externals.size} external workbook(s): ${[...externals].join(', ')} — external values are not loaded`
      );
    }
  }

  private warn(key: string, msg: string): void {
    if (this.warned.has(key) || this.warnings.length >= 200) return;
    this.warned.add(key);
    this.warnings.push(msg);
  }

  // ── Sheets & cells ────────────────────────────────────────────────────────

  /** First visible sheet (used to resolve unqualified addresses). */
  get defaultSheet(): string {
    for (const name of this.order) {
      if (!this.sheetStore.get(name.toLowerCase())!.hidden) return name;
    }
    return this.order[0];
  }

  sheetNames(): string[] {
    return [...this.order];
  }

  isSheetHidden(name: string): boolean {
    return this.sheet(name)?.hidden ?? false;
  }

  private sheet(name: string): SheetData | undefined {
    return this.sheetStore.get(name.toLowerCase());
  }

  /** Parse + validate an address against this workbook. Sheet names are case-insensitive. */
  resolveAddress(address: string): CellRef {
    const ref = parseCellAddress(address, this.defaultSheet);
    const sd = this.sheet(ref.sheet);
    if (!sd) {
      throw new Error(`Sheet "${ref.sheet}" not found. Sheets: ${this.order.join(', ')}`);
    }
    return { ...ref, sheet: sd.name };
  }

  getCell(target: string | CellRef): CellNode | undefined {
    const ref = typeof target === 'string' ? this.resolveAddress(target) : target;
    return this.sheet(ref.sheet)?.cells.get(cellKey(ref.row, ref.col));
  }

  /** Iterate every stored (non-empty) cell, optionally for one sheet. */
  *cells(sheetName?: string): Generator<CellNode> {
    const names = sheetName === undefined ? this.order : [sheetName];
    for (const name of names) {
      const sd = this.sheet(name);
      if (!sd) throw new Error(`Sheet "${name}" not found. Sheets: ${this.order.join(', ')}`);
      yield* sd.cells.values();
    }
  }

  sheets(): SheetInfo[] {
    return this.order.map((name, i) => {
      const sd = this.sheet(name)!;
      return {
        name: sd.name,
        index: i,
        hidden: sd.hidden,
        rowCount: sd.maxRow,
        colCount: sd.maxCol,
        cellCount: sd.cells.size,
        formulaCellCount: sd.formulaCellCount,
        usedRangeA1: sd.cells.size > 0 ? `A1:${colToLetter(sd.maxCol)}${sd.maxRow}` : undefined
      };
    });
  }

  /** Read-only grid view of a sheet (used by the analysis modules). */
  grid(name: string): SheetGrid | undefined {
    const sd = this.sheet(name);
    if (!sd) return undefined;
    return {
      name: sd.name,
      hidden: sd.hidden,
      maxRow: sd.maxRow,
      maxCol: sd.maxCol,
      cells: sd.cells,
      merges: sd.merges
    };
  }

  private readonly cellLookup: CellLookup = (sheet, row, col) =>
    this.sheet(sheet)?.cells.get(cellKey(row, col));

  // ── Regions ───────────────────────────────────────────────────────────────

  /** Detect (and cache) regions for one sheet, or all sheets when omitted. */
  detectRegions(sheetName?: string): Region[] {
    if (sheetName === undefined) {
      return this.order.flatMap((s) => this.detectRegions(s));
    }
    const sd = this.sheet(sheetName);
    if (!sd) throw new Error(`Sheet "${sheetName}" not found. Sheets: ${this.order.join(', ')}`);
    const key = sd.name.toLowerCase();
    const cached = this.regionsCache.get(key);
    if (cached) return cached;
    const regions = detectRegionsInGrid(this.grid(sd.name)!).map((d) => new Region(d));
    this.regionsCache.set(key, regions);
    for (const region of regions) {
      this.regionIndex.set(region.id, region);
      for (const cell of sd.cells.values()) {
        if (cell.regionId === undefined && region.contains(cell.ref)) cell.regionId = region.id;
      }
    }
    return regions;
  }

  allRegions(): Region[] {
    return this.detectRegions();
  }

  regionAt(target: string | CellRef): Region | undefined {
    const ref = typeof target === 'string' ? this.resolveAddress(target) : target;
    return this.detectRegions(ref.sheet).find((r) => r.contains(ref));
  }

  /** Look up a region by its stable id (rg_…). */
  getRegion(id: string): Region | undefined {
    if (!this.regionIndex.has(id)) this.detectRegions();
    return this.regionIndex.get(id);
  }

  // ── Summarisation & diffusion ────────────────────────────────────────────

  /**
   * Summarise a region at the requested granularity. `target` may be a
   * Region, a region id (`rg_…`) or any cell address inside the region.
   */
  summariseRegion(
    target: string | Region,
    mode: GranularityMode = 'summary',
    tokenBudget?: number
  ): RegionSummary {
    const region = this.toRegion(target);
    return summariseRegionImpl(region.data, this.cellLookup, mode, tokenBudget);
  }

  private toRegion(target: string | Region): Region {
    if (target instanceof Region) return target;
    if (target.startsWith('rg_')) {
      const byId = this.getRegion(target);
      if (byId) return byId;
      throw new Error(`Region id "${target}" not found`);
    }
    const region = this.regionAt(target);
    if (!region) {
      throw new Error(`No region found at "${target}" — the cell sits outside any detected region`);
    }
    return region;
  }

  /** Diffuse context outward from a seed cell. See DiffusionOptions. */
  expandContext(address: string, options?: DiffusionOptions): ContextPacket {
    return expandContextImpl(this, address, options);
  }

  /**
   * Zoom level 6: transformation-ready extraction. Typed rows keyed by
   * header names, subtotal rows excluded (no double counting), per-row
   * provenance, deterministic offset/limit paging.
   */
  extractTable(target: string | Region, opts: ExtractOptions = {}): ExtractedTable {
    const region = this.toRegion(target);
    return extractTableImpl(region.data, this.cellLookup, opts);
  }

  /** Question-aware narrowing: rank this workbook's regions against a question. */
  locate(question: string, opts?: LocateOptions): LocateHit[] {
    return locateRegions(this.allRegions().map((r) => ({ region: r.data })), question, opts);
  }

  /** Zoom level 1: a sheet's region inventory with purposes and notes. */
  sheetMap(sheetName: string): SheetMap {
    const sd = this.sheet(sheetName);
    if (!sd) throw new Error(`Sheet "${sheetName}" not found. Sheets: ${this.order.join(', ')}`);
    const regions = this.detectRegions(sd.name);
    const notes: string[] = [];
    for (const region of regions.filter((r) => r.kind === 'notes')) {
      for (let r = region.data.range.startRow; r <= region.data.range.endRow; r++) {
        for (let c = region.data.range.startCol; c <= region.data.range.endCol; c++) {
          const cell = sd.cells.get(cellKey(r, c));
          if (cell && typeof cell.value === 'string') notes.push(cell.value);
        }
      }
    }
    const purposes = [...new Set(regions.filter((r) => r.kind !== 'notes' && r.data.purpose).map((r) => r.data.purpose!))];
    return {
      sheet: sd.name,
      hidden: sd.hidden,
      usedRangeA1: sd.cells.size > 0 ? `A1:${colToLetter(sd.maxCol)}${sd.maxRow}` : undefined,
      purpose: purposes.length > 0 ? purposes.slice(0, 3).join(' · ') : undefined,
      regions: regions.map((r) => ({
        id: r.id,
        rangeA1: r.rangeA1,
        kind: r.kind,
        purpose: r.data.purpose,
        title: r.title,
        headers: r.headers?.slice(0, 10),
        rows: r.data.dataRowCount,
        cols: r.data.colCount,
        sections: r.data.sections?.filter((x) => x.kind === 'group').length,
        hasNotes: (r.data.notes?.length ?? 0) > 0,
        confidence: r.confidence
      })),
      notes
    };
  }

  /** Formula targets (resolved ranges) referenced from inside a region. */
  regionFormulaTargets(region: Region, cap = 500): RangeRef[] {
    const sd = this.sheet(region.sheet);
    if (!sd) return [];
    const out: RangeRef[] = [];
    let scanned = 0;
    for (const cell of sd.cells.values()) {
      if (scanned >= cap) break;
      if (cell.refs.length === 0 || !region.contains(cell.ref)) continue;
      scanned++;
      out.push(...this.targetRanges(cell.refs));
    }
    return out;
  }

  // ── Precedents & dependents ──────────────────────────────────────────────

  private addrKey(ref: CellRef): string {
    return `${ref.sheet.toLowerCase()}|${cellKey(ref.row, ref.col)}`;
  }

  private targetRanges(refs: FormulaRef[]): RangeRef[] {
    const out: RangeRef[] = [];
    for (const ref of refs) {
      if (ref.external || ref.range?.external) continue;
      if (ref.range) out.push(ref.range);
      else if (ref.resolved) out.push(...ref.resolved.filter((r) => !r.external));
    }
    return out;
  }

  private buildDeps(): void {
    const direct = new Map<string, string[]>();
    const ranged = new Map<string, { range: RangeRef; dep: string }[]>();
    for (const sd of this.sheetStore.values()) {
      for (const cell of sd.cells.values()) {
        if (cell.refs.length === 0) continue;
        for (const range of this.targetRanges(cell.refs)) {
          const target = this.sheet(range.sheet);
          if (!target) continue;
          const clamped = range.open ? clampRange(range, target.maxRow, target.maxCol) : range;
          if (rangeArea(clamped) <= RANGE_EXPAND_CAP) {
            for (const ref of iterateRange(clamped)) {
              const key = this.addrKey({ ...ref, sheet: target.name });
              const list = direct.get(key);
              if (list) list.push(cell.address);
              else direct.set(key, [cell.address]);
            }
          } else {
            const key = target.name.toLowerCase();
            const entry = { range: clamped, dep: cell.address };
            const list = ranged.get(key);
            if (list) list.push(entry);
            else ranged.set(key, [entry]);
          }
        }
      }
    }
    this.deps = { direct, ranged };
  }

  /** Addresses of cells whose formulas reference the given cell. */
  dependentsOf(target: string | CellRef): string[] {
    const ref = typeof target === 'string' ? this.resolveAddress(target) : target;
    if (!this.deps) this.buildDeps();
    const { direct, ranged } = this.deps!;
    const out = [...(direct.get(this.addrKey(ref)) ?? [])];
    for (const entry of ranged.get(ref.sheet.toLowerCase()) ?? []) {
      if (rangeContains(entry.range, ref)) out.push(entry.dep);
    }
    return uniq(out);
  }

  /** Apply the optional address qualifier from trace hooks. */
  private static hq(hooks: TraceHooks | undefined, s: string): string {
    return hooks?.qualify ? hooks.qualify(s) : s;
  }

  private hkey(hooks: TraceHooks | undefined, ref: CellRef): string {
    return (hooks?.stackKeyPrefix ?? '') + this.addrKey(ref);
  }

  /**
   * Recursively trace what a cell's formula depends on. `hooks` (optional)
   * let a Collection qualify addresses, share a cycle stack across
   * workbooks and resolve external references — see TraceHooks. Without
   * hooks, behavior is identical to v0.1.x.
   */
  tracePrecedents(address: string, depth = 2, hooks?: TraceHooks): TraceNode {
    const ref = this.resolveAddress(address);
    return this.precedentNode(ref, depth, new Set([this.hkey(hooks, ref)]), hooks);
  }

  /** Advanced: precedent trace continuing an existing cycle stack (Collections). */
  traceFrom(ref: CellRef, depth: number, stack: Set<string>, hooks?: TraceHooks): TraceNode {
    return this.precedentNode(ref, depth, stack, hooks);
  }

  /** Advanced: trace into a range target with an existing cycle stack (Collections). */
  traceRange(range: RangeRef, depth: number, stack: Set<string>, hooks?: TraceHooks): TraceNode {
    return this.rangeNode(range, depth, stack, hooks);
  }

  private precedentNode(
    ref: CellRef,
    depth: number,
    stack: Set<string>,
    hooks?: TraceHooks
  ): TraceNode {
    const cell = this.getCell(ref);
    const node: TraceNode = {
      address: WorkbookGraph.hq(hooks, formatCell(ref)),
      kind: 'cell',
      value: cell?.valueJson ?? null,
      children: []
    };
    if (cell?.formula) node.formula = cell.formula;
    if (!cell || cell.refs.length === 0) return node;
    if (depth <= 0) {
      node.truncated = true;
      node.note = `${cell.refs.length} reference(s) below depth limit`;
      return node;
    }
    for (const fr of cell.refs) {
      const ext = fr.external ?? fr.range?.external;
      if (ext) {
        const resolved = hooks?.resolveExternal?.(ext, fr.range, fr.raw, depth - 1, stack);
        if (resolved) {
          node.children.push(resolved);
          continue;
        }
        node.children.push({
          address: fr.raw,
          kind: 'external',
          children: [],
          note: `external workbook "${ext}" — not loaded`
        });
        continue;
      }
      if (fr.kind === 'structured') {
        node.children.push({
          address: fr.raw,
          kind: 'name',
          children: [],
          note: 'structured reference (not resolved)'
        });
        continue;
      }
      if (fr.kind === 'name') {
        const child: TraceNode = {
          address: fr.name ?? fr.raw,
          kind: 'name',
          children: [],
          note: fr.resolved?.length
            ? `defined name → ${fr.resolved.map((r) => WorkbookGraph.hq(hooks, formatRange(r))).join(', ')}`
            : 'undefined name'
        };
        for (const r of fr.resolved ?? []) {
          child.children.push(this.rangeNode(r, depth - 1, stack, hooks));
        }
        node.children.push(child);
        continue;
      }
      if (fr.range) node.children.push(this.rangeNode(fr.range, depth - 1, stack, hooks));
    }
    return node;
  }

  private rangeNode(
    range: RangeRef,
    depth: number,
    stack: Set<string>,
    hooks?: TraceHooks
  ): TraceNode {
    const sd = this.sheet(range.sheet);
    const clamped = sd && range.open ? clampRange(range, sd.maxRow, sd.maxCol) : range;
    if (rangeArea(clamped) === 1) {
      const ref: CellRef = { sheet: sd?.name ?? clamped.sheet, row: clamped.startRow, col: clamped.startCol };
      const key = this.hkey(hooks, ref);
      if (stack.has(key)) {
        return { address: WorkbookGraph.hq(hooks, formatCell(ref)), kind: 'cell', children: [], cycle: true };
      }
      stack.add(key);
      const node = this.precedentNode(ref, depth, stack, hooks);
      stack.delete(key);
      return node;
    }
    const node: TraceNode = {
      address: WorkbookGraph.hq(hooks, formatRange(clamped)),
      kind: 'range',
      cellCount: rangeArea(clamped),
      children: []
    };
    if (!sd) {
      node.note = `sheet "${range.sheet}" not found`;
      return node;
    }
    let nonEmpty = 0;
    const formulaCells: CellNode[] = [];
    for (const cr of iterateRange(clamped, TRACE_RANGE_SCAN_CAP)) {
      const c = sd.cells.get(cellKey(cr.row, cr.col));
      if (c) {
        nonEmpty++;
        if (c.formula) formulaCells.push(c);
      }
    }
    if (formulaCells.length === 0) {
      node.note = `${nonEmpty} value cell(s)`;
      return node;
    }
    const templates = new Map<string, { count: number; first: CellNode }>();
    for (const c of formulaCells.slice(0, 500)) {
      const t = canonicalizeFormula(c.formula!, c.sheet, c.row, c.col);
      const entry = templates.get(t);
      if (entry) entry.count++;
      else templates.set(t, { count: 1, first: c });
    }
    const top = [...templates.values()].sort((a, b) => b.count - a.count)[0];
    node.note = `${nonEmpty} cells, ${formulaCells.length} formula(s); dominant pattern =${top.first.formula} (${top.count}×)`;
    if (depth > 0) {
      const key = this.hkey(hooks, top.first.ref);
      if (!stack.has(key)) {
        stack.add(key);
        const rep = this.precedentNode(top.first.ref, depth - 1, stack, hooks);
        rep.note = `representative of ${top.count} cell(s) with this pattern`;
        node.children.push(rep);
        stack.delete(key);
      }
    } else {
      node.truncated = true;
    }
    return node;
  }

  /** Recursively trace which cells depend on the given cell. */
  traceDependents(address: string, depth = 2): TraceNode {
    const ref = this.resolveAddress(address);
    return this.dependentNode(ref, depth, new Set([this.addrKey(ref)]));
  }

  private dependentNode(ref: CellRef, depth: number, stack: Set<string>): TraceNode {
    const cell = this.getCell(ref);
    const node: TraceNode = {
      address: formatCell(ref),
      kind: 'cell',
      value: cell?.valueJson ?? null,
      children: []
    };
    if (cell?.formula) node.formula = cell.formula;
    const deps = this.dependentsOf(ref);
    if (deps.length === 0) return node;
    if (depth <= 0) {
      node.truncated = true;
      node.note = `${deps.length} dependent(s) below depth limit`;
      return node;
    }
    for (const addr of deps.slice(0, MAX_TRACE_CHILDREN)) {
      const childRef = this.resolveAddress(addr);
      const key = this.addrKey(childRef);
      if (stack.has(key)) {
        node.children.push({ address: addr, kind: 'cell', children: [], cycle: true });
        continue;
      }
      stack.add(key);
      node.children.push(this.dependentNode(childRef, depth - 1, stack));
      stack.delete(key);
    }
    if (deps.length > MAX_TRACE_CHILDREN) {
      node.truncated = true;
      node.note = `${deps.length - MAX_TRACE_CHILDREN} more dependent(s) omitted`;
    }
    return node;
  }

  // ── Inspection & search ──────────────────────────────────────────────────

  inspect(address: string): CellInspection {
    const ref = this.resolveAddress(address);
    const sd = this.sheet(ref.sheet)!;
    const cell = this.getCell(ref);
    const region = this.regionAt(ref);
    const deps = this.dependentsOf(ref);

    let merged: CellInspection['merged'];
    for (const m of sd.merges) {
      if (rangeContains(m, ref)) {
        const masterRef: CellRef = { sheet: sd.name, row: m.startRow, col: m.startCol };
        merged = {
          range: formatRange(m),
          master: formatCell(masterRef),
          masterValue: this.getCell(masterRef)?.valueJson ?? null
        };
        break;
      }
    }

    return {
      address: formatCell(ref),
      exists: cell !== undefined,
      sheetHidden: sd.hidden,
      value: cell?.valueJson ?? null,
      type: cell?.type ?? 'empty',
      formula: cell?.formula,
      style: cell?.style,
      hyperlink: cell?.hyperlink,
      merged,
      region: region ? { ...region.brief(), title: region.title } : undefined,
      precedents: (cell?.refs ?? []).map((r) => ({
        raw: r.raw,
        kind: r.kind,
        target: r.range
          ? formatRange(r.range)
          : r.resolved?.length
            ? r.resolved.map((x) => formatRange(x)).join(', ')
            : undefined,
        external: r.external ?? r.range?.external
      })),
      dependents: deps.slice(0, 20),
      dependentsTotal: deps.length,
      namedRanges: this.namesContaining(ref)
    };
  }

  /** Defined names whose target ranges contain the given cell. */
  namesContaining(ref: CellRef): string[] {
    return this.definedNames
      .filter((dn) => dn.resolved.some((r) => rangeContains(r, ref)))
      .map((dn) => dn.name);
  }

  /**
   * Search cell values. Strings match case-insensitively as substrings,
   * numbers match exactly, RegExp matches the string form of the value.
   */
  findValue(
    query: string | number | RegExp,
    opts?: { sheet?: string; limit?: number }
  ): FindHit[] {
    const limit = opts?.limit ?? 50;
    const sheetsToSearch = opts?.sheet ? [opts.sheet] : this.order;
    const out: FindHit[] = [];
    for (const name of sheetsToSearch) {
      const sd = this.sheet(name);
      if (!sd) throw new Error(`Sheet "${name}" not found. Sheets: ${this.order.join(', ')}`);
      for (const cell of sd.cells.values()) {
        if (out.length >= limit) return out;
        if (this.valueMatches(cell, query)) {
          out.push({ address: cell.address, value: cell.valueJson, type: cell.type, regionId: cell.regionId });
        }
      }
    }
    return out;
  }

  private valueMatches(cell: CellNode, q: string | number | RegExp): boolean {
    if (q instanceof RegExp) return q.test(String(cell.valueJson ?? ''));
    if (typeof q === 'number') {
      return cell.type === 'number' && Math.abs((cell.value as number) - q) < 1e-9;
    }
    return String(cell.valueJson ?? '')
      .toLowerCase()
      .includes(q.toLowerCase());
  }

  /** Workbook-level structural overview (sheets, regions, names, warnings). */
  overview(): WorkbookOverview {
    const sheets = this.sheets();
    const regionsBySheet: Record<string, RegionBrief[]> = {};
    for (const name of this.order) {
      regionsBySheet[name] = this.detectRegions(name).map((r) => r.brief());
    }
    return {
      sheets,
      definedNames: this.definedNames.map((d) => ({ name: d.name, ranges: d.ranges })),
      totalCells: sheets.reduce((a, s) => a + s.cellCount, 0),
      totalFormulaCells: sheets.reduce((a, s) => a + s.formulaCellCount, 0),
      externalRefs: [...this.externalRefs],
      regionsBySheet,
      warnings: [...this.warnings]
    };
  }
}

/**
 * Diffuser — expands context outward from a seed cell along weighted edges.
 *
 * Edge families (weights configurable):
 *   structural · seed → the region that contains it
 *   formula    · precedent/dependent links, region-to-region references
 *   semantic   · shared named ranges, overlapping column headers
 *   spatial    · physically adjacent regions on the same sheet
 *   sheet      · other regions on the seed's sheet
 *
 * Stopping rules: hop depth, per-edge weight decay with a minimum weight,
 * a region cap, and a global token budget. Deeper regions are summarised
 * at coarser granularity. The resulting ContextPacket always reports what
 * was left out (truncated flag + nextActions) so a consumer can decide
 * whether to spend more tokens.
 */
import { formatCell, formatRange, rangesOverlap } from '../core/address';
import { Region } from '../core/Region';
import {
  ContextPacket,
  DEFAULT_EDGE_WEIGHTS,
  DiffusionOptions,
  EdgeType,
  GranularityMode,
  RangeRef,
  RegionSummary,
  Relation,
  ResolvedDiffusionOptions,
  SeedSummary
} from '../core/types';
import { uniq } from '../core/util';
import { estimateTokens } from './Summariser';
import type { WorkbookGraph } from '../core/WorkbookGraph';

const MODE_RANK: Record<GranularityMode, number> = {
  summary: 0,
  compact: 1,
  evidence: 2,
  cells: 3,
  formulas: 3,
  audit: 4
};

/** Deeper hops get coarser summaries. */
function modeForDepth(requested: GranularityMode, depth: number): GranularityMode {
  if (depth <= 0) return requested;
  if (depth === 1) return MODE_RANK[requested] > 1 ? 'compact' : requested;
  return 'summary';
}

const round2 = (n: number): number => Math.round(n * 100) / 100;

interface Frontier {
  region: Region;
  depth: number;
  weight: number;
  type: EdgeType;
  why: string;
  from: string;
}

export function expandContext(
  graph: WorkbookGraph,
  seedAddress: string,
  options?: DiffusionOptions
): ContextPacket {
  const opts: ResolvedDiffusionOptions = {
    depth: options?.depth ?? 2,
    mode: options?.mode ?? 'compact',
    tokenBudget: options?.tokenBudget ?? 2000,
    maxRegions: options?.maxRegions ?? 8,
    minWeight: options?.minWeight ?? 0.15,
    decay: options?.decay ?? 0.75,
    weights: { ...DEFAULT_EDGE_WEIGHTS, ...options?.weights },
    includeTrace: options?.includeTrace ?? true
  };

  const warnings: string[] = [];
  const nextActions: string[] = [];
  const relations: Relation[] = [];
  const regions: RegionSummary[] = [];
  let truncated = false;

  const ref = graph.resolveAddress(seedAddress);
  const seed = formatCell(ref);
  const cell = graph.getCell(ref);
  const seedRegion = graph.regionAt(ref);

  if (!cell) warnings.push(`Seed cell ${seed} is empty`);
  if (graph.isSheetHidden(ref.sheet)) warnings.push(`Seed sheet "${ref.sheet}" is hidden`);
  for (const fr of cell?.refs ?? []) {
    const ext = fr.external ?? fr.range?.external;
    if (ext) warnings.push(`Seed formula references external workbook "${ext}" — values not loaded`);
    if (fr.kind === 'structured') warnings.push(`Seed formula uses a structured reference (${fr.raw}) that was not resolved`);
  }

  const seedCell: SeedSummary = {
    address: seed,
    value: cell?.valueJson ?? null,
    type: cell?.type ?? 'empty',
    formula: cell?.formula,
    style: cell?.style,
    regionId: seedRegion?.id,
    namedRanges: graph.namesContaining(ref)
  };

  let budget = opts.tokenBudget - estimateTokens(seedCell) - 60; // packet skeleton allowance

  // ── trace excerpts ───────────────────────────────────────────────────────
  let trace: ContextPacket['trace'];
  if (opts.includeTrace && cell?.formula && opts.depth > 0) {
    const prec = graph.tracePrecedents(seed, Math.min(2, opts.depth));
    const cost = estimateTokens(prec);
    if (cost <= budget * 0.4) {
      trace = { precedents: prec };
      budget -= cost;
    } else {
      truncated = true;
      nextActions.push(`tracePrecedents('${seed}') — trace omitted here to fit the token budget`);
    }
  }
  if (opts.includeTrace && cell && opts.depth > 0) {
    const depCount = graph.dependentsOf(ref).length;
    if (depCount > 0) {
      const dep = graph.traceDependents(seed, 1);
      const cost = estimateTokens(dep);
      if (cost <= budget * 0.25) {
        trace = { ...(trace ?? {}), dependents: dep };
        budget -= cost;
      } else {
        nextActions.push(`traceDependents('${seed}') to see ${depCount} dependent cell(s)`);
      }
    }
  }

  // ── frontier ─────────────────────────────────────────────────────────────
  const queue: Frontier[] = [];
  const seen = new Set<string>();

  const push = (
    region: Region | undefined,
    depth: number,
    weight: number,
    type: EdgeType,
    why: string,
    from: string
  ): void => {
    if (!region || seen.has(region.id)) return;
    if (depth > opts.depth || weight < opts.minWeight) return;
    const existing = queue.find((q) => q.region.id === region.id);
    if (existing) {
      if (weight > existing.weight) {
        existing.weight = weight;
        existing.depth = Math.min(existing.depth, depth);
        existing.type = type;
        existing.why = why;
        existing.from = from;
      }
      return;
    }
    queue.push({ region, depth, weight, type, why, from });
  };

  const regionsOverlapping = (range: RangeRef): Region[] => {
    try {
      return graph.detectRegions(range.sheet).filter((r) => rangesOverlap(r.data.range, range));
    } catch {
      return [];
    }
  };

  // Seed region (structural).
  if (seedRegion) {
    push(seedRegion, 0, opts.weights.structural, 'structural', 'contains the seed cell', seed);
  } else {
    // Empty/orphan seed: fall back to the nearest region on the sheet.
    const sheetRegions = graph.detectRegions(ref.sheet);
    const nearest = [...sheetRegions].sort((a, b) => {
      const da = Math.abs(a.data.range.startRow - ref.row) + Math.abs(a.data.range.startCol - ref.col);
      const db = Math.abs(b.data.range.startRow - ref.row) + Math.abs(b.data.range.startCol - ref.col);
      return da - db;
    })[0];
    if (nearest) {
      push(nearest, 0, opts.weights.spatial, 'spatial', 'nearest region to the (empty) seed cell', seed);
      nextActions.push(`Seed ${seed} is outside any region — nearest is ${nearest.id} at ${nearest.rangeA1}`);
    }
  }

  // Precedent regions (formula edges).
  if (cell && opts.depth >= 1) {
    for (const fr of cell.refs) {
      if (fr.external || fr.range?.external) continue;
      const targets = fr.range ? [fr.range] : (fr.resolved ?? []);
      for (const range of targets) {
        for (const region of regionsOverlapping(range)) {
          push(
            region,
            1,
            opts.weights.formula * opts.decay,
            'formula',
            `referenced by ${seed} via ${fr.raw}`,
            seed
          );
        }
      }
    }
    for (const depAddr of graph.dependentsOf(ref).slice(0, 25)) {
      const depRegion = graph.regionAt(depAddr);
      push(
        depRegion,
        1,
        opts.weights.formula * 0.9 * opts.decay,
        'formula',
        `${depAddr} depends on ${seed}`,
        seed
      );
    }
  }

  // Named ranges containing the seed (semantic edges).
  if (opts.depth >= 1) {
    for (const dn of graph.definedNames) {
      if (!seedCell.namedRanges.includes(dn.name)) continue;
      for (const range of dn.resolved) {
        for (const region of regionsOverlapping(range)) {
          push(
            region,
            1,
            opts.weights.semantic * opts.decay,
            'semantic',
            `shares defined name "${dn.name}" with the seed`,
            seed
          );
        }
      }
    }
  }

  // Header-similarity and same-sheet edges.
  if (seedRegion && opts.depth >= 1) {
    const seedHeaders = new Set((seedRegion.headers ?? []).map((h) => h.trim().toLowerCase()));
    if (seedHeaders.size > 0) {
      for (const region of graph.allRegions()) {
        if (region.id === seedRegion.id || !region.headers) continue;
        const shared = region.headers.filter((h) => seedHeaders.has(h.trim().toLowerCase()));
        if (shared.length >= 2) {
          push(
            region,
            1,
            opts.weights.semantic * opts.decay,
            'semantic',
            `shares headers (${shared.slice(0, 4).join(', ')}) with the seed's region`,
            seedRegion.id
          );
        }
      }
    }
    for (const region of graph.detectRegions(ref.sheet)) {
      if (region.id === seedRegion.id) continue;
      const a = seedRegion.data.range;
      const b = region.data.range;
      const rowGap = Math.max(a.startRow, b.startRow) - Math.min(a.endRow, b.endRow);
      const colGap = Math.max(a.startCol, b.startCol) - Math.min(a.endCol, b.endCol);
      const adjacent = rowGap <= 3 && colGap <= 3;
      push(
        region,
        1,
        (adjacent ? opts.weights.spatial : opts.weights.sheet) * opts.decay,
        adjacent ? 'spatial' : 'sheet',
        adjacent ? 'physically adjacent to the seed region' : `another region on sheet "${ref.sheet}"`,
        seedRegion.id
      );
    }
  }

  // ── diffusion loop ───────────────────────────────────────────────────────
  while (queue.length > 0) {
    if (regions.length >= opts.maxRegions || budget < 80) {
      truncated = true;
      const skipped = [...queue].sort((a, b) => b.weight - a.weight).slice(0, 2);
      for (const s of skipped) {
        nextActions.push(
          `summariseRegion('${s.region.id}') — candidate at ${s.region.rangeA1} skipped (${s.why})`
        );
      }
      break;
    }
    queue.sort((a, b) => b.weight - a.weight);
    const entry = queue.shift()!;
    if (seen.has(entry.region.id)) continue;
    seen.add(entry.region.id);

    const mode = modeForDepth(opts.mode, entry.depth);
    const regionBudget = Math.min(budget, Math.max(120, Math.floor(opts.tokenBudget / 3)));
    const summary = graph.summariseRegion(entry.region, mode, regionBudget);
    if (summary.tokens > budget) {
      truncated = true;
      nextActions.push(
        `summariseRegion('${entry.region.id}', '${mode}') — needs ~${summary.tokens} tokens, only ${Math.max(0, budget)} left`
      );
      continue;
    }
    budget -= summary.tokens;
    truncated = truncated || summary.truncated;
    regions.push(summary);
    relations.push({
      from: entry.from,
      to: entry.region.id,
      type: entry.type,
      weight: round2(entry.weight),
      why: entry.why
    });
    if (entry.region.data.hiddenSheet) {
      warnings.push(`Region ${entry.region.id} lives on hidden sheet "${entry.region.sheet}"`);
    }

    // Expand: regions referenced by this region's formulas.
    if (entry.depth < opts.depth) {
      for (const range of graph.regionFormulaTargets(entry.region)) {
        for (const target of regionsOverlapping(range)) {
          push(
            target,
            entry.depth + 1,
            opts.weights.formula * Math.pow(opts.decay, entry.depth + 1),
            'formula',
            `formulas in ${entry.region.id} reference ${formatRange(range)}`,
            entry.region.id
          );
        }
      }
    }
  }

  // ── next best actions ────────────────────────────────────────────────────
  const seedSummaryRegion = regions.find((r) => r.regionId === seedRegion?.id);
  const keyCol = seedRegion?.data.columns.find((c) => c.isKey);
  if (keyCol && keyCol.samples.length > 0) {
    nextActions.push(
      `findValue(${JSON.stringify(keyCol.samples[0])}) to locate rows keyed by ${keyCol.header ?? keyCol.letter} across sheets`
    );
  }
  if (cell?.formula && !trace?.precedents) {
    nextActions.push(`tracePrecedents('${seed}') to follow the formula chain`);
  }
  if (truncated) {
    nextActions.push(
      `re-run expandContext with a larger tokenBudget (used ~${opts.tokenBudget - Math.max(0, budget)} of ${opts.tokenBudget})`
    );
  }
  if (seedSummaryRegion && MODE_RANK[opts.mode] < MODE_RANK.cells) {
    nextActions.push(`summariseRegion('${seedSummaryRegion.regionId}', 'cells') for full cell-level detail`);
  }

  const sourceCells = uniq([seed, ...regions.flatMap((r) => r.sourceCells)]);
  const used = opts.tokenBudget - Math.max(0, budget);

  return {
    seed,
    seedCell,
    options: opts,
    regions,
    trace,
    relations,
    nextActions: uniq(nextActions).slice(0, 6),
    warnings: uniq(warnings),
    sourceCells,
    tokens: used,
    truncated
  };
}

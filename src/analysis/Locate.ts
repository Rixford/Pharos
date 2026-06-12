/**
 * Locate — question-aware region narrowing (deterministic).
 *
 * Scores every region against a natural-language question using a small
 * built-in business-synonym lexicon plus the region's own metadata (sheet
 * name, title, purpose tag, headers, section labels, category samples,
 * attached notes). No embeddings, no model calls — callers can extend the
 * lexicon (opts.synonyms) or re-rank the results with their own embedder.
 */
import { RegionData } from '../core/types';

const SYNONYMS: string[][] = [
  ['inflow', 'inflows', 'receipt', 'receipts', 'payment', 'payments', 'collected', 'collections', 'cash in', 'received', 'remittance'],
  ['outflow', 'outflows', 'spend', 'spending', 'cost', 'costs', 'expense', 'expenses', 'cash out'],
  ['vendor', 'vendors', 'supplier', 'suppliers'],
  ['payroll', 'salary', 'salaries', 'wages', 'compensation', 'headcount'],
  ['department', 'departments', 'dept', 'division'],
  ['cost center', 'cost centre', 'cost centers', 'cost-center', 'costcenter'],
  ['customer', 'customers', 'client', 'clients'],
  ['assumption', 'assumptions', 'policy', 'policies', 'basis', 'rules'],
  ['allocation', 'allocations', 'allocated', 'share', 'shares', 'split'],
  ['capex', 'capital', 'investment'],
  ['deferred', 'prepaid', 'recognition'],
  ['invoice', 'invoices', 'billed', 'billing'],
  ['credit', 'credits', 'adjustment', 'adjustments'],
  ['aging', 'ageing', 'overdue', 'receivable', 'receivables'],
  ['forecast', 'projection', 'budget'],
  ['month', 'monthly', 'months'],
  ['risk', 'risks', 'anomaly', 'anomalies', 'spike'],
  ['liquidity', 'net cash', 'cash position'],
  ['revenue', 'income', 'sales'],
  ['mapping', 'rates', 'lookup', 'reference'],
  ['hidden']
];

const STOP = new Set([
  'what', 'which', 'where', 'when', 'how', 'much', 'many', 'does', 'do', 'is', 'are', 'the', 'a', 'an',
  'of', 'by', 'for', 'to', 'in', 'on', 'and', 'or', 'with', 'from', 'that', 'this', 'each', 'per',
  'show', 'find', 'list', 'give', 'tell', 'their', 'there', 'into', 'out', 'about', 'expected', 'come', 'going'
]);

export interface LocateOptions {
  /** Maximum hits returned (default 8). */
  top?: number;
  /** Minimum score to include (default 3). */
  minScore?: number;
  /** Extra synonym groups merged with the built-in lexicon. */
  synonyms?: string[][];
}

export interface LocateTarget {
  region: RegionData;
  workbook?: string;
}

export interface LocateHit {
  workbook?: string;
  regionId: string;
  sheet: string;
  rangeA1: string;
  kind: RegionData['kind'];
  purpose?: string;
  title?: string;
  hiddenSheet: boolean;
  score: number;
  matched: string[];
  why: string;
}

interface Haystack {
  sheet: string;
  title: string;
  purpose: string;
  headers: string;
  extras: string; // section labels + category samples + notes
}

function haystack(d: RegionData): Haystack {
  const sectionLabels = (d.sections ?? []).map((s) => s.label ?? '').join(' ');
  const samples = d.columns
    .filter((c) => c.role === 'category' || c.role === 'key' || c.role === 'id' || d.kind === 'keyValue' || d.kind === 'notes' || d.kind === 'list')
    .flatMap((c) => c.samples.map((x) => String(x)))
    .join(' ');
  return {
    sheet: d.sheet.toLowerCase(),
    title: (d.title ?? '').toLowerCase(),
    purpose: (d.purpose ?? '').toLowerCase(),
    headers: (d.headers ?? []).join(' ').toLowerCase(),
    extras: `${sectionLabels} ${samples} ${(d.notes ?? []).join(' ')}`.toLowerCase()
  };
}

export function locateRegions(targets: LocateTarget[], question: string, opts: LocateOptions = {}): LocateHit[] {
  const lexicon = [...SYNONYMS, ...(opts.synonyms ?? [])];
  const q = ` ${question.toLowerCase()} `;
  const groups = lexicon.filter((g) => g.some((w) => q.includes(w)));
  const grouped = new Set(groups.flat());
  const rawTokens = [
    ...new Set(
      question
        .toLowerCase()
        .split(/[^a-z0-9-]+/)
        .filter((t) => t.length > 3 && !STOP.has(t) && !grouped.has(t))
    )
  ];

  const hits: LocateHit[] = [];
  for (const { region, workbook } of targets) {
    const hay = haystack(region);
    let score = 0;
    const matched = new Set<string>();
    const whys: string[] = [];

    const scoreWords = (words: string[], weights: { sheet: number; title: number; purpose: number; headers: number; extras: number }): void => {
      const fields: string[] = [];
      let hitWord: string | undefined;
      for (const w of words) {
        if (hay.sheet.includes(w)) {
          score += fields.includes('sheet') ? 0 : weights.sheet;
          fields.push('sheet');
          hitWord = hitWord ?? w;
        }
        if (hay.title.includes(w)) {
          score += fields.includes('title') ? 0 : weights.title;
          fields.push('title');
          hitWord = hitWord ?? w;
        }
        if (hay.purpose.includes(w)) {
          score += fields.includes('purpose') ? 0 : weights.purpose;
          fields.push('purpose');
          hitWord = hitWord ?? w;
        }
        if (hay.headers.includes(w)) {
          score += fields.includes('headers') ? 0 : weights.headers;
          fields.push('headers');
          hitWord = hitWord ?? w;
        }
        if (hay.extras.includes(w)) {
          score += fields.includes('extras') ? 0 : weights.extras;
          fields.push('extras');
          hitWord = hitWord ?? w;
        }
      }
      if (hitWord) {
        matched.add(hitWord);
        whys.push(`“${hitWord}” in ${[...new Set(fields)].join('/')}`);
      }
    };

    for (const group of groups) scoreWords(group, { sheet: 3, title: 3, purpose: 3, headers: 2, extras: 1 });
    for (const token of rawTokens) scoreWords([token], { sheet: 2, title: 2, purpose: 2, headers: 1, extras: 1 });

    if (/hidden/.test(q) && region.hiddenSheet) {
      score += 3;
      whys.push('region is on a hidden sheet');
    }

    if (score > 0) {
      hits.push({
        workbook,
        regionId: region.id,
        sheet: region.sheet,
        rangeA1: region.rangeA1,
        kind: region.kind,
        purpose: region.purpose,
        title: region.title,
        hiddenSheet: region.hiddenSheet,
        score,
        matched: [...matched],
        why: whys.slice(0, 4).join('; ')
      });
    }
  }

  // Equal scores: prefer data-bearing kinds (a notes block mentioning
  // "capex" must not outrank the capex table it annotates).
  const KIND_PREF: Record<RegionData['kind'], number> = {
    table: 0,
    matrix: 1,
    keyValue: 2,
    list: 3,
    block: 4,
    notes: 5
  };
  return hits
    .filter((h) => h.score >= (opts.minScore ?? 3))
    .sort(
      (a, b) =>
        b.score - a.score || KIND_PREF[a.kind] - KIND_PREF[b.kind] || a.rangeA1.localeCompare(b.rangeA1)
    )
    .slice(0, opts.top ?? 8);
}

// pharos:eof

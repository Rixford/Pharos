/**
 * Scoring utility — compares a candidate liquidity workbook against the
 * gold report. The scorer is the judge, so it may read the candidate file
 * directly (the *blind solver* may not).
 */
import * as ExcelJS from 'exceljs';
import { GoldReport } from './gold';

export interface SectionScore {
  name: string;
  score: number;
  threshold: number;
  pass: boolean;
  details: string[];
}

export interface ScoreReport {
  pass: boolean;
  sections: SectionScore[];
  /** Total Pharos-reported tokens the solver consumed (efficiency metric). */
  solverTokens?: number;
}

const TOL = 0.05;
const close = (a: number, b: number): boolean =>
  Number.isFinite(a) && Math.abs(a - b) <= Math.max(TOL, Math.abs(b) * 0.002);

type Row = (string | number | null)[];

interface SheetData {
  name: string;
  rows: Row[];
  text: string;
}

async function loadCandidate(path: string): Promise<SheetData[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path);
  const sheets: SheetData[] = [];
  wb.worksheets.forEach((ws) => {
    const rows: Row[] = [];
    ws.eachRow({ includeEmpty: true }, (row, rn) => {
      const out: Row = [];
      row.eachCell({ includeEmpty: true }, (cell, cn) => {
        const v = cell.value;
        let plain: string | number | null = null;
        if (typeof v === 'number' || typeof v === 'string') plain = v;
        else if (v && typeof v === 'object' && 'result' in v) plain = (v as { result: string | number }).result ?? null;
        else if (v && typeof v === 'object' && 'richText' in v) plain = (v as { richText: { text: string }[] }).richText.map((r) => r.text).join('');
        out[cn - 1] = plain;
      });
      rows[rn - 1] = out;
    });
    sheets.push({ name: ws.name, rows, text: rows.flat().filter((x) => x !== null).join(' | ').toLowerCase() });
  });
  return sheets;
}

const findSheet = (sheets: SheetData[], tokenSets: string[][]): SheetData | undefined =>
  sheets.find((s) => tokenSets.some((tokens) => tokens.every((t) => s.name.toLowerCase().includes(t))));

/** Parse a label→number map from a 2+ column sheet (skips Total rows). */
function twoColMap(sheet: SheetData): Map<string, number> {
  const out = new Map<string, number>();
  for (const row of sheet.rows) {
    if (!row) continue;
    const label = row.find((c) => typeof c === 'string' && c.trim() !== '') as string | undefined;
    const value = row.find((c) => typeof c === 'number') as number | undefined;
    if (label === undefined || value === undefined) continue;
    if (/^(total|grand total)$/i.test(label.trim())) continue;
    out.set(label.trim().toLowerCase(), value);
  }
  return out;
}

function monthlyTable(sheet: SheetData): Map<string, { inflows?: number; outflows?: number; net?: number }> {
  const out = new Map<string, { inflows?: number; outflows?: number; net?: number }>();
  let cols: { month: number; inflows: number; outflows: number; net: number } | undefined;
  for (const row of sheet.rows) {
    if (!row) continue;
    if (!cols) {
      const lower = row.map((c) => (typeof c === 'string' ? c.toLowerCase() : ''));
      const mi = lower.findIndex((c) => c.includes('month'));
      const ii = lower.findIndex((c) => c.includes('inflow'));
      const oi = lower.findIndex((c) => c.includes('outflow'));
      const ni = lower.findIndex((c) => c.startsWith('net'));
      if (mi >= 0 && ii >= 0 && oi >= 0 && ni >= 0) cols = { month: mi, inflows: ii, outflows: oi, net: ni };
      continue;
    }
    const month = row[cols.month];
    if (typeof month === 'string' && /^\d{4}-\d{2}$/.test(month.trim())) {
      out.set(month.trim(), {
        inflows: row[cols.inflows] as number,
        outflows: row[cols.outflows] as number,
        net: row[cols.net] as number
      });
    }
  }
  return out;
}

export async function scoreCandidate(
  candidatePath: string,
  gold: GoldReport,
  knownSheets: { billing: string[]; costcenter: string[] },
  solverTokens?: number
): Promise<ScoreReport> {
  const sheets = await loadCandidate(candidatePath);
  const sections: SectionScore[] = [];
  const section = (name: string, score: number, threshold: number, details: string[] = []): void => {
    sections.push({ name, score: Math.round(score * 1000) / 1000, threshold, pass: score >= threshold - 1e-9, details: details.slice(0, 6) });
  };

  // 1. Required sheets.
  const matchers: [string, string[][]][] = [
    ['Monthly Liquidity Summary', [['liquidity'], ['monthly', 'summary']]],
    ['Inflows by Customer', [['inflow', 'customer']]],
    ['Inflows by Billing Category', [['inflow', 'categor'], ['inflow', 'billing']]],
    ['Outflows by Department', [['outflow', 'depart']]],
    ['Outflows by Cost Center', [['outflow', 'cost']]],
    ['Outflows by Vendor', [['outflow', 'vendor']]],
    ['Risk Flags', [['risk']]],
    ['Key Assumptions', [['assumption']]],
    ['Source Mapping', [['source'], ['mapping'], ['provenance']]],
    ['Reconciliation Checks', [['reconcil']]]
  ];
  const found = new Map<string, SheetData | undefined>();
  const missing: string[] = [];
  for (const [label, tokenSets] of matchers) {
    const sheet = findSheet(sheets, tokenSets);
    found.set(label, sheet);
    if (!sheet) missing.push(label);
  }
  section('required sheets present', (matchers.length - missing.length) / matchers.length, 1, missing.map((m) => `missing: ${m}`));

  // 2 + 3. Monthly table: columns, month labels, numeric accuracy.
  const summarySheet = found.get('Monthly Liquidity Summary');
  const monthly = summarySheet ? monthlyTable(summarySheet) : new Map();
  section('required columns present', summarySheet && monthly.size > 0 ? 1 : 0, 1, monthly.size === 0 ? ['monthly table with Month/Inflows/Outflows/Net not found'] : []);
  const monthsOk = gold.months.every((mo) => monthly.has(mo)) && monthly.size === gold.months.length;
  section('month labels match', monthsOk ? 1 : 0, 1, monthsOk ? [] : [`expected ${gold.months.join(',')} got ${[...monthly.keys()].join(',')}`]);

  let totalChecks = 0;
  let totalHits = 0;
  const misses: string[] = [];
  for (const mo of gold.months) {
    const row = monthly.get(mo);
    const checks: [string, number | undefined, number][] = [
      [`inflows ${mo}`, row?.inflows, gold.inflowsByMonth[mo]],
      [`outflows ${mo}`, row?.outflows, gold.outflowsByMonth[mo]],
      [`net ${mo}`, row?.net, gold.netByMonth[mo]]
    ];
    for (const [label, got, want] of checks) {
      totalChecks++;
      if (got !== undefined && typeof got === 'number' && close(got, want)) totalHits++;
      else misses.push(`${label}: got ${got ?? 'missing'} want ${want}`);
    }
  }
  section('monthly totals accuracy', totalChecks ? totalHits / totalChecks : 0, 0.98, misses);

  // 4. Grouped summaries row-match.
  const groups: [string, Record<string, number>][] = [
    ['Inflows by Customer', gold.inflowsByCustomer],
    ['Inflows by Billing Category', gold.inflowsByCategory],
    ['Outflows by Department', gold.outflowsByDept],
    ['Outflows by Cost Center', gold.outflowsByCostCenter],
    ['Outflows by Vendor', gold.outflowsByVendor]
  ];
  let groupScoreSum = 0;
  const groupMisses: string[] = [];
  for (const [label, goldRec] of groups) {
    const sheet = found.get(label);
    const got = sheet ? twoColMap(sheet) : new Map<string, number>();
    const entries = Object.entries(goldRec);
    const hits = entries.filter(([k, v]) => {
      const g = got.get(k.toLowerCase());
      return g !== undefined && close(g, v);
    }).length;
    groupScoreSum += entries.length ? hits / entries.length : 0;
    if (hits < entries.length) groupMisses.push(`${label}: ${hits}/${entries.length} rows matched`);
  }
  section('grouped summary row match', groupScoreSum / groups.length, 0.95, groupMisses);

  // 5. Risk flags (critical: negative-net months + customer concentration).
  const riskText = found.get('Risk Flags')?.text ?? '';
  const critical = gold.riskFlags.filter((f) => f.startsWith('NEGATIVE_NET') || f.startsWith('CUSTOMER_CONCENTRATION'));
  const riskMisses: string[] = [];
  let riskHits = 0;
  for (const flag of gold.riskFlags) {
    const [kind, a, b] = flag.split(':');
    let ok = false;
    if (kind === 'NEGATIVE_NET') ok = riskText.includes(a) && /negative|net/.test(riskText);
    else if (kind === 'CUSTOMER_CONCENTRATION') ok = riskText.includes(a.toLowerCase()) && /concentr|top customer|30/.test(riskText);
    else if (kind === 'VENDOR_SPIKE') ok = riskText.includes(a.toLowerCase()) && (riskText.includes(b) || /spike|anomal|unusual/.test(riskText));
    else if (kind === 'OPEN_AR_HIGH') ok = /open ar|receivable|uncollected|overdue/.test(riskText);
    if (ok) riskHits++;
    else riskMisses.push(flag);
  }
  const criticalOk = critical.every((f) => !riskMisses.includes(f));
  section('risk flags (critical complete)', criticalOk ? Math.max(0.9, riskHits / Math.max(1, gold.riskFlags.length)) : 0, 0.9, riskMisses);

  // 6. Provenance coverage.
  const mapSheet = found.get('Source Mapping');
  const provMetrics = Object.keys(gold.provenance);
  const allKnown = [...knownSheets.billing, ...knownSheets.costcenter].map((s) => s.toLowerCase());
  let provHits = 0;
  const provMisses: string[] = [];
  for (const metric of provMetrics) {
    const row = mapSheet?.rows.find((r) => r?.some((c) => typeof c === 'string' && c.toLowerCase().replace(/[^a-z]/g, '').includes(metric.toLowerCase().replace(/[^a-z]/g, '').slice(0, 12))));
    const sources = row ? row.filter((c) => typeof c === 'string').join(' ').toLowerCase() : '';
    const valid = /(billing|costcenter)\.xlsx/.test(sources) && allKnown.some((s) => sources.includes(s.toLowerCase()));
    if (valid) provHits++;
    else provMisses.push(metric);
  }
  section('source provenance coverage', provMetrics.length ? provHits / provMetrics.length : 0, 0.9, provMisses);

  // 7. Reconciliation checks.
  const reconSheet = found.get('Reconciliation Checks');
  let reconCount = 0;
  let reconPass = 0;
  for (const row of reconSheet?.rows ?? []) {
    if (!row) continue;
    const nums = row.filter((c) => typeof c === 'number') as number[];
    if (typeof row[0] === 'string' && nums.length >= 2) {
      reconCount++;
      if (close(nums[0], nums[1])) reconPass++;
    }
  }
  section('reconciliation checks', reconCount >= 3 && reconPass === reconCount ? 1 : 0, 1, [`${reconPass}/${reconCount} checks reconcile (need ≥3, all passing)`]);

  return { pass: sections.every((s) => s.pass), sections, solverTokens };
}

export function formatScore(report: ScoreReport): string {
  const lines = report.sections.map(
    (s) =>
      `${s.pass ? 'PASS' : 'FAIL'}  ${s.name.padEnd(34)} ${(s.score * 100).toFixed(1).padStart(6)}% (need ${(s.threshold * 100).toFixed(0)}%)` +
      (s.details.length && !s.pass ? `\n        ${s.details.join('\n        ')}` : '')
  );
  lines.push(`${report.pass ? '✅ BENCHMARK PASS' : '❌ BENCHMARK FAIL'}${report.solverTokens ? `  (solver consumed ~${report.solverTokens} Pharos tokens)` : ''}`);
  return lines.join('\n');
}

// pharos:eof

/**
 * Blind reference solver — the scripted "Blind Context Agent".
 *
 * Hard rules:
 *  · Source workbooks are read ONLY through public Pharos APIs.
 *  · No fixture coordinates: navigation is overview → keyword-scored
 *    regions → extraction, so it works across seeds and layout shuffles.
 *  · Every Pharos call runs under a per-call token budget (CALL_BUDGET)
 *    to simulate a real agent's context discipline; consumed tokens are
 *    accumulated and reported as the efficiency metric.
 *  · Output workbook writing uses ExcelJS (writing is not parsing).
 *
 * If the installed Pharos exposes the v0.3 zoom APIs (extractTable), the
 * solver uses them; otherwise it falls back to v0.2 'cells' summaries —
 * that fallback is the measured baseline.
 */
import * as ExcelJS from 'exceljs';
import * as path from 'path';
import { Region, WorkbookGraph, letterToCol } from '../src';

const CALL_BUDGET = 1500;

const CONCEPTS: Record<string, string[]> = {
  payments: ['payment', 'receipt', 'received', 'collected', 'cash receipt', 'remittance'],
  credits: ['credit', 'adjustment'],
  invoices: ['invoice'],
  payroll: ['payroll', 'salary', 'salaries', 'compensation', 'wages'],
  vendorSpend: ['vendor', 'supplier'],
  capex: ['capex', 'capital'],
  assumptions: ['assumption', 'policy', 'basis'],
  allocation: ['alloc', 'share', 'split'],
  costCenter: ['cost center', 'cost centre', 'cc '],
  aging: ['aging', 'ageing', 'overdue']
};

interface SolveResult {
  outPath: string;
  tokens: number;
  warnings: string[];
}

interface ExtractedRow {
  byHeader: Record<string, string | number | boolean | null>;
  row: number;
}

interface Extraction {
  rows: ExtractedRow[];
  headers: string[];
  region: Region;
  source: string;
  complete: boolean;
}

const num = (v: unknown): number => (typeof v === 'number' && Number.isFinite(v) ? v : 0);
const r2 = (n: number): number => Math.round(n * 100) / 100;
const isMonth = (v: unknown): v is string => typeof v === 'string' && /^\d{4}-\d{2}$/.test(v.trim());

export class BlindSolver {
  tokens = 0;
  readonly warnings: string[] = [];

  constructor(
    private readonly graphs: { label: string; graph: WorkbookGraph }[]
  ) {}

  /** Score regions for a concept using only cheap region metadata. */
  private candidates(concept: string, opts?: { book?: string; requireMonths?: boolean }): { label: string; graph: WorkbookGraph; region: Region; score: number }[] {
    const words = CONCEPTS[concept] ?? [concept];
    const out: { label: string; graph: WorkbookGraph; region: Region; score: number }[] = [];
    for (const { label, graph } of this.graphs) {
      if (opts?.book && label !== opts.book) continue;
      const g = graph as WorkbookGraph & { locate?: (q: string, o?: object) => { regionId: string; score: number }[] };
      if (typeof g.locate === 'function') {
        // v0.3 path: engine-side question-aware narrowing.
        for (const hit of g.locate(words.join(' '), { top: 6, minScore: 1 })) {
          const region = graph.getRegion(hit.regionId);
          if (!region) continue;
          let score = hit.score;
          if (opts?.requireMonths) {
            const d = region.data;
            const monthsInHeaders = (d.headers ?? []).filter((h) => /^\d{4}-\d{2}$/.test(h)).length;
            const monthColumn = d.columns.some((c) => c.samples.some((s) => isMonth(s)));
            score += monthsInHeaders >= 3 || monthColumn ? 2 : -4;
          }
          if (score > 0) out.push({ label, graph, region, score });
        }
        continue;
      }
      for (const region of graph.detectRegions()) {
        const d = region.data;
        const hay = {
          sheet: d.sheet.toLowerCase(),
          title: (d.title ?? '').toLowerCase(),
          headers: (d.headers ?? []).join(' ').toLowerCase()
        };
        let score = 0;
        for (const w of words) {
          if (hay.sheet.includes(w)) score += 3;
          if (hay.title.includes(w)) score += 3;
          if (hay.headers.includes(w)) score += 2;
        }
        if (opts?.requireMonths) {
          const monthsInHeaders = (d.headers ?? []).filter((h) => /^\d{4}-\d{2}$/.test(h)).length;
          const monthColumn = d.columns.some((c) => c.samples.some((s) => isMonth(s)));
          if (monthsInHeaders >= 3 || monthColumn) score += 2;
          else score -= 2;
        }
        if (score > 0) out.push({ label, graph, region, score });
      }
    }
    return out.sort((a, b) => b.score - a.score);
  }

  private best(concept: string, opts?: { book?: string; requireMonths?: boolean }): { label: string; graph: WorkbookGraph; region: Region } | undefined {
    return this.candidates(concept, opts)[0];
  }

  /** Extract typed rows from a region through Pharos only. */
  private extract(label: string, graph: WorkbookGraph, region: Region): Extraction {
    const source = `[${label}]${region.rangeA1}`;
    const g = graph as WorkbookGraph & { extractTable?: (target: string | Region, opts?: object) => { columns: { name: string }[]; rows: Record<string, string | number | boolean | null>[]; tokens: number; complete: boolean } };
    if (typeof g.extractTable === 'function') {
      const table = g.extractTable(region, {});
      this.tokens += table.tokens;
      const headers = table.columns.map((c) => c.name);
      return {
        rows: table.rows.map((r, i) => ({ byHeader: r, row: i })),
        headers,
        region,
        source,
        complete: table.complete
      };
    }
    // v0.2 fallback: 'cells' summaries under the call budget.
    const summary = graph.summariseRegion(region, 'cells', CALL_BUDGET);
    this.tokens += summary.tokens;
    const data = summary.data as { cells?: { a: string; v: string | number | null; f?: string }[]; omitted?: number } | undefined;
    const cells = data?.cells ?? [];
    const startCol = region.data.range.startCol;
    const headers = region.data.headers ?? [];
    const byRow = new Map<number, Record<string, string | number | null>>();
    for (const cell of cells) {
      const m = /!([A-Z]+)(\d+)$/.exec(cell.a);
      if (!m) continue;
      const row = parseInt(m[2], 10);
      const col = letterToCol(m[1]);
      if (region.data.headerRow !== undefined && row <= region.data.headerRow) continue;
      if (row < region.data.dataStartRow || row > region.data.dataEndRow) continue;
      const header = headers[col - startCol] ?? `col${col}`;
      const entry = byRow.get(row) ?? {};
      entry[header] = cell.v;
      byRow.set(row, entry);
    }
    const complete = !summary.truncated && (data?.omitted ?? 0) === 0;
    if (!complete) this.warnings.push(`incomplete extraction for ${source} (truncated under ${CALL_BUDGET}-token call budget)`);
    const rows: ExtractedRow[] = [...byRow.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([row, byHeader]) => ({ row, byHeader }))
      .filter((r) => {
        const strings = Object.values(r.byHeader).filter((v): v is string => typeof v === 'string');
        return !strings.some((s) => /subtotal|grand total/i.test(s)) && !strings.some((s) => /^total\b/i.test(s.trim()));
      });
    return { rows, headers, region, source, complete };
  }

  private headerLike(headers: string[], ...words: string[]): string | undefined {
    // Word-priority (earlier words are stronger signals) and exact-first
    // ('Vendor' must beat 'Vendor ID'; 'amount' must beat 'Month Received').
    for (const w of words) {
      const exact = headers.find((h) => h.trim().toLowerCase() === w);
      if (exact) return exact;
    }
    for (const w of words) {
      const hit = headers.find((h) => h.toLowerCase().includes(w));
      if (hit) return hit;
    }
    return undefined;
  }

  /** Read assumption sheets and derive the cash rules. */
  private readRules(): { text: string; netOfCredits: boolean; outPayroll: boolean; outVendor: boolean; outCapex: boolean; sharedAllocated: boolean; sources: string[] } {
    let text = '';
    const sources: string[] = [];
    for (const { label } of this.graphs) {
      for (const cand of this.candidates('assumptions', { book: label }).slice(0, 2)) {
        const summary = cand.graph.summariseRegion(cand.region, 'cells', CALL_BUDGET);
        this.tokens += summary.tokens;
        const data = summary.data as { cells?: { v: string | number | null }[] } | undefined;
        text += ' ' + (data?.cells ?? []).map((c) => String(c.v ?? '')).join(' ');
        sources.push(`[${label}]${cand.region.rangeA1}`);
      }
    }
    const t = text.toLowerCase();
    return {
      text: t,
      netOfCredits: /minus credits|net of credits|credits applied/.test(t),
      outPayroll: /payroll/.test(t),
      outVendor: /vendor/.test(t),
      outCapex: /capex/.test(t),
      sharedAllocated: /allocat/.test(t),
      sources
    };
  }

  /** Hidden-sheet mappings (allocation shares + cost-center payroll shares). */
  private readHiddenMaps(): {
    deptShare: Map<string, number>;
    ccShare: Map<string, { dept: string; share: number }>;
    codeToName: Map<string, string>;
    sources: string[];
  } {
    const deptShare = new Map<string, number>();
    const ccShare = new Map<string, { dept: string; share: number }>();
    const codeToName = new Map<string, string>();
    const sources: string[] = [];
    for (const { label, graph } of this.graphs) {
      const hidden = graph.sheets().filter((s) => s.hidden);
      for (const sheet of hidden) {
        for (const region of graph.detectRegions(sheet.name)) {
          const ext = this.extract(label, graph, region);
          const h = ext.headers;
          const shareCol = this.headerLike(h, 'share', 'alloc', '%');
          const ccCol = this.headerLike(h, 'cost center', 'cost centre');
          const codeCol = this.headerLike(h, 'dept code', 'code', 'dept', 'department');
          const nameCol = h.find((x) => /department|name/i.test(x) && x !== codeCol);
          if (!shareCol || !codeCol) continue;
          sources.push(ext.source);
          for (const row of ext.rows) {
            const share = num(row.byHeader[shareCol]);
            const code = String(row.byHeader[codeCol] ?? '');
            if (ccCol && row.byHeader[ccCol]) {
              ccShare.set(String(row.byHeader[ccCol]), { dept: code, share });
            } else if (code) {
              deptShare.set(code, share);
              if (nameCol && row.byHeader[nameCol]) codeToName.set(code, String(row.byHeader[nameCol]));
            }
          }
        }
      }
    }
    return { deptShare, ccShare, codeToName, sources };
  }

  async solve(outPath: string): Promise<SolveResult> {
    const provenance: Record<string, string[]> = {};
    const rules = this.readRules();
    provenance.assumptions = rules.sources;

    // ── inflows ──────────────────────────────────────────────────────────
    const payments = this.best('payments', { requireMonths: true });
    if (!payments) throw new Error('blind solver: no payments-like region found');
    const payExt = this.extract(payments.label, payments.graph, payments.region);
    const ph = payExt.headers;
    const pMonth = this.headerLike(ph, 'month') ?? ph.find((h) => /^\d{4}-\d{2}$/.test(h)) ?? 'Month';
    const pAmount = this.headerLike(ph, 'amount', 'received') ?? 'Amount';
    const pCustomer = this.headerLike(ph, 'customer') ?? 'Customer';
    const pInvoice = this.headerLike(ph, 'invoice') ?? 'Invoice';

    const creditsCand = this.best('credits');
    const creditExt = creditsCand ? this.extract(creditsCand.label, creditsCand.graph, creditsCand.region) : undefined;

    const invoicesCand = this.best('invoices');
    const invExt = invoicesCand ? this.extract(invoicesCand.label, invoicesCand.graph, invoicesCand.region) : undefined;
    const invCategory = new Map<string, string>();
    let billedTotal = 0;
    if (invExt) {
      const ih = invExt.headers;
      const iId = this.headerLike(ih, 'invoice') ?? 'Invoice ID';
      const iCat = this.headerLike(ih, 'category') ?? 'Category';
      const iAmt = this.headerLike(ih, 'amount') ?? 'Amount';
      for (const row of invExt.rows) {
        invCategory.set(String(row.byHeader[iId]), String(row.byHeader[iCat] ?? 'Unknown'));
        billedTotal += num(row.byHeader[iAmt]);
      }
    }

    const months = new Set<string>();
    const inflowsByMonth = new Map<string, number>();
    const inflowsByCustomer = new Map<string, number>();
    const inflowsByCategory = new Map<string, number>();
    let collectedTotal = 0;
    for (const row of payExt.rows) {
      const mo = String(row.byHeader[pMonth] ?? '');
      if (!isMonth(mo)) continue;
      const amt = num(row.byHeader[pAmount]);
      months.add(mo);
      collectedTotal += amt;
      inflowsByMonth.set(mo, r2((inflowsByMonth.get(mo) ?? 0) + amt));
      const cust = String(row.byHeader[pCustomer] ?? 'Unknown');
      inflowsByCustomer.set(cust, r2((inflowsByCustomer.get(cust) ?? 0) + amt));
      const cat = invCategory.get(String(row.byHeader[pInvoice])) ?? 'Unknown';
      inflowsByCategory.set(cat, r2((inflowsByCategory.get(cat) ?? 0) + amt));
    }
    provenance.inflowsByMonth = [payExt.source];
    provenance.inflowsByCustomer = [payExt.source];
    provenance.inflowsByCategory = invExt ? [invExt.source, payExt.source] : [payExt.source];

    if (creditExt && rules.netOfCredits) {
      const ch = creditExt.headers;
      const cMonth = this.headerLike(ch, 'month') ?? 'Month';
      const cAmt = this.headerLike(ch, 'amount', 'credit') ?? 'Amount';
      const cCust = this.headerLike(ch, 'customer') ?? 'Customer';
      const cInv = this.headerLike(ch, 'invoice') ?? 'Invoice';
      for (const row of creditExt.rows) {
        const mo = String(row.byHeader[cMonth] ?? '');
        if (!isMonth(mo)) continue;
        const amt = num(row.byHeader[cAmt]);
        inflowsByMonth.set(mo, r2((inflowsByMonth.get(mo) ?? 0) - amt));
        const cust = String(row.byHeader[cCust] ?? 'Unknown');
        if (inflowsByCustomer.has(cust)) inflowsByCustomer.set(cust, r2(inflowsByCustomer.get(cust)! - amt));
        const cat = invCategory.get(String(row.byHeader[cInv]));
        if (cat && inflowsByCategory.has(cat)) inflowsByCategory.set(cat, r2(inflowsByCategory.get(cat)! - amt));
      }
      provenance.inflowsByMonth.push(creditExt.source);
      provenance.inflowsByCustomer.push(creditExt.source);
    }

    // ── outflows ─────────────────────────────────────────────────────────
    const payrollCand = this.best('payroll', { requireMonths: true });
    if (!payrollCand) throw new Error('blind solver: no payroll-like region found');
    const payrollExt = this.extract(payrollCand.label, payrollCand.graph, payrollCand.region);
    const payrollByDeptMonth = new Map<string, number>();
    const deptLabelCol = payrollExt.headers.find((h) => !/^\d{4}-\d{2}$/.test(h) && !/total/i.test(h)) ?? payrollExt.headers[0];
    const payrollMonths = payrollExt.headers.filter((h) => /^\d{4}-\d{2}$/.test(h));
    for (const row of payrollExt.rows) {
      const dept = String(row.byHeader[deptLabelCol] ?? '');
      if (!dept) continue;
      for (const mo of payrollMonths) {
        payrollByDeptMonth.set(`${dept}|${mo}`, num(row.byHeader[mo]));
        months.add(mo);
      }
    }
    provenance.outflowsByMonth = [payrollExt.source];

    const vendorCand = this.best('vendorSpend', { requireMonths: true });
    if (!vendorCand) throw new Error('blind solver: no vendor-spend region found');
    const vendorExt = this.extract(vendorCand.label, vendorCand.graph, vendorCand.region);
    const vh = vendorExt.headers;
    const vName = this.headerLike(vh, 'vendor') ?? 'Vendor';
    const vDept = this.headerLike(vh, 'dept') ?? 'Dept';
    const vMonth = this.headerLike(vh, 'month') ?? 'Month';
    const vAmt = this.headerLike(vh, 'amount') ?? 'Amount';
    provenance.outflowsByMonth.push(vendorExt.source);
    provenance.outflowsByVendor = [vendorExt.source];

    const capexCand = this.best('capex');
    const capexExt = capexCand ? this.extract(capexCand.label, capexCand.graph, capexCand.region) : undefined;
    if (capexExt) {
      provenance.outflowsByMonth.push(capexExt.source);
      (provenance.outflowsByVendor = provenance.outflowsByVendor ?? []).push(capexExt.source);
    }

    const hiddenMaps = this.readHiddenMaps();
    provenance.outflowsByDept = [vendorExt.source, payrollExt.source, ...hiddenMaps.sources];
    provenance.outflowsByCostCenter = hiddenMaps.sources.length ? [...hiddenMaps.sources] : [];

    const outflowsByMonth = new Map<string, number>();
    const outflowsByDept = new Map<string, number>();
    const outflowsByVendor = new Map<string, number>();
    const sharedByMonth = new Map<string, number>();
    const vendorMonthly = new Map<string, Map<string, number>>();
    const deptDirectMonthly = new Map<string, number>();

    for (const row of vendorExt.rows) {
      const mo = String(row.byHeader[vMonth] ?? '');
      if (!isMonth(mo)) continue;
      const amt = num(row.byHeader[vAmt]);
      const vendor = String(row.byHeader[vName] ?? 'Unknown');
      const dept = String(row.byHeader[vDept] ?? '');
      months.add(mo);
      outflowsByMonth.set(mo, r2((outflowsByMonth.get(mo) ?? 0) + amt));
      outflowsByVendor.set(vendor, r2((outflowsByVendor.get(vendor) ?? 0) + amt));
      const vm = vendorMonthly.get(vendor) ?? new Map<string, number>();
      vm.set(mo, r2((vm.get(mo) ?? 0) + amt));
      vendorMonthly.set(vendor, vm);
      if (dept.toUpperCase() === 'SHARED') sharedByMonth.set(mo, r2((sharedByMonth.get(mo) ?? 0) + amt));
      else deptDirectMonthly.set(`${dept}|${mo}`, r2((deptDirectMonthly.get(`${dept}|${mo}`) ?? 0) + amt));
    }
    for (const [key, amt] of payrollByDeptMonth) {
      const mo = key.split('|')[1];
      outflowsByMonth.set(mo, r2((outflowsByMonth.get(mo) ?? 0) + amt));
    }
    const capexRows: { dept: string; cc: string; vendor: string; month: string; amount: number }[] = [];
    if (capexExt) {
      const ch = capexExt.headers;
      const cMo = this.headerLike(ch, 'month') ?? 'Month';
      const cAmt = this.headerLike(ch, 'amount') ?? 'Amount';
      const cDept = this.headerLike(ch, 'dept') ?? 'Dept';
      const cCc = this.headerLike(ch, 'cost center', 'cost centre') ?? 'Cost Center';
      const cVendor = this.headerLike(ch, 'vendor') ?? 'Vendor';
      for (const row of capexExt.rows) {
        const mo = String(row.byHeader[cMo] ?? '');
        if (!isMonth(mo)) continue;
        const amt = num(row.byHeader[cAmt]);
        capexRows.push({ dept: String(row.byHeader[cDept] ?? ''), cc: String(row.byHeader[cCc] ?? ''), vendor: String(row.byHeader[cVendor] ?? ''), month: mo, amount: amt });
        outflowsByMonth.set(mo, r2((outflowsByMonth.get(mo) ?? 0) + amt));
        const vendor = String(row.byHeader[cVendor] ?? 'Unknown');
        outflowsByVendor.set(vendor, r2((outflowsByVendor.get(vendor) ?? 0) + amt));
      }
    }

    // Department rollup: payroll + direct vendor + allocated shared + capex.
    // Dept codes ↔ names: hidden dept map gives code→share; payroll matrix uses names.
    const deptCodes = [...hiddenMaps.deptShare.keys()];
    const deptNames = [...new Set([...payrollByDeptMonth.keys()].map((k) => k.split('|')[0]))];
    const codeToName = new Map<string, string>(hiddenMaps.codeToName);
    for (const code of deptCodes) {
      if (codeToName.has(code)) continue;
      const name = deptNames.find((n) => n.toLowerCase().startsWith(code.slice(0, 3).toLowerCase()) || n.toLowerCase().includes(code.toLowerCase()));
      codeToName.set(code, name ?? code);
    }
    // direct vendor dept values may be codes; payroll uses names — normalise to display name.
    const normDept = (raw: string): string => codeToName.get(raw) ?? raw;
    for (const name of deptNames) outflowsByDept.set(name, 0);
    for (const [key, amt] of payrollByDeptMonth) {
      const dept = key.split('|')[0];
      outflowsByDept.set(dept, r2((outflowsByDept.get(dept) ?? 0) + amt));
    }
    for (const [key, amt] of deptDirectMonthly) {
      const dept = normDept(key.split('|')[0]);
      outflowsByDept.set(dept, r2((outflowsByDept.get(dept) ?? 0) + amt));
    }
    if (rules.sharedAllocated && hiddenMaps.deptShare.size > 0) {
      for (const [mo, pool] of sharedByMonth) {
        void mo;
        for (const [code, share] of hiddenMaps.deptShare) {
          const dept = normDept(code);
          outflowsByDept.set(dept, r2((outflowsByDept.get(dept) ?? 0) + r2(pool * share)));
        }
      }
    }
    for (const cap of capexRows) {
      const dept = normDept(cap.dept);
      outflowsByDept.set(dept, r2((outflowsByDept.get(dept) ?? 0) + cap.amount));
    }

    // Cost-center rollup from hidden cc shares + vendor cc attribution is not
    // available in the registers; rebuild from cc payroll shares + capex +
    // direct-vendor cost centers when discoverable via a cost-center detail region.
    const ccCand = this.best('costCenter', { requireMonths: true });
    const outflowsByCC = new Map<string, number>();
    if (ccCand) {
      const ccExt = this.extract(ccCand.label, ccCand.graph, ccCand.region);
      const h = ccExt.headers;
      const ccCol = this.headerLike(h, 'cost center', 'cost centre') ?? h[0];
      const totalCol = this.headerLike(h, 'total');
      if (totalCol) {
        for (const row of ccExt.rows) {
          const cc = String(row.byHeader[ccCol] ?? '');
          if (!cc) continue;
          outflowsByCC.set(cc, r2((outflowsByCC.get(cc) ?? 0) + num(row.byHeader[totalCol])));
        }
        provenance.outflowsByCostCenter = [ccExt.source];
      }
    }

    // ── derived outputs ──────────────────────────────────────────────────
    const sortedMonths = [...months].filter(isMonth).sort();
    const netByMonth = new Map<string, number>();
    for (const mo of sortedMonths) {
      netByMonth.set(mo, r2((inflowsByMonth.get(mo) ?? 0) - (outflowsByMonth.get(mo) ?? 0)));
    }

    const totalIn = r2([...inflowsByMonth.values()].reduce((s, v) => s + v, 0));
    const totalOut = r2([...outflowsByMonth.values()].reduce((s, v) => s + v, 0));
    const risks: string[] = [];
    for (const mo of sortedMonths) {
      if ((netByMonth.get(mo) ?? 0) < 0) risks.push(`Negative net liquidity in ${mo} (outflows exceed inflows)`);
    }
    for (const [cust, amt] of inflowsByCustomer) {
      if (totalIn > 0 && amt / totalIn > 0.3) risks.push(`Customer concentration: ${cust} drives ${(100 * amt / totalIn).toFixed(1)}% of inflows (>30%)`);
    }
    for (const [vendor, vm] of vendorMonthly) {
      const vals = [...vm.values()].sort((a, b) => a - b);
      if (vals.length < 3) continue;
      const median = vals[Math.floor(vals.length / 2)];
      for (const [mo, amt] of vm) {
        if (median > 0 && amt > 1.8 * median) risks.push(`Vendor spend spike: ${vendor} in ${mo} (${r2(amt)} vs median ${r2(median)})`);
      }
    }
    if (billedTotal > 0 && billedTotal - collectedTotal > 0.1 * billedTotal) {
      risks.push(`Open AR is high: uncollected receivables exceed 10% of billed (overdue/open invoices)`);
    }

    const reconciliation: { name: string; a: number; b: number }[] = [
      { name: 'Total inflows equals sum of customer inflows', a: totalIn, b: r2([...inflowsByCustomer.values()].reduce((s, v) => s + v, 0)) },
      { name: 'Total inflows equals sum of category inflows', a: totalIn, b: r2([...inflowsByCategory.values()].reduce((s, v) => s + v, 0)) },
      { name: 'Total outflows equals sum of department outflows', a: totalOut, b: r2([...outflowsByDept.values()].reduce((s, v) => s + v, 0)) },
      { name: 'Total outflows equals sum of cost-center outflows', a: totalOut, b: r2([...outflowsByCC.values()].reduce((s, v) => s + v, 0)) },
      { name: 'Total outflows equals payroll plus vendor outflows', a: totalOut, b: r2([...outflowsByVendor.values()].reduce((s, v) => s + v, 0) + [...payrollByDeptMonth.values()].reduce((s, v) => s + v, 0)) }
    ];

    const assumptions = [
      'Cash basis: inflows are counted in the month payment is received.',
      rules.netOfCredits ? 'Inflows are net of credits applied in the month.' : 'Credits were not netted (no rule found).',
      'Outflows = payroll + vendor spend + capex, in the month incurred.',
      rules.sharedAllocated ? 'Shared vendor costs are allocated to departments via the hidden allocation shares.' : 'No shared-cost allocation rule found.',
      'Annual prepaid contracts collect full cash up front; deferred recognition is informational.',
      'Forecast figures are informational and excluded from actuals.'
    ];

    await this.write(outPath, sortedMonths, inflowsByMonth, outflowsByMonth, netByMonth, inflowsByCustomer, inflowsByCategory, outflowsByDept, outflowsByCC, outflowsByVendor, risks, assumptions, provenance, reconciliation);
    return { outPath, tokens: this.tokens, warnings: this.warnings };
  }

  private async write(
    outPath: string,
    months: string[],
    inflows: Map<string, number>,
    outflows: Map<string, number>,
    net: Map<string, number>,
    byCustomer: Map<string, number>,
    byCategory: Map<string, number>,
    byDept: Map<string, number>,
    byCC: Map<string, number>,
    byVendor: Map<string, number>,
    risks: string[],
    assumptions: string[],
    provenance: Record<string, string[]>,
    reconciliation: { name: string; a: number; b: number }[]
  ): Promise<void> {
    const wb = new ExcelJS.Workbook();
    const summary = wb.addWorksheet('Monthly Liquidity Summary');
    summary.addRow(['Month', 'Inflows', 'Outflows', 'Net']);
    for (const mo of months) summary.addRow([mo, inflows.get(mo) ?? 0, outflows.get(mo) ?? 0, net.get(mo) ?? 0]);

    const writeMap = (name: string, header: [string, string], map: Map<string, number>): void => {
      const ws = wb.addWorksheet(name);
      ws.addRow(header);
      for (const [k, v] of map) ws.addRow([k, v]);
      ws.addRow(['Total', r2([...map.values()].reduce((s, x) => s + x, 0))]);
    };
    writeMap('Inflows by Customer', ['Customer', 'Inflows'], byCustomer);
    writeMap('Inflows by Billing Category', ['Billing Category', 'Inflows'], byCategory);
    writeMap('Outflows by Department', ['Department', 'Outflows'], byDept);
    writeMap('Outflows by Cost Center', ['Cost Center', 'Outflows'], byCC);
    writeMap('Outflows by Vendor', ['Vendor', 'Outflows'], byVendor);

    const riskWs = wb.addWorksheet('Risk Flags');
    riskWs.addRow(['Flag']);
    risks.forEach((r) => riskWs.addRow([r]));

    const aWs = wb.addWorksheet('Key Assumptions');
    aWs.addRow(['#', 'Assumption']);
    assumptions.forEach((a, i) => aWs.addRow([i + 1, a]));

    const mWs = wb.addWorksheet('Source Mapping');
    mWs.addRow(['Metric', 'Sources']);
    Object.entries(provenance).forEach(([k, v]) => mWs.addRow([k, v.join(' ; ')]));

    const rWs = wb.addWorksheet('Reconciliation Checks');
    rWs.addRow(['Check', 'A', 'B', 'Diff']);
    reconciliation.forEach((c) => rWs.addRow([c.name, c.a, c.b, r2(c.a - c.b)]));

    await wb.xlsx.writeFile(outPath);
  }
}

export async function solveBlind(billingPath: string, costcenterPath: string, outPath: string): Promise<SolveResult> {
  const graphs = [
    { label: path.basename(billingPath), graph: await WorkbookGraph.load(billingPath) },
    { label: path.basename(costcenterPath), graph: await WorkbookGraph.load(costcenterPath) }
  ];
  return new BlindSolver(graphs).solve(outPath);
}

// pharos:eof

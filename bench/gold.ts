/**
 * Gold liquidity report — computed straight from the dataset (never from
 * the generated workbooks), written as xlsx + JSON with provenance taken
 * from the builder manifests.
 */
import * as ExcelJS from 'exceljs';
import { Dataset } from './data';
import { Manifest, headerRow, title, bold, r2 } from './xlsx-helpers';

export interface GoldReport {
  seed: number;
  months: string[];
  inflowsByMonth: Record<string, number>;
  outflowsByMonth: Record<string, number>;
  netByMonth: Record<string, number>;
  inflowsByCustomer: Record<string, number>;
  inflowsByCategory: Record<string, number>;
  outflowsByDept: Record<string, number>;
  outflowsByCostCenter: Record<string, number>;
  outflowsByVendor: Record<string, number>;
  riskFlags: string[];
  assumptions: string[];
  reconciliation: { name: string; a: number; b: number }[];
  provenance: Record<string, string[]>;
}

const sumVals = (rec: Record<string, number>): number => r2(Object.values(rec).reduce((s, v) => s + v, 0));

export function computeGold(ds: Dataset, billing: Manifest, costcenter: Manifest): GoldReport {

  const vendorById = new Map(ds.vendors.map((v) => [v.id, v]));
  const invoiceById = new Map(ds.invoices.map((i) => [i.id, i]));

  const inflowsByMonth: Record<string, number> = {};
  const outflowsByMonth: Record<string, number> = {};
  const netByMonth: Record<string, number> = {};
  for (const mo of ds.months) {
    const pay = ds.payments.filter((p) => p.month === mo).reduce((s, p) => s + p.amount, 0);
    const cr = ds.credits.filter((c) => c.month === mo).reduce((s, c) => s + c.amount, 0);
    inflowsByMonth[mo] = r2(pay - cr);
    const payroll = ds.payroll.filter((p) => p.month === mo).reduce((s, p) => s + p.amount, 0);
    const vendor = ds.vendorSpend.filter((v) => v.month === mo).reduce((s, v) => s + v.amount, 0);
    const capex = ds.capex.filter((c) => c.month === mo).reduce((s, c) => s + c.amount, 0);
    outflowsByMonth[mo] = r2(payroll + vendor + capex);
    netByMonth[mo] = r2(inflowsByMonth[mo] - outflowsByMonth[mo]);
  }

  const inflowsByCustomer: Record<string, number> = {};
  for (const c of ds.customers) {
    const pay = ds.payments.filter((p) => p.customerId === c.id).reduce((s, p) => s + p.amount, 0);
    const cr = ds.credits.filter((x) => x.customerId === c.id).reduce((s, x) => s + x.amount, 0);
    inflowsByCustomer[c.name] = r2(pay - cr);
  }

  const inflowsByCategory: Record<string, number> = {};
  for (const cat of ds.billingCategories) inflowsByCategory[cat.name] = 0;
  for (const p of ds.payments) {
    const cat = invoiceById.get(p.invoiceId)!.category;
    inflowsByCategory[cat] = r2(inflowsByCategory[cat] + p.amount);
  }
  for (const cr of ds.credits) {
    const cat = invoiceById.get(cr.invoiceId)!.category;
    inflowsByCategory[cat] = r2(inflowsByCategory[cat] - cr.amount);
  }

  const sharedByMonth = new Map(
    ds.months.map((mo) => [
      mo,
      r2(ds.vendorSpend.filter((v) => v.month === mo && vendorById.get(v.vendorId)!.dept === 'SHARED').reduce((s, v) => s + v.amount, 0))
    ])
  );
  const outflowsByDept: Record<string, number> = {};
  for (const d of ds.departments) {
    let total = 0;
    for (const mo of ds.months) {
      total +=
        ds.payroll.filter((p) => p.dept === d.code && p.month === mo).reduce((s, p) => s + p.amount, 0) +
        ds.vendorSpend.filter((v) => v.month === mo && vendorById.get(v.vendorId)!.dept === d.code).reduce((s, v) => s + v.amount, 0) +
        r2(sharedByMonth.get(mo)! * ds.allocShares[d.code]) +
        ds.capex.filter((c) => c.dept === d.code && c.month === mo).reduce((s, c) => s + c.amount, 0);
    }
    outflowsByDept[d.name] = r2(total);
  }

  const outflowsByCostCenter: Record<string, number> = {};
  for (const cc of ds.costCenters) {
    let total = 0;
    for (const mo of ds.months) {
      total +=
        r2(ds.payroll.filter((p) => p.dept === cc.dept && p.month === mo).reduce((s, p) => s + p.amount, 0) * cc.payrollShare) +
        ds.vendorSpend.filter((v) => v.month === mo && vendorById.get(v.vendorId)!.costCenter === cc.code).reduce((s, v) => s + v.amount, 0) +
        r2(sharedByMonth.get(mo)! * ds.allocShares[cc.dept] * cc.payrollShare) +
        ds.capex.filter((c) => c.costCenter === cc.code && c.month === mo).reduce((s, c) => s + c.amount, 0);
    }
    outflowsByCostCenter[cc.code] = r2(total);
  }

  const outflowsByVendor: Record<string, number> = {};
  for (const v of ds.vendors) {
    const opex = ds.vendorSpend.filter((x) => x.vendorId === v.id).reduce((s, x) => s + x.amount, 0);
    const capex = ds.capex.filter((c) => c.vendorId === v.id).reduce((s, c) => s + c.amount, 0);
    outflowsByVendor[v.name] = r2(opex + capex);
  }

  // ── risk flags ───────────────────────────────────────────────────────────
  const riskFlags: string[] = [];
  for (const mo of ds.months) if (netByMonth[mo] < 0) riskFlags.push(`NEGATIVE_NET:${mo}`);
  const totalIn = sumVals(inflowsByMonth);
  for (const [name, amount] of Object.entries(inflowsByCustomer)) {
    if (amount / totalIn > 0.3) riskFlags.push(`CUSTOMER_CONCENTRATION:${name}`);
  }
  for (const v of ds.vendors) {
    const monthly = ds.months.map((mo) =>
      ds.vendorSpend.filter((x) => x.vendorId === v.id && x.month === mo).reduce((s, x) => s + x.amount, 0)
    );
    const sorted = [...monthly].sort((a, b) => a - b);
    const median = sorted[Math.floor(sorted.length / 2)];
    monthly.forEach((amt, i) => {
      if (median > 0 && amt > 1.8 * median) riskFlags.push(`VENDOR_SPIKE:${v.name}:${ds.months[i]}`);
    });
  }
  const totalBilled = r2(ds.invoices.reduce((s, i) => s + i.amount, 0));
  const totalCollected = r2(ds.payments.reduce((s, p) => s + p.amount, 0));
  if (totalBilled - totalCollected > 0.1 * totalBilled) riskFlags.push('OPEN_AR_HIGH');

  const assumptions = [
    'Cash basis: inflows are counted in the month payment is received.',
    'Inflows are net of credits applied in the month.',
    'Outflows = payroll + vendor spend + capex, in the month incurred.',
    'Shared vendor costs are allocated to departments via the hidden AllocMap shares.',
    'Annual prepaid contracts collect full cash up front; deferred recognition is informational.',
    'The Forecast tab is informational and excluded from actuals.'
  ];

  const reconciliation = [
    { name: 'Total inflows equals sum of customer inflows', a: totalIn, b: sumVals(inflowsByCustomer) },
    { name: 'Total inflows equals sum of category inflows', a: totalIn, b: sumVals(inflowsByCategory) },
    { name: 'Total outflows equals sum of department outflows', a: sumVals(outflowsByMonth), b: sumVals(outflowsByDept) },
    { name: 'Total outflows equals sum of cost-center outflows', a: sumVals(outflowsByMonth), b: sumVals(outflowsByCostCenter) },
    { name: 'Total outflows equals payroll plus vendor outflows', a: sumVals(outflowsByMonth), b: r2(sumVals(outflowsByVendor) + ds.payroll.reduce((s, p) => s + p.amount, 0)) }
  ];

  const src = (man: Manifest, sheet: string, label: string): string => {
    const p = man.find(sheet, label);
    return `[${p.book}]${p.sheet}!${p.range}`;
  };
  const provenance: Record<string, string[]> = {
    inflowsByMonth: [src(billing, 'Payments', 'Payments register'), src(billing, 'Adjustments & Credits', 'Credits register')],
    inflowsByCustomer: [src(billing, 'Payments', 'Payments register'), src(billing, 'Adjustments & Credits', 'Credits register')],
    inflowsByCategory: [src(billing, 'Invoices', 'Invoice register'), src(billing, 'Payments', 'Payments register')],
    outflowsByMonth: [src(costcenter, 'Payroll', 'Payroll matrix'), src(costcenter, 'Vendor Spend', 'Vendor spend register'), src(costcenter, 'Capex', 'Capex purchases')],
    outflowsByDept: [src(costcenter, 'Department Spend', 'Department spend matrix'), src(costcenter, 'AllocMap', 'Department allocation shares')],
    outflowsByCostCenter: [src(costcenter, 'Cost Center Detail', 'Cost center detail')],
    outflowsByVendor: [src(costcenter, 'Vendor Spend', 'Vendor spend register'), src(costcenter, 'Capex', 'Capex purchases')],
    assumptions: [src(billing, 'Assumptions', 'Assumptions list'), src(costcenter, 'Assumptions', 'Assumptions list')]
  };

  return {
    seed: ds.seed,
    months: ds.months,
    inflowsByMonth,
    outflowsByMonth,
    netByMonth,
    inflowsByCustomer,
    inflowsByCategory,
    outflowsByDept,
    outflowsByCostCenter,
    outflowsByVendor,
    riskFlags,
    assumptions,
    reconciliation,
    provenance
  };
}

export async function writeGoldWorkbook(gold: GoldReport, outPath: string): Promise<void> {
  const wb = new ExcelJS.Workbook();

  const summary = wb.addWorksheet('Monthly Liquidity Summary');
  title(summary, 1, 1, 4, 'Monthly Liquidity Summary — H1 2026');
  headerRow(summary, 3, 1, ['Month', 'Inflows', 'Outflows', 'Net']);
  gold.months.forEach((mo, i) => {
    const r = 4 + i;
    summary.getCell(r, 1).value = mo;
    summary.getCell(r, 2).value = gold.inflowsByMonth[mo];
    summary.getCell(r, 3).value = gold.outflowsByMonth[mo];
    summary.getCell(r, 4).value = gold.netByMonth[mo];
  });

  const writeKv = (sheetName: string, header: [string, string], rec: Record<string, number>): void => {
    const ws = wb.addWorksheet(sheetName);
    headerRow(ws, 1, 1, header);
    Object.entries(rec).forEach(([k, v], i) => {
      ws.getCell(2 + i, 1).value = k;
      ws.getCell(2 + i, 2).value = v;
    });
    bold(ws, 2 + Object.keys(rec).length, 1, 'Total');
    ws.getCell(2 + Object.keys(rec).length, 2).value = sumVals(rec);
  };
  writeKv('Inflows by Customer', ['Customer', 'Inflows'], gold.inflowsByCustomer);
  writeKv('Inflows by Billing Category', ['Billing Category', 'Inflows'], gold.inflowsByCategory);
  writeKv('Outflows by Department', ['Department', 'Outflows'], gold.outflowsByDept);
  writeKv('Outflows by Cost Center', ['Cost Center', 'Outflows'], gold.outflowsByCostCenter);
  writeKv('Outflows by Vendor', ['Vendor', 'Outflows'], gold.outflowsByVendor);

  const risks = wb.addWorksheet('Risk Flags');
  headerRow(risks, 1, 1, ['Flag']);
  gold.riskFlags.forEach((flag, i) => (risks.getCell(2 + i, 1).value = flag));

  const assum = wb.addWorksheet('Key Assumptions');
  headerRow(assum, 1, 1, ['#', 'Assumption']);
  gold.assumptions.forEach((a, i) => {
    assum.getCell(2 + i, 1).value = i + 1;
    assum.getCell(2 + i, 2).value = a;
  });

  const mapping = wb.addWorksheet('Source Mapping');
  headerRow(mapping, 1, 1, ['Metric', 'Sources']);
  Object.entries(gold.provenance).forEach(([metric, sources], i) => {
    mapping.getCell(2 + i, 1).value = metric;
    mapping.getCell(2 + i, 2).value = sources.join(' ; ');
  });

  const recon = wb.addWorksheet('Reconciliation Checks');
  headerRow(recon, 1, 1, ['Check', 'A', 'B', 'Diff']);
  gold.reconciliation.forEach((c, i) => {
    recon.getCell(2 + i, 1).value = c.name;
    recon.getCell(2 + i, 2).value = c.a;
    recon.getCell(2 + i, 3).value = c.b;
    recon.getCell(2 + i, 4).value = r2(c.a - c.b);
  });

  await wb.xlsx.writeFile(outPath);
}

// pharos:eof

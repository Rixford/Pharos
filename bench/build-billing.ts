/**
 * Billing workbook builder — the "receivables view" of the mock company.
 * Deliberately messy-but-realistic: grouped 2-row headers, per-customer
 * subtotal rows inside one table, multiple tables per sheet, notes blocks,
 * a hidden mapping tab referenced by VLOOKUPs, named cells, totals rows.
 */
import * as ExcelJS from 'exceljs';
import { Dataset, Invoice } from './data';
import { Manifest, rangeA1, title, bold, headerRow, groupedHeader, notes, f, r2 } from './xlsx-helpers';

export async function buildBillingWorkbook(ds: Dataset, outPath: string): Promise<Manifest> {
  const wb = new ExcelJS.Workbook();
  const m = new Manifest('billing.xlsx');
  const custName = (id: string): string => ds.customers.find((c) => c.id === id)!.name;
  const collectedFor = (inv: Invoice): number =>
    r2(ds.payments.filter((p) => p.invoiceId === inv.id).reduce((s, p) => s + p.amount, 0));

  const totalBilled = r2(ds.invoices.reduce((s, i) => s + i.amount, 0));
  const totalCollected = r2(ds.payments.reduce((s, p) => s + p.amount, 0));
  const totalCredits = r2(ds.credits.reduce((s, c) => s + c.amount, 0));
  const openAR = r2(totalBilled - totalCollected);

  // Sheet order (Executive Summary first; filled after layout decisions).
  const exec = wb.addWorksheet('Executive Summary');
  const detail = wb.addWorksheet('Billing Detail');
  const contracts = wb.addWorksheet('Customer Contracts');
  const invoicesWs = wb.addWorksheet('Invoices');
  const paymentsWs = wb.addWorksheet('Payments');
  const aging = wb.addWorksheet('Aging');
  const creditsWs = wb.addWorksheet('Adjustments & Credits');
  const deferredWs = wb.addWorksheet('Deferred Revenue');
  const assumptions = wb.addWorksheet('Assumptions');
  const catMap = wb.addWorksheet('CategoryMap');
  catMap.state = 'hidden';

  // ── CategoryMap (hidden) ──────────────────────────────────────────────────
  headerRow(catMap, 1, 1, ['Category', 'Revenue Type', 'Code']);
  ds.billingCategories.forEach((c, i) => {
    catMap.getCell(2 + i, 1).value = c.name;
    catMap.getCell(2 + i, 2).value = c.revenueType;
    catMap.getCell(2 + i, 3).value = c.code;
  });
  catMap.getCell('A1').name = 'CategoryMap';
  m.add({ sheet: 'CategoryMap', label: 'Category mapping', range: rangeA1(1, 1, 1 + ds.billingCategories.length, 3), headerRow: 1, dataStart: 2, dataEnd: 1 + ds.billingCategories.length });

  // ── Invoices ─────────────────────────────────────────────────────────────
  title(invoicesWs, 1, 1, 8, 'Invoice Register — H1 2026');
  headerRow(invoicesWs, 3, 1, ['Invoice ID', 'Customer', 'Issue Month', 'Due Month', 'Category', 'Amount', 'Status', 'Memo']);
  ds.invoices.forEach((inv, i) => {
    const row = 4 + i;
    invoicesWs.getCell(row, 1).value = inv.id;
    invoicesWs.getCell(row, 2).value = custName(inv.customerId);
    invoicesWs.getCell(row, 3).value = inv.issueMonth;
    invoicesWs.getCell(row, 4).value = inv.dueMonth;
    invoicesWs.getCell(row, 5).value = inv.category;
    invoicesWs.getCell(row, 6).value = inv.amount;
    invoicesWs.getCell(row, 7).value = inv.status;
    if (inv.memo) invoicesWs.getCell(row, 8).value = inv.memo;
  });
  const invEnd = 3 + ds.invoices.length;
  bold(invoicesWs, invEnd + 2, 1, 'Total Billed');
  invoicesWs.getCell(invEnd + 2, 6).value = f(`SUM(F4:F${invEnd})`, totalBilled);
  notes(invoicesWs, invEnd + 4, 1, [
    'Status legend: Paid = cash fully received · Partial = part collected · Open = not yet due · Overdue = past due, no cash received.'
  ]);
  m.add({ sheet: 'Invoices', label: 'Invoice register', range: rangeA1(3, 1, invEnd + 2, 8), headerRow: 3, dataStart: 4, dataEnd: invEnd });

  // ── Payments ─────────────────────────────────────────────────────────────
  title(paymentsWs, 1, 1, 5, 'Cash Receipts (Payments) — H1 2026');
  headerRow(paymentsWs, 3, 1, ['Payment ID', 'Invoice ID', 'Customer', 'Month Received', 'Amount Received']);
  ds.payments.forEach((p, i) => {
    const row = 4 + i;
    paymentsWs.getCell(row, 1).value = p.id;
    paymentsWs.getCell(row, 2).value = p.invoiceId;
    paymentsWs.getCell(row, 3).value = custName(p.customerId);
    paymentsWs.getCell(row, 4).value = p.month;
    paymentsWs.getCell(row, 5).value = p.amount;
  });
  const payEnd = 3 + ds.payments.length;
  bold(paymentsWs, payEnd + 2, 1, 'Total Received');
  paymentsWs.getCell(payEnd + 2, 5).value = f(`SUM(E4:E${payEnd})`, totalCollected);
  m.add({ sheet: 'Payments', label: 'Payments register', range: rangeA1(3, 1, payEnd + 2, 5), headerRow: 3, dataStart: 4, dataEnd: payEnd });

  // Second table on the same sheet: receipts by month.
  title(paymentsWs, 3, 8, 2, 'Cash Receipts by Month');
  headerRow(paymentsWs, 5, 8, ['Month', 'Total Received']);
  ds.months.forEach((month, i) => {
    const row = 6 + i;
    const total = r2(ds.payments.filter((p) => p.month === month).reduce((s, p) => s + p.amount, 0));
    paymentsWs.getCell(row, 8).value = month;
    paymentsWs.getCell(row, 9).value = f(`SUMIF($D$4:$D$${payEnd},H${row},$E$4:$E$${payEnd})`, total);
  });
  m.add({ sheet: 'Payments', label: 'Cash receipts by month', range: rangeA1(5, 8, 5 + ds.months.length, 9), headerRow: 5, dataStart: 6, dataEnd: 5 + ds.months.length });

  // ── Billing Detail (grouped header + per-customer subtotal rows) ─────────
  title(detail, 1, 1, 8, 'Billing Detail by Customer — H1 2026');
  const subHeaderRow = groupedHeader(
    detail,
    3,
    1,
    [
      { label: 'Invoice Info', span: 3 },
      { label: 'Classification', span: 2 },
      { label: 'Amounts (USD)', span: 3 }
    ],
    ['Invoice', 'Customer', 'Month', 'Category', 'Rev Type', 'Billed', 'Collected', 'Balance']
  );
  let row = subHeaderRow + 1;
  const subtotalCells: number[] = [];
  for (const customer of ds.customers) {
    const invs = ds.invoices.filter((i) => i.customerId === customer.id);
    const blockStart = row;
    for (const inv of invs) {
      const collected = collectedFor(inv);
      detail.getCell(row, 1).value = inv.id;
      detail.getCell(row, 2).value = customer.name;
      detail.getCell(row, 3).value = inv.issueMonth;
      detail.getCell(row, 4).value = inv.category;
      detail.getCell(row, 5).value = f(`VLOOKUP(D${row},CategoryMap!$A$2:$B$6,2,FALSE)`, ds.billingCategories.find((c) => c.name === inv.category)!.revenueType);
      detail.getCell(row, 6).value = inv.amount;
      detail.getCell(row, 7).value = f(`SUMIF(Payments!$B$4:$B$${payEnd},A${row},Payments!$E$4:$E$${payEnd})`, collected);
      detail.getCell(row, 8).value = f(`F${row}-G${row}`, r2(inv.amount - collected));
      row++;
    }
    const blockEnd = row - 1;
    bold(detail, row, 1, `${customer.name} — Subtotal`);
    const bSum = r2(invs.reduce((s, i) => s + i.amount, 0));
    const cSum = r2(invs.reduce((s, i) => s + collectedFor(i), 0));
    detail.getCell(row, 6).value = f(`SUM(F${blockStart}:F${blockEnd})`, bSum);
    detail.getCell(row, 7).value = f(`SUM(G${blockStart}:G${blockEnd})`, cSum);
    detail.getCell(row, 8).value = f(`F${row}-G${row}`, r2(bSum - cSum));
    subtotalCells.push(row);
    row++;
  }
  bold(detail, row, 1, 'GRAND TOTAL');
  detail.getCell(row, 6).value = f(subtotalCells.map((r) => `F${r}`).join('+'), totalBilled);
  detail.getCell(row, 7).value = f(subtotalCells.map((r) => `G${r}`).join('+'), totalCollected);
  detail.getCell(row, 8).value = f(`F${row}-G${row}`, openAR);
  m.add({ sheet: 'Billing Detail', label: 'Billing detail by customer', range: rangeA1(3, 1, row, 8), headerRow: subHeaderRow, dataStart: subHeaderRow + 1, dataEnd: row - 1 });

  // ── Customer Contracts (two tables on one sheet) ─────────────────────────
  title(contracts, 1, 1, 6, 'Customer Contracts');
  headerRow(contracts, 3, 1, ['Customer ID', 'Customer', 'Tier', 'Contract Type', 'Start Month', 'Monthly Plan Value']);
  ds.customers.forEach((c, i) => {
    const r = 4 + i;
    contracts.getCell(r, 1).value = c.id;
    contracts.getCell(r, 2).value = c.name;
    contracts.getCell(r, 3).value = c.tier;
    contracts.getCell(r, 4).value = c.contractType;
    contracts.getCell(r, 5).value = '2026-01';
    contracts.getCell(r, 6).value = c.tier === 'Enterprise' ? 12000 : c.tier === 'Mid-Market' ? 6000 : 2500;
  });
  const contractsEnd = 3 + ds.customers.length;
  m.add({ sheet: 'Customer Contracts', label: 'Contracts table', range: rangeA1(3, 1, contractsEnd, 6), headerRow: 3, dataStart: 4, dataEnd: contractsEnd });

  title(contracts, contractsEnd + 3, 1, 2, 'Contract Type Legend');
  headerRow(contracts, contractsEnd + 5, 1, ['Type', 'Description']);
  [
    ['Subscription', 'Recurring monthly platform fees'],
    ['Project', 'Milestone-billed implementation work'],
    ['Hybrid', 'Subscription base plus project add-ons']
  ].forEach((pair, i) => {
    contracts.getCell(contractsEnd + 6 + i, 1).value = pair[0];
    contracts.getCell(contractsEnd + 6 + i, 2).value = pair[1];
  });

  // ── Aging ────────────────────────────────────────────────────────────────
  title(aging, 1, 1, 6, 'Accounts Receivable Aging — as of 2026-06-30');
  headerRow(aging, 3, 1, ['Customer', 'Current', '1-30 Days', '31-60 Days', '61+ Days', 'Total Open']);
  let agRow = 4;
  for (const customer of ds.customers) {
    const buckets = [0, 0, 0, 0];
    for (const inv of ds.invoices.filter((i) => i.customerId === customer.id)) {
      const balance = r2(inv.amount - collectedFor(inv));
      if (balance <= 0.005) continue;
      const monthsPastIssue = ds.months.length - 1 - ds.months.indexOf(inv.issueMonth);
      if (inv.status === 'Open') buckets[monthsPastIssue <= 1 ? 0 : 1] += balance;
      else if (inv.status === 'Partial') buckets[1] += balance;
      else buckets[monthsPastIssue >= 4 ? 3 : 2] += balance;
    }
    aging.getCell(agRow, 1).value = customer.name;
    buckets.forEach((b, i) => (aging.getCell(agRow, 2 + i).value = r2(b)));
    aging.getCell(agRow, 6).value = f(`SUM(B${agRow}:E${agRow})`, r2(buckets.reduce((s, b) => s + b, 0)));
    agRow++;
  }
  bold(aging, agRow + 1, 1, 'Total');
  aging.getCell(agRow + 1, 6).value = f(`SUM(F4:F${agRow - 1})`, openAR);
  m.add({ sheet: 'Aging', label: 'AR aging matrix', range: rangeA1(3, 1, agRow + 1, 6), headerRow: 3, dataStart: 4, dataEnd: agRow - 1 });

  // ── Adjustments & Credits ────────────────────────────────────────────────
  title(creditsWs, 1, 1, 6, 'Adjustments & Credits — H1 2026');
  headerRow(creditsWs, 3, 1, ['Credit ID', 'Customer', 'Invoice ID', 'Month Applied', 'Credit Amount', 'Reason']);
  ds.credits.forEach((c, i) => {
    const r = 4 + i;
    creditsWs.getCell(r, 1).value = c.id;
    creditsWs.getCell(r, 2).value = custName(c.customerId);
    creditsWs.getCell(r, 3).value = c.invoiceId;
    creditsWs.getCell(r, 4).value = c.month;
    creditsWs.getCell(r, 5).value = c.amount;
    creditsWs.getCell(r, 6).value = c.reason;
  });
  const crEnd = 3 + ds.credits.length;
  bold(creditsWs, crEnd + 2, 1, 'Total Credits');
  creditsWs.getCell(crEnd + 2, 5).value = f(`SUM(E4:E${crEnd})`, totalCredits);
  notes(creditsWs, crEnd + 4, 1, ['Credits reduce cash collected in the month they are applied.']);
  m.add({ sheet: 'Adjustments & Credits', label: 'Credits register', range: rangeA1(3, 1, crEnd + 2, 6), headerRow: 3, dataStart: 4, dataEnd: crEnd });

  // ── Deferred Revenue ─────────────────────────────────────────────────────
  title(deferredWs, 1, 1, 5, 'Deferred Revenue — Annual Prepaid Contracts');
  headerRow(deferredWs, 3, 1, ['Invoice ID', 'Customer', 'Total Contract', 'Cash Received Month', 'Monthly Recognition']);
  ds.deferred.forEach((d, i) => {
    const r = 4 + i;
    deferredWs.getCell(r, 1).value = d.invoiceId;
    deferredWs.getCell(r, 2).value = custName(d.customerId);
    deferredWs.getCell(r, 3).value = d.total;
    deferredWs.getCell(r, 4).value = d.startMonth;
    deferredWs.getCell(r, 5).value = d.monthlyRecognition;
  });
  const defEnd = 3 + ds.deferred.length;
  m.add({ sheet: 'Deferred Revenue', label: 'Deferred contracts', range: rangeA1(3, 1, defEnd, 5), headerRow: 3, dataStart: 4, dataEnd: defEnd });

  title(deferredWs, defEnd + 3, 1, 7, 'Recognition Schedule (informational — not cash)');
  headerRow(deferredWs, defEnd + 5, 1, ['Invoice ID', ...ds.months]);
  ds.deferred.forEach((d, i) => {
    const r = defEnd + 6 + i;
    deferredWs.getCell(r, 1).value = d.invoiceId;
    ds.months.forEach((_, k) => (deferredWs.getCell(r, 2 + k).value = d.monthlyRecognition));
  });
  notes(deferredWs, defEnd + 8 + ds.deferred.length, 1, [
    'Cash is received at invoice payment (January). Recognition spreads over the service period.',
    'Liquidity reporting uses the CASH month, not the recognition month.'
  ]);

  // ── Assumptions ──────────────────────────────────────────────────────────
  title(assumptions, 1, 1, 2, 'Reporting Assumptions — Billing View');
  headerRow(assumptions, 3, 1, ['#', 'Assumption']);
  [
    'Liquidity basis: CASH. Inflows are counted in the month the payment is received.',
    'Monthly inflows = Payments received minus Credits applied in that month.',
    'Payment terms are Net 30; late payments typically land one month after issue.',
    'Partial payments: the uncollected balance remains in Aging until collected.',
    'Annual prepaid contracts collect full cash up front (see Deferred Revenue tab).',
    'Revenue Type classification for each category comes from the hidden CategoryMap tab.'
  ].forEach((text, i) => {
    assumptions.getCell(4 + i, 1).value = i + 1;
    assumptions.getCell(4 + i, 2).value = text;
  });
  m.add({ sheet: 'Assumptions', label: 'Assumptions list', range: rangeA1(3, 1, 9, 2), headerRow: 3, dataStart: 4, dataEnd: 9 });

  // ── Executive Summary ────────────────────────────────────────────────────
  title(exec, 1, 1, 5, 'Apex Components Ltd — Billing & Receivables Summary (H1 2026)');
  bold(exec, 3, 2, 'Key Figures');
  exec.getCell(4, 2).value = 'Total Billed (H1)';
  exec.getCell(4, 3).value = f(`SUM(Invoices!F4:F${invEnd})`, totalBilled);
  exec.getCell(5, 2).value = 'Total Cash Collected';
  exec.getCell(5, 3).value = f(`SUM(Payments!E4:E${payEnd})`, totalCollected);
  exec.getCell(5, 3).name = 'TotalCollected';
  exec.getCell(6, 2).value = 'Credits Issued';
  exec.getCell(6, 3).value = f(`'Adjustments & Credits'!E${crEnd + 2}`, totalCredits);
  exec.getCell(7, 2).value = 'Open AR';
  exec.getCell(7, 3).value = f(`C4-C5`, openAR);
  notes(exec, 9, 2, [
    'Figures roll up from the Invoices, Payments, Aging and Adjustments & Credits tabs.',
    'See the Assumptions tab for the cash-basis rules used in liquidity reporting.'
  ]);
  m.add({ sheet: 'Executive Summary', label: 'KPI block', range: rangeA1(4, 2, 7, 3), dataStart: 4, dataEnd: 7 });

  await wb.xlsx.writeFile(outPath);
  return m;
}

// pharos:eof

/**
 * Cost-center workbook builder — the "spend view" of the same mock company.
 * Dept-grouped cost-center detail with subtotal rows, matrices, a shared
 * allocation tab driven by a hidden mapping, capex, forecast noise, and
 * an assumptions rulebook.
 */
import * as ExcelJS from 'exceljs';
import { Dataset } from './data';
import { Manifest, rangeA1, title, bold, headerRow, notes, f, r2 } from './xlsx-helpers';

export async function buildCostCenterWorkbook(ds: Dataset, outPath: string): Promise<Manifest> {
  const wb = new ExcelJS.Workbook();
  const m = new Manifest('costcenter.xlsx');

  const vendorById = new Map(ds.vendors.map((v) => [v.id, v]));

  const spend = (pred: (v: { vendorId: string; month: string; amount: number }) => boolean): number =>
    r2(ds.vendorSpend.filter(pred).reduce((s, x) => s + x.amount, 0));

  const sharedByMonth = new Map(ds.months.map((mo) => [mo, spend((x) => x.month === mo && vendorById.get(x.vendorId)!.dept === 'SHARED')]));
  const payrollByDeptMonth = new Map<string, number>();
  for (const p of ds.payroll) payrollByDeptMonth.set(`${p.dept}|${p.month}`, p.amount);
  const directDeptMonth = (dept: string, mo: string): number =>
    spend((x) => x.month === mo && vendorById.get(x.vendorId)!.dept === dept);
  const capexFor = (pred: (c: Dataset['capex'][0]) => boolean): number =>
    r2(ds.capex.filter(pred).reduce((s, c) => s + c.amount, 0));

  const totalPayroll = r2(ds.payroll.reduce((s, p) => s + p.amount, 0));
  const totalVendor = r2(ds.vendorSpend.reduce((s, v) => s + v.amount, 0));
  const totalCapex = r2(ds.capex.reduce((s, c) => s + c.amount, 0));
  const totalCashOut = r2(totalPayroll + totalVendor + totalCapex);

  const exec = wb.addWorksheet('Executive Summary');
  const detail = wb.addWorksheet('Cost Center Detail');
  const deptSpend = wb.addWorksheet('Department Spend');
  const payrollWs = wb.addWorksheet('Payroll');
  const vendorWs = wb.addWorksheet('Vendor Spend');
  const opexWs = wb.addWorksheet('Opex');
  const capexWs = wb.addWorksheet('Capex');
  const allocWs = wb.addWorksheet('Allocations');
  const forecastWs = wb.addWorksheet('Forecast');
  const assumptions = wb.addWorksheet('Assumptions');
  const allocMap = wb.addWorksheet('AllocMap');
  allocMap.state = 'hidden';

  // ── AllocMap (hidden): two mapping tables ────────────────────────────────
  headerRow(allocMap, 1, 1, ['Dept Code', 'Department', 'Shared Allocation %']);
  ds.departments.forEach((d, i) => {
    allocMap.getCell(2 + i, 1).value = d.code;
    allocMap.getCell(2 + i, 2).value = d.name;
    allocMap.getCell(2 + i, 3).value = ds.allocShares[d.code];
  });
  allocMap.getCell('A1').name = 'AllocMap';
  const ccMapStart = 4 + ds.departments.length;
  headerRow(allocMap, ccMapStart, 1, ['Cost Center', 'Dept Code', 'Payroll Share']);
  ds.costCenters.forEach((cc, i) => {
    allocMap.getCell(ccMapStart + 1 + i, 1).value = cc.code;
    allocMap.getCell(ccMapStart + 1 + i, 2).value = cc.dept;
    allocMap.getCell(ccMapStart + 1 + i, 3).value = cc.payrollShare;
  });
  m.add({ sheet: 'AllocMap', label: 'Department allocation shares', range: rangeA1(1, 1, 1 + ds.departments.length, 3), headerRow: 1, dataStart: 2, dataEnd: 1 + ds.departments.length });
  m.add({ sheet: 'AllocMap', label: 'Cost center payroll shares', range: rangeA1(ccMapStart, 1, ccMapStart + ds.costCenters.length, 3), headerRow: ccMapStart, dataStart: ccMapStart + 1, dataEnd: ccMapStart + ds.costCenters.length });

  // ── Vendor Spend (register + vendor totals table) ────────────────────────
  title(vendorWs, 1, 1, 6, 'Vendor Spend Register — H1 2026');
  headerRow(vendorWs, 3, 1, ['Vendor ID', 'Vendor', 'Dept', 'Expense Category', 'Month', 'Amount']);
  ds.vendorSpend.forEach((vs, i) => {
    const v = vendorById.get(vs.vendorId)!;
    const r = 4 + i;
    vendorWs.getCell(r, 1).value = vs.vendorId;
    vendorWs.getCell(r, 2).value = v.name;
    vendorWs.getCell(r, 3).value = v.dept;
    vendorWs.getCell(r, 4).value = v.category;
    vendorWs.getCell(r, 5).value = vs.month;
    vendorWs.getCell(r, 6).value = vs.amount;
  });
  const vsEnd = 3 + ds.vendorSpend.length;
  bold(vendorWs, vsEnd + 2, 1, 'Total Vendor Spend');
  vendorWs.getCell(vsEnd + 2, 6).value = f(`SUM(F4:F${vsEnd})`, totalVendor);
  notes(vendorWs, vsEnd + 4, 1, ['Dept = SHARED rows are allocated to departments — see the Allocations tab.']);
  m.add({ sheet: 'Vendor Spend', label: 'Vendor spend register', range: rangeA1(3, 1, vsEnd + 2, 6), headerRow: 3, dataStart: 4, dataEnd: vsEnd });

  title(vendorWs, 3, 9, 2, 'Vendor Totals (H1)');
  headerRow(vendorWs, 5, 9, ['Vendor', 'Total Spend']);
  ds.vendors.forEach((v, i) => {
    const r = 6 + i;
    vendorWs.getCell(r, 9).value = v.name;
    vendorWs.getCell(r, 10).value = f(`SUMIF($B$4:$B$${vsEnd},I${r},$F$4:$F$${vsEnd})`, spend((x) => x.vendorId === v.id));
  });
  m.add({ sheet: 'Vendor Spend', label: 'Vendor totals', range: rangeA1(5, 9, 5 + ds.vendors.length, 10), headerRow: 5, dataStart: 6, dataEnd: 5 + ds.vendors.length });

  // ── Payroll (matrix + headcount table) ───────────────────────────────────
  title(payrollWs, 1, 1, 8, 'Payroll by Department — H1 2026 (cash)');
  headerRow(payrollWs, 3, 1, ['Department', ...ds.months, 'Total']);
  ds.departments.forEach((d, i) => {
    const r = 4 + i;
    payrollWs.getCell(r, 1).value = d.name;
    ds.months.forEach((mo, k) => {
      payrollWs.getCell(r, 2 + k).value = payrollByDeptMonth.get(`${d.code}|${mo}`)!;
    });
    const rowTotal = r2(ds.months.reduce((s, mo) => s + payrollByDeptMonth.get(`${d.code}|${mo}`)!, 0));
    payrollWs.getCell(r, 2 + ds.months.length).value = f(`SUM(B${r}:G${r})`, rowTotal);
  });
  const pEnd = 3 + ds.departments.length;
  bold(payrollWs, pEnd + 1, 1, 'Total');
  ds.months.forEach((mo, k) => {
    const colTotal = r2(ds.departments.reduce((s, d) => s + payrollByDeptMonth.get(`${d.code}|${mo}`)!, 0));
    payrollWs.getCell(pEnd + 1, 2 + k).value = f(`SUM(${String.fromCharCode(66 + k)}4:${String.fromCharCode(66 + k)}${pEnd})`, colTotal);
  });
  payrollWs.getCell(pEnd + 1, 2 + ds.months.length).value = f(`SUM(H4:H${pEnd})`, totalPayroll);
  m.add({ sheet: 'Payroll', label: 'Payroll matrix', range: rangeA1(3, 1, pEnd + 1, 2 + ds.months.length), headerRow: 3, dataStart: 4, dataEnd: pEnd });

  title(payrollWs, 3, 11, 2, 'Avg Headcount');
  headerRow(payrollWs, 5, 11, ['Department', 'Heads']);
  ds.departments.forEach((d, i) => {
    const rows = ds.payroll.filter((p) => p.dept === d.code);
    payrollWs.getCell(6 + i, 11).value = d.name;
    payrollWs.getCell(6 + i, 12).value = Math.round(rows.reduce((s, p) => s + p.headcount, 0) / rows.length);
  });

  // ── Cost Center Detail (dept-grouped with subtotal rows) ─────────────────
  title(detail, 1, 1, 9, 'Cost Center Detail — Monthly Cash Spend');
  headerRow(detail, 3, 1, ['Cost Center', 'Dept Code', 'Department', 'Month', 'Payroll', 'Vendor Direct', 'Allocated Shared', 'Capex', 'Total Cash']);
  let row = 4;
  const deptSubtotalRows: number[] = [];
  for (const d of ds.departments) {
    const blockStart = row;
    for (const cc of ds.costCenters.filter((c) => c.dept === d.code)) {
      for (const mo of ds.months) {
        const pay = r2(payrollByDeptMonth.get(`${d.code}|${mo}`)! * cc.payrollShare);
        const direct = spend((x) => x.month === mo && vendorById.get(x.vendorId)!.costCenter === cc.code);
        const alloc = r2(sharedByMonth.get(mo)! * ds.allocShares[d.code] * cc.payrollShare);
        const cap = capexFor((c) => c.costCenter === cc.code && c.month === mo);
        detail.getCell(row, 1).value = cc.code;
        detail.getCell(row, 2).value = d.code;
        detail.getCell(row, 3).value = f(`VLOOKUP(B${row},AllocMap!$A$2:$B$6,2,FALSE)`, d.name);
        detail.getCell(row, 4).value = mo;
        detail.getCell(row, 5).value = pay;
        detail.getCell(row, 6).value = direct;
        detail.getCell(row, 7).value = alloc;
        detail.getCell(row, 8).value = cap;
        detail.getCell(row, 9).value = f(`SUM(E${row}:H${row})`, r2(pay + direct + alloc + cap));
        row++;
      }
    }
    const blockEnd = row - 1;
    bold(detail, row, 1, `${d.name} — Subtotal`);
    for (let c = 5; c <= 9; c++) {
      const col = String.fromCharCode(64 + c);
      const colSum = r2(
        Array.from({ length: blockEnd - blockStart + 1 }, (_, i) => detail.getCell(blockStart + i, c).value)
          .map((v) => (typeof v === 'object' && v !== null && 'result' in (v as object) ? ((v as { result: number }).result ?? 0) : (v as number) ?? 0))
          .reduce((s: number, x) => s + (typeof x === 'number' ? x : 0), 0)
      );
      detail.getCell(row, c).value = f(`SUM(${col}${blockStart}:${col}${blockEnd})`, colSum);
    }
    deptSubtotalRows.push(row);
    row++;
  }
  bold(detail, row, 1, 'GRAND TOTAL');
  detail.getCell(row, 9).value = f(deptSubtotalRows.map((r) => `I${r}`).join('+'), totalCashOut);
  m.add({ sheet: 'Cost Center Detail', label: 'Cost center detail', range: rangeA1(3, 1, row, 9), headerRow: 3, dataStart: 4, dataEnd: row - 1 });

  // ── Department Spend (matrix over detail) ────────────────────────────────
  title(deptSpend, 1, 1, 8, 'Department Cash Spend — H1 2026');
  headerRow(deptSpend, 3, 1, ['Department', ...ds.months, 'Total']);
  ds.departments.forEach((d, i) => {
    const r = 4 + i;
    deptSpend.getCell(r, 1).value = d.name;
    let rowTotal = 0;
    ds.months.forEach((mo, k) => {
      const val = r2(
        payrollByDeptMonth.get(`${d.code}|${mo}`)! +
          directDeptMonth(d.code, mo) +
          r2(sharedByMonth.get(mo)! * ds.allocShares[d.code]) +
          capexFor((c) => c.dept === d.code && c.month === mo)
      );
      rowTotal = r2(rowTotal + val);
      deptSpend.getCell(r, 2 + k).value = f(
        `SUMIFS('Cost Center Detail'!$I$4:$I$200,'Cost Center Detail'!$B$4:$B$200,"${d.code}",'Cost Center Detail'!$D$4:$D$200,${String.fromCharCode(66 + k)}$3)`,
        val
      );
    });
    deptSpend.getCell(r, 2 + ds.months.length).value = f(`SUM(B${r}:G${r})`, rowTotal);
  });
  const dsEnd = 3 + ds.departments.length;
  bold(deptSpend, dsEnd + 1, 1, 'Total');
  ds.months.forEach((mo, k) => {
    const colTotal = r2(
      ds.departments.reduce(
        (s, d) =>
          s +
          payrollByDeptMonth.get(`${d.code}|${mo}`)! +
          directDeptMonth(d.code, mo) +
          r2(sharedByMonth.get(mo)! * ds.allocShares[d.code]) +
          capexFor((c) => c.dept === d.code && c.month === mo),
        0
      )
    );
    deptSpend.getCell(dsEnd + 1, 2 + k).value = f(`SUM(${String.fromCharCode(66 + k)}4:${String.fromCharCode(66 + k)}${dsEnd})`, colTotal);
  });
  deptSpend.getCell(dsEnd + 1, 2 + ds.months.length).value = f(`SUM(H4:H${dsEnd})`, totalCashOut);
  m.add({ sheet: 'Department Spend', label: 'Department spend matrix', range: rangeA1(3, 1, dsEnd + 1, 2 + ds.months.length), headerRow: 3, dataStart: 4, dataEnd: dsEnd });

  // ── Opex (category × month, excludes capex) ──────────────────────────────
  title(opexWs, 1, 1, 8, 'Operating Expenses by Category — H1 2026');
  headerRow(opexWs, 3, 1, ['Expense Category', ...ds.months, 'Total']);
  const categories = [...new Set(ds.vendors.map((v) => v.category))];
  categories.forEach((cat, i) => {
    const r = 4 + i;
    opexWs.getCell(r, 1).value = cat;
    let rowTotal = 0;
    ds.months.forEach((mo, k) => {
      const val = spend((x) => x.month === mo && vendorById.get(x.vendorId)!.category === cat);
      rowTotal = r2(rowTotal + val);
      opexWs.getCell(r, 2 + k).value = val;
    });
    opexWs.getCell(r, 2 + ds.months.length).value = f(`SUM(B${r}:G${r})`, rowTotal);
  });
  const opEnd = 3 + categories.length;
  bold(opexWs, opEnd + 1, 1, 'Total Opex');
  opexWs.getCell(opEnd + 1, 2 + ds.months.length).value = f(`SUM(H4:H${opEnd})`, totalVendor);
  notes(opexWs, opEnd + 3, 1, ['Capex purchases are tracked separately on the Capex tab.']);
  m.add({ sheet: 'Opex', label: 'Opex by category', range: rangeA1(3, 1, opEnd + 1, 2 + ds.months.length), headerRow: 3, dataStart: 4, dataEnd: opEnd });

  // ── Capex ────────────────────────────────────────────────────────────────
  title(capexWs, 1, 1, 7, 'Capital Expenditure — H1 2026');
  headerRow(capexWs, 3, 1, ['ID', 'Item', 'Vendor', 'Dept', 'Cost Center', 'Month', 'Amount']);
  ds.capex.forEach((c, i) => {
    const r = 4 + i;
    capexWs.getCell(r, 1).value = c.id;
    capexWs.getCell(r, 2).value = c.item;
    capexWs.getCell(r, 3).value = vendorById.get(c.vendorId)!.name;
    capexWs.getCell(r, 4).value = c.dept;
    capexWs.getCell(r, 5).value = c.costCenter;
    capexWs.getCell(r, 6).value = c.month;
    capexWs.getCell(r, 7).value = c.amount;
  });
  const cxEnd = 3 + ds.capex.length;
  bold(capexWs, cxEnd + 2, 1, 'Total Capex');
  capexWs.getCell(cxEnd + 2, 7).value = f(`SUM(G4:G${cxEnd})`, totalCapex);
  notes(capexWs, cxEnd + 4, 1, ['Capex is cash in the purchase month and is excluded from the Opex tab.']);
  m.add({ sheet: 'Capex', label: 'Capex purchases', range: rangeA1(3, 1, cxEnd + 2, 7), headerRow: 3, dataStart: 4, dataEnd: cxEnd });

  // ── Allocations ──────────────────────────────────────────────────────────
  title(allocWs, 1, 1, 8, 'Shared Cost Allocations — H1 2026');
  headerRow(allocWs, 3, 1, ['Month', 'Shared Pool', ...ds.departments.map((d) => d.code), 'Check']);
  ds.months.forEach((mo, i) => {
    const r = 4 + i;
    const pool = sharedByMonth.get(mo)!;
    allocWs.getCell(r, 1).value = mo;
    allocWs.getCell(r, 2).value = f(`SUMIFS('Vendor Spend'!$F$4:$F$${vsEnd},'Vendor Spend'!$C$4:$C$${vsEnd},"SHARED",'Vendor Spend'!$E$4:$E$${vsEnd},A${r})`, pool);
    ds.departments.forEach((d, k) => {
      allocWs.getCell(r, 3 + k).value = f(`$B${r}*VLOOKUP("${d.code}",AllocMap!$A$2:$C$6,3,FALSE)`, r2(pool * ds.allocShares[d.code]));
    });
    allocWs.getCell(r, 3 + ds.departments.length).value = f(`SUM(C${r}:G${r})-B${r}`, 0);
  });
  notes(allocWs, 5 + ds.months.length, 1, [
    'Shared vendors (CloudNine Hosting, OfficeWorks) are allocated to departments using the hidden AllocMap shares.'
  ]);
  m.add({ sheet: 'Allocations', label: 'Shared allocation by month', range: rangeA1(3, 1, 3 + ds.months.length, 3 + ds.departments.length), headerRow: 3, dataStart: 4, dataEnd: 3 + ds.months.length });

  // ── Forecast (noise — clearly excluded from actuals) ─────────────────────
  title(forecastWs, 1, 1, 5, 'Spend Forecast Jul–Sep 2026 (informational)');
  headerRow(forecastWs, 3, 1, ['Department', ...ds.forecastMonths]);
  ds.departments.forEach((d, i) => {
    const r = 4 + i;
    forecastWs.getCell(r, 1).value = d.name;
    ds.forecastMonths.forEach((mo, k) => {
      forecastWs.getCell(r, 2 + k).value = ds.forecast.find((x) => x.dept === d.code && x.month === mo)!.amount;
    });
  });
  notes(forecastWs, 5 + ds.departments.length, 1, ['Forecast is informational only and is NOT part of H1 actuals.']);

  // ── Assumptions ──────────────────────────────────────────────────────────
  title(assumptions, 1, 1, 2, 'Reporting Assumptions — Spend View');
  headerRow(assumptions, 3, 1, ['#', 'Assumption']);
  [
    'Outflow basis: CASH. Monthly outflows = Payroll + Vendor spend + Capex in the month incurred.',
    'Shared vendor costs (Dept = SHARED) are allocated to departments via the hidden AllocMap shares.',
    'Department cash out = payroll + direct vendor spend + allocated shared costs + capex.',
    'Cost-center figures split department payroll and shared allocations using fixed payroll shares (AllocMap).',
    'Vendor totals include capex purchases made through that vendor.',
    'The Forecast tab (Jul–Sep) is informational and excluded from H1 actuals.'
  ].forEach((text, i) => {
    assumptions.getCell(4 + i, 1).value = i + 1;
    assumptions.getCell(4 + i, 2).value = text;
  });
  m.add({ sheet: 'Assumptions', label: 'Assumptions list', range: rangeA1(3, 1, 9, 2), headerRow: 3, dataStart: 4, dataEnd: 9 });

  // ── Executive Summary ────────────────────────────────────────────────────
  title(exec, 1, 1, 5, 'Apex Components Ltd — Cost & Cash-Out Summary (H1 2026)');
  bold(exec, 3, 2, 'Key Figures');
  exec.getCell(4, 2).value = 'Total Payroll';
  exec.getCell(4, 3).value = f(`Payroll!H${pEnd + 1}`, totalPayroll);
  exec.getCell(5, 2).value = 'Total Vendor Spend';
  exec.getCell(5, 3).value = f(`'Vendor Spend'!F${vsEnd + 2}`, totalVendor);
  exec.getCell(6, 2).value = 'Total Capex';
  exec.getCell(6, 3).value = f(`Capex!G${cxEnd + 2}`, totalCapex);
  exec.getCell(7, 2).value = 'Total Cash Out (H1)';
  exec.getCell(7, 3).value = f(`C4+C5+C6`, totalCashOut);
  exec.getCell(7, 3).name = 'TotalCashOut';
  notes(exec, 9, 2, [
    'Cash out rolls up Payroll, Vendor Spend and Capex. Shared vendor costs are allocated on the Allocations tab.',
    'See the Assumptions tab for the cash-basis rules used in liquidity reporting.'
  ]);
  m.add({ sheet: 'Executive Summary', label: 'KPI block', range: rangeA1(4, 2, 7, 3), dataStart: 4, dataEnd: 7 });

  await wb.xlsx.writeFile(outPath);
  return m;
}

// pharos:eof

/**
 * Generates test/fixtures/sample.xlsx — the integration-test workbook.
 *
 * Layout (deterministic, no randomness):
 *   Sales (visible)
 *     A1:F1   merged, bold title "ACME Q1 Sales"
 *     A3:F3   bold headers: Date | Region | Product | Units | Unit Price | Revenue
 *     A4:F33  30 data rows; F is a computed column =D{r}*E{r}
 *     row 34  blank
 *     A35     "Total" (bold)   F35  =SUM(F4:F33)   ← named range TotalRevenue
 *     H3:I7   side table: Region | Target
 *   Rates (hidden)
 *     A1:B5   Region | Rate lookup table
 *   Summary (visible)
 *     B2:C5   key/value block with cross-sheet formulas, a VLOOKUP into the
 *             hidden sheet, and an external-workbook reference
 */
import * as ExcelJS from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';

const REGIONS = ['North', 'South', 'East', 'West'];
const PRODUCTS = ['Anvil', 'Rocket', 'Magnet'];

async function main(): Promise<void> {
  const wb = new ExcelJS.Workbook();

  // ── Sales ────────────────────────────────────────────────────────────────
  const sales = wb.addWorksheet('Sales');
  sales.mergeCells('A1:F1');
  const title = sales.getCell('A1');
  title.value = 'ACME Q1 Sales';
  title.font = { bold: true, size: 14 };

  const headers = ['Date', 'Region', 'Product', 'Units', 'Unit Price', 'Revenue'];
  headers.forEach((h, i) => {
    const cell = sales.getCell(3, i + 1);
    cell.value = h;
    cell.font = { bold: true };
  });

  let totalRevenue = 0;
  for (let i = 0; i < 30; i++) {
    const row = 4 + i;
    const units = ((i * 7) % 50) + 5;
    const price = Math.round((9.99 + (i % 4) * 5) * 100) / 100;
    const revenue = Math.round(units * price * 100) / 100;
    totalRevenue += revenue;
    sales.getCell(row, 1).value = new Date(Date.UTC(2026, 0, 1 + i));
    sales.getCell(row, 2).value = REGIONS[i % 4];
    sales.getCell(row, 3).value = PRODUCTS[i % 3];
    sales.getCell(row, 4).value = units;
    sales.getCell(row, 5).value = price;
    sales.getCell(row, 6).value = { formula: `D${row}*E${row}`, result: revenue };
  }
  totalRevenue = Math.round(totalRevenue * 100) / 100;

  const totalLabel = sales.getCell('A35');
  totalLabel.value = 'Total';
  totalLabel.font = { bold: true };
  sales.getCell('F35').value = { formula: 'SUM(F4:F33)', result: totalRevenue };
  sales.getCell('F35').name = 'TotalRevenue';

  // Side table H3:I7.
  const tH = sales.getCell(3, 8);
  tH.value = 'Region';
  tH.font = { bold: true };
  const tI = sales.getCell(3, 9);
  tI.value = 'Target';
  tI.font = { bold: true };
  REGIONS.forEach((r, i) => {
    sales.getCell(4 + i, 8).value = r;
    sales.getCell(4 + i, 9).value = 5000 + i * 250;
  });

  // ── Rates (hidden) ───────────────────────────────────────────────────────
  const rates = wb.addWorksheet('Rates');
  rates.state = 'hidden';
  const rA = rates.getCell('A1');
  rA.value = 'Region';
  rA.font = { bold: true };
  const rB = rates.getCell('B1');
  rB.value = 'Rate';
  rB.font = { bold: true };
  REGIONS.forEach((r, i) => {
    rates.getCell(2 + i, 1).value = r;
    rates.getCell(2 + i, 2).value = Math.round((0.05 + i * 0.01) * 100) / 100;
  });

  // ── Summary ──────────────────────────────────────────────────────────────
  const summary = wb.addWorksheet('Summary');
  summary.getCell('B2').value = 'Total Revenue';
  summary.getCell('C2').value = { formula: 'Sales!F35', result: totalRevenue };
  summary.getCell('B3').value = 'Commission (10%)';
  summary.getCell('C3').value = { formula: 'C2*0.1', result: Math.round(totalRevenue * 10) / 100 };
  summary.getCell('B4').value = 'North Rate';
  summary.getCell('C4').value = { formula: 'VLOOKUP("North",Rates!A2:B5,2,FALSE)', result: 0.05 };
  summary.getCell('B5').value = 'External FY26';
  summary.getCell('C5').value = { formula: "'[Budget.xlsx]FY26'!B2", result: 0 };

  const outDir = path.join(__dirname, '..', 'test', 'fixtures');
  fs.mkdirSync(outDir, { recursive: true });
  const outPath = path.join(outDir, 'sample.xlsx');
  await wb.xlsx.writeFile(outPath);
  console.log(`fixture written: ${outPath} (total revenue ${totalRevenue})`);

  await makeCollectionFixtures();
}

/**
 * Collection fixtures: three linked workbooks under test/fixtures/collection/.
 *
 *   sales-2026.xlsx    Sales!A1:D15 table (Region|Product|Units|Revenue,
 *                      Revenue computed =C{r}*7.5), totals row 15 with
 *                      D15 =SUM(D2:D13) named GrandTotal
 *   targets.xlsx       Targets!A1:B7 table (Region|Target), totals row 7
 *                      with B7 =SUM(B2:B5) also named GrandTotal
 *   summary-2026.xlsx  Dash!B2:C5 key/value block referencing both books,
 *                      plus one reference to budget-2026.xlsx which is
 *                      deliberately NOT generated (unresolved-external case)
 */
async function makeCollectionFixtures(): Promise<void> {
  const outDir = path.join(__dirname, '..', 'test', 'fixtures', 'collection');
  fs.mkdirSync(outDir, { recursive: true });

  // ── sales-2026.xlsx ──────────────────────────────────────────────────────
  const salesWb = new ExcelJS.Workbook();
  const sales = salesWb.addWorksheet('Sales');
  ['Region', 'Product', 'Units', 'Revenue'].forEach((h, i) => {
    const cell = sales.getCell(1, i + 1);
    cell.value = h;
    cell.font = { bold: true };
  });
  let salesTotal = 0;
  for (let i = 0; i < 12; i++) {
    const row = 2 + i;
    const units = i * 3 + 4;
    const revenue = Math.round(units * 7.5 * 100) / 100;
    salesTotal += revenue;
    sales.getCell(row, 1).value = REGIONS[i % 4];
    sales.getCell(row, 2).value = PRODUCTS[i % 3];
    sales.getCell(row, 3).value = units;
    sales.getCell(row, 4).value = { formula: `C${row}*7.5`, result: revenue };
  }
  salesTotal = Math.round(salesTotal * 100) / 100;
  const salesLabel = sales.getCell('A15');
  salesLabel.value = 'Total';
  salesLabel.font = { bold: true };
  sales.getCell('D15').value = { formula: 'SUM(D2:D13)', result: salesTotal };
  sales.getCell('D15').name = 'GrandTotal';
  await salesWb.xlsx.writeFile(path.join(outDir, 'sales-2026.xlsx'));

  // ── targets.xlsx ─────────────────────────────────────────────────────────
  const targetsWb = new ExcelJS.Workbook();
  const targets = targetsWb.addWorksheet('Targets');
  ['Region', 'Target'].forEach((h, i) => {
    const cell = targets.getCell(1, i + 1);
    cell.value = h;
    cell.font = { bold: true };
  });
  let targetsTotal = 0;
  REGIONS.forEach((r, i) => {
    const value = 500 + i * 150;
    targetsTotal += value;
    targets.getCell(2 + i, 1).value = r;
    targets.getCell(2 + i, 2).value = value;
  });
  const targetsLabel = targets.getCell('A7');
  targetsLabel.value = 'Total';
  targetsLabel.font = { bold: true };
  targets.getCell('B7').value = { formula: 'SUM(B2:B5)', result: targetsTotal };
  targets.getCell('B7').name = 'GrandTotal';
  await targetsWb.xlsx.writeFile(path.join(outDir, 'targets.xlsx'));

  // ── summary-2026.xlsx ────────────────────────────────────────────────────
  const summaryWb = new ExcelJS.Workbook();
  const dash = summaryWb.addWorksheet('Dash');
  dash.getCell('B2').value = 'Total Sales';
  dash.getCell('C2').value = { formula: "'[sales-2026.xlsx]Sales'!D15", result: salesTotal };
  dash.getCell('B3').value = 'North Target';
  dash.getCell('C3').value = {
    formula: 'VLOOKUP("North",\'[targets.xlsx]Targets\'!A2:B5,2,FALSE)',
    result: 500
  };
  dash.getCell('B4').value = 'vs Budget';
  dash.getCell('C4').value = { formula: "C2-'[budget-2026.xlsx]FY'!B2", result: 0 };
  dash.getCell('B5').value = 'Commission';
  dash.getCell('C5').value = { formula: 'C2*0.1', result: Math.round(salesTotal * 10) / 100 };
  await summaryWb.xlsx.writeFile(path.join(outDir, 'summary-2026.xlsx'));

  console.log(`collection fixtures written: ${outDir} (sales total ${salesTotal})`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});

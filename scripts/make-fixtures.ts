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
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});

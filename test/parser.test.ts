import * as ExcelJS from 'exceljs';
import { WorkbookGraph } from '../src';
import { loadFixture } from './helpers';

describe('ExcelParser via WorkbookGraph', () => {
  test('parses all sheets including hidden ones', async () => {
    const graph = await loadFixture();
    const sheets = graph.sheets();
    expect(sheets.map((s) => s.name)).toEqual(['Sales', 'Rates', 'Summary']);
    expect(sheets.find((s) => s.name === 'Rates')?.hidden).toBe(true);
    expect(sheets.find((s) => s.name === 'Sales')?.hidden).toBe(false);
  });

  test('extracts values, types and formulas', async () => {
    const graph = await loadFixture();
    const date = graph.getCell('Sales!A4');
    expect(date?.type).toBe('date');
    expect(date?.valueJson).toBe('2026-01-01');

    const region = graph.getCell('Sales!B4');
    expect(region?.value).toBe('North');

    const revenue = graph.getCell('Sales!F4');
    expect(revenue?.formula).toBe('D4*E4');
    expect(revenue?.type).toBe('number');
    expect(revenue?.value).toBeCloseTo(49.95, 2);

    const total = graph.getCell('Sales!F35');
    expect(total?.formula).toBe('SUM(F4:F33)');
  });

  test('sheet lookups are case-insensitive', async () => {
    const graph = await loadFixture();
    expect(graph.getCell('sales!f4')?.address).toBe('Sales!F4');
  });

  test('captures styles used by heuristics', async () => {
    const graph = await loadFixture();
    expect(graph.getCell('Sales!A3')?.style?.bold).toBe(true);
    expect(graph.getCell('Sales!B7')?.style?.bold).toBeUndefined();
  });

  test('merged cells resolve to their master', async () => {
    const graph = await loadFixture();
    const inspection = graph.inspect('Sales!C1');
    expect(inspection.merged?.range).toBe('Sales!A1:F1');
    expect(inspection.merged?.master).toBe('Sales!A1');
    expect(inspection.merged?.masterValue).toBe('ACME Q1 Sales');
  });

  test('named ranges are read and resolvable', async () => {
    const graph = await loadFixture();
    const names = graph.namesContaining({ sheet: 'Sales', row: 35, col: 6 });
    expect(names).toContain('TotalRevenue');
  });

  test('external references are surfaced as warnings, not edges', async () => {
    const graph = await loadFixture();
    expect(graph.externalRefs).toContain('Budget.xlsx');
    expect(graph.warnings.some((w) => w.includes('Budget.xlsx'))).toBe(true);
  });

  test('shared formulas are translated onto member cells', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 1;
    ws.getCell('A2').value = 2;
    ws.getCell('A3').value = 3;
    ws.fillFormula('B1:B3', 'A1*2', [2, 4, 6]);
    const buf = Buffer.from(await wb.xlsx.writeBuffer());
    const graph = await WorkbookGraph.load(buf);
    expect(graph.getCell('Sheet1!B1')?.formula).toBe('A1*2');
    expect(graph.getCell('Sheet1!B2')?.formula).toBe('A2*2');
    expect(graph.getCell('Sheet1!B3')?.formula).toBe('A3*2');
  });
});

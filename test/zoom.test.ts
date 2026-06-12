import * as ExcelJS from 'exceljs';
import { Collection, WorkbookGraph } from '../src';

/**
 * Zoom-model tests: sections/subtotals, grouped headers, notes blocks,
 * column roles, extractTable (L6), locate (narrowing), sheetMap (L1),
 * question-aware diffusion.
 */

async function buildZoomFixture(): Promise<WorkbookGraph> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Spend');

  // Title.
  ws.mergeCells('A1:E1');
  ws.getCell('A1').value = 'Departmental Spend — H1';
  ws.getCell('A1').font = { bold: true };

  // Two-row grouped header (rows 3–4).
  ws.mergeCells(3, 1, 3, 2);
  ws.getCell(3, 1).value = 'Org';
  ws.getCell(3, 1).font = { bold: true, italic: true };
  ws.mergeCells(3, 3, 3, 5);
  ws.getCell(3, 3).value = 'Amounts';
  ws.getCell(3, 3).font = { bold: true, italic: true };
  ['Dept', 'Month', 'Payroll', 'Vendor', 'Total'].forEach((h, i) => {
    const cell = ws.getCell(4, 1 + i);
    cell.value = h;
    cell.font = { bold: true };
  });

  // Engineering group + subtotal, Sales group + subtotal, grand total.
  const rows: [string, string, number, number][] = [
    ['Engineering', '2026-01', 100, 50],
    ['Engineering', '2026-02', 110, 60],
    ['Sales', '2026-01', 80, 20],
    ['Sales', '2026-02', 90, 30]
  ];
  let r = 5;
  const writeRow = (dept: string, mo: string, pay: number, ven: number): void => {
    ws.getCell(r, 1).value = dept;
    ws.getCell(r, 2).value = mo;
    ws.getCell(r, 3).value = pay;
    ws.getCell(r, 4).value = ven;
    ws.getCell(r, 5).value = { formula: `C${r}+D${r}`, result: pay + ven };
    r++;
  };
  writeRow(...rows[0]);
  writeRow(...rows[1]);
  ws.getCell(r, 1).value = 'Engineering — Subtotal';
  ws.getCell(r, 3).value = { formula: `SUM(C5:C6)`, result: 210 };
  ws.getCell(r, 4).value = { formula: `SUM(D5:D6)`, result: 110 };
  ws.getCell(r, 5).value = { formula: `SUM(E5:E6)`, result: 320 };
  r++;
  writeRow(...rows[2]);
  writeRow(...rows[3]);
  ws.getCell(r, 1).value = 'Sales — Subtotal';
  ws.getCell(r, 3).value = { formula: `SUM(C8:C9)`, result: 170 };
  ws.getCell(r, 5).value = { formula: `SUM(E8:E9)`, result: 220 };
  r++;
  ws.getCell(r, 1).value = 'GRAND TOTAL';
  ws.getCell(r, 5).value = { formula: `E7+E10`, result: 540 };

  // Notes block two rows below the table.
  ws.getCell(r + 2, 1).value = 'Vendor figures exclude one-off capital purchases tracked elsewhere.';
  ws.getCell(r + 2, 1).font = { italic: true };

  // Hidden mapping sheet.
  const map = wb.addWorksheet('RateMap');
  map.state = 'hidden';
  ['Dept', 'Rate'].forEach((h, i) => {
    const cell = map.getCell(1, 1 + i);
    cell.value = h;
    cell.font = { bold: true };
  });
  map.getCell(2, 1).value = 'Engineering';
  map.getCell(2, 2).value = 0.6;
  map.getCell(3, 1).value = 'Sales';
  map.getCell(3, 2).value = 0.4;

  return WorkbookGraph.load(Buffer.from(await wb.xlsx.writeBuffer()));
}

describe('zoom model: sections, grouped headers, notes, roles', () => {
  test('two-row grouped headers combine group and sub labels', async () => {
    const graph = await buildZoomFixture();
    const region = graph.detectRegions('Spend').find((x) => x.kind === 'table')!;
    expect(region.data.headerRows).toEqual([3, 4]);
    expect(region.headers).toEqual(['Org · Dept', 'Org · Month', 'Amounts · Payroll', 'Amounts · Vendor', 'Amounts · Total']);
  });

  test('mid-sentence "totals" wording is not a subtotal row', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Assumptions');
    ['#', 'Assumption'].forEach((h, i) => {
      const cell = ws.getCell(1, 1 + i);
      cell.value = h;
      cell.font = { bold: true };
    });
    ws.getCell(2, 1).value = 1;
    ws.getCell(2, 2).value = 'Vendor totals include capex purchases made through that vendor.';
    ws.getCell(3, 1).value = 2;
    ws.getCell(3, 2).value = 'Forecast figures are informational only.';
    const graph = await WorkbookGraph.load(Buffer.from(await wb.xlsx.writeBuffer()));
    const region = graph.detectRegions('Assumptions')[0];
    expect(region.data.subtotalRows).toBeUndefined();
    const table = graph.extractTable(region);
    expect(table.rows).toHaveLength(2); // the capex rule row survives extraction
  });

  test('subtotal and grand-total rows become sections and are excluded from stats', async () => {
    const graph = await buildZoomFixture();
    const region = graph.detectRegions('Spend').find((x) => x.kind === 'table')!;
    expect(region.data.subtotalRows).toEqual([7, 10]);
    expect(region.data.totalsRow).toBe(11); // GRAND TOTAL row detected as the totals row
    const groups = region.data.sections!.filter((s) => s.kind === 'group');
    expect(groups.map((g) => g.label)).toEqual(['Engineering', 'Sales']);
    expect(region.data.dataRowCount).toBe(4);
    const payroll = region.data.columns.find((c) => c.header?.includes('Payroll'))!;
    expect(payroll.stats?.sum).toBe(380); // 100+110+80+90 — subtotals excluded
  });

  test('notes blocks are classified and attached to their host table', async () => {
    const graph = await buildZoomFixture();
    const regions = graph.detectRegions('Spend');
    expect(regions.some((x) => x.kind === 'notes')).toBe(true);
    const table = regions.find((x) => x.kind === 'table')!;
    expect(table.data.notes?.[0]).toContain('capital purchases');
  });

  test('column roles are inferred', async () => {
    const graph = await buildZoomFixture();
    const region = graph.detectRegions('Spend').find((x) => x.kind === 'table')!;
    const roles = Object.fromEntries(region.data.columns.map((c) => [c.header, c.role]));
    expect(roles['Org · Month']).toBe('month');
    expect(roles['Amounts · Payroll']).toBe('measure');
    expect(roles['Amounts · Total']).toBe('computed');
    expect(roles['Org · Dept']).toBe('category');
  });
});

describe('extractTable (zoom level 6)', () => {
  test('returns complete typed rows with provenance, subtotals excluded', async () => {
    const graph = await buildZoomFixture();
    const region = graph.detectRegions('Spend').find((x) => x.kind === 'table')!;
    const table = graph.extractTable(region);
    expect(table.complete).toBe(true);
    expect(table.rows).toHaveLength(4);
    expect(table.rows[0]['Org · Dept']).toBe('Engineering');
    expect(table.rows[0]['Amounts · Total']).toBe(150);
    expect(table.excludedSubtotalRows).toEqual([7, 10]);
    expect(table.subtotals?.map((x) => x.label)).toEqual(['Engineering — Subtotal', 'Sales — Subtotal']);
    expect(table.subtotals?.[0].values['Amounts · Total']).toBe(320);
    expect(table.rowProvenance[0]).toBe('Spend!A5:E5');
    expect(table.sections?.map((s) => s.label)).toEqual(['Engineering', 'Sales']);
  });

  test('supports paging and column filters', async () => {
    const graph = await buildZoomFixture();
    const region = graph.detectRegions('Spend').find((x) => x.kind === 'table')!;
    const page = graph.extractTable(region, { offset: 2, limit: 1, columns: ['dept', 'total'] });
    expect(page.rows).toHaveLength(1);
    expect(page.complete).toBe(false);
    expect(Object.keys(page.rows[0])).toEqual(['Org · Dept', 'Amounts · Total']);
    expect(page.rows[0]['Org · Dept']).toBe('Sales');
  });

  test('cells on the sample fixture extract fully (no truncation)', async () => {
    const graph = await WorkbookGraph.load(`${__dirname}/fixtures/sample.xlsx`);
    const table = graph.extractTable('Sales!B7');
    expect(table.complete).toBe(true);
    expect(table.totalDataRows).toBe(30);
    expect(table.rows[0]['Region']).toBe('North');
  });
});

describe('locate (question-aware narrowing)', () => {
  test('ranks the payroll-like table first for a spend question', async () => {
    const graph = await buildZoomFixture();
    const hits = graph.locate('how much payroll spend by month per department?');
    expect(hits.length).toBeGreaterThan(0);
    expect(hits[0].kind).toBe('table');
    expect(hits[0].rangeA1).toContain('Spend!A3');
    expect(hits[0].why.length).toBeGreaterThan(0);
  });

  test('data kinds outrank notes at equal relevance; hidden boost works', async () => {
    const graph = await buildZoomFixture();
    const vendorHits = graph.locate('vendor costs');
    expect(vendorHits[0].kind).toBe('table'); // not the notes block mentioning vendors
    const hiddenHits = graph.locate('hidden rates mapping');
    expect(hiddenHits[0].hiddenSheet).toBe(true);
  });

  test('collection locate spans workbooks with labels', async () => {
    const graph = await buildZoomFixture();
    const c = new Collection();
    c.add('a.xlsx', graph);
    c.add('b.xlsx', await buildZoomFixture());
    const hits = c.locate('payroll by department');
    expect(hits[0].workbook).toBeDefined();
    expect(hits.some((h) => h.workbook === 'a.xlsx') && hits.some((h) => h.workbook === 'b.xlsx')).toBe(true);
  });
});

describe('sheetMap (zoom level 1) and question-aware diffusion', () => {
  test('sheetMap lists regions with purposes and notes', async () => {
    const graph = await buildZoomFixture();
    const map = graph.sheetMap('Spend');
    expect(map.regions.length).toBeGreaterThanOrEqual(2);
    const table = map.regions.find((x) => x.kind === 'table')!;
    expect(table.sections).toBe(2);
    expect(map.notes[0]).toContain('capital purchases');
    expect(map.purpose).toBeDefined();
  });

  test('expandContext with a question pulls in semantically matching regions', async () => {
    const graph = await WorkbookGraph.load(`${__dirname}/fixtures/sample.xlsx`);
    const without = graph.expandContext('Summary!C3', { depth: 1, includeTrace: false });
    const withQ = graph.expandContext('Summary!C3', {
      depth: 1,
      includeTrace: false,
      question: 'what are the sales targets per region?'
    });
    const targetRange = 'Sales!H3:I7';
    const hadBefore = without.regions.some((x) => x.rangeA1 === targetRange);
    expect(withQ.regions.some((x) => x.rangeA1 === targetRange)).toBe(true);
    expect(withQ.relations.some((rel) => rel.why.includes('question'))).toBe(true);
    expect(hadBefore).toBe(false); // the question is what pulled it in
  });
});

// pharos:eof

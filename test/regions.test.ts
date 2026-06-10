import { WorkbookGraph } from '../src';
import { FIXTURE, loadFixture } from './helpers';

describe('RegionDetector', () => {
  test('finds the two tables on Sales', async () => {
    const graph = await loadFixture();
    const regions = graph.detectRegions('Sales');
    expect(regions).toHaveLength(2);
    expect(regions.map((r) => r.rangeA1).sort()).toEqual(['Sales!A3:F35', 'Sales!H3:I7']);
  });

  test('main table: headers, title, totals, key column, computed column', async () => {
    const graph = await loadFixture();
    const main = graph.detectRegions('Sales').find((r) => r.rangeA1 === 'Sales!A3:F35')!;
    expect(main.kind).toBe('table');
    expect(main.headers).toEqual(['Date', 'Region', 'Product', 'Units', 'Unit Price', 'Revenue']);
    expect(main.title).toBe('ACME Q1 Sales');
    expect(main.data.totalsRow).toBe(35);
    expect(main.data.dataRowCount).toBe(30);
    expect(main.data.dataStartRow).toBe(4);
    expect(main.data.dataEndRow).toBe(33);

    const dateCol = main.data.columns.find((c) => c.header === 'Date')!;
    expect(dateCol.isKey).toBe(true);
    expect(dateCol.dateRange).toEqual({ min: '2026-01-01', max: '2026-01-30' });

    const revenueCol = main.data.columns.find((c) => c.header === 'Revenue')!;
    expect(revenueCol.formulaTemplate).toBeDefined();
    expect(revenueCol.formulaExample).toBe('=D4*E4');
    expect(revenueCol.stats?.sum).toBeGreaterThan(0);

    expect(main.confidence).toBeGreaterThan(0.7);
  });

  test('side table is detected separately with its own key column', async () => {
    const graph = await loadFixture();
    const side = graph.detectRegions('Sales').find((r) => r.rangeA1 === 'Sales!H3:I7')!;
    expect(side.kind).toBe('table');
    expect(side.headers).toEqual(['Region', 'Target']);
    expect(side.data.columns[0].isKey).toBe(true);
  });

  test('Summary block is classified as keyValue', async () => {
    const graph = await loadFixture();
    const kv = graph.detectRegions('Summary');
    expect(kv).toHaveLength(1);
    expect(kv[0].rangeA1).toBe('Summary!B2:C5');
    expect(kv[0].kind).toBe('keyValue');
  });

  test('regions on hidden sheets are flagged', async () => {
    const graph = await loadFixture();
    const rates = graph.detectRegions('Rates');
    expect(rates).toHaveLength(1);
    expect(rates[0].data.hiddenSheet).toBe(true);
    expect(rates[0].kind).toBe('table');
    expect(rates[0].headers).toEqual(['Region', 'Rate']);
  });

  test('region ids are stable across separate loads', async () => {
    const graph = await loadFixture();
    const again = await WorkbookGraph.load(FIXTURE);
    expect(again.detectRegions('Sales').map((r) => r.id)).toEqual(
      graph.detectRegions('Sales').map((r) => r.id)
    );
  });

  test('regionAt resolves cells to their region; getRegion resolves ids', async () => {
    const graph = await loadFixture();
    const region = graph.regionAt('Sales!D10');
    expect(region?.rangeA1).toBe('Sales!A3:F35');
    expect(graph.getRegion(region!.id)?.id).toBe(region!.id);
    expect(graph.regionAt('Sales!A50')).toBeUndefined();
  });
});

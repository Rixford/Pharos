import { estimateTokens, Region } from '../src';
import { loadFixture } from './helpers';

const getMain = async (): Promise<{ main: Region; graph: Awaited<ReturnType<typeof loadFixture>> }> => {
  const graph = await loadFixture();
  const main = graph.detectRegions('Sales').find((r) => r.rangeA1 === 'Sales!A3:F35')!;
  return { main, graph };
};

describe('Summariser modes', () => {
  test('summary mode: short English description with source range', async () => {
    const { main, graph } = await getMain();
    const s = graph.summariseRegion(main, 'summary');
    expect(s.text).toContain('30 data rows');
    expect(s.text).toContain('ACME Q1 Sales');
    expect(s.text).toContain('Revenue');
    expect(s.text).toContain('totals row at 35');
    expect(s.sourceCells).toContain('Sales!A3:F35');
    expect(s.truncated).toBe(false);
    expect(s.tokens).toBeGreaterThan(0);
  });

  test('compact mode adds per-column statistics', async () => {
    const { main, graph } = await getMain();
    const s = graph.summariseRegion(main, 'compact');
    expect(s.text).toContain('columns:');
    expect(s.text).toMatch(/Units \(D\): number.*sum/);
    expect(s.text).toMatch(/Region \(B\): string.*4 distinct/);
  });

  test('evidence mode cites specific cell addresses', async () => {
    const { main, graph } = await getMain();
    const s = graph.summariseRegion(main, 'evidence');
    expect(s.text).toMatch(/max Revenue [\d,.]+ at Sales!F\d+/);
    expect(s.text).toContain('headers at Sales!A3:F3');
    expect(s.text).toContain('sample row 4:');
    expect(s.text).toMatch(/totals: Sales!F35 =SUM\(F4:F33\)/);
    expect(s.sourceCells.length).toBeGreaterThan(2);
  });

  test('cells mode returns every cell with addresses', async () => {
    const { main, graph } = await getMain();
    const s = graph.summariseRegion(main, 'cells');
    const data = s.data as { cells: { a: string; v: unknown; f?: string }[]; omitted: number };
    expect(data.cells).toHaveLength(188); // 6 headers + 180 data + 2 totals cells
    expect(data.cells[0].a).toBe('Sales!A3');
    expect(data.cells.some((c) => c.f === 'D4*E4')).toBe(true);
    expect(data.omitted).toBe(0);
  });

  test('formulas mode returns formulas, references and templates', async () => {
    const { main, graph } = await getMain();
    const s = graph.summariseRegion(main, 'formulas');
    const data = s.data as {
      formulas: { a: string; f: string; refs: string[] }[];
      templates: { column: string; template: string }[];
    };
    expect(data.formulas).toHaveLength(31); // 30 computed + 1 SUM
    const sum = data.formulas.find((f) => f.a === 'Sales!F35')!;
    expect(sum.refs).toContain('F4:F33');
    expect(data.templates.some((t) => t.column === 'Revenue')).toBe(true);
  });

  test('audit mode includes region metadata and styles', async () => {
    const { main, graph } = await getMain();
    const s = graph.summariseRegion(main, 'audit');
    const data = s.data as { region: { id: string }; cells: { a: string; s?: { bold?: boolean } }[] };
    expect(data.region.id).toBe(main.id);
    const header = data.cells.find((c) => c.a === 'Sales!A3');
    expect(header?.s?.bold).toBe(true);
  });

  test('keyValue regions list their pairs', async () => {
    const graph = await loadFixture();
    const s = graph.summariseRegion('Summary!B2', 'summary');
    expect(s.kind).toBe('keyValue');
    expect(s.text).toContain('pairs:');
    expect(s.text).toContain('Total Revenue');
  });

  test('token budgets force truncation and flag it', async () => {
    const { main, graph } = await getMain();
    const full = graph.summariseRegion(main, 'cells');
    const tiny = graph.summariseRegion(main, 'cells', 60);
    expect(full.truncated).toBe(false);
    expect(tiny.truncated).toBe(true);
    expect(tiny.tokens).toBeLessThan(full.tokens);

    const minimal = graph.summariseRegion(main, 'evidence', 15);
    expect(minimal.truncated).toBe(true);
    expect(minimal.tokens).toBeLessThanOrEqual(40);
  });

  test('estimateTokens approximates 4 chars per token', () => {
    expect(estimateTokens('abcd')).toBe(1);
    expect(estimateTokens('a'.repeat(40))).toBe(10);
    expect(estimateTokens({ a: 1 })).toBeGreaterThan(0);
    expect(estimateTokens('')).toBe(0);
  });
});

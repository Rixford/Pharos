import { loadFixture } from './helpers';

describe('Diffuser.expandContext', () => {
  test('expands from a formula seed across formula edges', async () => {
    const graph = await loadFixture();
    const packet = graph.expandContext('Summary!C2', { depth: 2 });

    expect(packet.seed).toBe('Summary!C2');
    expect(packet.seedCell.formula).toBe('Sales!F35');
    expect(packet.seedCell.regionId).toBeDefined();

    const ranges = packet.regions.map((r) => r.rangeA1);
    expect(ranges).toContain('Summary!B2:C5'); // structural: seed's own region
    expect(ranges).toContain('Sales!A3:F35'); // formula: precedent region

    expect(packet.relations.some((r) => r.type === 'formula')).toBe(true);
    expect(packet.relations.some((r) => r.type === 'structural')).toBe(true);
    expect(packet.trace?.precedents).toBeDefined();
    expect(packet.nextActions.length).toBeGreaterThan(0);
    expect(packet.sourceCells).toContain('Summary!C2');
    expect(packet.tokens).toBeGreaterThan(0);
    expect(packet.tokens).toBeLessThanOrEqual(packet.options.tokenBudget);
  });

  test('depth 0 stays at the seed region', async () => {
    const graph = await loadFixture();
    const packet = graph.expandContext('Summary!C2', { depth: 0 });
    expect(packet.regions).toHaveLength(1);
    expect(packet.regions[0].rangeA1).toBe('Summary!B2:C5');
  });

  test('tiny token budgets truncate and say so', async () => {
    const graph = await loadFixture();
    const packet = graph.expandContext('Summary!C2', { depth: 2, tokenBudget: 150 });
    expect(packet.truncated).toBe(true);
    expect(packet.nextActions.some((a) => /budget|summariseRegion|tracePrecedents/i.test(a))).toBe(true);
  });

  test('maxRegions caps expansion and flags truncation', async () => {
    const graph = await loadFixture();
    const packet = graph.expandContext('Sales!F4', { depth: 2, maxRegions: 1 });
    expect(packet.regions).toHaveLength(1);
    expect(packet.truncated).toBe(true);
  });

  test('empty seeds warn and fall back to the nearest region', async () => {
    const graph = await loadFixture();
    const packet = graph.expandContext('Sales!A50');
    expect(packet.warnings.some((w) => w.includes('empty'))).toBe(true);
    expect(packet.regions.length).toBeGreaterThanOrEqual(1);
  });

  test('hidden-sheet regions reached via formulas are flagged', async () => {
    const graph = await loadFixture();
    const packet = graph.expandContext('Summary!C4', { depth: 2 });
    expect(packet.regions.some((r) => r.rangeA1.startsWith('Rates!'))).toBe(true);
    expect(packet.warnings.some((w) => /hidden/i.test(w))).toBe(true);
  });

  test('deeper regions get coarser summaries', async () => {
    const graph = await loadFixture();
    const packet = graph.expandContext('Summary!C2', { depth: 2, mode: 'evidence' });
    const seedRegion = packet.regions.find((r) => r.rangeA1 === 'Summary!B2:C5');
    const hop = packet.regions.find((r) => r.rangeA1 === 'Sales!A3:F35');
    expect(seedRegion?.mode).toBe('evidence');
    expect(hop?.mode).toBe('compact');
  });

  test('external references surface as warnings', async () => {
    const graph = await loadFixture();
    const packet = graph.expandContext('Summary!C5', { depth: 1 });
    expect(packet.warnings.some((w) => w.includes('Budget.xlsx'))).toBe(true);
  });
});

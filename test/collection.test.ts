import { execFileSync } from 'child_process';
import * as path from 'path';
import * as ExcelJS from 'exceljs';
import { Collection, TraceNode, WorkbookGraph } from '../src';

const DIR = path.join(__dirname, 'fixtures', 'collection');
const SALES = path.join(DIR, 'sales-2026.xlsx');
const TARGETS = path.join(DIR, 'targets.xlsx');
const SUMMARY = path.join(DIR, 'summary-2026.xlsx');
const CLI = path.join(__dirname, '..', 'dist', 'cli', 'index.js');

const collect = (node: TraceNode): TraceNode[] => [node, ...node.children.flatMap(collect)];

let cached: Promise<Collection> | undefined;
const loadCollection = (): Promise<Collection> => {
  cached ??= Collection.load([SUMMARY, SALES, TARGETS]);
  return cached;
};

describe('Collection: loading & addressing', () => {
  test('loads multiple workbooks and resolves qualified addresses', async () => {
    const c = await loadCollection();
    expect(c.workbooks()).toEqual(['summary-2026.xlsx', 'sales-2026.xlsx', 'targets.xlsx']);
    expect(c.defaultWorkbook).toBe('summary-2026.xlsx');

    const q = c.resolveAddress('[sales-2026.xlsx]Sales!D15');
    expect(q.book).toBe('sales-2026.xlsx');
    expect(q.ref).toMatchObject({ sheet: 'Sales', row: 15, col: 4 });

    // Unqualified addresses fall back to the first workbook.
    expect(c.resolveAddress('Dash!C2').book).toBe('summary-2026.xlsx');
    expect(() => c.resolveAddress('[nope.xlsx]A!A1')).toThrow(/not in the collection/);
  });

  test('resolveExternalName matches basenames, with and without paths', async () => {
    const c = await loadCollection();
    expect(c.resolveExternalName('sales-2026.xlsx')).toBe('sales-2026.xlsx');
    expect(c.resolveExternalName('C:\\models\\sales-2026.xlsx')).toBe('sales-2026.xlsx');
    expect(c.resolveExternalName('budget-2026.xlsx')).toBeUndefined();
    expect(c.resolveExternalName('1')).toBeUndefined(); // numeric link-table index
  });
});

describe('Collection: link graph', () => {
  test('overview reports workbooks, resolved links and unresolved externals', async () => {
    const c = await loadCollection();
    const o = c.overview();

    expect(o.workbooks).toHaveLength(3);
    expect(o.workbooks[0]).toMatchObject({ key: 'summary-2026.xlsx', regions: 1 });

    const toSales = o.formulaLinks.find((l) => l.toBook === 'sales-2026.xlsx');
    expect(toSales).toBeDefined();
    expect(toSales!.fromBook).toBe('summary-2026.xlsx');
    expect(toSales!.cells).toContain('[summary-2026.xlsx]Dash!C2');
    expect(toSales!.targets).toContain('Sales!D15');

    const toTargets = o.formulaLinks.find((l) => l.toBook === 'targets.xlsx');
    expect(toTargets!.targets[0]).toContain('Targets!A2:B5');

    expect(o.unresolved).toHaveLength(1);
    expect(o.unresolved[0]).toMatchObject({ external: 'budget-2026.xlsx', refCount: 1 });
    expect(o.warnings.some((w) => w.includes('budget-2026.xlsx'))).toBe(true);
  });

  test('shared defined names across workbooks are detected', async () => {
    const c = await loadCollection();
    const shared = c.sharedNames();
    const grandTotal = shared.find((s) => s.name.toLowerCase() === 'grandtotal');
    expect(grandTotal).toBeDefined();
    expect(grandTotal!.books.map((b) => b.book).sort()).toEqual(['sales-2026.xlsx', 'targets.xlsx']);
  });

  test('data links find lookup-style key overlap across workbooks', async () => {
    const c = await loadCollection();
    const links = c.dataLinks();
    const regionLink = links.find(
      (l) =>
        [l.a.book, l.b.book].sort().join('|') === 'sales-2026.xlsx|targets.xlsx' &&
        l.a.column === 'Region' &&
        l.b.column === 'Region'
    );
    expect(regionLink).toBeDefined();
    expect(regionLink!.shared).toBe(4);
    expect(Math.max(regionLink!.coverageA, regionLink!.coverageB)).toBe(1);
  });

  test('cross-workbook dependents are indexed', async () => {
    const c = await loadCollection();
    expect(c.crossDependentsOf('[sales-2026.xlsx]Sales!D15')).toEqual(['[summary-2026.xlsx]Dash!C2']);
    expect(c.crossDependentsOf('[targets.xlsx]Targets!B3')).toEqual(['[summary-2026.xlsx]Dash!C3']);
    expect(c.crossDependentsOf('[sales-2026.xlsx]Sales!A2')).toEqual([]);
  });
});

describe('Collection: cross-workbook tracing', () => {
  test('precedent traces follow external references into loaded workbooks', async () => {
    const c = await loadCollection();
    const tree = c.tracePrecedents('[summary-2026.xlsx]Dash!C5', 4);
    const addresses = collect(tree).map((n) => n.address);

    expect(tree.address).toBe('[summary-2026.xlsx]Dash!C5');
    expect(addresses).toContain('[summary-2026.xlsx]Dash!C2');
    expect(addresses).toContain('[sales-2026.xlsx]Sales!D15');
    expect(addresses).toContain('[sales-2026.xlsx]Sales!D2:D13');

    const range = collect(tree).find((n) => n.address === '[sales-2026.xlsx]Sales!D2:D13')!;
    expect(range.cellCount).toBe(12);
    expect(range.note).toContain('=C2*7.5');
  });

  test('unresolved externals still appear as stub nodes', async () => {
    const c = await loadCollection();
    const tree = c.tracePrecedents('[summary-2026.xlsx]Dash!C4', 3);
    const external = collect(tree).find((n) => n.kind === 'external');
    expect(external).toBeDefined();
    expect(external!.note).toContain('budget-2026.xlsx');
  });

  test('dependent traces cross workbook boundaries', async () => {
    const c = await loadCollection();
    const tree = c.traceDependents('[sales-2026.xlsx]Sales!D15', 3);
    const addresses = collect(tree).map((n) => n.address);
    expect(addresses).toContain('[summary-2026.xlsx]Dash!C2');
    expect(addresses).toContain('[summary-2026.xlsx]Dash!C4'); // C4 = C2 - external
    expect(addresses).toContain('[summary-2026.xlsx]Dash!C5'); // C5 = C2*0.1
  });

  test('cross-workbook reference cycles terminate with a cycle flag', async () => {
    const mk = async (formula: string): Promise<Buffer> => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = { formula, result: 0 };
      return Buffer.from(await wb.xlsx.writeBuffer());
    };
    const c = new Collection();
    c.add('a.xlsx', await WorkbookGraph.load(await mk("'[b.xlsx]S'!A1+1")));
    c.add('b.xlsx', await WorkbookGraph.load(await mk("'[a.xlsx]S'!A1+1")));
    const tree = c.tracePrecedents('[a.xlsx]S!A1', 8);
    expect(collect(tree).some((n) => n.cycle)).toBe(true);
  });
});

describe('Collection: inspect, find, context', () => {
  test('inspect adds workbook and cross-workbook dependents', async () => {
    const c = await loadCollection();
    const i = c.inspect('[sales-2026.xlsx]Sales!D15');
    expect(i.workbook).toBe('sales-2026.xlsx');
    expect(i.address).toBe('[sales-2026.xlsx]Sales!D15');
    expect(i.formula).toBe('SUM(D2:D13)');
    expect(i.crossDependents).toContain('[summary-2026.xlsx]Dash!C2');
    expect(i.namedRanges).toContain('GrandTotal');
  });

  test('findValue searches every workbook with qualified addresses', async () => {
    const c = await loadCollection();
    const hits = c.findValue('North');
    const books = new Set(hits.map((h) => h.workbook));
    expect(books.has('sales-2026.xlsx')).toBe(true);
    expect(books.has('targets.xlsx')).toBe(true);
    expect(hits.every((h) => h.address.startsWith('['))).toBe(true);
  });

  test('expandContext pulls regions from other workbooks across formula edges', async () => {
    const c = await loadCollection();
    const packet = c.expandContext('[summary-2026.xlsx]Dash!C2', { depth: 2, tokenBudget: 2500 });

    expect(packet.seed).toBe('[summary-2026.xlsx]Dash!C2');
    expect(packet.workbooks).toHaveLength(3);

    const byBook = new Map(packet.regions.map((r) => [r.workbook, r]));
    expect(byBook.has('summary-2026.xlsx')).toBe(true); // structural: seed's region
    expect(byBook.has('sales-2026.xlsx')).toBe(true); // formula: external ref target

    const crossRelation = packet.relations.find((r) => r.to.startsWith('[sales-2026.xlsx]'));
    expect(crossRelation).toBeDefined();
    expect(crossRelation!.type).toBe('formula');

    expect(packet.crossLinks.some((l) => l.external === 'budget-2026.xlsx')).toBe(true);
    expect(packet.nextActions.some((a) => a.includes('budget-2026.xlsx'))).toBe(true);

    // The trace excerpt is the cross-workbook one.
    const traceAddresses = packet.trace?.precedents ? collect(packet.trace.precedents).map((n) => n.address) : [];
    expect(traceAddresses).toContain('[sales-2026.xlsx]Sales!D15');

    expect(packet.sourceCells).toContain('[summary-2026.xlsx]Dash!C2');
    expect(packet.tokens).toBeLessThanOrEqual(2500);
  });

  test('tiny budgets truncate collection packets too', async () => {
    const c = await loadCollection();
    const packet = c.expandContext('[summary-2026.xlsx]Dash!C2', { tokenBudget: 200 });
    expect(packet.truncated).toBe(true);
  });
});

describe('Collection: CLI', () => {
  const runCli = (...args: string[]): string =>
    execFileSync(process.execPath, [CLI, ...args], { encoding: 'utf8' });

  test('collection overview lists workbooks, links and the unresolved external', () => {
    const out = runCli('collection', SUMMARY, SALES, TARGETS);
    expect(out).toContain('Workbooks (3)');
    expect(out).toContain('summary-2026.xlsx → sales-2026.xlsx');
    expect(out).toContain('budget-2026.xlsx');
    expect(out).toContain('GrandTotal');
    expect(out).toMatch(/Data links/);
  });

  test('--links --json is parseable', () => {
    const out = runCli('collection', SUMMARY, SALES, TARGETS, '--links', '--json');
    const parsed = JSON.parse(out);
    expect(parsed.formulaLinks.length).toBeGreaterThanOrEqual(2);
    expect(parsed.unresolved).toHaveLength(1);
    expect(parsed.dataLinks.length).toBeGreaterThanOrEqual(1);
  });

  test('--precedents follows the chain into another workbook', () => {
    const out = runCli('collection', SUMMARY, SALES, '--precedents', '[summary-2026.xlsx]Dash!C5', '--depth', '4');
    expect(out).toContain('[sales-2026.xlsx]Sales!D15');
    expect(out).toContain('[sales-2026.xlsx]Sales!D2:D13');
  });

  test('--context --json returns a collection packet', () => {
    const out = runCli('collection', SUMMARY, SALES, TARGETS, '--context', '[summary-2026.xlsx]Dash!C2', '--json');
    const packet = JSON.parse(out);
    expect(packet.workbooks).toHaveLength(3);
    expect(packet.regions.some((r: { workbook?: string }) => r.workbook === 'sales-2026.xlsx')).toBe(true);
  });

  test('--dependents crosses books; bad workbook errors cleanly', () => {
    const out = runCli('collection', SUMMARY, SALES, '--dependents', '[sales-2026.xlsx]Sales!D15');
    expect(out).toContain('[summary-2026.xlsx]Dash!C2');
    expect(() => runCli('collection', SUMMARY, SALES, '--inspect', '[nope.xlsx]A!A1')).toThrow(/not in the collection/);
  });
});

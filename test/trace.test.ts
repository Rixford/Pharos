import * as ExcelJS from 'exceljs';
import { WorkbookGraph, TraceNode } from '../src';
import { loadFixture } from './helpers';

const collect = (node: TraceNode): TraceNode[] => [node, ...node.children.flatMap(collect)];

describe('formula tracing', () => {
  test('precedents follow chains across sheets down to ranges', async () => {
    const graph = await loadFixture();
    const tree = graph.tracePrecedents('Summary!C3', 4);
    const nodes = collect(tree);
    const addresses = nodes.map((n) => n.address);

    expect(tree.formula).toBe('C2*0.1');
    expect(addresses).toContain('Summary!C2');
    expect(addresses).toContain('Sales!F35');

    const rangeNode = nodes.find((n) => n.kind === 'range' && n.address === 'Sales!F4:F33')!;
    expect(rangeNode.cellCount).toBe(30);
    expect(rangeNode.note).toContain('=D4*E4');

    const representative = rangeNode.children[0];
    expect(representative.formula).toBe('D4*E4');
    expect(representative.note).toContain('representative');
  });

  test('depth limits truncate the trace', async () => {
    const graph = await loadFixture();
    const tree = graph.tracePrecedents('Summary!C3', 1);
    const nodes = collect(tree);
    expect(nodes.some((n) => n.truncated)).toBe(true);
    expect(nodes.map((n) => n.address)).not.toContain('Sales!F4:F33');
  });

  test('external references appear as explicit external nodes', async () => {
    const graph = await loadFixture();
    const tree = graph.tracePrecedents('Summary!C5', 2);
    const external = collect(tree).find((n) => n.kind === 'external');
    expect(external?.note).toContain('Budget.xlsx');
  });

  test('dependents are indexed, including through ranges', async () => {
    const graph = await loadFixture();
    expect(graph.dependentsOf('Sales!F35')).toContain('Summary!C2');
    expect(graph.dependentsOf('Sales!F4')).toContain('Sales!F35');
    expect(graph.dependentsOf('Sales!D4')).toContain('Sales!F4');

    const tree = graph.traceDependents('Sales!F35', 2);
    const addresses = collect(tree).map((n) => n.address);
    expect(addresses).toContain('Summary!C2');
    expect(addresses).toContain('Summary!C3');
  });

  test('circular references terminate with a cycle flag', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = { formula: 'B1', result: 0 };
    ws.getCell('B1').value = { formula: 'A1', result: 0 };
    const graph = await WorkbookGraph.load(Buffer.from(await wb.xlsx.writeBuffer()));
    const tree = graph.tracePrecedents('Sheet1!A1', 6);
    expect(collect(tree).some((n) => n.cycle)).toBe(true);
  });
});

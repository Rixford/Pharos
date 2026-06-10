import { canonicalizeFormula, extractRefs, offsetFormula } from '../src/parser/FormulaParser';
import { FormulaRef } from '../src/core/types';

const refs = (f: string, sheet = 'S1'): FormulaRef[] => extractRefs(f, sheet).refs;

describe('FormulaParser.extractRefs', () => {
  test('simple cell references resolve to the current sheet', () => {
    const r = refs('A1+B2');
    expect(r).toHaveLength(2);
    expect(r[0].range).toMatchObject({ sheet: 'S1', startRow: 1, startCol: 1 });
    expect(r[1].range).toMatchObject({ sheet: 'S1', startRow: 2, startCol: 2 });
  });

  test('absolute references parse correctly', () => {
    const r = refs('$A$1+$B2+C$3');
    expect(r.map((x) => x.raw)).toEqual(['$A$1', '$B2', 'C$3']);
    expect(r[0].range).toMatchObject({ startRow: 1, startCol: 1 });
  });

  test('ranges, whole columns and whole rows', () => {
    const r1 = refs('SUM(A1:B10)');
    expect(r1).toHaveLength(1);
    expect(r1[0].kind).toBe('range');
    expect(r1[0].range).toMatchObject({ startRow: 1, endRow: 10, startCol: 1, endCol: 2 });

    const r2 = refs('SUM(A:A)');
    expect(r2[0].range?.open).toBe('columns');

    const r3 = refs('SUM(2:5)');
    expect(r3[0].range?.open).toBe('rows');
    expect(r3[0].range).toMatchObject({ startRow: 2, endRow: 5 });

    const r4 = refs('SUM($2:$5)');
    expect(r4[0].range?.open).toBe('rows');
  });

  test('sheet-qualified references, quoted and unquoted', () => {
    const r = refs("'Sheet 2'!C3+Sheet3!D4");
    expect(r[0].range?.sheet).toBe('Sheet 2');
    expect(r[1].range?.sheet).toBe('Sheet3');
  });

  test('function names that look like cells are not references', () => {
    expect(refs('LOG10(5)')).toHaveLength(0);
    const r = refs('LOG10(A1)+SUM(B2)');
    expect(r.map((x) => x.raw)).toEqual(['A1', 'B2']);
  });

  test('text inside string literals is ignored', () => {
    expect(refs('"A1"&"B2:C3"')).toHaveLength(0);
    const r = refs('IF(A1="x:y",B2,C3)');
    expect(r.map((x) => x.raw)).toEqual(['A1', 'B2', 'C3']);
    expect(refs('"she said ""A1"""')).toHaveLength(0);
  });

  test('external workbook references', () => {
    const r1 = refs('[Budget.xlsx]FY26!B2');
    expect(r1[0].range?.external).toBe('Budget.xlsx');
    expect(r1[0].range?.sheet).toBe('FY26');

    const r2 = refs("'[Budget.xlsx]FY26'!B2");
    expect(r2[0].range?.external).toBe('Budget.xlsx');
    expect(r2[0].range?.sheet).toBe('FY26');

    const r3 = refs("'C:\\models\\[Book1.xlsx]Sheet1'!A1:A5");
    expect(r3[0].range?.external).toBe('C:\\models\\Book1.xlsx');
    expect(r3[0].range?.sheet).toBe('Sheet1');
  });

  test('defined names are captured as name refs', () => {
    const r = refs('TotalRevenue*2+My_Range.v1');
    expect(r.map((x) => [x.kind, x.name])).toEqual([
      ['name', 'TotalRevenue'],
      ['name', 'My_Range.v1']
    ]);
  });

  test('TRUE/FALSE and numbers are not references', () => {
    const r = refs('IF(TRUE,FALSE,A1)+1.5E+3');
    expect(r.map((x) => x.raw)).toEqual(['A1']);
  });

  test('structured references are recorded with a warning', () => {
    const { refs: r, warnings } = extractRefs('SUM(Table1[[#All],[Revenue]])', 'S1');
    expect(r[0].kind).toBe('structured');
    expect(r[0].name).toBe('Table1');
    expect(warnings.some((w) => w.includes('Structured reference'))).toBe(true);
  });

  test('3-D references traverse the first sheet with a warning', () => {
    const { refs: r, warnings } = extractRefs('SUM(Sheet1:Sheet3!A1)', 'S1');
    expect(r[0].range?.sheet).toBe('Sheet1');
    expect(warnings.some((w) => w.includes('3-D'))).toBe(true);
  });
});

describe('FormulaParser.offsetFormula', () => {
  test('shifts relative references', () => {
    expect(offsetFormula('D4*E4', 'S', 1, 0)).toBe('D5*E5');
    expect(offsetFormula('D4*E4', 'S', 0, 2)).toBe('F4*G4');
  });

  test('preserves absolute parts and sheet prefixes', () => {
    expect(offsetFormula('$D$4*E4', 'S', 2, 3)).toBe('$D$4*H6');
    expect(offsetFormula("Sales!A1+'My Sheet'!B2", 'S', 1, 1)).toBe("Sales!B2+'My Sheet'!C3");
  });

  test('shifts ranges and leaves names/strings alone', () => {
    expect(offsetFormula('SUM(F4:F33)', 'S', 0, 1)).toBe('SUM(G4:G33)');
    expect(offsetFormula('TotalRevenue+"A1"&B2', 'S', 1, 0)).toBe('TotalRevenue+"A1"&B3');
  });
});

describe('FormulaParser.canonicalizeFormula', () => {
  test('same column pattern canonicalises identically across rows', () => {
    const a = canonicalizeFormula('D4*E4', 'S', 4, 6);
    const b = canonicalizeFormula('D9*E9', 'S', 9, 6);
    const c = canonicalizeFormula('D4+E4', 'S', 4, 6);
    expect(a).toBe(b);
    expect(a).not.toBe(c);
  });

  test('is case-insensitive outside references', () => {
    expect(canonicalizeFormula('sum(A1)', 'S', 1, 2)).toBe(canonicalizeFormula('SUM(A1)', 'S', 1, 2));
  });

  test('absolute references stay absolute', () => {
    const a = canonicalizeFormula('$D$4*E4', 'S', 4, 6);
    expect(a).toContain('R4C4');
    expect(a).toContain('R[0]C[-1]');
  });
});

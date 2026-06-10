import {
  colToLetter,
  letterToCol,
  parseCellAddress,
  parseRange,
  formatCell,
  formatRange,
  rangeContains,
  quoteSheet
} from '../src/core/address';

describe('address utilities', () => {
  test('column letters round-trip', () => {
    expect(colToLetter(1)).toBe('A');
    expect(colToLetter(26)).toBe('Z');
    expect(colToLetter(27)).toBe('AA');
    expect(colToLetter(703)).toBe('AAA');
    expect(letterToCol('A')).toBe(1);
    expect(letterToCol('aa')).toBe(27);
    expect(letterToCol('XFD')).toBe(16384);
    expect(() => colToLetter(0)).toThrow();
  });

  test('parses plain, quoted and absolute addresses', () => {
    expect(parseCellAddress('Sales!F35')).toEqual({ sheet: 'Sales', row: 35, col: 6 });
    expect(parseCellAddress("'My Sheet'!B4")).toEqual({ sheet: 'My Sheet', row: 4, col: 2 });
    expect(parseCellAddress("'O''Brien'!A1")).toEqual({ sheet: "O'Brien", row: 1, col: 1 });
    expect(parseCellAddress('$B$4', 'S')).toEqual({ sheet: 'S', row: 4, col: 2 });
    expect(() => parseCellAddress('F35')).toThrow(/no sheet/);
    expect(() => parseCellAddress('!!nope', 'S')).toThrow(/Invalid cell address/);
  });

  test('parses and formats ranges', () => {
    const r = parseRange('Sales!B2:D10');
    expect(r).toMatchObject({ sheet: 'Sales', startRow: 2, startCol: 2, endRow: 10, endCol: 4 });
    expect(formatRange(r)).toBe('Sales!B2:D10');
    expect(rangeContains(r, { sheet: 'sales', row: 5, col: 3 })).toBe(true);
    expect(rangeContains(r, { sheet: 'Sales', row: 1, col: 3 })).toBe(false);

    expect(parseRange('A:B', 'S').open).toBe('columns');
    expect(parseRange('2:5', 'S').open).toBe('rows');

    const reversed = parseRange('D10:B2', 'S');
    expect(reversed.startRow).toBe(2);
    expect(reversed.startCol).toBe(2);

    const single = parseRange('C3', 'S');
    expect(single.startRow).toBe(3);
    expect(single.endRow).toBe(3);
  });

  test('quotes sheet names only when needed', () => {
    expect(quoteSheet('Sales')).toBe('Sales');
    expect(quoteSheet('My Sheet')).toBe("'My Sheet'");
    expect(formatCell({ sheet: 'P&L 2026', row: 1, col: 1 })).toBe("'P&L 2026'!A1");
    expect(formatCell({ sheet: 'Sales', row: 35, col: 6 })).toBe('Sales!F35');
  });
});

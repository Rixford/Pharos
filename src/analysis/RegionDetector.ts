/**
 * RegionDetector — finds coherent tables/blocks on a sheet.
 *
 * Heuristics, in order:
 *  1. Build an occupancy grid (merged ranges count as occupied when their
 *     master cell has content).
 *  2. Flood-fill 8-connected components of non-empty cells; merge any
 *     components whose bounding boxes overlap.
 *  3. Attach satellites: a short all-string single-row component 1–2 blank
 *     rows above a block becomes its *title*; a single-row component with
 *     formulas or a "Total" label exactly one blank row below becomes its
 *     *totals row*.
 *  4. Classify each block (table / keyValue / matrix / list / block),
 *     detect a header row (string coverage + type discontinuity + bold),
 *     profile columns (types, stats, distinct counts, key columns,
 *     repeated-formula templates) and score confidence.
 *
 * Region ids are stable: rg_<hash of sheet|range>.
 */
import { cellKey, colToLetter, formatCell, formatRange } from '../core/address';
import { GridCell, SheetGrid } from '../core/grid';
import {
  CellValueType,
  ColumnProfile,
  JsonScalar,
  NumericStats,
  RangeRef,
  RegionData,
  RegionKind
} from '../core/types';
import { fmtDate, hash36 } from '../core/util';
import { canonicalizeFormula, extractRefs } from '../parser/FormulaParser';

interface Component {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

const NEIGHBOURS = [
  [-1, -1], [-1, 0], [-1, 1],
  [0, -1], [0, 1],
  [1, -1], [1, 0], [1, 1]
];

function jsonValue(cell: GridCell): JsonScalar {
  return cell.value instanceof Date ? fmtDate(cell.value) : cell.value;
}

export function detectRegions(grid: SheetGrid): RegionData[] {
  // 1. Occupancy (merge-aware).
  const occupied = new Set<string>(grid.cells.keys());
  const shadowMaster = new Map<string, string>();
  for (const m of grid.merges) {
    const masterKey = cellKey(m.startRow, m.startCol);
    if (!grid.cells.has(masterKey)) continue;
    for (let r = m.startRow; r <= m.endRow; r++) {
      for (let c = m.startCol; c <= m.endCol; c++) {
        const key = cellKey(r, c);
        occupied.add(key);
        if (key !== masterKey) shadowMaster.set(key, masterKey);
      }
    }
  }

  /** Cell content at (row, col), seeing merged masters through their shadows. */
  const cellAt = (row: number, col: number): GridCell | undefined => {
    const key = cellKey(row, col);
    const direct = grid.cells.get(key);
    if (direct) return direct;
    const master = shadowMaster.get(key);
    return master ? grid.cells.get(master) : undefined;
  };

  // 2. Connected components (8-neighbourhood).
  const visited = new Set<string>();
  const comps: Component[] = [];
  for (const startKey of occupied) {
    if (visited.has(startKey)) continue;
    visited.add(startKey);
    const stack = [startKey];
    const comp: Component = { top: Infinity, bottom: -1, left: Infinity, right: -1 };
    while (stack.length > 0) {
      const key = stack.pop()!;
      const [r, c] = key.split(',').map(Number);
      if (r < comp.top) comp.top = r;
      if (r > comp.bottom) comp.bottom = r;
      if (c < comp.left) comp.left = c;
      if (c > comp.right) comp.right = c;
      for (const [dr, dc] of NEIGHBOURS) {
        const nk = cellKey(r + dr, c + dc);
        if (occupied.has(nk) && !visited.has(nk)) {
          visited.add(nk);
          stack.push(nk);
        }
      }
    }
    comps.push(comp);
  }

  // 2b. Merge components whose bounding boxes overlap.
  let mergedSomething = true;
  while (mergedSomething) {
    mergedSomething = false;
    outer: for (let i = 0; i < comps.length; i++) {
      for (let j = i + 1; j < comps.length; j++) {
        const a = comps[i];
        const b = comps[j];
        if (a.top <= b.bottom && b.top <= a.bottom && a.left <= b.right && b.left <= a.right) {
          comps[i] = {
            top: Math.min(a.top, b.top),
            bottom: Math.max(a.bottom, b.bottom),
            left: Math.min(a.left, b.left),
            right: Math.max(a.right, b.right)
          };
          comps.splice(j, 1);
          mergedSomething = true;
          break outer;
        }
      }
    }
  }
  comps.sort((a, b) => a.top - b.top || a.left - b.left);

  const storedCellsIn = (c: Component): GridCell[] => {
    const out: GridCell[] = [];
    for (const cell of grid.cells.values()) {
      if (cell.row >= c.top && cell.row <= c.bottom && cell.col >= c.left && cell.col <= c.right) {
        out.push(cell);
      }
    }
    return out;
  };

  // 3. Attach titles and totals satellites.
  //
  // Satellites are usually *fragments*: a totals row like `A35:"Total" …
  // F35:=SUM(…)` is several disconnected single-cell components on the same
  // row, so fragments are grouped by row before being matched to a host.
  const isSingleRow = (c: Component): boolean => c.top === c.bottom;
  const colOverlap = (a: Component, b: Component): number =>
    Math.min(a.right, b.right) - Math.max(a.left, b.left) + 1;

  const titles = new Map<Component, Component>();
  const totals = new Map<Component, Component>();
  const consumed = new Set<Component>();

  const singlesByRow = new Map<number, Component[]>();
  for (const c of comps) {
    if (!isSingleRow(c)) continue;
    const list = singlesByRow.get(c.top);
    if (list) list.push(c);
    else singlesByRow.set(c.top, [c]);
  }

  const composite = (members: Component[]): Component => ({
    top: members[0].top,
    bottom: members[0].bottom,
    left: Math.min(...members.map((m) => m.left)),
    right: Math.max(...members.map((m) => m.right))
  });

  for (const host of comps) {
    if (isSingleRow(host) || consumed.has(host)) continue;

    // Totals: fragments exactly one blank row below, within the host's columns.
    if (!totals.has(host)) {
      const row = host.bottom + 2;
      const frags = (singlesByRow.get(row) ?? []).filter(
        (f) => !consumed.has(f) && f.left >= host.left - 1 && f.right <= host.right + 1
      );
      if (frags.length > 0) {
        const merged = composite(frags);
        const cells = frags.flatMap((f) => storedCellsIn(f));
        const hasFormula = cells.some((c) => c.formula);
        const label = cells.find((c) => typeof c.value === 'string');
        const looksTotal =
          hasFormula || (label !== undefined && /total|sum|grand/i.test(String(label.value)));
        if (looksTotal && colOverlap(merged, host) >= 1) {
          totals.set(host, merged);
          for (const f of frags) consumed.add(f);
        }
      }
    }

    // Title: a short all-string row 1–2 blank rows above the host.
    if (!titles.has(host)) {
      for (const offset of [2, 3]) {
        const row = host.top - offset;
        if (row < 1) break;
        const frags = (singlesByRow.get(row) ?? []).filter(
          (f) =>
            !consumed.has(f) &&
            colOverlap(f, host) >= 1 &&
            f.left >= host.left - 1 &&
            f.right <= host.right + 1
        );
        if (frags.length === 0) continue;
        const cells = frags.flatMap((f) => storedCellsIn(f));
        const allStrings =
          cells.length > 0 && cells.length <= 4 && cells.every((c) => c.type === 'string' && !c.formula);
        if (allStrings) {
          titles.set(host, composite(frags));
          for (const f of frags) consumed.add(f);
          break;
        }
      }
    }
  }

  // 4. Analyse each remaining component into a RegionData.
  const regions: RegionData[] = [];
  for (const comp of comps) {
    if (consumed.has(comp)) continue;
    regions.push(analyse(grid, comp, titles.get(comp), totals.get(comp), cellAt, storedCellsIn));
  }
  regions.sort((a, b) => a.range.startRow - b.range.startRow || a.range.startCol - b.range.startCol);
  return regions;
}

function analyse(
  grid: SheetGrid,
  comp: Component,
  titleComp: Component | undefined,
  totalsComp: Component | undefined,
  cellAt: (row: number, col: number) => GridCell | undefined,
  storedCellsIn: (c: Component) => GridCell[]
): RegionData {
  const sheet = grid.name;
  const left = Math.min(comp.left, totalsComp?.left ?? Infinity);
  const right = Math.max(comp.right, totalsComp?.right ?? -1);
  const top = comp.top;
  const bottom = totalsComp ? totalsComp.bottom : comp.bottom;
  const colCount = right - left + 1;
  const boxRows = comp.bottom - comp.top + 1;

  // Header detection on the first row of the main block.
  const headerCells: (GridCell | undefined)[] = [];
  for (let c = left; c <= right; c++) headerCells.push(cellAt(top, c));
  const presentHeader = headerCells.filter((c): c is GridCell => c !== undefined);
  const headerStrings = presentHeader.filter((c) => c.type === 'string');
  const headerDistinct =
    new Set(headerStrings.map((c) => String(c.value).trim().toLowerCase())).size === headerStrings.length;

  let nextRow = -1;
  for (let r = top + 1; r <= comp.bottom; r++) {
    let any = false;
    for (let c = left; c <= right; c++) {
      if (cellAt(r, c)) {
        any = true;
        break;
      }
    }
    if (any) {
      nextRow = r;
      break;
    }
  }
  let belowNonString = 0;
  let belowAllBold = true;
  if (nextRow > 0) {
    const below: GridCell[] = [];
    for (let c = left; c <= right; c++) {
      const cell = cellAt(nextRow, c);
      if (cell) below.push(cell);
    }
    if (below.length > 0) {
      belowNonString = below.filter((c) => c.type !== 'string').length / below.length;
      belowAllBold = below.every((c) => c.style?.bold);
    }
  }
  const headerAllBold = presentHeader.length > 0 && presentHeader.every((c) => c.style?.bold);
  const boldSignal = headerAllBold && !belowAllBold;

  const isHeader =
    colCount >= 2 &&
    boxRows >= 2 &&
    presentHeader.length / colCount >= 0.6 &&
    presentHeader.length > 0 &&
    headerStrings.length / presentHeader.length >= 0.8 &&
    headerDistinct &&
    (belowNonString >= 0.4 || boldSignal);

  const headers = isHeader
    ? Array.from({ length: colCount }, (_, i) => {
        const cell = cellAt(top, left + i);
        const text = cell ? String(cell.value ?? '').trim() : '';
        return text || colToLetter(left + i);
      })
    : undefined;

  const dataStartRow = top + (isHeader ? 1 : 0);
  let dataEndRow = comp.bottom;
  let totalsRow = totalsComp?.top;

  // Internal totals row (contiguous with the data, no blank gap).
  if (totalsRow === undefined && boxRows >= 3 && dataEndRow > dataStartRow) {
    const lastCells: GridCell[] = [];
    for (let c = left; c <= right; c++) {
      const cell = cellAt(dataEndRow, c);
      if (cell) lastCells.push(cell);
    }
    const labelCell = lastCells.find((c) => typeof c.value === 'string');
    const labelSaysTotal = labelCell !== undefined && /^\s*(grand\s+)?(sub)?total/i.test(String(labelCell.value));
    const sumsColumn = lastCells.some((cell) => {
      if (!cell.formula || !/^(SUM|SUBTOTAL)\(/i.test(cell.formula)) return false;
      const { refs } = extractRefs(cell.formula, sheet);
      return refs.some(
        (r) =>
          r.range !== undefined &&
          r.range.startCol === cell.col &&
          r.range.endCol === cell.col &&
          r.range.endRow - r.range.startRow + 1 >= Math.max(2, (dataEndRow - dataStartRow) / 2)
      );
    });
    if (labelSaysTotal || sumsColumn) {
      totalsRow = dataEndRow;
      dataEndRow -= 1;
    }
  }

  const dataRowCount = Math.max(0, dataEndRow - dataStartRow + 1);

  // Column profiles.
  const columns: ColumnProfile[] = [];
  for (let c = left; c <= right; c++) {
    const typeCounts = new Map<CellValueType, number>();
    const distinct = new Set<string>();
    const samples: JsonScalar[] = [];
    let nonEmpty = 0;
    let stats: NumericStats | undefined;
    let dateMin: Date | undefined;
    let dateMax: Date | undefined;
    const templates = new Map<string, { count: number; example: string }>();
    let formulaCount = 0;

    for (let r = dataStartRow; r <= dataEndRow; r++) {
      const cell = grid.cells.get(cellKey(r, c));
      if (!cell) continue;
      nonEmpty++;
      typeCounts.set(cell.type, (typeCounts.get(cell.type) ?? 0) + 1);
      const jv = jsonValue(cell);
      const sv = String(jv);
      if (distinct.size < 10000) distinct.add(sv);
      if (samples.length < 3 && !samples.some((s) => String(s) === sv)) samples.push(jv);
      if (cell.type === 'number' && typeof cell.value === 'number') {
        if (!stats) {
          stats = { sum: 0, min: Infinity, max: -Infinity, mean: 0 };
        }
        stats.sum += cell.value;
        if (cell.value < stats.min) {
          stats.min = cell.value;
          stats.minAt = formatCell({ sheet, row: r, col: c });
        }
        if (cell.value > stats.max) {
          stats.max = cell.value;
          stats.maxAt = formatCell({ sheet, row: r, col: c });
        }
      }
      if (cell.type === 'date' && cell.value instanceof Date) {
        if (!dateMin || cell.value < dateMin) dateMin = cell.value;
        if (!dateMax || cell.value > dateMax) dateMax = cell.value;
      }
      if (cell.formula) {
        formulaCount++;
        const t = canonicalizeFormula(cell.formula, sheet, r, c);
        const entry = templates.get(t);
        if (entry) entry.count++;
        else templates.set(t, { count: 1, example: cell.formula });
      }
    }

    if (stats) {
      const numberCount = typeCounts.get('number') ?? 0;
      stats.mean = numberCount > 0 ? stats.sum / numberCount : 0;
    }

    let dominant: CellValueType | 'mixed' = 'empty';
    let best = 0;
    for (const [t, count] of typeCounts) {
      if (count > best) {
        best = count;
        dominant = t;
      }
    }
    if (nonEmpty > 0 && best / nonEmpty < 0.7) dominant = 'mixed';

    let formulaTemplate: string | undefined;
    let formulaExample: string | undefined;
    if (formulaCount >= 2) {
      const top = [...templates.entries()].sort((a, b) => b[1].count - a[1].count)[0];
      if (top[1].count / formulaCount >= 0.6) {
        formulaTemplate = top[0];
        formulaExample = `=${top[1].example}`;
      }
    }

    columns.push({
      col: c,
      letter: colToLetter(c),
      header: headers?.[c - left],
      type: dominant,
      nonEmpty,
      distinct: distinct.size,
      stats,
      formulaTemplate,
      formulaExample,
      samples,
      dateRange:
        dateMin && dateMax ? { min: fmtDate(dateMin), max: fmtDate(dateMax) } : undefined
    });
  }

  // Key column: leftmost fully-populated, all-distinct column.
  for (const profile of columns) {
    if (
      dataRowCount >= 2 &&
      profile.nonEmpty === dataRowCount &&
      profile.distinct === profile.nonEmpty &&
      profile.type !== 'empty'
    ) {
      profile.isKey = true;
      break;
    }
  }

  // Kind classification.
  let kind: RegionKind;
  if (colCount === 1) kind = 'list';
  else if (isHeader) kind = 'table';
  else if (
    colCount === 2 &&
    columns[0].type === 'string' &&
    columns[1].nonEmpty > 0 &&
    columns[1].type !== 'string'
  ) {
    kind = 'keyValue';
  } else if (colCount >= 2 && boxRows >= 2) {
    let firstColStrings = 0;
    let firstColNonEmpty = 0;
    for (let r = dataStartRow; r <= dataEndRow; r++) {
      const cell = cellAt(r, left);
      if (cell) {
        firstColNonEmpty++;
        if (cell.type === 'string') firstColStrings++;
      }
    }
    let interiorNumeric = 0;
    let interiorNonEmpty = 0;
    for (let r = dataStartRow; r <= dataEndRow; r++) {
      for (let c = left + 1; c <= right; c++) {
        const cell = grid.cells.get(cellKey(r, c));
        if (cell) {
          interiorNonEmpty++;
          if (cell.type === 'number') interiorNumeric++;
        }
      }
    }
    kind =
      firstColNonEmpty > 0 &&
      firstColStrings / firstColNonEmpty >= 0.8 &&
      interiorNonEmpty > 0 &&
      interiorNumeric / interiorNonEmpty >= 0.6
        ? 'matrix'
        : 'block';
  } else {
    kind = 'block';
  }

  // Title.
  let title: string | undefined;
  let titleRange: string | undefined;
  if (titleComp) {
    const titleCells = storedCellsIn(titleComp).sort((a, b) => a.col - b.col);
    title = titleCells
      .map((c) => String(c.value ?? '').trim())
      .filter((s) => s.length > 0)
      .join(' ');
    titleRange = formatRange(
      {
        sheet,
        startRow: titleComp.top,
        endRow: titleComp.bottom,
        startCol: titleComp.left,
        endCol: titleComp.right
      }
    );
  }

  const range: RangeRef = { sheet, startRow: top, endRow: bottom, startCol: left, endCol: right };
  const stored = storedCellsIn({ top, bottom, left, right });
  const cellCount = stored.length;
  const formulaCellCount = stored.filter((c) => c.formula).length;
  const density = cellCount / ((bottom - top + 1) * colCount);

  let typeConsistency = 0;
  const dataCols = columns.filter((c) => c.nonEmpty > 0);
  if (dataCols.length > 0) {
    typeConsistency =
      dataCols.reduce((acc, c) => acc + (c.type === 'mixed' ? 0.4 : 1), 0) / dataCols.length;
  }

  let confidence =
    0.3 +
    (isHeader ? 0.2 : 0) +
    (density >= 0.5 ? 0.1 : 0) +
    (dataRowCount >= 3 ? 0.1 : 0) +
    0.2 * typeConsistency +
    (title ? 0.05 : 0) +
    (kind !== 'block' ? 0.05 : 0);
  if (cellCount === 1) confidence -= 0.35;
  confidence = Math.round(Math.max(0.05, Math.min(0.99, confidence)) * 100) / 100;

  const rangeA1 = formatRange(range);
  return {
    id: `rg_${hash36(`${sheet.toLowerCase()}|${rangeA1}`)}`,
    sheet,
    range,
    rangeA1,
    kind,
    title: title || undefined,
    titleRange,
    headerRow: isHeader ? top : undefined,
    headers,
    columns,
    rowCount: bottom - top + 1,
    colCount,
    dataRowCount,
    dataStartRow,
    dataEndRow,
    totalsRow,
    hiddenSheet: grid.hidden,
    density: Math.round(density * 100) / 100,
    cellCount,
    formulaCellCount,
    confidence
  };
}

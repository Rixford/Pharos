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
  ColumnRole,
  JsonScalar,
  NumericStats,
  RangeRef,
  RegionData,
  RegionKind,
  RegionSection
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

  // 5. Attach notes blocks to the nearest table above them.
  for (const note of regions.filter((r) => r.kind === 'notes')) {
    const host = regions
      .filter(
        (r) =>
          r.kind !== 'notes' &&
          r.range.endRow < note.range.startRow &&
          note.range.startRow - r.range.endRow <= 4 &&
          r.range.startCol <= note.range.endCol &&
          note.range.startCol <= r.range.endCol
      )
      .sort((a, b) => b.range.endRow - a.range.endRow)[0];
    if (host && note.columns.length > 0) {
      const lines: string[] = [];
      for (let rr = note.range.startRow; rr <= note.range.endRow; rr++) {
        for (let cc = note.range.startCol; cc <= note.range.endCol; cc++) {
          const cell = grid.cells.get(cellKey(rr, cc));
          if (cell && typeof cell.value === 'string') lines.push(cell.value);
        }
      }
      if (lines.length > 0) host.notes = [...(host.notes ?? []), ...lines];
    }
  }
  return regions;
}

const PURPOSE_RULES: [RegExp, string][] = [
  [/executive|kpi|key figures/i, 'summary / KPIs'],
  [/payment|receipt|collect|remittance/i, 'payments / cash receipts'],
  [/invoice/i, 'invoice register'],
  [/credit|adjustment/i, 'credits & adjustments'],
  [/aging|ageing|overdue/i, 'receivables aging'],
  [/deferred|recognition/i, 'deferred revenue'],
  [/contract/i, 'customer contracts'],
  [/assumption|policy|basis/i, 'assumptions & policies'],
  [/payroll|salar|wage|headcount|compensation/i, 'payroll'],
  [/vendor|supplier/i, 'vendor spend'],
  [/capex|capital/i, 'capital expenditure'],
  [/alloc/i, 'cost allocations'],
  [/forecast|projection/i, 'forecast'],
  [/cost\s*cent/i, 'cost-center spend'],
  [/department|dept/i, 'department spend'],
  [/map|rate/i, 'reference mapping']
];

function inferPurpose(hay: string, monthHeaders: number, kind: RegionKind): string | undefined {
  for (const [re, label] of PURPOSE_RULES) {
    if (re.test(hay)) return monthHeaders >= 3 ? `${label} (by month)` : label;
  }
  if (kind === 'notes') return 'notes';
  return undefined;
}

function inferRole(p: {
  header?: string;
  type: CellValueType | 'mixed';
  distinct: number;
  nonEmpty: number;
  isKey?: boolean;
  formulaTemplate?: string;
  monthish: number;
}): ColumnRole | undefined {
  if (p.nonEmpty === 0) return undefined;
  if (p.isKey) return 'key';
  if (p.formulaTemplate) return 'computed';
  if (p.type === 'date') return 'date';
  if (p.monthish / p.nonEmpty >= 0.6) return 'month';
  if (p.type === 'number') return 'measure';
  if (p.type === 'string' || p.type === 'mixed') {
    if (p.header && (/(^|\s)(id|code|ref)\b|\bid$/i.test(p.header)) && p.distinct === p.nonEmpty) return 'id';
    if (p.distinct <= Math.max(12, Math.ceil(p.nonEmpty / 3))) return 'category';
    return 'text';
  }
  return undefined;
}

const cleanSectionLabel = (label: string): string =>
  label.replace(/\s*[—–-]*\s*(grand\s*)?(sub)?\s*totals?\s*$/i, '').trim() || label.trim();

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

  const presentRow = (r: number): GridCell[] => {
    const out: GridCell[] = [];
    for (let c = left; c <= right; c++) {
      const cell = cellAt(r, c);
      if (cell) out.push(cell);
    }
    return out;
  };
  const nextNonEmptyRow = (after: number): number => {
    for (let r = after + 1; r <= comp.bottom; r++) {
      if (presentRow(r).length > 0) return r;
    }
    return -1;
  };

  /** Header test: string coverage + distinctness + type discontinuity/bold vs the row below. */
  const testHeader = (hr: number): boolean => {
    if (colCount < 2 || hr >= comp.bottom) return false;
    const cells = presentRow(hr);
    if (cells.length === 0 || cells.length / colCount < 0.6) return false;
    const strings = cells.filter((c) => c.type === 'string');
    if (strings.length / cells.length < 0.8) return false;
    const distinct = new Set(strings.map((c) => String(c.value).trim().toLowerCase())).size === strings.length;
    if (!distinct) return false;
    const nr = nextNonEmptyRow(hr);
    if (nr < 0) return false;
    const below = presentRow(nr);
    if (below.length === 0) return false;
    const belowNonString = below.filter((c) => c.type !== 'string').length / below.length;
    const headerAllBold = cells.every((c) => c.style?.bold);
    const belowAllBold = below.every((c) => c.style?.bold);
    return belowNonString >= 0.4 || (headerAllBold && !belowAllBold);
  };

  // ── header detection: single row, then two-row grouped headers ───────────
  let isHeader = testHeader(top);
  let headerRowIdx = isHeader ? top : undefined;
  let headerRows: number[] | undefined = isHeader ? [top] : undefined;
  let groupRow: number | undefined;
  if (!isHeader && boxRows >= 3) {
    const topRow = presentRow(top);
    const topAllStrings = topRow.length > 0 && topRow.every((c) => c.type === 'string');
    const hasTopMerge = grid.merges.some(
      (mg) =>
        mg.startRow === top &&
        mg.endRow === top &&
        mg.endCol > mg.startCol &&
        mg.startCol >= left &&
        mg.endCol <= right
    );
    if (topAllStrings && hasTopMerge && testHeader(top + 1)) {
      isHeader = true;
      headerRowIdx = top + 1;
      headerRows = [top, top + 1];
      groupRow = top;
    }
  }

  const headers = isHeader
    ? Array.from({ length: colCount }, (_, i) => {
        const col = left + i;
        const sub = cellAt(headerRowIdx!, col);
        const subText = sub ? String(sub.value ?? '').trim() : '';
        let text = subText || colToLetter(col);
        if (groupRow !== undefined) {
          const group = cellAt(groupRow, col);
          const groupText = group ? String(group.value ?? '').trim() : '';
          if (groupText && groupText.toLowerCase() !== subText.toLowerCase()) {
            text = `${groupText} · ${text}`;
          }
        }
        return text;
      })
    : undefined;

  const dataStartRow = isHeader ? headerRowIdx! + 1 : top;
  let dataEndRow = comp.bottom;
  let totalsRow = totalsComp?.top;

  // Internal totals row (contiguous with the data, no blank gap).
  if (totalsRow === undefined && boxRows >= 3 && dataEndRow > dataStartRow) {
    const lastCells = presentRow(dataEndRow);
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

  // ── subtotal rows + grouped sections (zoom level 3) ─────────────────────
  const subtotalRows: number[] = [];
  const subtotalKind = new Map<number, 'subtotal' | 'grandTotal'>();
  if (dataEndRow - dataStartRow >= 2) {
    for (let r = dataStartRow; r <= dataEndRow; r++) {
      const cells = presentRow(r);
      const label = cells.find((c) => typeof c.value === 'string' && String(c.value).trim() !== '');
      const labelText = label ? String(label.value) : '';
      // The label must END at "total/subtotal" — a sentence like "Vendor
      // totals include capex…" in an assumptions list is NOT a subtotal row.
      const labelSub = /(^|\s|—|–|-)(grand\s*)?(sub)?totals?\s*$/i.test(labelText.trim());
      let formulaSub = false;
      if (!labelSub) {
        formulaSub = cells.some((cell) => {
          if (!cell.formula || !/^(SUM|SUBTOTAL)\(/i.test(cell.formula)) return false;
          const { refs } = extractRefs(cell.formula, sheet);
          return refs.some(
            (ref) =>
              ref.range !== undefined &&
              ref.range.startCol === cell.col &&
              ref.range.endCol === cell.col &&
              ref.range.endRow === r - 1 &&
              ref.range.startRow >= dataStartRow &&
              ref.range.endRow - ref.range.startRow >= 1
          );
        });
      }
      if (labelSub || formulaSub) {
        subtotalRows.push(r);
        subtotalKind.set(r, /grand/i.test(labelText) ? 'grandTotal' : 'subtotal');
      }
    }
  }
  const subtotalSet = new Set(subtotalRows);

  const sections: RegionSection[] = [];
  if (subtotalRows.length > 0) {
    let groupStart = dataStartRow;
    for (const sr of subtotalRows) {
      const kind = subtotalKind.get(sr)!;
      const labelCell = presentRow(sr).find((c) => typeof c.value === 'string' && String(c.value).trim() !== '');
      const label = labelCell ? cleanSectionLabel(String(labelCell.value)) : undefined;
      if (kind === 'grandTotal') {
        sections.push({ kind: 'grandTotal', label, startRow: sr, endRow: sr, subtotalRow: sr });
      } else if (sr > groupStart) {
        sections.push({ kind: 'group', label, startRow: groupStart, endRow: sr - 1, subtotalRow: sr });
      } else {
        sections.push({ kind: 'subtotal', label, startRow: sr, endRow: sr, subtotalRow: sr });
      }
      groupStart = sr + 1;
    }
    if (groupStart <= dataEndRow) {
      sections.push({ kind: 'group', startRow: groupStart, endRow: dataEndRow });
    }
  }

  // ── column profiles (subtotal rows excluded from all statistics) ────────
  const columns: ColumnProfile[] = [];
  for (let c = left; c <= right; c++) {
    const typeCounts = new Map<CellValueType, number>();
    const distinct = new Set<string>();
    const samples: JsonScalar[] = [];
    let nonEmpty = 0;
    let monthish = 0;
    let stats: NumericStats | undefined;
    let dateMin: Date | undefined;
    let dateMax: Date | undefined;
    const templates = new Map<string, { count: number; example: string }>();
    let formulaCount = 0;

    for (let r = dataStartRow; r <= dataEndRow; r++) {
      if (subtotalSet.has(r)) continue;
      const cell = grid.cells.get(cellKey(r, c));
      if (!cell) continue;
      nonEmpty++;
      typeCounts.set(cell.type, (typeCounts.get(cell.type) ?? 0) + 1);
      const jv = jsonValue(cell);
      const sv = String(jv);
      if (typeof jv === 'string' && /^\d{4}-\d{2}$/.test(jv.trim())) monthish++;
      if (distinct.size < 10000) distinct.add(sv);
      if (samples.length < 3 && !samples.some((x) => String(x) === sv)) samples.push(jv);
      if (cell.type === 'number' && typeof cell.value === 'number') {
        if (!stats) stats = { sum: 0, min: Infinity, max: -Infinity, mean: 0 };
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
      stats.sum = Math.round(stats.sum * 10000) / 10000;
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
      const topT = [...templates.entries()].sort((a, b) => b[1].count - a[1].count)[0];
      if (topT[1].count / formulaCount >= 0.6) {
        formulaTemplate = topT[0];
        formulaExample = `=${topT[1].example}`;
      }
    }

    const profile: ColumnProfile = {
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
      dateRange: dateMin && dateMax ? { min: fmtDate(dateMin), max: fmtDate(dateMax) } : undefined
    };
    profile.role = inferRole({ ...profile, monthish });
    columns.push(profile);
  }

  const dataRowCount = Math.max(0, dataEndRow - dataStartRow + 1 - subtotalRows.length);

  for (const profile of columns) {
    if (
      dataRowCount >= 2 &&
      profile.nonEmpty === dataRowCount &&
      profile.distinct === profile.nonEmpty &&
      profile.type !== 'empty' &&
      profile.type !== 'number'
    ) {
      profile.isKey = true;
      profile.role = 'key';
      break;
    }
  }

  // ── kind classification ──────────────────────────────────────────────────
  const stored = storedCellsIn({ top, bottom, left, right });
  let kind: RegionKind;
  const allStrings = stored.length > 0 && stored.every((c) => c.type === 'string' && !c.formula);
  const avgLen = allStrings ? stored.reduce((s, c) => s + String(c.value).length, 0) / stored.length : 0;
  const allItalic = allStrings && stored.every((c) => c.style?.italic);
  if (!isHeader && allStrings && colCount <= 2 && bottom - top + 1 <= 4 && (avgLen >= 35 || allItalic)) {
    kind = 'notes';
  } else if (colCount === 1) {
    kind = 'list';
  } else if (isHeader) {
    kind = 'table';
  } else if (colCount === 2 && columns[0].type === 'string' && columns[1].nonEmpty > 0 && columns[1].type !== 'string') {
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

  // ── title ────────────────────────────────────────────────────────────────
  let title: string | undefined;
  let titleRange: string | undefined;
  if (titleComp) {
    const titleCells = storedCellsIn(titleComp).sort((a, b) => a.col - b.col);
    title = titleCells
      .map((c) => String(c.value ?? '').trim())
      .filter((x) => x.length > 0)
      .join(' ');
    titleRange = formatRange({
      sheet,
      startRow: titleComp.top,
      endRow: titleComp.bottom,
      startCol: titleComp.left,
      endCol: titleComp.right
    });
  }

  const range: RangeRef = { sheet, startRow: top, endRow: bottom, startCol: left, endCol: right };
  const cellCount = stored.length;
  const formulaCellCount = stored.filter((c) => c.formula).length;
  const density = cellCount / ((bottom - top + 1) * colCount);

  let typeConsistency = 0;
  const dataCols = columns.filter((c) => c.nonEmpty > 0);
  if (dataCols.length > 0) {
    typeConsistency = dataCols.reduce((acc, c) => acc + (c.type === 'mixed' ? 0.4 : 1), 0) / dataCols.length;
  }

  let confidence =
    0.3 +
    (isHeader ? 0.2 : 0) +
    (density >= 0.5 ? 0.1 : 0) +
    (dataRowCount >= 3 ? 0.1 : 0) +
    0.2 * typeConsistency +
    (title ? 0.05 : 0) +
    (kind !== 'block' ? 0.05 : 0) +
    (sections.length > 0 ? 0.05 : 0);
  if (cellCount === 1) confidence -= 0.35;
  confidence = Math.round(Math.max(0.05, Math.min(0.99, confidence)) * 100) / 100;

  const monthHeaders = (headers ?? []).filter((h) => /^\d{4}-\d{2}$/.test(h)).length;
  const kvSamples =
    kind === 'keyValue' || kind === 'notes' || kind === 'list'
      ? columns[0]?.samples.map((x) => String(x)).join(' ') ?? ''
      : '';
  const purpose =
    kind === 'notes'
      ? 'notes'
      : inferPurpose(
          `${sheet} ${title ?? ''} ${(headers ?? []).join(' ')} ${kvSamples}`.toLowerCase(),
          monthHeaders,
          kind
        );

  const rangeA1 = formatRange(range);
  return {
    id: `rg_${hash36(`${sheet.toLowerCase()}|${rangeA1}`)}`,
    sheet,
    range,
    rangeA1,
    kind,
    title: title || undefined,
    titleRange,
    headerRow: headerRowIdx,
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
    confidence,
    headerRows,
    subtotalRows: subtotalRows.length > 0 ? subtotalRows : undefined,
    sections: sections.length > 0 ? sections : undefined,
    purpose
  };
}

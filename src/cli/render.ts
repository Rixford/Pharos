/**
 * Human-readable renderers for CLI output. JSON mode bypasses these.
 */
import { CellInspection } from '../core/WorkbookGraph';
import { CollectionOverview } from '../core/Collection';
import { ContextPacket, RegionSummary, TraceNode, WorkbookOverview } from '../core/types';

const indent = (text: string, pad: string): string =>
  text
    .split('\n')
    .map((l) => pad + l)
    .join('\n');

export function renderOverview(o: WorkbookOverview): string {
  const lines: string[] = [];
  lines.push(`Sheets (${o.sheets.length}):`);
  for (const s of o.sheets) {
    const vis = s.hidden ? 'hidden ' : 'visible';
    const used = s.usedRangeA1 ?? '(empty)';
    lines.push(
      `  ${s.name.padEnd(18)} ${vis}  ${used.padEnd(12)} ${String(s.cellCount).padStart(6)} cells  ${String(
        s.formulaCellCount
      ).padStart(5)} formulas`
    );
    const regions = o.regionsBySheet[s.name] ?? [];
    for (const r of regions) {
      const title = r.title ? `  “${r.title}”` : '';
      lines.push(
        `      [${r.id}] ${r.rangeA1}  ${r.kind}  ${r.rows}×${r.cols}  conf ${r.confidence}${title}`
      );
    }
  }
  if (o.definedNames.length > 0) {
    lines.push(`Named ranges (${o.definedNames.length}):`);
    for (const dn of o.definedNames.slice(0, 20)) {
      lines.push(`  ${dn.name} → ${dn.ranges.join(', ')}`);
    }
  }
  if (o.externalRefs.length > 0) {
    lines.push(`External workbooks referenced: ${o.externalRefs.join(', ')}`);
  }
  lines.push(`Totals: ${o.totalCells} cells, ${o.totalFormulaCells} formulas across ${o.sheets.length} sheet(s)`);
  if (o.warnings.length > 0) {
    lines.push('Warnings:');
    for (const w of o.warnings.slice(0, 8)) lines.push(`  ! ${w}`);
    if (o.warnings.length > 8) lines.push(`  … ${o.warnings.length - 8} more`);
  }
  return lines.join('\n');
}

export function renderInspection(i: CellInspection): string {
  const lines: string[] = [];
  lines.push(`${i.address}${i.sheetHidden ? '  (on hidden sheet)' : ''}`);
  lines.push(`  value:    ${i.value === null ? '(empty)' : String(i.value)}  [${i.type}]`);
  if (i.formula) lines.push(`  formula:  =${i.formula}`);
  if (i.style) {
    const bits: string[] = [];
    if (i.style.bold) bits.push('bold');
    if (i.style.italic) bits.push('italic');
    if (i.style.numFmt) bits.push(`numFmt "${i.style.numFmt}"`);
    if (i.style.fillColor) bits.push(`fill ${i.style.fillColor}`);
    if (bits.length > 0) lines.push(`  style:    ${bits.join(', ')}`);
  }
  if (i.hyperlink) lines.push(`  link:     ${i.hyperlink}`);
  if (i.merged) {
    lines.push(`  merged:   ${i.merged.range} (master ${i.merged.master} = ${String(i.merged.masterValue)})`);
  }
  if (i.region) {
    const title = i.region.title ? ` — “${i.region.title}”` : '';
    lines.push(`  region:   [${i.region.id}] ${i.region.rangeA1} ${i.region.kind}${title}`);
  }
  if (i.precedents.length > 0) {
    lines.push(`  precedents:`);
    for (const p of i.precedents) {
      const target = p.external ? `external "${p.external}"` : (p.target ?? '(unresolved)');
      lines.push(`    ${p.raw}  →  ${target}`);
    }
  }
  if (i.dependentsTotal > 0) {
    const extra = i.dependentsTotal > i.dependents.length ? ` (+${i.dependentsTotal - i.dependents.length} more)` : '';
    lines.push(`  dependents: ${i.dependents.join(', ')}${extra}`);
  }
  if (i.namedRanges.length > 0) lines.push(`  named ranges: ${i.namedRanges.join(', ')}`);
  if (!i.exists) lines.push('  note: cell is empty (no stored content)');
  return lines.join('\n');
}

export function renderTrace(node: TraceNode): string {
  const lines: string[] = [];
  const label = (n: TraceNode): string => {
    let s = n.address;
    if (n.value !== undefined && n.value !== null && n.kind !== 'range') s += ` = ${String(n.value)}`;
    if (n.formula) s += `  =${n.formula}`;
    if (n.cellCount !== undefined) s += ` (${n.cellCount} cells)`;
    if (n.cycle) s += '  [cycle]';
    if (n.truncated) s += '  […]';
    if (n.note) s += `  — ${n.note}`;
    return s;
  };
  const walk = (n: TraceNode, prefix: string, childPrefix: string): void => {
    lines.push(prefix + label(n));
    n.children.forEach((c, idx) => {
      const last = idx === n.children.length - 1;
      walk(c, childPrefix + (last ? '└─ ' : '├─ '), childPrefix + (last ? '   ' : '│  '));
    });
  };
  walk(node, '', '');
  return lines.join('\n');
}

export function renderRegionSummary(s: RegionSummary, ordinal?: number): string {
  const lines: string[] = [];
  const where = s.workbook ? `[${s.workbook}]${s.rangeA1}` : s.rangeA1;
  const head = `${ordinal !== undefined ? `${ordinal}. ` : ''}[${s.regionId}] ${s.kind} ${where}  (mode ${s.mode}, ~${s.tokens} tokens${s.truncated ? ', TRUNCATED' : ''})`;
  lines.push(head);
  lines.push(indent(s.text, '   '));
  if (s.data !== undefined) {
    lines.push(indent(JSON.stringify(s.data, null, 2), '   '));
  }
  lines.push(indent(`sources: ${s.sourceCells.join(', ')}`, '   '));
  return lines.join('\n');
}

export function renderPacket(p: ContextPacket): string {
  const lines: string[] = [];
  let seedLine = `Seed ${p.seed} = ${p.seedCell.value === null ? '(empty)' : String(p.seedCell.value)} [${p.seedCell.type}]`;
  if (p.seedCell.formula) seedLine += `  =${p.seedCell.formula}`;
  if (p.seedCell.regionId) seedLine += `  region ${p.seedCell.regionId}`;
  if (p.seedCell.namedRanges.length > 0) seedLine += `  names: ${p.seedCell.namedRanges.join(', ')}`;
  lines.push(seedLine);
  lines.push(
    `Options: depth ${p.options.depth} · mode ${p.options.mode} · budget ${p.options.tokenBudget} (used ~${p.tokens}) · truncated: ${p.truncated ? 'YES' : 'no'}`
  );
  lines.push('');
  lines.push(`Regions (${p.regions.length}):`);
  p.regions.forEach((r, i) => {
    lines.push(renderRegionSummary(r, i + 1));
  });
  if (p.trace?.precedents) {
    lines.push('');
    lines.push('Precedent trace:');
    lines.push(indent(renderTrace(p.trace.precedents), '  '));
  }
  if (p.trace?.dependents) {
    lines.push('');
    lines.push('Dependent trace:');
    lines.push(indent(renderTrace(p.trace.dependents), '  '));
  }
  if (p.relations.length > 0) {
    lines.push('');
    lines.push('Relations:');
    for (const r of p.relations) {
      lines.push(`  ${r.from} →(${r.type} ${r.weight}) ${r.to} — ${r.why}`);
    }
  }
  if (p.nextActions.length > 0) {
    lines.push('');
    lines.push('Next actions:');
    p.nextActions.forEach((a, i) => lines.push(`  ${i + 1}. ${a}`));
  }
  if (p.warnings.length > 0) {
    lines.push('');
    lines.push('Warnings:');
    for (const w of p.warnings) lines.push(`  ! ${w}`);
  }
  return lines.join('\n');
}

export function renderCollectionLinks(o: CollectionOverview): string {
  const lines: string[] = [];
  lines.push(`Cross-workbook formula links (${o.formulaLinks.length}):`);
  if (o.formulaLinks.length === 0) lines.push('  (none)');
  for (const l of o.formulaLinks) {
    const target = l.toBook ?? `${l.external}  ⚠ NOT LOADED`;
    lines.push(`  ${l.fromBook} → ${target}   ${l.refCount} ref(s)`);
    lines.push(`      from: ${l.cells.join(', ')}`);
    if (l.targets.length > 0) lines.push(`      into: ${l.targets.join(', ')}`);
  }
  if (o.sharedNames.length > 0) {
    lines.push(`Shared defined names (${o.sharedNames.length}):`);
    for (const sn of o.sharedNames) {
      lines.push(`  ${sn.name}: ${sn.books.map((b) => `${b.book} (${b.ranges.join(', ')})`).join(' · ')}`);
    }
  }
  if (o.dataLinks.length > 0) {
    lines.push(`Data links — regions sharing key values (${o.dataLinks.length}):`);
    for (const dl of o.dataLinks.slice(0, 10)) {
      lines.push(
        `  [${dl.a.book}]${dl.a.rangeA1} (${dl.a.column}) ⋈ [${dl.b.book}]${dl.b.rangeA1} (${dl.b.column}) — ${dl.shared} shared, e.g. ${dl.sample.join(', ')}`
      );
    }
  }
  if (o.unresolved.length > 0) {
    lines.push('Unresolved external workbooks:');
    for (const u of o.unresolved) {
      lines.push(`  ! "${u.external}" — ${u.refCount} ref(s) from ${u.fromBooks.join(', ')} (e.g. ${u.sampleCells.join(', ')})`);
    }
  }
  return lines.join('\n');
}

export function renderCollectionOverview(o: CollectionOverview): string {
  const lines: string[] = [];
  lines.push(`Workbooks (${o.workbooks.length}):`);
  for (const w of o.workbooks) {
    const hidden = w.hiddenSheets > 0 ? ` (${w.hiddenSheets} hidden)` : '';
    lines.push(
      `  ${w.key.padEnd(22)} ${w.sheets} sheet${w.sheets === 1 ? '' : 's'}${hidden}  ${String(w.cells).padStart(6)} cells  ${String(
        w.formulaCells
      ).padStart(4)} formulas  ${w.regions} region${w.regions === 1 ? '' : 's'}`
    );
  }
  lines.push(renderCollectionLinks(o));
  if (o.warnings.length > 0) {
    lines.push('Warnings:');
    for (const w of o.warnings.slice(0, 8)) lines.push(`  ! ${w}`);
  }
  return lines.join('\n');
}

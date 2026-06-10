#!/usr/bin/env node
/**
 * pharos — explore Excel workbooks as navigable graphs.
 *
 *   pharos load      book.xlsx
 *   pharos inspect   book.xlsx 'Sales!F35'
 *   pharos region    book.xlsx --list
 *   pharos region    book.xlsx 'Sales!B7' --mode evidence
 *   pharos context   book.xlsx 'Summary!C2' --depth 2 --budget 1500
 *   pharos precedents book.xlsx 'Summary!C3' --depth 3
 *   pharos dependents book.xlsx 'Sales!F35'
 *   pharos find      book.xlsx "North" --sheet Sales
 */
import { Command, InvalidArgumentError } from 'commander';
import * as fs from 'fs';
import { Collection } from '../core/Collection';
import { WorkbookGraph } from '../core/WorkbookGraph';
import { GRANULARITY_MODES, GranularityMode } from '../core/types';
import {
  renderCollectionLinks,
  renderCollectionOverview,
  renderInspection,
  renderOverview,
  renderPacket,
  renderRegionSummary,
  renderTrace
} from './render';

const pkg = require('../../package.json') as { version: string };

const program = new Command();

program
  .name('pharos')
  .description(
    'Explore Excel workbooks as navigable graphs: detect regions, trace formulas,\n' +
      'and extract token-budgeted context packets for AI agents and analysts.'
  )
  .version(pkg.version);

const parseIntOption = (value: string): number => {
  const n = parseInt(value, 10);
  if (Number.isNaN(n) || n < 0) throw new InvalidArgumentError('expected a non-negative integer');
  return n;
};

const parseMode = (value: string): GranularityMode => {
  if (!GRANULARITY_MODES.includes(value as GranularityMode)) {
    throw new InvalidArgumentError(`expected one of: ${GRANULARITY_MODES.join(', ')}`);
  }
  return value as GranularityMode;
};

async function loadGraph(file: string): Promise<WorkbookGraph> {
  if (!fs.existsSync(file)) {
    throw new Error(`file not found: ${file}`);
  }
  return WorkbookGraph.load(file);
}

const emit = (json: boolean, value: unknown, human: () => string): void => {
  console.log(json ? JSON.stringify(value, null, 2) : human());
};

function run(fn: () => Promise<void>): void {
  fn().catch((err: unknown) => {
    console.error(`pharos: ${err instanceof Error ? err.message : String(err)}`);
    process.exitCode = 1;
  });
}

program
  .command('load')
  .description('parse a workbook and print a structural overview')
  .argument('<file>', 'path to .xlsx workbook')
  .option('--json', 'output JSON')
  .action((file: string, opts: { json?: boolean }) =>
    run(async () => {
      const graph = await loadGraph(file);
      const overview = graph.overview();
      emit(!!opts.json, overview, () => renderOverview(overview));
    })
  );

program
  .command('inspect')
  .description("inspect a cell's value, formula, style, region and references")
  .argument('<file>', 'path to .xlsx workbook')
  .argument('<address>', "cell address, e.g. Sales!F35 or 'My Sheet'!B4")
  .option('--json', 'output JSON')
  .action((file: string, address: string, opts: { json?: boolean }) =>
    run(async () => {
      const graph = await loadGraph(file);
      const inspection = graph.inspect(address);
      emit(!!opts.json, inspection, () => renderInspection(inspection));
    })
  );

program
  .command('region')
  .description('list detected regions, or summarise the region at an address / region id')
  .argument('<file>', 'path to .xlsx workbook')
  .argument('[target]', 'cell address (Sales!B7) or region id (rg_…)')
  .option('--list', 'list all detected regions')
  .option('-s, --sheet <name>', 'restrict --list to one sheet')
  .option('-m, --mode <mode>', `granularity: ${GRANULARITY_MODES.join('|')}`, parseMode, 'compact')
  .option('-b, --budget <tokens>', 'token budget for the summary', parseIntOption)
  .option('--json', 'output JSON')
  .action(
    (
      file: string,
      target: string | undefined,
      opts: { list?: boolean; sheet?: string; mode: GranularityMode; budget?: number; json?: boolean }
    ) =>
      run(async () => {
        const graph = await loadGraph(file);
        if (opts.list || target === undefined) {
          const regions = graph.detectRegions(opts.sheet);
          const briefs = regions.map((r) => r.brief());
          emit(!!opts.json, briefs, () =>
            briefs
              .map(
                (b) =>
                  `[${b.id}] ${b.rangeA1}  ${b.kind}  ${b.rows}×${b.cols}  conf ${b.confidence}${b.title ? `  “${b.title}”` : ''}`
              )
              .join('\n')
          );
          return;
        }
        const summary = graph.summariseRegion(target, opts.mode, opts.budget);
        emit(!!opts.json, summary, () => renderRegionSummary(summary));
      })
  );

program
  .command('context')
  .description('diffuse context outward from a seed cell into a token-budgeted packet')
  .argument('<file>', 'path to .xlsx workbook')
  .argument('<address>', 'seed cell address, e.g. Summary!C2')
  .option('-d, --depth <n>', 'maximum hops from the seed', parseIntOption, 2)
  .option('-m, --mode <mode>', `granularity: ${GRANULARITY_MODES.join('|')}`, parseMode, 'compact')
  .option('-b, --budget <tokens>', 'token budget for the packet', parseIntOption, 2000)
  .option('--max-regions <n>', 'cap on regions included', parseIntOption, 8)
  .option('--no-trace', 'skip the precedent/dependent trace excerpt')
  .option('--json', 'output JSON')
  .action(
    (
      file: string,
      address: string,
      opts: {
        depth: number;
        mode: GranularityMode;
        budget: number;
        maxRegions: number;
        trace: boolean;
        json?: boolean;
      }
    ) =>
      run(async () => {
        const graph = await loadGraph(file);
        const packet = graph.expandContext(address, {
          depth: opts.depth,
          mode: opts.mode,
          tokenBudget: opts.budget,
          maxRegions: opts.maxRegions,
          includeTrace: opts.trace
        });
        emit(!!opts.json, packet, () => renderPacket(packet));
      })
  );

program
  .command('precedents')
  .description('recursively trace what a formula depends on')
  .argument('<file>', 'path to .xlsx workbook')
  .argument('<address>', 'cell address, e.g. Summary!C3')
  .option('-d, --depth <n>', 'recursion depth', parseIntOption, 3)
  .option('--json', 'output JSON')
  .action((file: string, address: string, opts: { depth: number; json?: boolean }) =>
    run(async () => {
      const graph = await loadGraph(file);
      const tree = graph.tracePrecedents(address, opts.depth);
      emit(!!opts.json, tree, () => renderTrace(tree));
    })
  );

program
  .command('dependents')
  .description('recursively trace which cells depend on this one')
  .argument('<file>', 'path to .xlsx workbook')
  .argument('<address>', 'cell address, e.g. Sales!F35')
  .option('-d, --depth <n>', 'recursion depth', parseIntOption, 2)
  .option('--json', 'output JSON')
  .action((file: string, address: string, opts: { depth: number; json?: boolean }) =>
    run(async () => {
      const graph = await loadGraph(file);
      const tree = graph.traceDependents(address, opts.depth);
      emit(!!opts.json, tree, () => renderTrace(tree));
    })
  );

program
  .command('find')
  .description('search cell values (case-insensitive substring, exact number, or regex)')
  .argument('<file>', 'path to .xlsx workbook')
  .argument('<query>', 'text, number, or regex with --regex')
  .option('-s, --sheet <name>', 'restrict to one sheet')
  .option('--regex', 'treat query as a regular expression')
  .option('--limit <n>', 'maximum hits', parseIntOption, 50)
  .option('--json', 'output JSON')
  .action(
    (
      file: string,
      query: string,
      opts: { sheet?: string; regex?: boolean; limit: number; json?: boolean }
    ) =>
      run(async () => {
        const graph = await loadGraph(file);
        let q: string | number | RegExp = query;
        if (opts.regex) q = new RegExp(query, 'i');
        else if (/^-?\d+(\.\d+)?$/.test(query)) q = Number(query);
        const hits = graph.findValue(q, { sheet: opts.sheet, limit: opts.limit });
        emit(!!opts.json, hits, () =>
          hits.length === 0
            ? 'no matches'
            : hits.map((h) => `${h.address}  ${String(h.value)}  [${h.type}]`).join('\n')
        );
      })
  );

program
  .command('collection')
  .description('multi-workbook graph: link overview, cross-workbook traces and context')
  .argument('<files...>', 'two or more .xlsx files (first is the default for unqualified addresses)')
  .option('--links', 'list cross-workbook links (formula refs, shared names, data links)')
  .option('--inspect <address>', 'inspect a cell, e.g. "[book.xlsx]Sheet!A1"')
  .option('--context <address>', 'cross-workbook context packet from a seed cell')
  .option('--precedents <address>', 'precedent trace that follows external refs into loaded workbooks')
  .option('--dependents <address>', 'dependent trace across workbooks')
  .option('--find <query>', 'search values across every workbook')
  .option('-d, --depth <n>', 'trace/diffusion depth', parseIntOption, 3)
  .option('-m, --mode <mode>', `granularity: ${GRANULARITY_MODES.join('|')}`, parseMode, 'compact')
  .option('-b, --budget <tokens>', 'token budget', parseIntOption, 2500)
  .option('--json', 'output JSON')
  .action(
    (
      files: string[],
      opts: {
        links?: boolean;
        inspect?: string;
        context?: string;
        precedents?: string;
        dependents?: string;
        find?: string;
        depth: number;
        mode: GranularityMode;
        budget: number;
        json?: boolean;
      }
    ) =>
      run(async () => {
        for (const f of files) {
          if (!fs.existsSync(f)) throw new Error(`file not found: ${f}`);
        }
        const collection = await Collection.load(files);
        if (opts.inspect) {
          const inspection = collection.inspect(opts.inspect);
          emit(!!opts.json, inspection, () => {
            let text = renderInspection(inspection);
            if (inspection.crossDependents.length > 0) {
              text += `\n  cross-workbook dependents: ${inspection.crossDependents.join(', ')}`;
            }
            return text;
          });
          return;
        }
        if (opts.context) {
          const packet = collection.expandContext(opts.context, {
            depth: opts.depth,
            mode: opts.mode,
            tokenBudget: opts.budget
          });
          emit(!!opts.json, packet, () => renderPacket(packet));
          return;
        }
        if (opts.precedents) {
          const tree = collection.tracePrecedents(opts.precedents, opts.depth);
          emit(!!opts.json, tree, () => renderTrace(tree));
          return;
        }
        if (opts.dependents) {
          const tree = collection.traceDependents(opts.dependents, opts.depth);
          emit(!!opts.json, tree, () => renderTrace(tree));
          return;
        }
        if (opts.find) {
          const hits = collection.findValue(opts.find);
          emit(!!opts.json, hits, () =>
            hits.length === 0
              ? 'no matches'
              : hits.map((h) => `${h.address}  ${String(h.value)}  [${h.type}]`).join('\n')
          );
          return;
        }
        const overview = collection.overview();
        if (opts.links) {
          emit(
            !!opts.json,
            {
              formulaLinks: overview.formulaLinks,
              sharedNames: overview.sharedNames,
              dataLinks: overview.dataLinks,
              unresolved: overview.unresolved
            },
            () => renderCollectionLinks(overview)
          );
          return;
        }
        emit(!!opts.json, overview, () => renderCollectionOverview(overview));
      })
  );

program.addHelpText(
  'after',
  `
Examples:
  pharos load report.xlsx
  pharos inspect report.xlsx 'Sales!F35'
  pharos region report.xlsx --list
  pharos region report.xlsx rg_1a2b3c --mode cells --budget 800
  pharos context report.xlsx 'Summary!C2' --depth 2 --mode evidence --budget 1500
  pharos precedents report.xlsx 'Summary!C3' --depth 4
  pharos dependents report.xlsx 'Sales!F35'
  pharos find report.xlsx 24118.2

Collections (multi-workbook):
  pharos collection summary.xlsx sales.xlsx targets.xlsx
  pharos collection summary.xlsx sales.xlsx --links
  pharos collection summary.xlsx sales.xlsx --context '[summary.xlsx]Dash!C2'
  pharos collection summary.xlsx sales.xlsx --precedents '[summary.xlsx]Dash!C5' --depth 4
  pharos collection summary.xlsx sales.xlsx --dependents '[sales.xlsx]Sales!D15'`
);

program.parseAsync(process.argv);

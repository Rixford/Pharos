/**
 * Benchmark orchestrator — the closed loop:
 *   dataset(seed) → source workbooks → gold (from raw data, never the
 *   files) → blind solve (Pharos only) → score → report.
 *
 * The gold report and manifests live under out/<seed>/gold/ and are only
 * produced *after* the solver inputs; the scripted solver receives just
 * the two source paths.
 */
import * as fs from 'fs';
import * as path from 'path';
import { generateDataset } from './data';
import { buildBillingWorkbook } from './build-billing';
import { buildCostCenterWorkbook } from './build-costcenter';
import { computeGold, writeGoldWorkbook, GoldReport } from './gold';
import { scoreCandidate, formatScore, ScoreReport } from './score';
import { solveBlind } from './solve';

export interface BenchResult {
  seed: number;
  dir: string;
  gold: GoldReport;
  score: ScoreReport;
  solverTokens: number;
  solverWarnings: string[];
}

export async function runBench(seed: number, outRoot = path.join(__dirname, 'out')): Promise<BenchResult> {
  const dir = path.join(outRoot, `seed-${seed}`);
  const sourcesDir = path.join(dir, 'sources');
  const goldDir = path.join(dir, 'gold');
  const blindDir = path.join(dir, 'blind');
  fs.rmSync(dir, { recursive: true, force: true });
  for (const d of [sourcesDir, goldDir, blindDir]) fs.mkdirSync(d, { recursive: true });

  const ds = generateDataset(seed);
  const billingPath = path.join(sourcesDir, 'billing.xlsx');
  const costcenterPath = path.join(sourcesDir, 'costcenter.xlsx');
  const billingManifest = await buildBillingWorkbook(ds, billingPath);
  const costcenterManifest = await buildCostCenterWorkbook(ds, costcenterPath);

  // Blind phase first: nothing in gold/ exists while the solver runs.
  const blindOut = path.join(blindDir, 'blind-liquidity.xlsx');
  const solved = await solveBlind(billingPath, costcenterPath, blindOut);

  // Gold phase (from raw dataset + builder manifests only).
  const gold = computeGold(ds, billingManifest, costcenterManifest);
  await writeGoldWorkbook(gold, path.join(goldDir, 'gold-liquidity.xlsx'));
  fs.writeFileSync(path.join(goldDir, 'gold.json'), JSON.stringify(gold, null, 2));
  fs.writeFileSync(
    path.join(goldDir, 'manifests.json'),
    JSON.stringify({ billing: billingManifest.placements, costcenter: costcenterManifest.placements }, null, 2)
  );

  const knownSheets = {
    billing: [...new Set(billingManifest.placements.map((p) => p.sheet))],
    costcenter: [...new Set(costcenterManifest.placements.map((p) => p.sheet))]
  };
  const score = await scoreCandidate(blindOut, gold, knownSheets, solved.tokens);
  fs.writeFileSync(path.join(dir, 'score.json'), JSON.stringify({ score, warnings: solved.warnings }, null, 2));
  return { seed, dir, gold, score, solverTokens: solved.tokens, solverWarnings: solved.warnings };
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const seeds = args.length > 0 ? args.map(Number).filter((n) => Number.isFinite(n)) : [42];
  let allPass = true;
  for (const seed of seeds) {
    const result = await runBench(seed);
    console.log(`\n=== Benchmark seed ${seed} ===`);
    console.log(formatScore(result.score));
    if (result.solverWarnings.length > 0) {
      console.log(`solver warnings (${result.solverWarnings.length}):`);
      for (const w of result.solverWarnings.slice(0, 6)) console.log(`  ! ${w}`);
    }
    allPass = allPass && result.score.pass;
  }
  process.exitCode = allPass ? 0 : 1;
}

if (require.main === module) {
  main().catch((err) => {
    console.error(err);
    process.exit(1);
  });
}

// pharos:eof

/** Score an externally produced liquidity workbook against a seed's gold. */
import * as os from 'os';
import * as path from 'path';
import * as fs from 'fs';
import { generateDataset } from './data';
import { buildBillingWorkbook } from './build-billing';
import { buildCostCenterWorkbook } from './build-costcenter';
import { computeGold } from './gold';
import { formatScore, scoreCandidate } from './score';

(async () => {
  const seed = Number(process.argv[2]);
  const candidate = process.argv[3];
  if (!Number.isFinite(seed) || !candidate) {
    console.error('usage: ts-node bench/score-blind.ts <seed> <candidate.xlsx>');
    process.exit(2);
  }
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'pharos-gold-'));
  const ds = generateDataset(seed);
  const billing = await buildBillingWorkbook(ds, path.join(tmp, 'billing.xlsx'));
  const costcenter = await buildCostCenterWorkbook(ds, path.join(tmp, 'costcenter.xlsx'));
  const gold = computeGold(ds, billing, costcenter);
  const known = {
    billing: [...new Set(billing.placements.map((p) => p.sheet))],
    costcenter: [...new Set(costcenter.placements.map((p) => p.sheet))]
  };
  const score = await scoreCandidate(candidate, gold, known);
  console.log(formatScore(score));
  process.exitCode = score.pass ? 0 : 1;
})();

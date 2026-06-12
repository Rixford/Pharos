/** Generate only the two source workbooks for a seed (no gold artifacts). */
import * as fs from 'fs';
import * as path from 'path';
import { generateDataset } from './data';
import { buildBillingWorkbook } from './build-billing';
import { buildCostCenterWorkbook } from './build-costcenter';

(async () => {
  const seed = Number(process.argv[2] ?? 42);
  const outDir = process.argv[3] ?? '.';
  fs.mkdirSync(outDir, { recursive: true });
  const ds = generateDataset(seed);
  await buildBillingWorkbook(ds, path.join(outDir, 'billing.xlsx'));
  await buildCostCenterWorkbook(ds, path.join(outDir, 'costcenter.xlsx'));
  console.log(`sources for seed ${seed} written to ${outDir}`);
})();

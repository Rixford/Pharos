import * as path from 'path';
import { WorkbookGraph } from '../src';

export const FIXTURE = path.join(__dirname, 'fixtures', 'sample.xlsx');

let cached: Promise<WorkbookGraph> | undefined;

/** Shared fixture graph (per test file — Jest isolates files in workers). */
export function loadFixture(): Promise<WorkbookGraph> {
  cached ??= WorkbookGraph.load(FIXTURE);
  return cached;
}

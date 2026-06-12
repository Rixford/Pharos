import { runBench } from '../bench/run';

/**
 * Closed-loop benchmark as a regression test: generate the two complex
 * source workbooks from the deterministic dataset, run the blind solver
 * (public Pharos APIs only), build the gold report from the raw dataset,
 * and require every scoring section to pass its threshold.
 */
describe('liquidity benchmark (closed loop)', () => {
  test(
    'blind solver reconstructs the gold liquidity report (seed 42)',
    async () => {
      const result = await runBench(42);
      const failed = result.score.sections.filter((s) => !s.pass);
      expect(failed.map((s) => `${s.name}: ${(s.score * 100).toFixed(1)}%`)).toEqual([]);
      expect(result.score.pass).toBe(true);
      expect(result.solverTokens).toBeGreaterThan(0);
    },
    120000
  );
});

// pharos:eof

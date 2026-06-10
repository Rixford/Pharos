import { execFileSync } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import { FIXTURE } from './helpers';

const CLI = path.join(__dirname, '..', 'dist', 'cli', 'index.js');

const runCli = (...args: string[]): string =>
  execFileSync(process.execPath, [CLI, ...args], { encoding: 'utf8' });

describe('pharos CLI (built dist)', () => {
  beforeAll(() => {
    expect(fs.existsSync(CLI)).toBe(true); // `npm test` builds first (pretest)
  });

  test('--help lists all subcommands', () => {
    const out = runCli('--help');
    for (const cmd of ['load', 'inspect', 'region', 'context', 'precedents', 'dependents', 'find']) {
      expect(out).toContain(cmd);
    }
  });

  test('load prints sheets, regions, names and hidden flags', () => {
    const out = runCli('load', FIXTURE);
    expect(out).toContain('Sales');
    expect(out).toContain('Rates');
    expect(out).toContain('hidden');
    expect(out).toContain('TotalRevenue');
    expect(out).toContain('rg_');
    expect(out).toContain('Budget.xlsx');
  });

  test('inspect shows formula, region and dependents', () => {
    const out = runCli('inspect', FIXTURE, 'Sales!F4');
    expect(out).toContain('=D4*E4');
    expect(out).toContain('region');
    expect(out).toContain('Sales!F35'); // dependent
  });

  test('region --list and region summary', () => {
    const list = runCli('region', FIXTURE, '--list');
    expect(list).toContain('Sales!A3:F35');
    expect(list).toContain('rg_');

    const summary = runCli('region', FIXTURE, 'Sales!B7', '--mode', 'summary');
    expect(summary).toContain('30 data rows');
  });

  test('context --json returns a parseable packet', () => {
    const out = runCli('context', FIXTURE, 'Summary!C2', '--depth', '2', '--json');
    const packet = JSON.parse(out);
    expect(packet.seed).toBe('Summary!C2');
    expect(Array.isArray(packet.regions)).toBe(true);
    expect(packet.regions.length).toBeGreaterThanOrEqual(2);
    expect(packet.nextActions.length).toBeGreaterThan(0);
  });

  test('precedents and dependents render trees', () => {
    const prec = runCli('precedents', FIXTURE, 'Summary!C3', '--depth', '4');
    expect(prec).toContain('Sales!F4:F33');
    expect(prec).toContain('└─');

    const deps = runCli('dependents', FIXTURE, 'Sales!F35');
    expect(deps).toContain('Summary!C2');
  });

  test('find locates values', () => {
    const out = runCli('find', FIXTURE, 'North', '--sheet', 'Sales');
    expect(out).toContain('Sales!B4');
  });

  test('bad input exits non-zero with a friendly message', () => {
    expect(() => runCli('load', 'no-such-file.xlsx')).toThrow(/file not found/);
    expect(() => runCli('inspect', FIXTURE, 'Nope!A1')).toThrow(/not found/);
    expect(() => runCli('region', FIXTURE, 'Sales!B7', '--mode', 'bogus')).toThrow();
  });
});

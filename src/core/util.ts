/** Small shared utilities. */

/** FNV-1a 32-bit hash, returned in base36 — used for stable region ids. */
export function hash36(input: string): string {
  let h = 0x811c9dc5;
  for (let i = 0; i < input.length; i++) {
    h ^= input.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(36);
}

export function uniq<T>(items: Iterable<T>): T[] {
  return [...new Set(items)];
}

export function clampNum(n: number, lo: number, hi: number): number {
  return Math.max(lo, Math.min(hi, n));
}

/** Deterministic, locale-independent number formatting with thousands separators. */
export function fmtNumber(n: number): string {
  if (!Number.isFinite(n)) return String(n);
  const rounded = Math.abs(n) >= 100 ? Math.round(n * 100) / 100 : Math.round(n * 10000) / 10000;
  const [int, frac] = String(rounded).split('.');
  const neg = int.startsWith('-');
  const digits = neg ? int.slice(1) : int;
  const grouped = digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (neg ? '-' : '') + grouped + (frac ? '.' + frac : '');
}

/** Render a Date as ISO date (or datetime when it has a time component). */
export function fmtDate(d: Date): string {
  const iso = d.toISOString();
  return iso.endsWith('T00:00:00.000Z') ? iso.slice(0, 10) : iso.replace('.000Z', 'Z');
}

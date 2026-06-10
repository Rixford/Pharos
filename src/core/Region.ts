/**
 * Region — a detected table/block with a stable identifier.
 *
 * Identifiers are derived from `sheet|range`, so repeated runs over the
 * same workbook produce the same ids and different calls can refer to the
 * same region.
 */
import { rangeContains } from './address';
import { CellRef, RegionBrief, RegionData } from './types';

export class Region {
  constructor(readonly data: RegionData) {}

  get id(): string {
    return this.data.id;
  }

  get sheet(): string {
    return this.data.sheet;
  }

  /** Sheet-qualified A1 range, e.g. `Sales!A3:F35`. */
  get rangeA1(): string {
    return this.data.rangeA1;
  }

  get kind(): RegionData['kind'] {
    return this.data.kind;
  }

  get title(): string | undefined {
    return this.data.title;
  }

  get headers(): string[] | undefined {
    return this.data.headers;
  }

  get confidence(): number {
    return this.data.confidence;
  }

  contains(ref: CellRef): boolean {
    return rangeContains(this.data.range, ref);
  }

  brief(): RegionBrief {
    return {
      id: this.data.id,
      rangeA1: this.data.rangeA1,
      kind: this.data.kind,
      title: this.data.title,
      rows: this.data.rowCount,
      cols: this.data.colCount,
      confidence: this.data.confidence
    };
  }

  toJSON(): RegionData {
    return this.data;
  }
}

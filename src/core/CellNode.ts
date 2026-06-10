/**
 * CellNode — a single cell in the workbook graph.
 */
import { formatCell } from './address';
import { CellRef, CellScalar, CellStyle, CellValueType, FormulaRef, JsonScalar } from './types';
import { fmtDate } from './util';

export class CellNode {
  readonly sheet: string;
  readonly row: number;
  readonly col: number;
  readonly value: CellScalar;
  readonly type: CellValueType;
  /** Formula text without the leading '='. */
  readonly formula?: string;
  readonly style?: CellStyle;
  readonly hyperlink?: string;
  /** Precedent references extracted from the formula (names resolved). */
  readonly refs: FormulaRef[];
  /** Region id assigned once region detection has run for this sheet. */
  regionId?: string;

  constructor(init: {
    sheet: string;
    row: number;
    col: number;
    value: CellScalar;
    type: CellValueType;
    formula?: string;
    style?: CellStyle;
    hyperlink?: string;
    refs?: FormulaRef[];
  }) {
    this.sheet = init.sheet;
    this.row = init.row;
    this.col = init.col;
    this.value = init.value;
    this.type = init.type;
    this.formula = init.formula;
    this.style = init.style;
    this.hyperlink = init.hyperlink;
    this.refs = init.refs ?? [];
  }

  get ref(): CellRef {
    return { sheet: this.sheet, row: this.row, col: this.col };
  }

  /** Fully qualified address, e.g. `Sales!F4`. */
  get address(): string {
    return formatCell(this.ref);
  }

  get isFormula(): boolean {
    return this.formula !== undefined;
  }

  /** JSON-safe value (Dates become ISO strings). */
  get valueJson(): JsonScalar {
    return this.value instanceof Date ? fmtDate(this.value) : this.value;
  }

  /** Raw precedent reference strings as written in the formula. */
  precedentRefs(): string[] {
    return this.refs.map((r) => r.raw);
  }

  toJSON(): Record<string, unknown> {
    const out: Record<string, unknown> = {
      address: this.address,
      value: this.valueJson,
      type: this.type
    };
    if (this.formula !== undefined) out.formula = this.formula;
    if (this.style) out.style = this.style;
    if (this.hyperlink) out.hyperlink = this.hyperlink;
    if (this.refs.length) out.precedents = this.precedentRefs();
    if (this.regionId) out.regionId = this.regionId;
    return out;
  }
}

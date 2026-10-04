import type { WorkbookModel } from "./model";
import type { Cell, WorkbookSnapshot } from "./types";

/** Keeps a point-in-time view while later transactions replace cells. */
export class SnapshotReader {
  readonly revision: number;
  readonly snapshot: WorkbookSnapshot;
  private sheets;
  private sheetIndex = 0;
  private position = 0;
  constructor(model: WorkbookModel) {
    this.revision = model.revision;
    this.snapshot = {
      version: 1,
      sheets: model.sheets.map((s) => ({
        ...structuredClone(s.meta),
        cells: [],
      })),
      styles: structuredClone(model.styles),
      names: structuredClone(model.names),
      dateSystem: model.dateSystem,
      diagnostics: structuredClone(model.diagnostics),
    };
    this.sheets = model.sheets.map((s) => {
      // Keep insertion order, including competing spill anchors. Filling a typed
      // buffer directly avoids Float64Array.from's million-entry temporary array.
      const keys = new Float64Array(s.cells.size);
      let index = 0;
      for (const key of s.cells.keys()) keys[index++] = key;
      return {
        cells: s.cells,
        keys,
        before: new Map<number, Cell | undefined>(),
      };
    });
  }
  beforeWrite(cells: Map<number, Cell>, key: number): void {
    const sheet = this.sheets.find((s) => s.cells === cells);
    if (sheet?.keys.length && !sheet.before.has(key))
      sheet.before.set(key, cells.get(key));
  }
  read(): { sheetIndex: number; cells: [number, Cell][]; done: boolean } {
    const sheetIndex = this.sheetIndex;
    const sheet = this.sheets[sheetIndex];
    if (!sheet) return { sheetIndex, cells: [], done: true };
    const cells: [number, Cell][] = [];
    const end = Math.min(this.position + 2048, sheet.keys.length);
    for (; this.position < end; this.position++) {
      const key = sheet.keys[this.position];
      const cell = sheet.before.has(key)
        ? sheet.before.get(key)
        : sheet.cells.get(key);
      if (cell)
        cells.push([
          key,
          cell.value && typeof cell.value === "object"
            ? { ...cell, value: { ...cell.value } }
            : { ...cell },
        ]);
    }
    if (this.position === sheet.keys.length) {
      this.sheetIndex++;
      this.position = 0;
      sheet.before.clear();
      sheet.keys = new Float64Array();
    }
    return { sheetIndex, cells, done: this.sheetIndex === this.sheets.length };
  }
}

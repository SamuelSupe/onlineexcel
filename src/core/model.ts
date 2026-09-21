import {
  address,
  columnOf,
  contains,
  intersects,
  keyOf,
  MAX_COLUMNS,
  MAX_ROWS,
  rowOf,
  validateRect,
} from "./address";
import type {
  Cell,
  CellStyle,
  ChangeEvent,
  Command,
  Diagnostic,
  NamedRange,
  Rect,
  Region,
  Scalar,
  SheetMeta,
  WorkbookOptions,
  WorkbookSnapshot,
} from "./types";
import {
  validateEntry,
  validateRule,
  conditionalStyle,
  checkProtection,
} from "./rules";
import { filterText } from "./queries";
import { FormulaEngine } from "../formula/engine";

import { compare, isError } from "../formula/values";
import { executeCommand, validName } from "./commands";
export interface SheetState {
  meta: SheetMeta;
  cells: Map<number, Cell>;
}
interface CellPatch {
  before?: Cell;
  after?: Cell;
}
interface HistoryEntry {
  styleCount: number;
  cells: Map<string, Map<number, CellPatch>>;
  metadata: Map<string, { before: SheetMeta; after?: SheetMeta }>;
  beforeSheets: SheetState[];
  afterSheets: SheetState[];
  beforeNames: Record<string, NamedRange>;
  afterNames: Record<string, NamedRange>;
  rebuild: boolean;
}
const copy = <T>(value: T): T => structuredClone(value);

export class WorkbookModel {
  sheets: SheetState[] = [];
  styles: CellStyle[] = [{}];
  names: Record<string, NamedRange> = {};
  dateSystem: 1900 | 1904 = 1900;
  revision = 0;
  diagnostics: Diagnostic[] = [];
  readonly engine: FormulaEngine;
  private styleIds = new Map<string, number>();
  private history: HistoryEntry[] = [];
  private future: HistoryEntry[] = [];
  private active?: HistoryEntry;
  private historyLimit: number;
  constructor(options: WorkbookOptions = {}) {
    if (
      options.historyLimit !== undefined &&
      (!Number.isInteger(options.historyLimit) || options.historyLimit < 0)
    )
      throw new Error("Invalid history limit");
    if (
      options.dateSystem !== undefined &&
      ![1900, 1904].includes(options.dateSystem)
    )
      throw new Error("Invalid date system");
    this.historyLimit = options.historyLimit ?? 100;
    this.dateSystem = options.dateSystem ?? 1900;
    this.engine = new FormulaEngine({
      get dateSystem() {
        return model.dateSystem;
      },
      cell: (id, key) =>
        this.sheets.find((s) => s.meta.id === id)?.cells.get(key),
      sheetId: (name) =>
        this.sheets.find(
          (s) => s.meta.name.toUpperCase() === name.toUpperCase(),
        )?.meta.id,
      extent: (id) => {
        const s = this.sheet(id);
        return { rows: s.meta.rowCount, columns: s.meta.columnCount };
      },
      canSpill: (id, range) =>
        !this.sheet(id).meta.merges.some((m) => intersects(m, range)),
      named: (name) => this.names[name.toUpperCase()],
      formulas: () => this.formulas(),
    });
    const model = this;
    if (options.snapshot) this.load(options.snapshot);
    else {
      for (const s of options.sheets ?? [{ name: "Sheet1" }])
        this.sheets.push(this.newSheet(s.name, s.rows, s.columns));
      if (!this.sheets.length)
        throw new Error("A workbook requires at least one sheet");
    }
  }
  newSheet(
    name: string,
    rows = 100000,
    columns = 100,
    id: string = crypto.randomUUID(),
  ): SheetState {
    if (
      !validName(name) ||
      this.sheets.some((s) => s.meta.name.toUpperCase() === name.toUpperCase())
    )
      throw new Error("Invalid or duplicate sheet name");
    if (
      !Number.isInteger(rows) ||
      !Number.isInteger(columns) ||
      rows < 1 ||
      rows > MAX_ROWS ||
      columns < 1 ||
      columns > MAX_COLUMNS
    )
      throw new Error("Invalid sheet dimensions");
    return {
      meta: {
        id,
        name,
        rowCount: rows,
        columnCount: columns,
        rowHeights: {},
        columnWidths: {},
        hiddenRows: [],
        hiddenColumns: [],
        frozenRows: 0,
        frozenColumns: 0,
        merges: [],
      },
      cells: new Map(),
    };
  }
  sheet(id: string): SheetState {
    const s = this.sheets.find((s) => s.meta.id === id);
    if (!s) throw new Error(`Sheet not found: ${id}`);
    return s;
  }
  *formulas(): Generator<{ sheetId: string; key: number; formula: string }> {
    for (const s of this.sheets)
      for (const [key, cell] of s.cells)
        if (cell.formula)
          yield { sheetId: s.meta.id, key, formula: cell.formula };
  }
  metadata(): {
    sheets: SheetMeta[];
    revision: number;
    canUndo: boolean;
    canRedo: boolean;
    dateSystem: 1900 | 1904;
    names: Record<string, NamedRange>;
  } {
    return {
      sheets: this.sheets.map((s) => copy(s.meta)),
      revision: this.revision,
      canUndo: this.history.length > 0,
      canRedo: this.future.length > 0,
      dateSystem: this.dateSystem,
      names: copy(this.names),
    };
  }
  touchMeta(sheet: SheetState): void {
    if (!this.active) throw new Error("Mutation outside a transaction");
    if (!this.active.metadata.has(sheet.meta.id))
      this.active.metadata.set(sheet.meta.id, { before: copy(sheet.meta) });
  }
  rebuildAfterTransaction(): void {
    if (this.active) this.active.rebuild = true;
  }
  setCell(sheet: SheetState, key: number, cell?: Cell): void {
    if (!this.active) throw new Error("Mutation outside a transaction");
    let patches = this.active.cells.get(sheet.meta.id);
    if (!patches) {
      patches = new Map();
      this.active.cells.set(sheet.meta.id, patches);
    }
    let patch = patches.get(key);
    if (!patch) {
      patch = { before: sheet.cells.get(key) };
      patches.set(key, patch);
    }
    if (
      cell &&
      (cell.formula ||
        (cell.value !== null && cell.value !== undefined) ||
        cell.style)
    ) {
      sheet.cells.set(key, cell);
      patch.after = cell;
    } else {
      sheet.cells.delete(key);
      patch.after = undefined;
    }
  }
  validateArea(sheet: SheetState, range: Rect): void {
    validateRect(range);
    if (range.r2 >= sheet.meta.rowCount || range.c2 >= sheet.meta.columnCount)
      throw new Error("Range exceeds sheet dimensions");
  }
  ensureWritable(sheet: SheetState, range: Rect): void {
    this.validateArea(sheet, range);
    checkProtection(this, sheet, range);
    for (let r = range.r1; r <= range.r2; r++)
      for (let c = range.c1; c <= range.c2; c++) {
        const owner = this.engine.spillOwner(sheet.meta.id, keyOf(r, c));
        if (owner && !contains(range, owner.row, owner.column))
          throw new Error(
            "Cannot edit part of a spilled array; edit its anchor cell",
          );
      }
  }
  styleId(style: CellStyle): number {
    const color = (value: string): string => {
      if (!/^#(?:[0-9a-f]{3}|[0-9a-f]{6})$/i.test(value))
        throw new Error("Style colors must be #RGB or #RRGGBB");
      return value.length === 4
        ? "#" + [...value.slice(1)].map((c) => c + c).join("")
        : value;
    };
    if (style.color) style.color = color(style.color);
    if (style.background) style.background = color(style.background);
    if (style.border)
      style.border = Object.fromEntries(
        Object.entries(style.border).map(([side, value]) => [
          side,
          value === "" ? "" : color(value),
        ]),
      );
    if (
      style.fontSize !== undefined &&
      (!Number.isFinite(style.fontSize) ||
        style.fontSize < 6 ||
        style.fontSize > 200)
    )
      throw new Error("Font size must be between 6 and 200 points");
    const normalized = Object.fromEntries(
      Object.entries(style)
        .filter(([, value]) => value !== undefined)
        .sort(([a], [b]) => a.localeCompare(b)),
    );
    const key = JSON.stringify(normalized),
      existing = this.styleIds.get(key);
    if (existing !== undefined) return existing;
    const id = this.styles.length;
    this.styles.push(normalized);
    this.styleIds.set(key, id);
    return id;
  }
  private begin(): HistoryEntry {
    if (this.active) throw new Error("Nested model transaction");
    return (this.active = {
      styleCount: this.styles.length,
      cells: new Map(),
      metadata: new Map(),
      beforeSheets: [...this.sheets],
      afterSheets: [],
      beforeNames: copy(this.names),
      afterNames: {},
      rebuild: false,
    });
  }
  private changedCells(
    entry: HistoryEntry,
  ): { sheetId: string; key: number }[] {
    const changed: { sheetId: string; key: number }[] = [];
    for (const [sheetId, patches] of entry.cells)
      for (const key of patches.keys()) changed.push({ sheetId, key });
    return changed;
  }
  private calculate(entry: HistoryEntry): void {
    // Bulk edits invalidate enough dependencies that rebuilding is cheaper than indexing each cell.
    const changedCount = [...entry.cells.values()].reduce(
      (n, patches) => n + patches.size,
      0,
    );
    if (entry.rebuild || changedCount > 5000) this.engine.rebuild();
    else {
      const changed = this.changedCells(entry);
      this.engine.invalidate(changed);
      for (const { sheetId, key } of changed)
        this.engine.updateFormula(
          sheetId,
          key,
          this.sheet(sheetId).cells.get(key)?.formula,
        );
    }
    this.engine.recalculate();
    for (const [id, patches] of entry.cells) {
      const sheet = this.sheets.find((s) => s.meta.id === id);
      if (!sheet?.meta.validations?.length) continue;
      for (const [key, patch] of patches) {
        if (
          patch.before?.value === patch.after?.value &&
          patch.before?.formula === patch.after?.formula
        )
          continue;
        validateEntry(
          sheet.meta,
          rowOf(key),
          columnOf(key),
          this.engine.get(id, key),
        );
      }
    }
    this.refreshFilters();
  }
  private commit(entry: HistoryEntry, calculated = false): ChangeEvent {
    entry.afterSheets = [...this.sheets];
    entry.afterNames = copy(this.names);
    for (const [id, patch] of entry.metadata) {
      const sheet = this.sheets.find((s) => s.meta.id === id);
      if (sheet) patch.after = copy(sheet.meta);
    }
    if (!calculated) this.calculate(entry);
    this.active = undefined;
    this.history.push(entry);
    if (this.history.length > this.historyLimit) this.history.shift();
    this.future = [];
    this.revision++;
    return this.change(entry, "edit");
  }
  execute(commands: Command[]): ChangeEvent {
    const entry = this.begin();
    try {
      for (const command of commands) this.command(command);
      return this.commit(entry);
    } catch (e) {
      this.restore(entry, false);
      this.styles.length = entry.styleCount;
      this.styleIds.clear();
      this.styles.forEach((style, index) =>
        this.styleIds.set(JSON.stringify(style), index),
      );
      this.active = undefined;
      throw e;
    }
  }
  async executeAsync(
    commands: Command[],
    checkpoint: (progress: number) => Promise<void>,
    validate?: () => void,
  ): Promise<ChangeEvent> {
    const entry = this.begin();
    try {
      for (let index = 0; index < commands.length; index++) {
        const command = commands[index];
        if (command.type === "setValues" && command.values.length > 2000) {
          const rows = command.range.r2 - command.range.r1 + 1,
            cols = command.range.c2 - command.range.c1 + 1;
          if (
            command.values.length !== rows ||
            command.values.some((row) => row.length !== cols)
          )
            throw new Error("Values must match range dimensions");
          // Spill ownership is checked against the whole public write, not each checkpoint chunk.
          this.ensureWritable(this.sheet(command.sheetId), command.range);
          for (let offset = 0; offset < rows; offset += 2000) {
            const values = command.values.slice(offset, offset + 2000);
            this.command(
              {
                ...command,
                values,
                range: {
                  ...command.range,
                  r1: command.range.r1 + offset,
                  r2: command.range.r1 + offset + values.length - 1,
                },
              },
              true,
            );
            await checkpoint(
              (index + Math.min(rows, offset + 2000) / rows) / commands.length,
            );
          }
        } else {
          this.command(command);
          await checkpoint((index + 1) / commands.length);
        }
      }
      this.calculate(entry);
      validate?.();
      await checkpoint(1);
      return this.commit(entry, true);
    } catch (e) {
      this.restore(entry, false);
      this.styles.length = entry.styleCount;
      this.styleIds.clear();
      this.styles.forEach((style, index) =>
        this.styleIds.set(JSON.stringify(style), index),
      );
      this.active = undefined;
      throw e;
    }
  }
  private change(
    entry: HistoryEntry,
    source: ChangeEvent["source"],
  ): ChangeEvent {
    const changes = new Map<string, Rect | undefined>();
    for (const [id, patches] of entry.cells) {
      let r1 = Infinity,
        c1 = Infinity,
        r2 = 0,
        c2 = 0;
      for (const key of patches.keys()) {
        const r = rowOf(key),
          c = columnOf(key);
        r1 = Math.min(r1, r);
        c1 = Math.min(c1, c);
        r2 = Math.max(r2, r);
        c2 = Math.max(c2, c);
      }
      if (patches.size) changes.set(id, { r1, c1, r2, c2 });
    }
    for (const id of entry.metadata.keys()) changes.set(id, undefined);
    return {
      revision: this.revision,
      changes: [...changes].map(([sheetId, range]) => ({ sheetId, range })),
      source,
    };
  }
  private restore(entry: HistoryEntry, forward: boolean): void {
    this.sheets = [...(forward ? entry.afterSheets : entry.beforeSheets)];
    this.names = copy(forward ? entry.afterNames : entry.beforeNames);
    for (const [id, patch] of entry.metadata) {
      const s = this.sheets.find((s) => s.meta.id === id),
        meta = forward ? patch.after : patch.before;
      if (s && meta) s.meta = copy(meta);
    }
    for (const [id, patches] of entry.cells) {
      const s = this.sheets.find((s) => s.meta.id === id);
      if (!s) continue;
      for (const [key, patch] of patches) {
        const cell = forward ? patch.after : patch.before;
        if (cell) s.cells.set(key, cell);
        else s.cells.delete(key);
      }
    }
    this.engine.rebuild();
    this.engine.recalculate();
    this.refreshFilters();
  }
  undo(): ChangeEvent | null {
    const entry = this.history.pop();
    if (!entry) return null;
    this.restore(entry, false);
    this.future.push(entry);
    this.revision++;
    return this.change(entry, "undo");
  }
  redo(): ChangeEvent | null {
    const entry = this.future.pop();
    if (!entry) return null;
    this.restore(entry, true);
    this.history.push(entry);
    this.revision++;
    return this.change(entry, "redo");
  }
  private command(command: Command, writableChecked = false): void {
    // These commands read calculated values and must see earlier writes in this transaction.
    if (
      (this.active?.cells.size || this.active?.rebuild) &&
      (command.type === "sort" ||
        command.type === "copy" ||
        command.type === "fill")
    ) {
      this.calculate(this.active);
    }
    executeCommand(this, command, writableChecked);
  }
  refreshFilters(): void {
    for (const sheet of this.sheets) {
      const filter = sheet.meta.filter;
      if (!filter) {
        delete sheet.meta.filteredRows;
        continue;
      }
      const lists = new Map(
        filter.rules
          .filter((rule) => rule.operator === "in")
          .map((rule) => [
            rule.column,
            new Set(rule.values.map((value) => value.toLowerCase())),
          ]),
      );
      const hidden: number[] = [];
      for (let r = filter.range.r1 + 1; r <= filter.range.r2; r++)
        if (
          !filter.rules.every((rule) => {
            const value = this.engine.get(sheet.meta.id, keyOf(r, rule.column));
            if (rule.operator === "in")
              return lists
                .get(rule.column)!
                .has(filterText(value).toLowerCase());
            if (isError(value)) return false;
            if (rule.operator === "contains")
              return String(value ?? "")
                .toLowerCase()
                .includes(String(rule.value ?? "").toLowerCase());
            const order = compare(value, rule.value);
            return rule.operator === "eq"
              ? order === 0
              : rule.operator === "neq"
                ? order !== 0
                : rule.operator === "gt"
                  ? order > 0
                  : rule.operator === "lt"
                    ? order < 0
                    : rule.operator === "gte"
                      ? order >= 0
                      : order <= 0;
          })
        )
          hidden.push(r);
      sheet.meta.filteredRows = hidden;
    }
  }
  region(sheetId: string, range: Rect): Region {
    const sheet = this.sheet(sheetId);
    this.validateArea(sheet, range);
    if ((range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1) > 200_000)
      throw new Error(
        "Read regions are limited to 200,000 cells; request chunks",
      );
    const cells: Region["cells"] = [];
    for (let r = range.r1; r <= range.r2; r++)
      for (let c = range.c1; c <= range.c2; c++) {
        const key = keyOf(r, c),
          cell = sheet.cells.get(key),
          value = this.engine.get(sheetId, key);
        if (cell || value !== null)
          cells.push({
            row: r,
            column: c,
            value,
            formula: cell?.formula,
            style: this.styles[cell?.style ?? 0] ?? {},
            displayStyle: conditionalStyle(
              sheet.meta,
              r,
              c,
              value,
              this.styles[cell?.style ?? 0] ?? {},
            ),
            spill: this.engine.spillOwner(sheetId, key),
          });
      }
    return { range, cells, revision: this.revision };
  }
  find(
    sheetId: string,
    search: string,
    matchCase = false,
  ): { address: string; value: Scalar }[] {
    const sheet = this.sheet(sheetId),
      results: { address: string; value: Scalar }[] = [];
    for (const [key] of sheet.cells) {
      const value = this.engine.get(sheetId, key),
        content = isError(value) ? value.error : String(value ?? "");
      if (
        (matchCase ? content : content.toLowerCase()).includes(
          matchCase ? search : search.toLowerCase(),
        )
      )
        results.push({ address: address(rowOf(key), columnOf(key)), value });
      if (results.length >= 1000) break;
    }
    return results;
  }
  snapshot(): WorkbookSnapshot {
    return {
      version: 1,
      sheets: this.sheets.map((s) => ({
        ...copy(s.meta),
        cells: [...s.cells].map(([key, cell]) => [key, { ...cell }]),
      })),
      styles: copy(this.styles),
      names: copy(this.names),
      dateSystem: this.dateSystem,
      diagnostics: copy(this.diagnostics),
    };
  }
  load(snapshot: WorkbookSnapshot): void {
    if (
      snapshot.version !== 1 ||
      !Array.isArray(snapshot.sheets) ||
      !snapshot.sheets.length ||
      ![1900, 1904].includes(snapshot.dateSystem)
    )
      throw new Error("Unsupported workbook snapshot");
    const sheets: SheetState[] = [],
      ids = new Set<string>(),
      names = new Set<string>();
    for (const input of snapshot.sheets) {
      if (
        !validName(input.name) ||
        ids.has(input.id) ||
        names.has(input.name.toUpperCase())
      )
        throw new Error("Invalid snapshot sheets");
      ids.add(input.id);
      names.add(input.name.toUpperCase());
      validateRect({
        r1: 0,
        c1: 0,
        r2: input.rowCount - 1,
        c2: input.columnCount - 1,
      });
      const { cells, ...meta } = input;
      const validIndex = (index: number, limit: number) =>
        Number.isInteger(index) && index >= 0 && index < limit;
      if (
        !validIndex(meta.frozenRows, meta.rowCount) ||
        !validIndex(meta.frozenColumns, meta.columnCount) ||
        meta.hiddenRows.some((r) => !validIndex(r, meta.rowCount)) ||
        meta.hiddenColumns.some((c) => !validIndex(c, meta.columnCount))
      )
        throw new Error("Invalid snapshot dimensions");
      for (const [sizes, limit] of [
        [meta.rowHeights, meta.rowCount],
        [meta.columnWidths, meta.columnCount],
      ] as const) {
        if (
          Object.entries(sizes).some(
            ([index, size]) =>
              !validIndex(Number(index), limit) ||
              !Number.isFinite(size) ||
              size < 0 ||
              size > 4096,
          )
        )
          throw new Error("Invalid snapshot sizes");
      }
      const data = new Map<number, Cell>();
      for (const [key, cell] of cells) {
        if (
          !Number.isSafeInteger(key) ||
          key < 0 ||
          rowOf(key) >= meta.rowCount ||
          columnOf(key) >= meta.columnCount ||
          (cell.formula !== undefined &&
            (typeof cell.formula !== "string" || !cell.formula.startsWith("=")))
        )
          throw new Error("Invalid snapshot cell");
        if (
          cell.style !== undefined &&
          (!Number.isInteger(cell.style) ||
            cell.style < 0 ||
            cell.style >= snapshot.styles.length)
        )
          throw new Error("Invalid style reference");
        if (
          cell.value !== undefined &&
          cell.value !== null &&
          !["string", "boolean"].includes(typeof cell.value) &&
          !(typeof cell.value === "number" && Number.isFinite(cell.value)) &&
          !(
            typeof cell.value === "object" &&
            typeof cell.value.error === "string"
          )
        )
          throw new Error("Invalid snapshot value");
        data.set(key, { ...cell });
      }
      for (const rule of meta.validations ?? []) {
        validateRule(rule);
        this.validateArea({ meta, cells: data }, rule.range);
      }
      for (const rule of meta.conditionalFormats ?? [])
        this.validateArea({ meta, cells: data }, rule.range);
      for (const merge of meta.merges) {
        validateRect(merge);
        if (merge.r2 >= meta.rowCount || merge.c2 >= meta.columnCount)
          throw new Error("Invalid snapshot merge");
      }
      if (meta.filter) {
        validateRect(meta.filter.range);
        if (
          meta.filter.range.r2 >= meta.rowCount ||
          meta.filter.range.c2 >= meta.columnCount
        )
          throw new Error("Invalid snapshot filter");
      }
      sheets.push({ meta: copy(meta), cells: data });
    }
    if (sheets.every((s) => s.meta.hidden))
      throw new Error("Cannot hide the last visible sheet");
    this.sheets = sheets;
    this.styles = copy(snapshot.styles.length ? snapshot.styles : [{}]);
    this.names = copy(snapshot.names ?? {});
    this.dateSystem = snapshot.dateSystem;
    this.diagnostics = copy(snapshot.diagnostics ?? []);
    this.styleIds.clear();
    this.styles.forEach((s, i) => this.styleIds.set(JSON.stringify(s), i));
    this.history = [];
    this.future = [];
    this.engine.rebuild();
    this.engine.recalculate();
    this.refreshFilters();
    this.revision++;
  }
  getDiagnostics(): Diagnostic[] {
    const diagnostics = [...this.diagnostics, ...this.engine.diagnostics()];
    for (const sheet of this.sheets)
      for (const rule of sheet.meta.validations ?? [])
        if (
          rule.type === "list" &&
          (rule.values!.join(",").length > 255 ||
            rule.values!.some((v) => v.includes(",")))
        )
          diagnostics.push({
            code: "XLSX_VALIDATION_LIST",
            severity: "warning",
            lossy: true,
            sheetId: sheet.meta.id,
            message:
              "XLSX inline validation lists are limited to 255 characters and cannot contain comma-bearing items. JSON retains the full list.",
          });
    return diagnostics;
  }
}

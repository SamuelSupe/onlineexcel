import { pasteCells } from "./paste";
import { pasteRules } from "./rules";
import {
  columnOf,
  contains,
  intersects,
  keyOf,
  rowOf,
  MAX_ROWS,
  MAX_COLUMNS,
  validateRect,
} from "./address";
import type { WorkbookModel, SheetState } from "./model";
import type { Cell, Command, Rect } from "./types";
import {
  mapReferences,
  parseFormula,
  printFormula,
  shiftFormula,
} from "../formula/parser";
import { compare, isError } from "../formula/values";
export function executeStructural(
  model: WorkbookModel,
  sheet: SheetState,
  command: Extract<
    Command,
    { type: "structure" | "sort" | "copy" | "fill" | "paste" }
  >,
): void {
  const assertRange = (s: SheetState, r: Rect) => {
    validateRect(r);
    if (r.r2 >= s.meta.rowCount || r.c2 >= s.meta.columnCount)
      throw new Error("Range exceeds sheet dimensions");
  };
  const noComplexArea = (s: SheetState, range: Rect) => {
    if (s.meta.merges.some((m) => intersects(m, range)))
      throw new Error("Unmerge this range before sorting, moving or filling");
    for (let r = range.r1; r <= range.r2; r++)
      for (let c = range.c1; c <= range.c2; c++)
        if (model.engine.spillOwner(s.meta.id, keyOf(r, c)))
          throw new Error(
            "Cannot move or sort spilled results; operate on the formula anchor",
          );
  };
  if (command.type === "paste") {
    pasteCells(model, sheet, command);
    return;
  }
  if (command.type === "structure") {
    if (sheet.meta.objects)
      throw new Error(
        "Remove preserved objects before changing sheet structure",
      );
    const axis = command.axis,
      limit = axis === "row" ? sheet.meta.rowCount : sheet.meta.columnCount;
    if (
      !Number.isInteger(command.index) ||
      !Number.isInteger(command.count) ||
      command.index < 0 ||
      command.count < 1 ||
      command.index > limit ||
      (command.delete && command.index + command.count > limit) ||
      (command.delete && command.count >= limit)
    )
      throw new Error("Invalid row/column operation");
    const delta = command.delete ? -command.count : command.count,
      newLimit = limit + delta;
    if (newLimit > (axis === "row" ? MAX_ROWS : MAX_COLUMNS))
      throw new Error("Excel dimension limit exceeded");
    const move = (index: number): number | undefined =>
      index < command.index
        ? index
        : command.delete && index < command.index + command.count
          ? undefined
          : index + delta;
    const originals = [...sheet.cells];
    for (const [key] of originals) model.setCell(sheet, key, undefined);
    for (const [key, cell] of originals) {
      const r = rowOf(key),
        c = columnOf(key),
        next = move(axis === "row" ? r : c);
      if (next !== undefined)
        model.setCell(
          sheet,
          keyOf(axis === "row" ? next : r, axis === "column" ? next : c),
          cell,
        );
    }
    model.touchMeta(sheet);
    if (axis === "row") sheet.meta.rowCount = newLimit;
    else sheet.meta.columnCount = newLimit;
    const sizes =
        axis === "row" ? sheet.meta.rowHeights : sheet.meta.columnWidths,
      mappedSizes: Record<number, number> = {};
    for (const [index, size] of Object.entries(sizes)) {
      const target = move(Number(index));
      if (target !== undefined) mappedSizes[target] = size;
    }
    if (axis === "row") {
      sheet.meta.rowHeights = mappedSizes;
      sheet.meta.hiddenRows = sheet.meta.hiddenRows
        .map(move)
        .filter((n): n is number => n !== undefined);
      sheet.meta.frozenRows = Math.min(sheet.meta.frozenRows, newLimit - 1);
    } else {
      sheet.meta.columnWidths = mappedSizes;
      sheet.meta.hiddenColumns = sheet.meta.hiddenColumns
        .map(move)
        .filter((n): n is number => n !== undefined);
      sheet.meta.frozenColumns = Math.min(
        sheet.meta.frozenColumns,
        newLimit - 1,
      );
    }
    const transformRange = (range: Rect): Rect | undefined => {
      const first = axis === "row" ? range.r1 : range.c1,
        last = axis === "row" ? range.r2 : range.c2;
      let a = move(first),
        b = move(last);
      if (command.delete) {
        if (a === undefined) a = command.index;
        if (b === undefined) b = command.index - 1;
      }
      if (a === undefined || b === undefined || b < a) return;
      return axis === "row"
        ? { ...range, r1: a, r2: b }
        : { ...range, c1: a, c2: b };
    };
    sheet.meta.merges = sheet.meta.merges
      .map(transformRange)
      .filter((r): r is Rect => !!r);
    for (const property of ["validations", "conditionalFormats"] as const) {
      sheet.meta[property] = sheet.meta[property]?.flatMap((rule) => {
        const range = transformRange(rule.range);
        return range ? [{ ...rule, range }] : [];
      }) as never;
    }
    if (sheet.meta.filter) {
      const range = transformRange(sheet.meta.filter.range);
      if (!range) delete sheet.meta.filter;
      else {
        sheet.meta.filter.range = range;
        if (axis === "column")
          sheet.meta.filter.rules = sheet.meta.filter.rules
            .map((rule) => ({ ...rule, column: move(rule.column) ?? -1 }))
            .filter((rule) => rule.column >= 0);
      }
    }
    for (const [name, ref] of Object.entries(model.names))
      if (ref.sheetId === sheet.meta.id) {
        const range = transformRange(ref.range);
        if (range) model.names[name] = { ...ref, range };
        else delete model.names[name];
      }
    for (const s of model.sheets)
      for (const [key, cell] of s.cells)
        if (cell.formula) {
          try {
            const ast = mapReferences(
              parseFormula(cell.formula),
              (ref) => {
                const targets = ref.sheet
                  ? ref.sheet.toUpperCase() === sheet.meta.name.toUpperCase()
                  : s.meta.id === sheet.meta.id;
                if (!targets) return ref;
                const next = move(axis === "row" ? ref.row : ref.column);
                return next === undefined
                  ? { type: "literal", value: { error: "#REF!" } }
                  : {
                      ...ref,
                      row: axis === "row" ? next : ref.row,
                      column: axis === "column" ? next : ref.column,
                    };
              },
              (range) => {
                const targets = range.start.sheet
                  ? range.start.sheet.toUpperCase() ===
                    sheet.meta.name.toUpperCase()
                  : s.meta.id === sheet.meta.id;
                if (
                  !targets ||
                  (range.whole === "column" && axis === "row") ||
                  (range.whole === "row" && axis === "column")
                )
                  return range;
                const rect = transformRange({
                  r1: Math.min(range.start.row, range.end.row),
                  c1: Math.min(range.start.column, range.end.column),
                  r2: Math.max(range.start.row, range.end.row),
                  c2: Math.max(range.start.column, range.end.column),
                });
                return rect
                  ? {
                      ...range,
                      start: {
                        ...range.start,
                        row:
                          range.start.row <= range.end.row ? rect.r1 : rect.r2,
                        column:
                          range.start.column <= range.end.column
                            ? rect.c1
                            : rect.c2,
                      },
                      end: {
                        ...range.end,
                        row:
                          range.start.row <= range.end.row ? rect.r2 : rect.r1,
                        column:
                          range.start.column <= range.end.column
                            ? rect.c2
                            : rect.c1,
                      },
                    }
                  : { type: "literal", value: { error: "#REF!" } };
              },
            );
            model.setCell(s, key, { ...cell, formula: printFormula(ast) });
          } catch {
            /* Keep unsupported formulas verbatim. */
          }
        }
    model.rebuildAfterTransaction();
    return;
  }
  if (command.type === "sort") {
    assertRange(sheet, command.range);
    noComplexArea(sheet, command.range);
    if (
      !command.keys.length ||
      command.keys.some(
        (k) => k.column < command.range.c1 || k.column > command.range.c2,
      )
    )
      throw new Error("Invalid sort keys");
    const start = command.range.r1 + (command.header ? 1 : 0),
      rows = Array.from(
        { length: Math.max(0, command.range.r2 - start + 1) },
        (_, i) => start + i,
      );
    rows.sort((a, b) => {
      for (const sort of command.keys) {
        const x = model.engine.get(sheet.meta.id, keyOf(a, sort.column)),
          y = model.engine.get(sheet.meta.id, keyOf(b, sort.column));
        const cmp = isError(x)
          ? isError(y)
            ? 0
            : 1
          : isError(y)
            ? -1
            : compare(x, y);
        if (cmp) return sort.direction === "asc" ? cmp : -cmp;
      }
      return a - b;
    });
    const originals = new Map(sheet.cells);
    for (let i = 0; i < rows.length; i++)
      for (let c = command.range.c1; c <= command.range.c2; c++) {
        const from = rows[i],
          to = start + i,
          cell = originals.get(keyOf(from, c));
        model.setCell(
          sheet,
          keyOf(to, c),
          cell?.formula
            ? { ...cell, formula: shiftFormula(cell.formula, to - from, 0) }
            : cell,
        );
      }
    return;
  }
  const source = command.type === "fill" ? command.source : command.range;
  const targetSheet =
    command.type === "fill" ? sheet : model.sheet(command.targetSheetId);
  const target =
    command.type === "fill"
      ? command.target
      : {
          r1: command.targetRow,
          c1: command.targetColumn,
          r2: command.targetRow + source.r2 - source.r1,
          c2: command.targetColumn + source.c2 - source.c1,
        };
  assertRange(sheet, source);
  assertRange(targetSheet, target);
  model.ensureWritable(targetSheet, target);
  if (command.type === "copy" && command.cut)
    model.ensureWritable(sheet, source);
  if (
    targetSheet.meta.protected &&
    !(command.type === "copy" && command.valuesOnly)
  )
    throw new Error("Unprotect the sheet before copying formats");
  noComplexArea(sheet, source);
  noComplexArea(targetSheet, target);
  const originals = new Map<number, Cell>();
  const targetStyles = new Map<number, number | undefined>();
  if (command.type === "copy" && command.valuesOnly)
    for (const [key, cell] of targetSheet.cells)
      if (contains(target, rowOf(key), columnOf(key)))
        targetStyles.set(key, cell.style);
  for (const [key, cell] of sheet.cells)
    if (contains(source, rowOf(key), columnOf(key)))
      originals.set(
        key,
        command.type === "copy" && command.valuesOnly
          ? { value: model.engine.get(sheet.meta.id, key) }
          : { ...cell },
      );
  const rows = source.r2 - source.r1 + 1,
    cols = source.c2 - source.c1 + 1;
  const vertical = cols === 1 && rows > 1;
  const length = vertical ? rows : cols;
  let series: { first: number; step: number } | undefined;
  if (command.type === "fill" && (rows === 1 || cols === 1) && length > 1) {
    const values = Array.from(
      { length },
      (_, i) =>
        originals.get(
          keyOf(source.r1 + (vertical ? i : 0), source.c1 + (vertical ? 0 : i)),
        )?.value,
    );
    if (values.every((value): value is number => typeof value === "number")) {
      const step = values[1] - values[0];
      if (
        values.every((value, i) => {
          const expected = values[0] + step * i;
          return (
            Math.abs(value - expected) <=
            Number.EPSILON *
              16 *
              Math.max(1, Math.abs(value), Math.abs(expected))
          );
        })
      )
        series = { first: values[0], step };
    }
  }
  const modulo = (value: number, size: number) =>
    ((value % size) + size) % size;
  if (command.type === "copy" && command.cut)
    for (const key of originals.keys()) model.setCell(sheet, key, undefined);
  const movedFormulaKeys = new Set<number>();
  for (let r = target.r1; r <= target.r2; r++)
    for (let c = target.c1; c <= target.c2; c++) {
      if (command.type === "fill" && contains(source, r, c)) continue;
      const origin = command.type === "fill" ? source : target;
      const sr = source.r1 + modulo(r - origin.r1, rows),
        sc = source.c1 + modulo(c - origin.c1, cols),
        original = originals.get(keyOf(sr, sc));
      let cell = original ? { ...original } : undefined;
      if (cell?.formula && !(command.type === "copy" && command.cut))
        cell.formula = shiftFormula(cell.formula, r - sr, c - sc);
      if (command.type === "copy" && command.cut && cell?.formula) {
        try {
          cell.formula = printFormula(
            mapReferences(parseFormula(cell.formula), (ref) => {
              const origin = ref.sheet ?? sheet.meta.name;
              if (
                origin.toUpperCase() === sheet.meta.name.toUpperCase() &&
                contains(source, ref.row, ref.column)
              )
                return {
                  ...ref,
                  row: ref.row + target.r1 - source.r1,
                  column: ref.column + target.c1 - source.c1,
                  sheet: ref.sheet ? targetSheet.meta.name : undefined,
                };
              return !ref.sheet && targetSheet !== sheet
                ? { ...ref, sheet: sheet.meta.name }
                : ref;
            }),
          );
          movedFormulaKeys.add(keyOf(r, c));
        } catch {
          /* Unparsed formulas are kept with a diagnostic. */
        }
      }
      if (command.type === "copy" && command.valuesOnly)
        cell = {
          value: original?.value ?? null,
          style: targetStyles.get(keyOf(r, c)),
        };
      if (series && cell)
        cell.value =
          series.first +
          series.step * (vertical ? r - source.r1 : c - source.c1);
      model.setCell(targetSheet, keyOf(r, c), cell);
    }
  if (command.type === "copy" && !command.valuesOnly) {
    const sourceRows = Array.from({ length: rows }, (_, i) => source.r1 + i),
      sourceColumns = Array.from({ length: cols }, (_, i) => source.c1 + i),
      rules = structuredClone({
        validations: sheet.meta.validations ?? [],
        conditionalFormats: sheet.meta.conditionalFormats ?? [],
      });
    if (command.cut)
      pasteRules(
        model,
        sheet,
        { validations: [], conditionalFormats: [] },
        sourceRows,
        sourceColumns,
        sourceRows,
        sourceColumns,
      );
    pasteRules(
      model,
      targetSheet,
      rules,
      sourceRows,
      sourceColumns,
      sourceRows.map((r) => r + target.r1 - source.r1),
      sourceColumns.map((c) => c + target.c1 - source.c1),
    );
  }
  if (command.type === "copy" && command.cut) {
    const dr = target.r1 - source.r1,
      dc = target.c1 - source.c1;
    for (const s of model.sheets)
      for (const [key, cell] of s.cells)
        if (cell.formula && !(s === targetSheet && movedFormulaKeys.has(key)))
          try {
            const formula = printFormula(
              mapReferences(parseFormula(cell.formula), (ref) => {
                const targets = ref.sheet
                  ? ref.sheet.toUpperCase() === sheet.meta.name.toUpperCase()
                  : s.meta.id === sheet.meta.id;
                if (!targets || !contains(source, ref.row, ref.column))
                  return ref;
                return {
                  ...ref,
                  row: ref.row + dr,
                  column: ref.column + dc,
                  sheet:
                    targetSheet.meta.id === s.meta.id && !ref.sheet
                      ? undefined
                      : targetSheet.meta.name,
                };
              }),
            );
            if (formula !== cell.formula)
              model.setCell(s, key, { ...cell, formula });
          } catch {
            /* Unsupported formulas remain verbatim. */
          }
    model.rebuildAfterTransaction();
  }
}

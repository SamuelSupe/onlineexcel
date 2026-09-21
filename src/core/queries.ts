import type { WorkbookModel } from "./model";
import type { Rect, Scalar } from "./types";
import { address, columnOf, contains, keyOf, rowOf } from "./address";
import { isArray, isError } from "../formula/values";

function* computedCells(
  model: WorkbookModel,
  sheetId: string,
): Generator<[number, Scalar]> {
  const sheet = model.sheet(sheetId);
  for (const [key, cell] of sheet.cells) {
    const value = model.engine.get(sheetId, key);
    if (value !== null || cell.formula) yield [key, value];
    if (!cell.formula) continue;
    const result = model.engine.arrayResult(sheetId, key);
    if (!isArray(result)) continue;
    for (let r = 0; r < result.rows; r++)
      for (let c = 0; c < result.columns; c++) {
        const target = keyOf(rowOf(key) + r, columnOf(key) + c);
        if (sheet.cells.has(target)) continue;
        const value = result.get(r, c);
        yield [target, value];
      }
  }
}

export const filterText = (value: Scalar): string =>
  isError(value)
    ? value.error
    : typeof value === "boolean"
      ? value
        ? "TRUE"
        : "FALSE"
      : String(value ?? "");

export function dataRegion(
  model: WorkbookModel,
  sheetId: string,
  selection: Rect,
): Rect {
  model.validateArea(model.sheet(sheetId), selection);
  const rows = new Map<number, number[]>(),
    columns = new Map<number, number[]>();
  let used: Rect | undefined;
  for (const [key] of computedCells(model, sheetId)) {
    const r = rowOf(key),
      c = columnOf(key);
    if (!used) used = { r1: r, r2: r, c1: c, c2: c };
    else {
      used.r1 = Math.min(used.r1, r);
      used.r2 = Math.max(used.r2, r);
      used.c1 = Math.min(used.c1, c);
      used.c2 = Math.max(used.c2, c);
    }
    if (!rows.has(r)) rows.set(r, []);
    if (!columns.has(c)) columns.set(c, []);
    rows.get(r)!.push(c);
    columns.get(c)!.push(r);
  }
  if (!used) return { ...selection };
  const result = {
    r1: Math.max(used.r1, selection.r1),
    r2: Math.min(used.r2, selection.r2),
    c1: Math.max(used.c1, selection.c1),
    c2: Math.min(used.c2, selection.c2),
  };
  if (result.r1 > result.r2 || result.c1 > result.c2) return { ...selection };
  for (const index of [rows, columns])
    for (const values of index.values()) values.sort((a, b) => a - b);
  const has = (values: number[] | undefined, min: number, max: number) => {
    if (!values) return false;
    let left = 0,
      right = values.length;
    while (left < right) {
      const mid = (left + right) >>> 1;
      if (values[mid] < min) left = mid + 1;
      else right = mid;
    }
    return left < values.length && values[left] <= max;
  };
  let changed = true;
  while (changed) {
    changed = false;
    for (const edge of ["r1", "r2", "c1", "c2"] as const) {
      const row = edge[0] === "r",
        step = edge.endsWith("1") ? -1 : 1;
      while (
        has(
          (row ? rows : columns).get(result[edge] + step),
          row ? result.c1 : result.r1,
          row ? result.c2 : result.r2,
        )
      ) {
        result[edge] += step;
        changed = true;
      }
    }
  }
  return result;
}

export function findNext(
  model: WorkbookModel,
  sheetId: string,
  search: string,
  after?: number,
  matchCase = false,
  entireCell = false,
): { address: string; value: Scalar } | null {
  if (!search) throw new Error("Search text is empty");
  const needle = matchCase ? search : search.toLowerCase();
  let first: [number, Scalar] | undefined, next: [number, Scalar] | undefined;
  for (const [key, value] of computedCells(model, sheetId)) {
    const content = matchCase
      ? filterText(value)
      : filterText(value).toLowerCase();
    if (entireCell ? content !== needle : !content.includes(needle)) continue;
    if (!first || key < first[0]) first = [key, value];
    if ((after === undefined || key > after) && (!next || key < next[0]))
      next = [key, value];
  }
  const match = next ?? first;
  return match
    ? { address: address(rowOf(match[0]), columnOf(match[0])), value: match[1] }
    : null;
}

export function distinctValues(
  model: WorkbookModel,
  sheetId: string,
  range: Rect,
  column: number,
): { values: string[]; truncated: boolean } {
  model.validateArea(model.sheet(sheetId), range);
  if (!Number.isInteger(column) || !contains(range, range.r1, column))
    throw new Error("Filter column outside range");
  const values = new Map<string, string>();
  for (let r = range.r1 + 1; r <= range.r2; r++) {
    const value = filterText(model.engine.get(sheetId, keyOf(r, column)));
    values.set(value.toLowerCase(), value);
    if (values.size > 2000) return { values: [], truncated: true };
  }
  return {
    values: [...values.values()].sort((a, b) => a.localeCompare(b)),
    truncated: false,
  };
}

export function navigationTarget(
  model: WorkbookModel,
  sheetId: string,
  row: number,
  column: number,
  direction: "up" | "down" | "left" | "right",
): { row: number; column: number } {
  const sheet = model.sheet(sheetId);
  model.validateArea(sheet, { r1: row, r2: row, c1: column, c2: column });
  if (!["up", "down", "left", "right"].includes(direction))
    throw new Error("Invalid navigation direction");
  const vertical = direction === "up" || direction === "down";
  const step = direction === "up" || direction === "left" ? -1 : 1;
  const count = vertical ? sheet.meta.rowCount : sheet.meta.columnCount;
  const hidden = new Set(
    vertical
      ? [...sheet.meta.hiddenRows, ...(sheet.meta.filteredRows ?? [])]
      : sheet.meta.hiddenColumns,
  );
  const next = (index: number) => {
    let candidate = index + step;
    while (candidate >= 0 && candidate < count && hidden.has(candidate))
      candidate += step;
    return candidate >= 0 && candidate < count ? candidate : index;
  };
  const filled = (index: number) => {
    let r = vertical ? index : row,
      c = vertical ? column : index;
    const merge = sheet.meta.merges.find((range) => contains(range, r, c));
    if (merge) {
      r = merge.r1;
      c = merge.c1;
    }
    const key = keyOf(r, c);
    if (sheet.cells.get(key)?.formula) return true;
    const value = model.engine.get(sheetId, key);
    return value !== null && value !== "";
  };
  let current = vertical ? row : column;
  const adjacent = next(current);
  if (adjacent !== current) {
    const contiguous = filled(current) && filled(adjacent);
    current = adjacent;
    if (contiguous) {
      while (next(current) !== current && filled(next(current)))
        current = next(current);
    } else {
      while (!filled(current) && next(current) !== current)
        current = next(current);
    }
  }
  return vertical ? { row: current, column } : { row, column: current };
}

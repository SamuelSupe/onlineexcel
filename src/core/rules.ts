import type { WorkbookModel, SheetState } from "./model";
import type {
  CellStyle,
  Rect,
  Scalar,
  SheetMeta,
  ValidationRule,
} from "./types";
import { address, contains, keyOf, validateRect } from "./address";
import { compare, isError } from "../formula/values";

export function checkMergedWrite(
  meta: SheetMeta,
  row: number,
  column: number,
  value: Scalar | undefined,
  formula?: string,
) {
  if (!formula && (value === null || value === undefined)) return;
  if (
    meta.merges.some(
      (merge) =>
        contains(merge, row, column) &&
        (row !== merge.r1 || column !== merge.c1),
    )
  )
    throw new Error(
      "Cannot write hidden content into a merged cell; edit its top-left cell",
    );
}

export function checkProtection(
  model: WorkbookModel,
  sheet: SheetState,
  range: Rect,
) {
  if (!sheet.meta.protected) return;
  for (let r = range.r1; r <= range.r2; r++)
    for (let c = range.c1; c <= range.c2; c++)
      if (
        model.styles[sheet.cells.get(keyOf(r, c))?.style ?? 0]?.locked !== false
      )
        throw new Error(`Protected cell: ${address(r, c)}`);
}
export function validateRule(rule: Omit<ValidationRule, "range">) {
  if (!["list", "decimal", "whole", "date", "textLength"].includes(rule.type))
    throw new Error("Invalid validation rule");
  if (
    rule.type === "list" &&
    (!rule.values?.length || rule.values.some((v) => typeof v !== "string"))
  )
    throw new Error("Validation list cannot be empty");
  if (
    [rule.minimum, rule.maximum].some(
      (n) => n !== undefined && !Number.isFinite(n),
    ) ||
    (rule.minimum !== undefined &&
      rule.maximum !== undefined &&
      rule.minimum > rule.maximum)
  )
    throw new Error("Invalid validation limits");
}
export function validateEntry(
  meta: SheetMeta,
  row: number,
  column: number,
  value: Scalar,
) {
  for (const rule of meta.validations ?? []) {
    if (!contains(rule.range, row, column)) continue;
    if (value === null || value === "") {
      if (rule.allowBlank !== false) continue;
      throw new Error(`Validation failed: ${address(row, column)}`);
    }
    const n = rule.type === "textLength" ? String(value).length : value;
    const valid =
      rule.type === "list"
        ? !isError(value) &&
          rule.values!.some(
            (v) => v.toLowerCase() === String(value).toLowerCase(),
          )
        : typeof n === "number" &&
          Number.isFinite(n) &&
          (rule.type !== "whole" || Number.isInteger(n)) &&
          (rule.minimum === undefined || n >= rule.minimum) &&
          (rule.maximum === undefined || n <= rule.maximum);
    if (!valid) throw new Error(`Validation failed: ${address(row, column)}`);
  }
}
export function conditionalStyle(
  meta: SheetMeta,
  row: number,
  column: number,
  value: Scalar,
  base: CellStyle,
): CellStyle | undefined {
  let result: CellStyle | undefined;
  for (const rule of meta.conditionalFormats ?? []) {
    if (!contains(rule.range, row, column) || isError(value) || value === null)
      continue;
    const order = compare(value, rule.value);
    const match =
      rule.operator === "contains"
        ? String(value)
            .toLowerCase()
            .includes(String(rule.value ?? "").toLowerCase())
        : rule.operator === "eq"
          ? order === 0
          : rule.operator === "neq"
            ? order !== 0
            : rule.operator === "gt"
              ? order > 0
              : rule.operator === "gte"
                ? order >= 0
                : rule.operator === "lt"
                  ? order < 0
                  : rule.operator === "lte"
                    ? order <= 0
                    : order >= 0 &&
                      typeof value === "number" &&
                      value <= rule.second!;
    if (match) result = { ...(result ?? base), ...rule.style };
  }
  return result;
}

export function subtractRange(source: Rect, removed: Rect): Rect[] {
  const r1 = Math.max(source.r1, removed.r1),
    r2 = Math.min(source.r2, removed.r2),
    c1 = Math.max(source.c1, removed.c1),
    c2 = Math.min(source.c2, removed.c2);
  if (r1 > r2 || c1 > c2) return [source];
  return [
    { ...source, r2: r1 - 1 },
    { ...source, r1: r2 + 1 },
    { r1, r2, c1: source.c1, c2: c1 - 1 },
    { r1, r2, c1: c2 + 1, c2: source.c2 },
  ].filter((rect) => rect.r1 <= rect.r2 && rect.c1 <= rect.c2);
}

function contiguousRuns(indices: number[]): [number, number][] {
  const runs: [number, number][] = [];
  for (const index of indices) {
    const last = runs.at(-1);
    if (last && last[1] + 1 === index) last[1] = index;
    else runs.push([index, index]);
  }
  return runs;
}

export function pasteRules(
  model: WorkbookModel,
  sheet: SheetState,
  source: Pick<SheetMeta, "validations" | "conditionalFormats">,
  sourceRows: number[],
  sourceColumns: number[],
  targetRows: number[],
  targetColumns: number[],
  transpose = false,
) {
  const rectangles = (rows: number[], columns: number[]): Rect[] =>
    contiguousRuns(rows).flatMap(([r1, r2]) =>
      contiguousRuns(columns).map(([c1, c2]) => ({ r1, r2, c1, c2 })),
    );
  const destinations = rectangles(targetRows, targetColumns);
  for (const property of ["validations", "conditionalFormats"] as const) {
    const incoming = source[property];
    // Omitted metadata is an external/value-only clipboard, not a request to clear rules.
    if (!incoming) continue;
    for (const rule of incoming) {
      validateRect(rule.range);
      if ("type" in rule) validateRule(rule);
      else model.styleId(rule.style);
    }
    const mapped = incoming.flatMap((rule) => {
      const rows = sourceRows.flatMap((n, i) =>
        n >= rule.range.r1 && n <= rule.range.r2
          ? [(transpose ? targetColumns : targetRows)[i]]
          : [],
      );
      const columns = sourceColumns.flatMap((n, i) =>
        n >= rule.range.c1 && n <= rule.range.c2
          ? [(transpose ? targetRows : targetColumns)[i]]
          : [],
      );
      return rectangles(
        transpose ? columns : rows,
        transpose ? rows : columns,
      ).map((range) => ({ ...structuredClone(rule), range }));
    });
    model.touchMeta(sheet);
    const retained = (sheet.meta[property] ?? []).flatMap((rule) => {
      let ranges = [rule.range];
      for (const destination of destinations)
        ranges = ranges.flatMap((range) => subtractRange(range, destination));
      return ranges.map((range) => ({ ...rule, range }));
    });
    sheet.meta[property] = [...retained, ...mapped] as never;
  }
}

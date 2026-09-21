import type { CellError, ErrorCode, Scalar } from "../core/types";
export interface ArrayValue {
  rows: number;
  columns: number;
  get(row: number, column: number): Scalar;
  reference?: { sheetId: string; row: number; column: number };
}
export type Value = Scalar | ArrayValue;
export const error = (code: ErrorCode): CellError => ({ error: code });
export const isError = (value: unknown): value is CellError =>
  !!value && typeof value === "object" && "error" in value;
export const isArray = (value: Value | undefined): value is ArrayValue =>
  !!value && typeof value === "object" && "get" in value;
export const scalar = (value: Value): Scalar =>
  isArray(value) ? value.get(0, 0) : value;
export function number(value: Value): number {
  const v = scalar(value);
  if (isError(v)) throw v;
  if (v === null || v === "") return 0;
  const n = Number(v);
  if (!Number.isFinite(n)) throw error("#VALUE!");
  return n;
}
export function text(value: Value): string {
  const v = scalar(value);
  if (isError(v)) throw v;
  return v === null
    ? ""
    : typeof v === "boolean"
      ? v
        ? "TRUE"
        : "FALSE"
      : String(v);
}
export function bool(value: Value): boolean {
  const v = scalar(value);
  if (isError(v)) throw v;
  if (typeof v === "string" && /^(true|false)$/i.test(v))
    return v.toUpperCase() === "TRUE";
  return number(v) !== 0;
}
export function* flatten(values: Value[]): Generator<Scalar> {
  for (const v of values) {
    if (isArray(v)) {
      for (let r = 0; r < v.rows; r++)
        for (let c = 0; c < v.columns; c++) yield v.get(r, c);
    } else yield v;
  }
}
export function numbers(values: Value[], includeLogical = false): number[] {
  const out: number[] = [];
  for (const value of values) {
    const ranged = isArray(value);
    for (const v of flatten([value])) {
      if (isError(v)) throw v;
      if (typeof v === "number") out.push(v);
      else if (includeLogical && v !== null)
        out.push(typeof v === "boolean" ? +v : 0);
      else if (!ranged && v !== null && v !== "") out.push(number(v));
    }
  }
  return out;
}
export function matrix(rows: Scalar[][]): ArrayValue {
  return {
    rows: rows.length,
    columns: rows[0]?.length ?? 0,
    get: (r, c) => rows[r]?.[c] ?? null,
  };
}
export function array(value: Value): ArrayValue {
  return isArray(value) ? value : matrix([[value]]);
}
export function finite(value: number): Scalar {
  return Number.isFinite(value) ? value : error("#NUM!");
}
export function compare(a: Scalar, b: Scalar): number {
  if (isError(a)) throw a;
  if (isError(b)) throw b;
  if (a === null) a = typeof b === "string" ? "" : 0;
  if (b === null) b = typeof a === "string" ? "" : 0;
  if (typeof a !== typeof b) {
    const rank = (v: Scalar) =>
      typeof v === "boolean" ? 3 : typeof v === "string" ? 2 : 1;
    return rank(a) - rank(b);
  }
  if (typeof a === "string" && typeof b === "string")
    return a.toUpperCase() < b.toUpperCase()
      ? -1
      : a.toUpperCase() > b.toUpperCase()
        ? 1
        : 0;
  return a < b ? -1 : a > b ? 1 : 0;
}
export function wildcardPattern(operand: string): RegExp {
  let regex = "";
  for (let i = 0; i < operand.length; i++) {
    const c = operand[i];
    if (c === "~" && i + 1 < operand.length)
      regex += operand[++i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    else
      regex +=
        c === "*"
          ? ".*"
          : c === "?"
            ? "."
            : c.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }
  return new RegExp("^" + regex + "$", "i");
}
export function criteria(condition: Value): (value: Scalar) => boolean {
  const value = scalar(condition);
  if (isError(value)) throw value;
  if (typeof value !== "string")
    return (v) => !isError(v) && compare(v, value) === 0;
  const match = /^(<=|>=|<>|=|<|>)(.*)$/.exec(value),
    op = match?.[1] ?? "=",
    operand = match?.[2] ?? value;
  const numeric = operand.trim() !== "" && Number.isFinite(Number(operand));
  const pattern = wildcardPattern(operand);
  return (v) => {
    if (isError(v)) return false;
    const order = compare(v, numeric ? Number(operand) : operand);
    const equal = numeric
      ? order === 0
      : pattern.test(v === null ? "" : String(v));
    return op === "="
      ? equal
      : op === "<>"
        ? !equal
        : op === "<"
          ? order < 0
          : op === ">"
            ? order > 0
            : op === "<="
              ? order <= 0
              : order >= 0;
  };
}
export const DAY = 86_400_000;
export function serialDate(
  serial: number,
  dateSystem: 1900 | 1904 = 1900,
): Date {
  return new Date(
    (dateSystem === 1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31)) +
      (serial - (dateSystem === 1900 && serial >= 60 ? 1 : 0)) * DAY,
  );
}
export function dateSerial(date: Date, dateSystem: 1900 | 1904 = 1900): number {
  let n =
    (date.getTime() -
      (dateSystem === 1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31))) /
    DAY;
  if (dateSystem === 1900 && n >= 60) n++;
  return n;
}

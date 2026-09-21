import type { Rect } from "./types";
export const MAX_ROWS = 1_048_576;
export const MAX_COLUMNS = 16_384;
export const keyOf = (row: number, column: number): number =>
  row * MAX_COLUMNS + column;
export const rowOf = (key: number): number => Math.floor(key / MAX_COLUMNS);
export const columnOf = (key: number): number => key % MAX_COLUMNS;
export function columnName(column: number): string {
  let name = "";
  for (let n = column + 1; n > 0; n = Math.floor((n - 1) / 26))
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}
export function columnIndex(name: string): number {
  let n = 0;
  for (const c of name.toUpperCase()) n = n * 26 + c.charCodeAt(0) - 64;
  return n - 1;
}
export function address(row: number, column: number): string {
  return `${columnName(column)}${row + 1}`;
}
export function parseCell(input: string): {
  row: number;
  column: number;
  absoluteRow: boolean;
  absoluteColumn: boolean;
} {
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)$/.exec(input);
  if (!match) throw new Error(`Invalid cell address: ${input}`);
  const row = Number(match[4]) - 1,
    column = columnIndex(match[2]);
  if (row >= MAX_ROWS || column >= MAX_COLUMNS)
    throw new Error(`Address outside Excel limits: ${input}`);
  return { row, column, absoluteRow: !!match[3], absoluteColumn: !!match[1] };
}
export function parseRange(input: string | Rect): Rect {
  if (typeof input !== "string") {
    validateRect(input);
    return { ...input };
  }
  const parts = input.split(":");
  if (parts.length > 2) throw new Error(`Invalid range: ${input}`);
  const a = parseCell(parts[0]),
    b = parseCell(parts[1] ?? parts[0]);
  return {
    r1: Math.min(a.row, b.row),
    c1: Math.min(a.column, b.column),
    r2: Math.max(a.row, b.row),
    c2: Math.max(a.column, b.column),
  };
}
export function validateRect(r: Rect): void {
  if (
    ![r.r1, r.r2, r.c1, r.c2].every(Number.isInteger) ||
    r.r1 < 0 ||
    r.c1 < 0 ||
    r.r2 < r.r1 ||
    r.c2 < r.c1 ||
    r.r2 >= MAX_ROWS ||
    r.c2 >= MAX_COLUMNS
  )
    throw new Error("Invalid range bounds");
}
export const contains = (r: Rect, row: number, col: number): boolean =>
  row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;
export const intersects = (a: Rect, b: Rect): boolean =>
  a.r1 <= b.r2 && a.r2 >= b.r1 && a.c1 <= b.c2 && a.c2 >= b.c1;
export const rangeName = (r: Rect): string =>
  `${address(r.r1, r.c1)}:${address(r.r2, r.c2)}`;

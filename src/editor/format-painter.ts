import type { Command, ConditionalRule, Rect, RegionCell } from "../core/types";
import { keyOf } from "../core/address";

export interface PaintSource {
  rows: number[];
  columns: number[];
  cells: RegionCell[];
  conditionalFormats: ConditionalRule[];
}
export function paintCommand(
  source: PaintSource,
  sheetId: string,
  targetRows: number[],
  targetColumns: number[],
): Command {
  const height = targetRows.length,
    width = targetColumns.length;
  if (
    !height ||
    !width ||
    height * width > 200000 ||
    !source.rows.length ||
    !source.columns.length
  )
    throw new Error("Invalid clipboard dimensions");
  const cells = new Map(
    source.cells.map((cell) => [keyOf(cell.row, cell.column), cell]),
  );
  const rows = source.rows.length,
    columns = source.columns.length;
  const conditionalFormats: ConditionalRule[] = [];
  for (const rule of source.conditionalFormats) {
    const rs = source.rows
      .map((row, i) => (row >= rule.range.r1 && row <= rule.range.r2 ? i : -1))
      .filter((i) => i >= 0);
    const cs = source.columns
      .map((col, i) => (col >= rule.range.c1 && col <= rule.range.c2 ? i : -1))
      .filter((i) => i >= 0);
    if (!rs.length || !cs.length) continue;
    for (let r = 0; r < height; r += rows)
      for (let c = 0; c < width; c += columns) {
        const range: Rect = {
          r1: r + rs[0],
          r2: Math.min(height - 1, r + rs.at(-1)!),
          c1: c + cs[0],
          c2: Math.min(width - 1, c + cs.at(-1)!),
        };
        if (range.r1 <= range.r2 && range.c1 <= range.c2)
          conditionalFormats.push({ ...structuredClone(rule), range });
      }
  }
  return {
    type: "paste",
    sheetId,
    source: { r1: 0, r2: height - 1, c1: 0, c2: width - 1 },
    targetRow: targetRows[0],
    targetColumn: targetColumns[0],
    targetRows,
    targetColumns,
    mode: "formats",
    conditionalFormats,
    cells: Array.from({ length: height * width }, (_, i) => {
      const row = Math.floor(i / width),
        column = i % width;
      return {
        row,
        column,
        value: null,
        style: structuredClone(
          cells.get(
            keyOf(source.rows[row % rows], source.columns[column % columns]),
          )?.style ?? {},
        ),
      };
    }),
  };
}

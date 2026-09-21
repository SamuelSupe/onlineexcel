import type { Workbook } from "../runtime/client";
import type { Command, Rect, SheetMeta } from "../core/types";
import { contains } from "../core/address";
import { formatValue } from "../core/format";
import { DEFAULT_COLUMN, DEFAULT_ROW } from "./geometry";

export async function autoFit(
  workbook: Workbook,
  sheet: SheetMeta,
  selection: Rect,
  axis: "row" | "column",
  signal: AbortSignal,
) {
  const dateSystem = (await workbook.getMetadata()).dateSystem;
  const range = await workbook.getDataRegion(
    sheet.id,
    axis === "row"
      ? { ...selection, c1: 0, c2: sheet.columnCount - 1 }
      : { ...selection, r1: 0, r2: sheet.rowCount - 1 },
  );
  if (axis === "row") {
    range.r1 = selection.r1;
    range.r2 = selection.r2;
  } else {
    range.c1 = selection.c1;
    range.c2 = selection.c2;
  }
  const ctx = document.createElement("canvas").getContext("2d")!;
  const sizes = new Map<number, number>(),
    step = Math.max(1, Math.floor(50000 / (range.c2 - range.c1 + 1)));
  for (let start = range.r1; start <= range.r2; start += step) {
    signal.throwIfAborted();
    const region = await workbook.getRegion(sheet.id, {
      ...range,
      r1: start,
      r2: Math.min(range.r2, start + step - 1),
    });
    for (const cell of region.cells) {
      if (
        cell.value === null ||
        sheet.merges.some((merge) => contains(merge, cell.row, cell.column))
      )
        continue;
      const style = cell.style,
        fontSize = ((style.fontSize ?? 10) * 4) / 3;
      ctx.font = `${style.italic ? "italic " : ""}${style.bold ? "600 " : ""}${fontSize}px ${style.fontFamily ?? "Arial, sans-serif"}`;
      const lines = formatValue(cell.value, style, dateSystem).split("\n");
      let size: number;
      if (axis === "column")
        size =
          Math.max(...lines.map((line) => ctx.measureText(line).width)) + 18;
      else {
        let count = 0;
        const width = (sheet.columnWidths[cell.column] ?? DEFAULT_COLUMN) - 16;
        for (const original of lines) {
          let line = "";
          count++;
          if (style.wrap)
            for (const char of original) {
              if (line && ctx.measureText(line + char).width > width) {
                count++;
                line = "";
              }
              line += char;
            }
        }
        size = count * fontSize * 1.35 + 6;
      }
      const index = axis === "row" ? cell.row : cell.column;
      sizes.set(index, Math.max(sizes.get(index) ?? 0, size));
    }
  }
  const grouped = new Map<number, number[]>(),
    start = axis === "row" ? selection.r1 : selection.c1,
    end = axis === "row" ? selection.r2 : selection.c2;
  for (let index = start; index <= end; index++) {
    // Merged cells do not have an unambiguous independent row/column size.
    if (
      sheet.merges.some((m) =>
        axis === "row"
          ? index >= m.r1 && index <= m.r2
          : index >= m.c1 && index <= m.c2,
      )
    )
      continue;
    const size = Math.min(
      4096,
      Math.max(
        8,
        Math.ceil(
          sizes.get(index) ?? (axis === "row" ? DEFAULT_ROW : DEFAULT_COLUMN),
        ),
      ),
    );
    if (!grouped.has(size)) grouped.set(size, []);
    grouped.get(size)!.push(index);
  }
  const commands: Command[] = [...grouped].map(([size, indexes]) => ({
    type: "dimensions",
    sheetId: sheet.id,
    axis,
    indexes,
    size,
  }));
  if (commands.length) await workbook.transaction(commands, { signal });
}

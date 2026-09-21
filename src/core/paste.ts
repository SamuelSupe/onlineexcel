import type { WorkbookModel, SheetState } from "./model";
import type { Command } from "./types";
import { intersects, keyOf, validateRect } from "./address";
import { shiftFormula } from "../formula/parser";
import { checkProtection, checkMergedWrite, pasteRules } from "./rules";
import { isError } from "../formula/values";

export function pasteCells(
  model: WorkbookModel,
  sheet: SheetState,
  command: Extract<Command, { type: "paste" }>,
) {
  validateRect(command.source);
  const sequence = (start: number, end: number) =>
    Array.from({ length: end - start + 1 }, (_, i) => start + i);
  const sourceRows =
    command.sourceRows ?? sequence(command.source.r1, command.source.r2);
  const sourceColumns =
    command.sourceColumns ?? sequence(command.source.c1, command.source.c2);
  const height = command.transpose ? sourceColumns.length : sourceRows.length;
  const width = command.transpose ? sourceRows.length : sourceColumns.length;
  if (!height || !width || height * width > 200000)
    throw new Error("Invalid clipboard dimensions");
  const targetRows =
    command.targetRows ??
    sequence(command.targetRow, command.targetRow + height - 1);
  const targetColumns =
    command.targetColumns ??
    sequence(command.targetColumn, command.targetColumn + width - 1);
  for (const [indices, count, lo, hi] of [
    [sourceRows, sourceRows.length, command.source.r1, command.source.r2],
    [sourceColumns, sourceColumns.length, command.source.c1, command.source.c2],
    [targetRows, height, 0, sheet.meta.rowCount - 1],
    [targetColumns, width, 0, sheet.meta.columnCount - 1],
  ] as const)
    if (
      indices.length !== count ||
      indices.some(
        (n, i) =>
          !Number.isInteger(n) ||
          n < lo ||
          n > hi ||
          (i > 0 && n <= indices[i - 1]),
      )
    )
      throw new Error("Invalid clipboard coordinates");
  const source = new Map(
    command.cells.map((cell) => [keyOf(cell.row, cell.column), cell]),
  );
  const mode = command.mode ?? "all";
  const covered = (indices: number[], first: number, last: number) =>
    indices.filter((n) => n >= first && n <= last).length;
  const merged = sheet.meta.merges.filter(
    (m) =>
      covered(targetRows, m.r1, m.r2) && covered(targetColumns, m.c1, m.c2),
  );
  if (
    merged.some(
      (m) =>
        covered(targetRows, m.r1, m.r2) !== m.r2 - m.r1 + 1 ||
        covered(targetColumns, m.c1, m.c2) !== m.c2 - m.c1 + 1,
    )
  )
    throw new Error("Cannot paste into part of a merged range");
  if (sheet.meta.protected && mode !== "values" && mode !== "formulas")
    throw new Error("Unprotect the sheet before changing formats");
  for (let r = 0; r < height; r++)
    for (let c = 0; c < width; c++) {
      const row = targetRows[r],
        column = targetColumns[c];
      checkProtection(model, sheet, {
        r1: row,
        r2: row,
        c1: column,
        c2: column,
      });
      const owner = model.engine.spillOwner(sheet.meta.id, keyOf(row, column));
      if (
        owner &&
        (!targetRows.includes(owner.row) ||
          !targetColumns.includes(owner.column))
      )
        throw new Error(
          "Cannot edit part of a spilled array; edit its anchor cell",
        );
      const sr = sourceRows[command.transpose ? c : r],
        sc = sourceColumns[command.transpose ? r : c];
      const cell = source.get(keyOf(sr, sc));
      if (
        cell &&
        ((cell.formula && !cell.formula.startsWith("=")) ||
          (cell.value !== null &&
            !["number", "string", "boolean"].includes(typeof cell.value) &&
            !isError(cell.value)))
      )
        throw new Error("Invalid clipboard cell");
      if (
        cell?.spill &&
        mode !== "values" &&
        mode !== "formats" &&
        !source.get(keyOf(cell.spill.row, cell.spill.column))?.formula
      )
        throw new Error("Cannot copy part of a spilled array");
      const key = keyOf(row, column),
        old = sheet.cells.get(key);
      if (mode === "values" || mode === "formulas")
        checkMergedWrite(
          sheet.meta,
          row,
          column,
          cell?.spill && mode !== "values" ? undefined : cell?.value,
          mode === "formulas" ? cell?.formula : undefined,
        );
      const style =
        mode === "values" || mode === "formulas"
          ? old?.style
          : cell
            ? model.styleId(structuredClone(cell.style))
            : 0;
      if (mode === "formats") model.setCell(sheet, key, { ...old, style });
      else if (cell?.spill && mode !== "values")
        model.setCell(sheet, key, { style });
      else
        model.setCell(
          sheet,
          key,
          cell?.formula && mode !== "values"
            ? {
                formula: shiftFormula(cell.formula, row - sr, column - sc),
                style,
              }
            : { value: structuredClone(cell?.value ?? null), style },
        );
    }
  if (mode === "all") {
    model.touchMeta(sheet);
    sheet.meta.merges = sheet.meta.merges.filter((m) => !merged.includes(m));
    for (const merge of command.merges ?? []) {
      const rs = sourceRows.filter((r) => r >= merge.r1 && r <= merge.r2);
      const cs = sourceColumns.filter((c) => c >= merge.c1 && c <= merge.c2);
      if (
        rs.length !== merge.r2 - merge.r1 + 1 ||
        cs.length !== merge.c2 - merge.c1 + 1
      )
        throw new Error("Cannot copy part of a merged range");
      const ri = sourceRows.indexOf(merge.r1),
        ci = sourceColumns.indexOf(merge.c1);
      const tr = command.transpose ? ci : ri,
        tc = command.transpose ? ri : ci;
      const rh = command.transpose ? cs.length : rs.length,
        cw = command.transpose ? rs.length : cs.length;
      const next = {
        r1: targetRows[tr],
        r2: targetRows[tr + rh - 1],
        c1: targetColumns[tc],
        c2: targetColumns[tc + cw - 1],
      };
      if (next.r2 - next.r1 + 1 !== rh || next.c2 - next.c1 + 1 !== cw)
        throw new Error("Merged cells require contiguous paste destinations");
      if (sheet.meta.merges.some((merge) => intersects(merge, next)))
        throw new Error("Overlapping merged cells");
      sheet.meta.merges.push(next);
      for (let row = next.r1; row <= next.r2; row++)
        for (let column = next.c1; column <= next.c2; column++) {
          const cell = sheet.cells.get(keyOf(row, column));
          checkMergedWrite(sheet.meta, row, column, cell?.value, cell?.formula);
        }
    }
  }
  if (mode === "all" || mode === "formats")
    pasteRules(
      model,
      sheet,
      {
        validations: mode === "all" ? command.validations : undefined,
        conditionalFormats: command.conditionalFormats,
      },
      sourceRows,
      sourceColumns,
      targetRows,
      targetColumns,
      command.transpose,
    );
}

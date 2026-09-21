import { sheetCommand } from "./sheet-commands";
import { checkProtection, checkMergedWrite } from "./rules";
import type { WorkbookModel } from "./model";
import type { Command } from "./types";
import { keyOf, contains, rowOf, columnOf, intersects } from "./address";
import { mapReferences, parseFormula, printFormula } from "../formula/parser";
import { executeStructural } from "./structural";
export const validName = (name: string): boolean =>
  name.length > 0 &&
  name.length <= 31 &&
  !/[\\/?*\[\]:]/.test(name) &&
  !name.startsWith("'") &&
  !name.endsWith("'");
export function executeCommand(
  model: WorkbookModel,
  command: Command,
  writableChecked = false,
): void {
  if (sheetCommand(model, command)) return;
  if (
    command.type === "renameName" ||
    command.type === "deleteName" ||
    command.type === "duplicateSheet" ||
    command.type === "sheetVisibility" ||
    command.type === "protect" ||
    command.type === "validation" ||
    command.type === "conditionalFormat"
  )
    return;
  if (command.type === "addSheet") {
    model.sheets.push(
      model.newSheet(command.name, command.rows, command.columns, command.id),
    );
    model.rebuildAfterTransaction();
    return;
  }
  if (command.type === "defineName") {
    if (
      !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(command.name) ||
      /^\$?[A-Za-z]{1,3}\$?\d+$/.test(command.name)
    )
      throw new Error("Invalid named range");
    model.validateArea(model.sheet(command.sheetId), command.range);
    model.names[command.name.toUpperCase()] = {
      sheetId: command.sheetId,
      range: { ...command.range },
    };
    model.rebuildAfterTransaction();
    return;
  }
  const sheet = model.sheet(command.sheetId);
  if (
    sheet.meta.protected &&
    ![
      "setValues",
      "setFormula",
      "clear",
      "copy",
      "paste",
      "fill",
      "replace",
      "filter",
    ].includes(command.type)
  )
    throw new Error(
      "Unprotect the sheet before changing its structure or formats",
    );
  switch (command.type) {
    case "setValues": {
      if (!writableChecked) model.ensureWritable(sheet, command.range);
      const { r1, r2, c1, c2 } = command.range;
      if (
        command.values.length !== r2 - r1 + 1 ||
        command.values.some((row) => row.length !== c2 - c1 + 1)
      )
        throw new Error("Values must match range dimensions");
      for (let r = r1; r <= r2; r++)
        for (let c = c1; c <= c2; c++) {
          const value = command.values[r - r1][c - c1];
          if (
            (value !== null &&
              !["string", "number", "boolean"].includes(typeof value)) ||
            (typeof value === "number" && !Number.isFinite(value))
          )
            throw new Error("Invalid cell value");
          checkMergedWrite(sheet.meta, r, c, value);
          const key = keyOf(r, c),
            style = sheet.cells.get(key)?.style;
          model.setCell(
            sheet,
            key,
            typeof value === "string" &&
              value.startsWith("=") &&
              command.parseFormulas !== false
              ? { formula: value, style }
              : { value, style },
          );
        }
      break;
    }
    case "setFormula": {
      model.ensureWritable(sheet, {
        r1: command.row,
        r2: command.row,
        c1: command.column,
        c2: command.column,
      });
      if (!command.formula.startsWith("="))
        throw new Error("Formula must begin with =");
      checkMergedWrite(
        sheet.meta,
        command.row,
        command.column,
        undefined,
        command.formula,
      );
      const key = keyOf(command.row, command.column);
      model.setCell(sheet, key, {
        formula: command.formula,
        style: sheet.cells.get(key)?.style,
      });
      break;
    }
    case "clear": {
      if (sheet.meta.protected && command.formats)
        throw new Error("Unprotect the sheet before changing formats");
      model.ensureWritable(sheet, command.range);
      for (const [key, cell] of sheet.cells)
        if (contains(command.range, rowOf(key), columnOf(key)))
          model.setCell(
            sheet,
            key,
            command.formats ? undefined : { style: cell.style },
          );
      break;
    }
    case "style": {
      model.validateArea(sheet, command.range);
      if (
        (command.range.r2 - command.range.r1 + 1) *
          (command.range.c2 - command.range.c1 + 1) >
        1_000_000
      )
        throw new Error(
          "Format ranges are limited to 1,000,000 cells per transaction",
        );
      const ids = new Map<number, number>();
      for (let r = command.range.r1; r <= command.range.r2; r++)
        for (let c = command.range.c1; c <= command.range.c2; c++) {
          const key = keyOf(r, c),
            cell = sheet.cells.get(key) ?? {},
            old = cell.style ?? 0;
          if (!ids.has(old))
            ids.set(
              old,
              model.styleId({
                ...model.styles[old],
                ...command.style,
                border: command.style.border
                  ? { ...model.styles[old]?.border, ...command.style.border }
                  : model.styles[old]?.border,
              }),
            );
          model.setCell(sheet, key, { ...cell, style: ids.get(old) });
        }
      break;
    }
    case "deleteSheet": {
      if (
        model.sheets.length === 1 ||
        !model.sheets.some((s) => s !== sheet && !s.meta.hidden)
      )
        throw new Error("Cannot delete the last sheet");
      for (const other of model.sheets)
        if (other !== sheet)
          for (const [key, cell] of other.cells)
            if (cell.formula) {
              try {
                const formula = printFormula(
                  mapReferences(parseFormula(cell.formula), (ref) =>
                    ref.sheet?.toUpperCase() === sheet.meta.name.toUpperCase()
                      ? { type: "literal", value: { error: "#REF!" } }
                      : ref,
                  ),
                );
                if (formula !== cell.formula)
                  model.setCell(other, key, { ...cell, formula });
              } catch {
                /* Unsupported formulas remain diagnosed. */
              }
            }
      model.sheets = model.sheets.filter((s) => s !== sheet);
      for (const [name, value] of Object.entries(model.names))
        if (value.sheetId === sheet.meta.id) delete model.names[name];
      model.rebuildAfterTransaction();
      break;
    }
    case "renameSheet": {
      if (
        !validName(command.name) ||
        model.sheets.some(
          (s) =>
            s !== sheet &&
            s.meta.name.toUpperCase() === command.name.toUpperCase(),
        )
      )
        throw new Error("Invalid or duplicate sheet name");
      const oldName = sheet.meta.name;
      model.touchMeta(sheet);
      sheet.meta.name = command.name;
      for (const s of model.sheets)
        for (const [key, cell] of s.cells)
          if (cell.formula) {
            try {
              const formula = printFormula(
                mapReferences(parseFormula(cell.formula), (ref) =>
                  ref.sheet?.toUpperCase() === oldName.toUpperCase()
                    ? { ...ref, sheet: command.name }
                    : ref,
                ),
              );
              if (formula !== cell.formula)
                model.setCell(s, key, { ...cell, formula });
            } catch {
              /* Unparsed formulas are preserved and diagnosed. */
            }
          }
      model.rebuildAfterTransaction();
      break;
    }
    case "reorderSheet": {
      if (
        !Number.isInteger(command.index) ||
        command.index < 0 ||
        command.index >= model.sheets.length
      )
        throw new Error("Invalid sheet index");
      model.sheets.splice(model.sheets.indexOf(sheet), 1);
      model.sheets.splice(command.index, 0, sheet);
      break;
    }
    case "dimensions": {
      model.touchMeta(sheet);
      const max =
          command.axis === "row" ? sheet.meta.rowCount : sheet.meta.columnCount,
        sizes =
          command.axis === "row"
            ? sheet.meta.rowHeights
            : sheet.meta.columnWidths,
        hidden = new Set(
          command.axis === "row"
            ? sheet.meta.hiddenRows
            : sheet.meta.hiddenColumns,
        );
      if (
        command.size !== undefined &&
        (!Number.isFinite(command.size) ||
          command.size < 8 ||
          command.size > 4096)
      )
        throw new Error("Dimension size must be between 8 and 4096");
      for (const index of command.indexes) {
        if (!Number.isInteger(index) || index < 0 || index >= max)
          throw new Error("Invalid dimension index");
        if (command.size !== undefined) sizes[index] = command.size;
        if (command.hidden === true) hidden.add(index);
        else if (command.hidden === false) hidden.delete(index);
      }
      if (command.axis === "row") sheet.meta.hiddenRows = [...hidden];
      else sheet.meta.hiddenColumns = [...hidden];
      break;
    }
    case "merge": {
      model.rebuildAfterTransaction();
      model.validateArea(sheet, command.range);
      model.touchMeta(sheet);
      if (command.unmerge)
        sheet.meta.merges = sheet.meta.merges.filter(
          (m) => !intersects(m, command.range),
        );
      else {
        if (sheet.meta.merges.some((m) => intersects(m, command.range)))
          throw new Error("Overlapping merged cells");
        model.ensureWritable(sheet, command.range);
        for (const [key, cell] of sheet.cells)
          if (
            contains(command.range, rowOf(key), columnOf(key)) &&
            key !== keyOf(command.range.r1, command.range.c1) &&
            (cell.formula || (cell.value !== undefined && cell.value !== null))
          )
            throw new Error(
              "Merge would discard nonempty cells; clear them explicitly first",
            );
        sheet.meta.merges.push({ ...command.range });
      }
      break;
    }
    case "freeze": {
      if (
        ![command.rows, command.columns].every(Number.isInteger) ||
        command.rows < 0 ||
        command.columns < 0 ||
        command.rows >= sheet.meta.rowCount ||
        command.columns >= sheet.meta.columnCount
      )
        throw new Error("Invalid frozen pane");
      model.touchMeta(sheet);
      sheet.meta.frozenRows = command.rows;
      sheet.meta.frozenColumns = command.columns;
      break;
    }
    case "filter": {
      model.touchMeta(sheet);
      if (command.range) {
        model.validateArea(sheet, command.range);
        if (
          command.rules.some(
            (r) => r.column < command.range!.c1 || r.column > command.range!.c2,
          )
        )
          throw new Error("Filter column outside range");
        const perColumn = new Map<number, number>();
        for (const rule of command.rules) {
          if (
            rule.operator === "in" &&
            (!Array.isArray(rule.values) ||
              rule.values.some((value) => typeof value !== "string"))
          )
            throw new Error("Invalid filter values");
          if (
            rule.operator === "in" &&
            command.rules.some(
              (other) => other !== rule && other.column === rule.column,
            )
          )
            throw new Error(
              "Value lists cannot be combined with conditions on the same column",
            );
        }
        for (const rule of command.rules) {
          const count = (perColumn.get(rule.column) ?? 0) + 1;
          if (count > 2)
            throw new Error(
              "A filter supports at most two conditions per column",
            );
          perColumn.set(rule.column, count);
        }
        sheet.meta.filter = {
          range: command.range,
          rules: structuredClone(command.rules),
        };
      } else delete sheet.meta.filter;
      break;
    }
    case "replace": {
      if (!command.search) throw new Error("Search text is empty");
      const escape = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
        regex = new RegExp(
          (command.entireCell ? "^" : "") +
            escape(command.search) +
            (command.entireCell ? "$" : ""),
          command.matchCase ? "g" : "gi",
        );
      for (const [key, cell] of sheet.cells) {
        if (cell.formula || typeof cell.value !== "string") continue;
        const value = cell.value.replace(regex, () => command.replacement);
        if (value !== cell.value) {
          checkProtection(model, sheet, {
            r1: rowOf(key),
            r2: rowOf(key),
            c1: columnOf(key),
            c2: columnOf(key),
          });
          model.setCell(sheet, key, { ...cell, value });
        }
      }
      break;
    }
    default:
      executeStructural(model, sheet, command);
  }
}

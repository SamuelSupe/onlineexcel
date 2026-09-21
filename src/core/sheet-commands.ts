import type { WorkbookModel } from "./model";
import type { Command } from "./types";
import {
  mapReferences,
  parseFormula,
  printFormula,
  type AST,
} from "../formula/parser";
import { subtractRange, validateRule } from "./rules";

export function sheetCommand(model: WorkbookModel, command: Command): boolean {
  if (command.type === "renameName" || command.type === "deleteName") {
    const name = command.name.toUpperCase();
    if (!model.names[name]) throw new Error("Named range not found");
    const next =
      command.type === "renameName" ? command.newName.toUpperCase() : undefined;
    if (
      next &&
      (!/^[A-Z_][A-Z0-9_.]*$/.test(next) ||
        /^[A-Z]{1,3}\d+$/.test(next) ||
        (next !== name && model.names[next]))
    )
      throw new Error("Invalid or duplicate named range");
    function visit(ast: AST): AST {
      if (ast.type === "name" && ast.name.toUpperCase() === name)
        return next
          ? { ...ast, name: next }
          : { type: "literal", value: { error: "#REF!" } };
      if (ast.type === "call") return { ...ast, args: ast.args.map(visit) };
      if (ast.type === "binary")
        return { ...ast, left: visit(ast.left), right: visit(ast.right) };
      if (ast.type === "unary") return { ...ast, value: visit(ast.value) };
      if (ast.type === "array")
        return { ...ast, rows: ast.rows.map((row) => row.map(visit)) };
      return ast;
    }
    for (const sheet of model.sheets)
      for (const [key, cell] of sheet.cells)
        if (cell.formula) {
          try {
            model.setCell(sheet, key, {
              ...cell,
              formula: printFormula(visit(parseFormula(cell.formula))),
            });
          } catch {
            /* Unparsed formulas remain available in diagnostics. */
          }
        }
    if (next) model.names[next] = model.names[name];
    if (next !== name) delete model.names[name];
    model.rebuildAfterTransaction();
    return true;
  }
  if (
    ![
      "duplicateSheet",
      "sheetVisibility",
      "protect",
      "validation",
      "conditionalFormat",
    ].includes(command.type)
  )
    return false;
  if (!("sheetId" in command)) return false;
  const sheet = model.sheet(command.sheetId);
  switch (command.type) {
    case "duplicateSheet": {
      if (sheet.meta.objects)
        throw new Error(
          "Sheets with preserved file objects cannot be duplicated yet",
        );
      const target = model.newSheet(
        command.name,
        sheet.meta.rowCount,
        sheet.meta.columnCount,
        command.id,
      );
      target.meta = {
        ...structuredClone(sheet.meta),
        id: command.id,
        name: command.name,
        hidden: false,
      };
      target.cells = new Map(
        [...sheet.cells].map(([key, cell]) => {
          let formula = cell.formula;
          if (formula)
            try {
              formula = printFormula(
                mapReferences(parseFormula(formula), (ref) =>
                  ref.sheet?.toUpperCase() === sheet.meta.name.toUpperCase()
                    ? { ...ref, sheet: command.name }
                    : ref,
                ),
              );
            } catch {}
          return [key, { ...structuredClone(cell), formula }];
        }),
      );
      model.sheets.splice(model.sheets.indexOf(sheet) + 1, 0, target);
      model.rebuildAfterTransaction();
      return true;
    }
    case "sheetVisibility":
      if (
        command.hidden &&
        !model.sheets.some((s) => s !== sheet && !s.meta.hidden)
      )
        throw new Error("Cannot hide the last visible sheet");
      model.touchMeta(sheet);
      sheet.meta.hidden = command.hidden;
      return true;
    case "protect":
      model.touchMeta(sheet);
      sheet.meta.protected = command.enabled;
      return true;
    case "validation":
    case "conditionalFormat": {
      if (sheet.meta.protected)
        throw new Error("Unprotect the sheet before changing its rules");
      model.validateArea(sheet, command.range);
      model.touchMeta(sheet);
      if (command.type === "validation") {
        if (command.rule) validateRule(command.rule);
        sheet.meta.validations = (sheet.meta.validations ?? []).flatMap(
          (rule) =>
            subtractRange(rule.range, command.range).map((range) => ({
              ...rule,
              range,
            })),
        );
        if (command.rule)
          sheet.meta.validations.push({
            ...structuredClone(command.rule),
            range: { ...command.range },
          });
      } else {
        if (command.rule) {
          if (
            ![
              "eq",
              "neq",
              "gt",
              "gte",
              "lt",
              "lte",
              "between",
              "contains",
            ].includes(command.rule.operator)
          )
            throw new Error("Invalid conditional format");
          if (
            command.rule.operator === "between" &&
            !Number.isFinite(command.rule.second)
          )
            throw new Error("Invalid conditional format");
          model.styleId(command.rule.style);
        }
        sheet.meta.conditionalFormats = (
          sheet.meta.conditionalFormats ?? []
        ).flatMap((rule) =>
          subtractRange(rule.range, command.range).map((range) => ({
            ...rule,
            range,
          })),
        );
        if (command.rule)
          sheet.meta.conditionalFormats.push({
            ...structuredClone(command.rule),
            range: { ...command.range },
          });
      }
      return true;
    }
  }
  return false;
}

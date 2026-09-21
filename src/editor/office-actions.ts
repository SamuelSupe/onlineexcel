import { parseInputValue } from "../core/input";
import { formatValue } from "../core/format";
import type { ActionContext } from "./actions";
import { contains, parseRange, rangeName } from "../core/address";
import type { ConditionalRule, ValidationRule } from "../core/types";
export async function officeAction(
  ctx: ActionContext,
  name: string,
): Promise<boolean> {
  const {
    workbook,
    sheetId,
    sheet,
    sheets,
    selection,
    ask,
    t,
    activate,
    select,
    notify,
  } = ctx;
  if (name === "duplicateSheet") {
    const result = await ask(t.duplicateSheet, [
      { key: "name", label: t.name, value: sheet.name.slice(0, 20) + " (2)" },
    ]);
    if (result)
      await activate(await workbook.duplicateSheet(sheetId, result.name));
  } else if (name === "hideSheet") {
    await workbook.setSheetHidden(sheetId, true);
    await activate(sheets.find((s) => s.id !== sheetId && !s.hidden)!.id);
  } else if (name === "showSheet") {
    const hidden = sheets.filter((s) => s.hidden);
    if (!hidden.length) notify(t.noHiddenSheets);
    else {
      const result = await ask(t.showSheet, [
        {
          key: "id",
          label: t.showSheet,
          options: hidden.map((s) => ({ value: s.id, label: s.name })),
        },
      ]);
      if (result) {
        await workbook.setSheetHidden(result.id, false);
        await activate(result.id);
      }
    }
  } else if (name === "lockCells" || name === "unlockCells")
    await workbook.setStyle(sheetId, selection, {
      locked: name === "lockCells",
    });
  else if (name === "protectSheet" || name === "unprotectSheet") {
    if (await ask(t[name], [], t.protectionHelp))
      await workbook.protectSheet(sheetId, name === "protectSheet");
  } else if (name === "nameManager") {
    const metadata = await workbook.getMetadata(),
      names = Object.keys(metadata.names);
    if (!names.length) {
      notify(t.noNames);
      return true;
    }
    const chosen = await ask(t.nameManager, [
      {
        key: "name",
        label: t.name,
        options: names.map((name) => ({
          value: name,
          label: `${name} — ${sheets.find((s) => s.id === metadata.names[name].sheetId)?.name}!${rangeName(metadata.names[name].range)}`,
        })),
      },
      {
        key: "operation",
        label: t.operation,
        options: [
          { value: "edit", label: t.editName },
          { value: "delete", label: t.deleteName },
          { value: "goto", label: t.goToName },
        ],
      },
    ]);
    if (!chosen) return true;
    const entry = metadata.names[chosen.name];
    if (chosen.operation === "delete") await workbook.deleteName(chosen.name);
    else if (chosen.operation === "goto") {
      if (sheets.find((s) => s.id === entry.sheetId)?.hidden)
        await workbook.setSheetHidden(entry.sheetId, false);
      await select(entry.sheetId, entry.range);
    } else {
      const edited = await ask(t.editName, [
        { key: "name", label: t.name, value: chosen.name },
        {
          key: "sheet",
          label: t.sheet,
          options: sheets.map((s) => ({ value: s.id, label: s.name })),
          value: entry.sheetId,
        },
        { key: "range", label: t.range, value: rangeName(entry.range) },
      ]);
      if (edited)
        await workbook.transaction([
          { type: "renameName", name: chosen.name, newName: edited.name },
          {
            type: "defineName",
            name: edited.name,
            sheetId: edited.sheet,
            range: parseRange(edited.range),
          },
        ]);
    }
  } else if (name === "validation") {
    const metadata = await workbook.getMetadata();
    const old = metadata.sheets
      .find((item) => item.id === sheetId)
      ?.validations?.find((rule) =>
        contains(rule.range, selection.r1, selection.c1),
      );
    const system = metadata.dateSystem;
    const boundText = (value: number | undefined) =>
      value === undefined
        ? ""
        : old?.type === "date"
          ? formatValue(
              value,
              {
                numberFormat: Number.isInteger(value)
                  ? "yyyy-mm-dd"
                  : "yyyy-mm-dd hh:mm:ss",
              },
              system,
            )
          : String(value);
    const minimum = boundText(old?.minimum),
      maximum = boundText(old?.maximum);
    const result = await ask(t.validation, [
      {
        key: "type",
        label: t.validationType,
        value: old?.type ?? "list",
        options: [
          { value: "list", label: t.listType },
          { value: "decimal", label: t.number },
          { value: "whole", label: t.wholeType },
          { value: "date", label: t.date },
          { value: "textLength", label: t.lengthType },
          { value: "none", label: t.clearRule },
        ],
      },
      {
        key: "values",
        label: t.validationList,
        value: old?.values?.join(",") ?? "",
      },
      { key: "minimum", label: t.minimum, value: minimum },
      { key: "maximum", label: t.maximum, value: maximum },
      {
        key: "blank",
        label: t.allowBlank,
        type: "checkbox",
        value: String(old?.allowBlank !== false),
      },
    ]);
    const bound = (key: "minimum" | "maximum", value: string) =>
      old?.type === result?.type &&
      value === (key === "minimum" ? minimum : maximum)
        ? old?.[key]
        : value
          ? result?.type === "date"
            ? (parseInputValue(value, "date", system).value as number)
            : Number(value)
          : undefined;
    if (result)
      await workbook.setValidation(
        sheetId,
        selection,
        result.type === "none"
          ? undefined
          : {
              type: result.type as ValidationRule["type"],
              values:
                result.type === "list"
                  ? result.values.split(",").map((s) => s.trim())
                  : undefined,
              minimum: bound("minimum", result.minimum),
              maximum: bound("maximum", result.maximum),
              allowBlank: result.blank === "true",
            },
      );
  } else if (name === "conditionalFormat") {
    const metadata = await workbook.getMetadata();
    const old = metadata.sheets
      .find((item) => item.id === sheetId)
      ?.conditionalFormats?.find((rule) =>
        contains(rule.range, selection.r1, selection.c1),
      );
    const result = await ask(t.conditionalFormat, [
      {
        key: "operator",
        label: t.byCondition,
        value: old?.operator ?? "gt",
        options: [
          ...["eq", "neq", "gt", "gte", "lt", "lte", "contains"].map(
            (operator) => ({
              value: operator,
              label: t[operator as keyof typeof t] ?? operator,
            }),
          ),
          { value: "none", label: t.clearRule },
        ],
      },
      { key: "value", label: t.value, value: String(old?.value ?? 0) },
      {
        key: "color",
        label: t.fill,
        type: "color",
        value: old?.style.background ?? "#ffdddd",
      },
    ]);
    if (result)
      await workbook.setConditionalFormat(
        sheetId,
        selection,
        result.operator === "none"
          ? undefined
          : {
              operator: result.operator as ConditionalRule["operator"],
              value: ctx.parseInput(result.value),
              style: { background: result.color },
            },
      );
  } else return false;
  return true;
}

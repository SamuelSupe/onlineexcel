import type {
  ConditionalRule,
  Diagnostic,
  SheetSnapshot,
  ValidationRule,
} from "../core/types";
import { parseRange, rangeName } from "../core/address";
import { child, children, xmlEscape as x, xmlTree, type XmlNode } from "./xml";
import { readDifferential } from "./xlsx-differential";
export const serializeNode = (n: XmlNode): string =>
  `<${n.name}${Object.entries(n.attributes)
    .map(([k, v]) => ` ${k}="${x(v)}"`)
    .join(
      "",
    )}>${x(n.text)}${n.children.map(serializeNode).join("")}</${n.name}>`;
const operators: Record<string, ConditionalRule["operator"]> = {
  equal: "eq",
  notEqual: "neq",
  greaterThan: "gt",
  greaterThanOrEqual: "gte",
  lessThan: "lt",
  lessThanOrEqual: "lte",
  between: "between",
  containsText: "contains",
};
export function readRules(
  nodes: XmlNode[],
  sheet: SheetSnapshot,
  styleBytes: Uint8Array,
  diagnostics: Diagnostic[],
) {
  const dxfs = children(child(xmlTree(styleBytes), "dxfs"), "dxf");
  const report = (message: string) =>
    diagnostics.push({
      code: "RULE_VARIANT",
      severity: "warning",
      lossy: true,
      message,
      sheetId: sheet.id,
    });
  for (const node of nodes) {
    if (node.name === "sheetProtection") {
      sheet.protected = node.attributes.sheet !== "0";
      if (node.attributes.password || node.attributes.hashValue)
        report(
          "Password protection is retained as edit protection without a password.",
        );
      continue;
    }
    if (node.name === "dataValidations")
      for (const validation of children(node, "dataValidation")) {
        const type = validation.attributes.type as ValidationRule["type"],
          f1 = child(validation, "formula1")?.text ?? "",
          f2 = child(validation, "formula2")?.text;
        const op = validation.attributes.operator ?? "between";
        if (
          !["list", "decimal", "whole", "date", "textLength"].includes(type) ||
          (type === "list" && !/^".*"$/s.test(f1)) ||
          (type !== "list" &&
            (!Number.isFinite(Number(f1)) ||
              ![
                "between",
                "greaterThanOrEqual",
                "lessThanOrEqual",
                "equal",
              ].includes(op)))
        ) {
          report("Unsupported data validation formula or operator.");
          continue;
        }
        const rule: Omit<ValidationRule, "range"> = {
          type,
          allowBlank: validation.attributes.allowBlank !== "0",
        };
        if (type === "list")
          rule.values = f1.slice(1, -1).replace(/""/g, '"').split(",");
        else {
          if (op !== "lessThanOrEqual") rule.minimum = Number(f1);
          if (op !== "greaterThanOrEqual")
            rule.maximum = Number(op === "between" ? f2 : f1);
        }
        for (const ref of (validation.attributes.sqref ?? "")
          .split(/\s+/)
          .filter(Boolean))
          (sheet.validations ??= []).push({ ...rule, range: parseRange(ref) });
      }
    if (node.name === "conditionalFormatting")
      for (const cf of children(node, "cfRule")) {
        const operator = operators[cf.attributes.operator],
          formulas = children(cf, "formula"),
          dxf = dxfs[Number(cf.attributes.dxfId)];
        if (
          !operator ||
          (cf.attributes.type !== "cellIs" &&
            cf.attributes.type !== "containsText") ||
          !dxf
        ) {
          report("Unsupported conditional formatting rule.");
          continue;
        }
        const style = readDifferential(dxf, report);
        const raw = formulas[0]?.text ?? "";
        const value =
          operator === "contains"
            ? (cf.attributes.text ?? "")
            : /^".*"$/s.test(raw)
              ? raw.slice(1, -1).replace(/""/g, '"')
              : Number(raw);
        if (typeof value === "number" && !Number.isFinite(value)) {
          report("Conditional formulas with references are not evaluated.");
          continue;
        }
        for (const ref of (node.attributes.sqref ?? "")
          .split(/\s+/)
          .filter(Boolean))
          (sheet.conditionalFormats ??= []).push({
            range: parseRange(ref),
            operator,
            value,
            second: formulas[1] ? Number(formulas[1].text) : undefined,
            style,
          });
      }
  }
}
export function writeValidations(
  sheet: SheetSnapshot | { validations?: ValidationRule[] },
) {
  const rules = sheet.validations ?? [];
  if (!rules.length) return "";
  return `<dataValidations count="${rules.length}">${rules
    .map((rule) => {
      const operator =
        rule.minimum !== undefined && rule.maximum !== undefined
          ? "between"
          : rule.minimum !== undefined
            ? "greaterThanOrEqual"
            : "lessThanOrEqual";
      const first =
        rule.type === "list"
          ? '"' + rule.values!.join(",").replace(/"/g, '""') + '"'
          : String(rule.minimum ?? rule.maximum ?? 0);
      return `<dataValidation type="${rule.type}" operator="${operator}" allowBlank="${rule.allowBlank !== false ? 1 : 0}" showErrorMessage="1" errorStyle="stop" sqref="${rangeName(rule.range)}"><formula1>${x(first)}</formula1>${operator === "between" && rule.type !== "list" ? `<formula2>${rule.maximum}</formula2>` : ""}</dataValidation>`;
    })
    .join("")}</dataValidations>`;
}
export function writeConditional(rules: ConditionalRule[], offset: number) {
  return rules
    .map((rule, i) => {
      const op = Object.entries(operators).find(
        ([, value]) => value === rule.operator,
      )![0];
      const formula =
        typeof rule.value === "string"
          ? '"' + rule.value.replace(/"/g, '""') + '"'
          : String(rule.value ?? 0);
      return `<conditionalFormatting sqref="${rangeName(rule.range)}"><cfRule type="${rule.operator === "contains" ? "containsText" : "cellIs"}" operator="${op}" priority="${offset + i + 1}" dxfId="${offset + i}"${rule.operator === "contains" ? ` text="${x(rule.value)}"` : ""}><formula>${x(rule.operator === "contains" ? `NOT(ISERROR(SEARCH(${formula},${rangeName(rule.range).split(":")[0]})))` : formula)}</formula>${rule.second === undefined ? "" : `<formula>${rule.second}</formula>`}</cfRule></conditionalFormatting>`;
    })
    .join("");
}

import type { Workbook } from "../runtime/client";
import type { Diagnostic } from "../core/types";
import { parseRange } from "../core/address";
import { element, formDialog, labeledControl } from "./dom";
import { icon } from "./icons";
import { labels, type Locale } from "./locale";

export async function showDiagnostics(
  parent: HTMLElement,
  workbook: Workbook,
  locale: Locale,
  signal: AbortSignal,
  navigate: (sheetId: string, range: string) => Promise<void>,
) {
  const t = labels(locale);
  const [metadata, issues] = await Promise.all([
    workbook.getMetadata(),
    workbook.getDiagnostics(),
  ]);
  let selected: Diagnostic | undefined;
  const result = await formDialog(
    parent,
    t.diagnostics,
    t,
    { signal, confirm: t.close },
    (form) => {
      form.classList.add("diagnostics-dialog");
      const filters = element("div", "issue-filters");
      const level = element("select"),
        sheet = element("select");
      for (const [value, name] of [
        ["", t.allSeverities],
        ["error", t.errors],
        ["warning", t.warnings],
      ]) {
        const option = element("option", "", name);
        option.value = value;
        level.append(option);
      }
      for (const [value, name] of [
        ["", t.allSheets],
        ["__global", t.globalScope],
        ...metadata.sheets.map((s) => [s.id, s.name]),
      ]) {
        const option = element("option", "", name);
        option.value = value;
        sheet.append(option);
      }
      labeledControl(filters, t.allSeverities, level);
      labeledControl(filters, t.allSheets, sheet);
      const summary = element("p"),
        list = element("div", "issue-list"),
        more = element("button", "", t.showMore);
      more.type = "button";
      form.append(filters, summary, list, more);
      let limit = 100;
      const render = () => {
        const filtered = issues.filter(
          (issue) =>
            (!level.value || issue.severity === level.value) &&
            (!sheet.value ||
              (sheet.value === "__global"
                ? !issue.sheetId
                : issue.sheetId === sheet.value)),
        );
        summary.textContent = `${t.errors}: ${filtered.filter((i) => i.severity === "error").length} · ${t.warnings}: ${filtered.filter((i) => i.severity === "warning").length}`;
        list.replaceChildren();
        const groups = new Map<string, HTMLElement>();
        for (const issue of filtered.slice(0, limit)) {
          const meta = metadata.sheets.find((s) => s.id === issue.sheetId);
          const groupId = issue.sheetId ?? "__global";
          if (!groups.has(groupId)) {
            const group = element("section");
            group.append(
              element("h3", "", meta?.name ?? issue.sheetId ?? t.globalScope),
            );
            groups.set(groupId, group);
            list.append(group);
          }
          const card = element("article", "issue-card");
          const formula = [
            "INVALID_FORMULA",
            "UNSUPPORTED_FORMULA",
            "CIRCULAR_REFERENCE",
            "FORMULA_ERROR",
            "SPILL_CONFLICT",
          ].includes(issue.code);
          card.append(
            element(
              "strong",
              "",
              `${issue.severity === "error" ? t.errors : t.warnings}${issue.range ? " · " + issue.range : ""}`,
            ),
          );
          const descriptions: Record<string, string> = {
            INVALID_FORMULA: t.invalidFormulaIssue,
            UNSUPPORTED_FORMULA: t.unknownFormulaIssue,
            CIRCULAR_REFERENCE: t.cyclicFormulaIssue,
            SPILL_CONFLICT: t.spillIssue,
            FORMULA_ERROR: `${t.formulaErrorIssue}: ${issue.message}`,
          };
          const description = descriptions[issue.code];
          if (description) card.append(element("p", "", description));
          card.append(
            element(
              "p",
              "issue-impact",
              issue.lossy
                ? t.exportLoss
                : formula
                  ? t.calculationImpact
                  : t.reviewImpact,
            ),
          );
          card.append(
            element(
              "p",
              "",
              issue.code === "SPILL_CONFLICT"
                ? t.spillFix
                : formula
                  ? t.formulaFix
                  : issue.code === "CUSTOM_FUNCTION_REQUIRES_MODULE"
                    ? t.customFix
                    : t.compatibilityFix,
            ),
          );
          const details = element("details");
          details.append(
            element("summary", "", t.diagnosticDetails),
            element("pre", "", `${issue.code}\n${issue.message}`),
          );
          card.append(details);
          if (issue.sheetId) {
            const go = element("button", "issue-locate");
            go.type = "button";
            go.append(icon("find"), element("span", "", t.locate));
            let valid = !!meta && !meta.hidden;
            try {
              const range = parseRange(issue.range ?? "A1");
              valid &&=
                range.r2 < meta!.rowCount && range.c2 < meta!.columnCount;
            } catch {
              valid = false;
            }
            go.disabled = !valid;
            if (meta?.hidden) card.append(element("p", "", t.hiddenLocation));
            go.onclick = () => {
              selected = issue;
              form.requestSubmit();
            };
            card.append(go);
          }
          groups.get(groupId)!.append(card);
        }
        if (!filtered.length) list.append(element("p", "", t.noIssues));
        more.hidden = filtered.length <= limit;
      };
      more.onclick = () => {
        limit += 100;
        render();
      };
      level.onchange = sheet.onchange = () => {
        limit = 100;
        render();
      };
      render();
      return () => selected;
    },
  );
  if (result?.sheetId) await navigate(result.sheetId, result.range ?? "A1");
}

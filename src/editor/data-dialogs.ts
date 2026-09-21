import type { FilterRule, Rect, SortKey } from "../core/types";
import { columnName, rangeName } from "../core/address";
import type { Workbook } from "../runtime/client";
import type { Label } from "./locale";
import { element, formDialog, labeledControl } from "./dom";

interface Context {
  parent: HTMLElement;
  signal: AbortSignal;
  t: Record<Label, string>;
  workbook: Workbook;
  sheetId: string;
  selection: Rect;
}
const select = (values: [string, string][], value?: string) => {
  const input = element("select");
  for (const [key, text] of values) {
    const option = element("option", "", text);
    option.value = key;
    input.append(option);
  }
  if (value !== undefined) input.value = value;
  return input;
};
const button = (text: string, action: () => void) => {
  const input = element("button", "dialog-action", text);
  input.type = "button";
  input.onclick = action;
  return input;
};
async function columnLabels(
  context: Context,
  range: Rect,
): Promise<[string, string][]> {
  const row = await context.workbook.getValues(context.sheetId, {
    ...range,
    r2: range.r1,
  });
  return row[0].map((value, i) => [
    String(range.c1 + i),
    `${columnName(range.c1 + i)}${value === null ? "" : ` — ${typeof value === "object" ? value.error : value}`}`,
  ]);
}

export async function sortDialog(context: Context, direction: "asc" | "desc") {
  const { parent, signal, t, workbook, sheetId, selection } = context;
  const expanded = await workbook.getDataRegion(sheetId, selection);
  const columns = await columnLabels(context, expanded);
  const same = rangeName(selection) === rangeName(expanded);
  return formDialog(parent, t.sort, t, { signal }, (form) => {
    form.classList.add("sort-dialog");
    const scope = select(
      same
        ? [["selection", rangeName(selection)]]
        : [
            ["expand", `${t.expandSelection}: ${rangeName(expanded)}`],
            ["selection", `${t.keepSelection}: ${rangeName(selection)}`],
          ],
    );
    labeledControl(form, t.sortRange, scope);
    const header = element("input");
    header.type = "checkbox";
    header.checked = true;
    labeledControl(form, t.header, header);
    const levels = element("div", "dialog-levels");
    form.append(levels);
    const rows: {
      node: HTMLElement;
      column: HTMLSelectElement;
      direction: HTMLSelectElement;
    }[] = [];
    const available = () =>
      scope.value === "expand"
        ? columns
        : Array.from(
            { length: selection.c2 - selection.c1 + 1 },
            (_, i) =>
              columns.find(([c]) => Number(c) === selection.c1 + i) ??
              ([String(selection.c1 + i), columnName(selection.c1 + i)] as [
                string,
                string,
              ]),
          );
    const add = (column?: number) => {
      const node = element("fieldset", "dialog-level"),
        number = rows.length + 1;
      const col = select(
        available(),
        String(
          column ??
            available().find(
              ([c]) => !rows.some((r) => r.column.value === c),
            )?.[0] ??
            selection.c1,
        ),
      );
      if (col.selectedIndex < 0) col.selectedIndex = 0;
      const order = select(
        [
          ["asc", t.sortAsc],
          ["desc", t.sortDesc],
        ],
        direction,
      );
      labeledControl(node, `${t.sortColumn} ${number}`, col);
      labeledControl(node, `${t.direction} ${number}`, order);
      const row = { node, column: col, direction: order };
      rows.push(row);
      node.append(
        button(t.removeLevel, () => {
          if (rows.length > 1) {
            rows.splice(rows.indexOf(row), 1);
            node.remove();
          }
        }),
      );
      levels.append(node);
    };
    add(selection.c1);
    form.append(button(`＋ ${t.addLevel}`, () => add()));
    scope.onchange = () => {
      rows.length = 0;
      levels.replaceChildren();
      add(selection.c1);
    };
    return () => {
      const keys = rows.map(
        (r): SortKey => ({
          column: Number(r.column.value),
          direction: r.direction.value as "asc" | "desc",
        }),
      );
      if (new Set(keys.map((k) => k.column)).size !== keys.length)
        throw new Error(t.sortDuplicate);
      return {
        range: scope.value === "expand" ? expanded : selection,
        keys,
        header: header.checked,
      };
    };
  });
}

export async function filterDialog(
  context: Context,
  current?: { range: Rect; rules: FilterRule[] },
) {
  const { parent, signal, t, workbook, sheetId, selection } = context;
  const range =
    current?.range ?? (await workbook.getDataRegion(sheetId, selection));
  const columns = await columnLabels(context, range);
  let rules = structuredClone(current?.rules ?? []);
  return formDialog(parent, t.filter, t, { signal }, (form) => {
    form.append(
      element("p", "", `${rangeName(range)} · ${t.header}`),
      element("p", "", t.filterAnd),
    );
    const col = select(
      columns,
      String(Math.max(range.c1, Math.min(range.c2, selection.c1))),
    );
    labeledControl(form, t.filterColumn, col);
    const mode = select([
      ["all", t.allValues],
      ["values", t.byValues],
      ["conditions", t.byCondition],
    ]);
    labeledControl(form, t.filterMode, mode);
    const conditions = element("div"),
      valuesBox = element("div"),
      info = element("p");
    const operators = [
      "eq",
      "neq",
      "contains",
      "gt",
      "lt",
      "gte",
      "lte",
    ] as const;
    const inputs = Array.from({ length: 2 }, (_, i) => {
      const op = select([
        ["", t.noCondition],
        ...operators.map((value): [string, string] => [value, t[value]]),
      ]);
      const value = element("input");
      labeledControl(conditions, `${t.operator} ${i + 1}`, op);
      labeledControl(conditions, `${t.value} ${i + 1}`, value);
      return { op, value };
    });
    const search = element("input");
    labeledControl(valuesBox, t.valuesSearch, search);
    const list = element("div", "filter-values");
    const chosen = new Set<string>();
    let loadError: Error | undefined;
    let candidates: string[] = [],
      truncated = false,
      pending = Promise.resolve(),
      generation = 0,
      column = Number(col.value);
    const render = () => {
      list.replaceChildren();
      for (const value of candidates.filter((v) =>
        v.toLowerCase().includes(search.value.toLowerCase()),
      )) {
        const label = element("label", "filter-value"),
          check = element("input");
        check.type = "checkbox";
        check.checked = chosen.has(value);
        check.onchange = () => {
          if (check.checked) chosen.add(value);
          else chosen.delete(value);
        };
        label.append(check, document.createTextNode(value || t.blank));
        list.append(label);
      }
    };
    valuesBox.append(
      button(t.selectAll, () => {
        candidates
          .filter((v) => v.toLowerCase().includes(search.value.toLowerCase()))
          .forEach((v) => chosen.add(v));
        render();
      }),
      button(t.selectNone, () => {
        candidates
          .filter((v) => v.toLowerCase().includes(search.value.toLowerCase()))
          .forEach((v) => chosen.delete(v));
        render();
      }),
      info,
      list,
    );
    search.oninput = render;
    form.append(conditions, valuesBox);
    const visibility = () => {
      conditions.hidden = mode.value !== "conditions";
      valuesBox.hidden = mode.value !== "values";
    };
    const save = () => {
      if (mode.value === "values" && truncated)
        throw new Error(t.tooManyValues);
      rules = rules.filter((r) => r.column !== column);
      if (
        mode.value === "values" &&
        (chosen.size !== candidates.length ||
          candidates.some((value) => !chosen.has(value)))
      )
        rules.push({ column, operator: "in", values: [...chosen] });
      if (mode.value === "conditions")
        for (const { op, value } of inputs)
          if (op.value) {
            const text = value.value;
            rules.push({
              column,
              operator: op.value as Exclude<FilterRule["operator"], "in">,
              value: text.startsWith("'")
                ? text.slice(1)
                : text !== "" && Number.isFinite(Number(text))
                  ? Number(text)
                  : /^(true|false)$/i.test(text)
                    ? text.toLowerCase() === "true"
                    : text,
            });
          }
    };
    const load = async () => {
      loadError = undefined;
      const request = ++generation;
      col.disabled = true;
      column = Number(col.value);
      const existing = rules.filter((r) => r.column === column),
        chosenRule = existing.find((r) => r.operator === "in");
      mode.value = chosenRule
        ? "values"
        : existing.length
          ? "conditions"
          : "values";
      inputs.forEach(({ op, value }, i) => {
        const rule = existing[i];
        op.value = rule && rule.operator !== "in" ? rule.operator : "";
        value.value =
          rule && rule.operator !== "in" ? String(rule.value ?? "") : "";
      });
      chosen.clear();
      search.value = "";
      candidates = [];
      list.replaceChildren();
      info.textContent = t.loading;
      visibility();
      const result = await workbook.getDistinctValues(sheetId, range, column);
      if (request !== generation || signal.aborted) return;
      truncated = result.truncated;
      candidates = [
        ...new Set([
          ...result.values,
          ...(chosenRule?.operator === "in" ? chosenRule.values : []),
        ]),
      ];
      (chosenRule?.operator === "in" ? chosenRule.values : candidates).forEach(
        (v) => chosen.add(v),
      );
      info.textContent = truncated ? t.tooManyValues : "";
      render();
      col.disabled = false;
    };
    const loadSafely = () =>
      load().catch((error: Error) => {
        loadError = error;
        info.textContent = error.message;
        col.disabled = false;
      });
    pending = loadSafely();
    col.onchange = () => {
      try {
        save();
        pending = loadSafely();
      } catch (error) {
        col.value = String(column);
        info.textContent = (error as Error).message;
      }
    };
    mode.onchange = visibility;
    return async () => {
      await pending;
      if (loadError) throw loadError;
      save();
      return { range, rules };
    };
  });
}

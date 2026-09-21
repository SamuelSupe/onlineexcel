import type { EditorAction } from "./index";
import type { CellStyle } from "../core/types";
import { element } from "./dom";
import type { Label } from "./locale";
import { icon, type IconName } from "./icons";
export function createToolbar(
  labels: Record<Label, string>,
  action: (name: string, value?: string) => void,
  readOnly: boolean,
  signal: AbortSignal,
  config: { items?: readonly string[]; actions?: readonly EditorAction[] } = {},
): {
  menu: HTMLElement;
  ribbon: HTMLElement;
  setTab(tab: string): void;
  update(styles: CellStyle[], blanks?: boolean): void;
} {
  const menu = element("div", "menubar"),
    ribbon = element("div", "toolbar");
  let currentStyles: CellStyle[] = [{}];
  let includesBlanks = false;
  const tabs = ["home", "insert", "data", "view"] as const;
  function button(parent: HTMLElement, name: IconName, display?: string) {
    const b = element("button", "tool-button");
    b.append(
      icon(name),
      element("span", "tool-text", display ?? labels[name as Label] ?? name),
    );
    b.type = "button";
    b.title = labels[name as Label] ?? name;
    b.setAttribute("aria-label", b.title);
    b.dataset.action = name;
    b.disabled =
      readOnly &&
      ![
        "saveXlsx",
        "saveCsv",
        "saveJson",
        "find",
        "findNext",
        "copy",
        "diagnostics",
        "shortcuts",
        "saveNow",
      ].includes(name);
    b.addEventListener("click", () => action(name), { signal });
    parent.append(b);
    return b;
  }
  function group(): HTMLElement {
    const g = element("div", "group");
    ribbon.append(g);
    return g;
  }
  function select(
    parent: HTMLElement,
    name: "font" | "fontSize" | "format" | "verticalAlign",
    values: [string, string][],
    selected?: string,
  ) {
    const input = element("select");
    input.dataset.style = name;
    input.title = labels[name as Label];
    input.setAttribute("aria-label", labels[name as Label]);
    input.disabled = readOnly;
    for (const [value, label] of values) {
      const option = element("option", "", label);
      option.value = value;
      input.append(option);
    }
    if (selected) input.value = selected;
    input.onchange = () => action(name, input.value);
    const control = element("label", "tool-control");
    control.append(
      icon(name),
      element("span", "tool-text", labels[name]),
      input,
    );
    parent.append(control);
  }
  function setTab(tab: string) {
    menu
      .querySelectorAll("[data-tab]")
      .forEach((b) =>
        b.classList.toggle("selected", (b as HTMLElement).dataset.tab === tab),
      );
    ribbon.replaceChildren();
    if (tab === "home") {
      let g = group();
      button(g, "undo");
      button(g, "redo");
      button(g, "copy");
      button(g, "formatPainter");
      button(g, "pasteSpecial");
      g = group();
      select(g, "font", [
        ["Arial", "Arial"],
        ["Calibri", "Calibri"],
        ["Microsoft YaHei", "Microsoft YaHei"],
        ["Georgia", "Georgia"],
      ]);
      select(
        g,
        "fontSize",
        ["9", "10", "11", "12", "14", "16", "18", "24", "32"].map((n) => [
          n,
          n,
        ]),
        "11",
      );
      g = group();
      button(g, "bold");
      button(g, "italic");
      button(g, "underline");
      g = group();
      for (const name of ["color", "fill"] as const) {
        const input = element("input");
        input.type = "color";
        input.dataset.style = name;
        input.value = name === "color" ? "#20342b" : "#e9f4ee";
        input.title = labels[name];
        input.setAttribute("aria-label", labels[name]);
        input.disabled = readOnly;
        input.oninput = () => action(name, input.value);
        const control = element("label", "tool-control");
        control.append(
          icon(name),
          element("span", "tool-text", labels[name]),
          input,
        );
        g.append(control);
      }
      g = group();
      button(g, "left");
      button(g, "center");
      button(g, "right");
      button(g, "wrap");
      button(g, "border");
      button(g, "borderOptions");
      select(
        g,
        "verticalAlign",
        [
          ["top", labels.top],
          ["center", labels.center],
          ["bottom", labels.bottom],
        ],
        "center",
      );
      g = group();
      select(g, "format", [
        ["General", labels.general],
        ["#,##0.00", labels.number],
        ['"¥"#,##0.00', labels.currency],
        ["0.00%", labels.percent],
        ["yyyy-mm-dd", labels.date],
        ["@", labels.text],
      ]);
      button(g, "customFormat");
      button(g, "merge");
      button(g, "unmerge");
      g = group();
      button(g, "find");
      button(g, "findNext");
    } else if (tab === "insert") {
      let g = group();
      button(g, "addRow");
      button(g, "addColumn");
      button(g, "deleteRow");
      button(g, "deleteColumn");
      g = group();
      button(g, "addSheet");
      button(g, "rename");
      button(g, "deleteSheet");
      g = group();
      button(g, "namedRange");
      button(g, "nameManager");
      g = group();
      button(g, "duplicateSheet");
      button(g, "hideSheet");
      button(g, "showSheet");
    } else if (tab === "data") {
      let g = group();
      button(g, "sort");
      button(g, "sortAsc");
      button(g, "sortDesc");
      button(g, "filter");
      button(g, "clearFilter");
      g = group();
      button(g, "validation");
      button(g, "conditionalFormat");
      g = group();
      button(g, "find");
      button(g, "clear");
      g = group();
      button(g, "saveCsv");
      button(g, "saveJson");
    } else {
      let g = group();
      button(g, "freeze");
      button(g, "unfreeze");
      g = group();
      button(g, "rowHeight");
      button(g, "columnWidth");
      button(g, "autoFitRows");
      button(g, "autoFitColumns");
      g = group();
      button(g, "hideRows");
      button(g, "hideColumns");
      button(g, "showRows");
      button(g, "showColumns");
      g = group();
      button(g, "unlockCells");
      button(g, "lockCells");
      button(g, "protectSheet");
      button(g, "unprotectSheet");
      g = group();
      button(g, "diagnostics");
      button(g, "shortcuts");
    }
    filterItems(ribbon);
    update(currentStyles, includesBlanks);
  }
  function update(styles: CellStyle[], blanks = false) {
    currentStyles = styles.length ? styles : [{}];
    includesBlanks = blanks;
    const all = blanks ? [...currentStyles, {}] : currentStyles;
    const defaults: Record<string, unknown> = {
      fontFamily: "Arial",
      fontSize: 11,
      numberFormat: "General",
      verticalAlign: "center",
      color: "#20342b",
      background: "#ffffff",
      align: "left",
    };
    const uniform = (key: keyof CellStyle) => {
      const values = all.map((style) => style[key] ?? defaults[key] ?? false);
      return values.every((value) => value === values[0])
        ? values[0]
        : undefined;
    };
    const mapping = {
      font: "fontFamily",
      fontSize: "fontSize",
      format: "numberFormat",
      verticalAlign: "verticalAlign",
      color: "color",
      fill: "background",
    } as const;
    for (const [name, key] of Object.entries(mapping)) {
      const control = ribbon.querySelector<
        HTMLSelectElement | HTMLInputElement
      >(`[data-style="${name}"]`);
      if (!control) continue;
      const value = uniform(key);
      if (control instanceof HTMLSelectElement) {
        control
          .querySelectorAll("[data-dynamic]")
          .forEach((option) => option.remove());
        const text = value === undefined ? "—" : String(value);
        if (![...control.options].some((option) => option.value === text)) {
          const option = element("option", "", text);
          option.value = text;
          option.dataset.dynamic = "true";
          control.append(option);
        }
        control.value = text;
      } else {
        control.value = typeof value === "string" ? value : "#000000";
        control.title =
          value === undefined
            ? labels[name as Label] + " —"
            : labels[name as Label];
      }
    }
    for (const name of [
      "bold",
      "italic",
      "underline",
      "wrap",
      "left",
      "center",
      "right",
    ] as const) {
      const button = ribbon.querySelector(`[data-action="${name}"]`);
      const value = uniform(
        ["left", "center", "right"].includes(name)
          ? "align"
          : (name as keyof CellStyle),
      );
      const active = typeof value === "string" ? value === name : value;
      button?.setAttribute(
        "aria-pressed",
        value === undefined ? "mixed" : String(!!active),
      );
      button?.classList.toggle("selected", !!active);
    }
  }
  for (const tab of tabs) {
    const b = element("button", "tool-button");
    b.append(icon(tab), element("span", "tool-text", labels[tab]));
    b.dataset.tab = tab;
    b.type = "button";
    b.onclick = () => setTab(tab);
    menu.append(b);
  }
  function filterItems(parent: HTMLElement) {
    if (!config.items) return;
    parent
      .querySelectorAll<HTMLElement>("[data-action],[data-style]")
      .forEach((control) => {
        const name = control.dataset.action ?? control.dataset.style!;
        if (!config.items!.includes(name))
          (control.closest(".tool-control") ?? control).remove();
      });
    parent.querySelectorAll(".group").forEach((group) => {
      if (!group.children.length) group.remove();
    });
  }
  const end = element("div", "end");
  for (const item of config.actions ?? []) {
    const b = element("button", "tool-button");
    b.type = "button";
    b.append(icon(item.icon), element("span", "tool-text", item.label));
    b.disabled = readOnly && !item.allowReadOnly;
    b.addEventListener("click", () => action("custom:" + item.id), { signal });
    end.append(b);
  }
  button(end, "open");
  button(end, "saveXlsx", labels.save).classList.add("primary");
  filterItems(end);
  menu.append(end);
  setTab("home");
  return { menu, ribbon, setTab, update };
}

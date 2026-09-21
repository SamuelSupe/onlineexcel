import { parseCsv } from "../io/csv";
import { parseInputValue, type ColumnType } from "../core/input";
import { columnName } from "../core/address";
import { element, formDialog, labeledControl } from "./dom";
import type { Label } from "./locale";
export async function csvImportDialog(
  parent: HTMLElement,
  file: File,
  t: Record<Label, string>,
  signal: AbortSignal,
) {
  const bytes = await file.arrayBuffer();
  return formDialog(parent, t.importPreview, t, { signal }, (form) => {
    const encoding = element("select"),
      delimiter = element("select"),
      header = element("input"),
      preview = element("div", "import-preview");
    for (const value of [
      "utf-8",
      "utf-16le",
      "gb18030",
      "big5",
      "shift_jis",
      "euc-kr",
    ]) {
      const option = element("option", "", value);
      option.value = value;
      encoding.append(option);
    }
    for (const [value, label] of [
      [",", ","],
      ["\t", "Tab"],
      [";", ";"],
      ["|", "|"],
    ]) {
      const option = element("option", "", label);
      option.value = value;
      delimiter.append(option);
    }
    header.type = "checkbox";
    header.checked = true;
    labeledControl(form, t.encoding, encoding);
    labeledControl(form, t.delimiter, delimiter);
    labeledControl(form, t.header, header);
    form.append(preview);
    let types: HTMLSelectElement[] = [],
      text = "";
    const render = () => {
      preview.replaceChildren();
      try {
        text = new TextDecoder(encoding.value, { fatal: true }).decode(bytes);
        const rows = parseCsv(text, delimiter.value);
        const table = element("table"),
          headings = element("tr");
        types = (rows[0] ?? []).map((_, col) => {
          const th = element("th"),
            select = element("select");
          select.setAttribute(
            "aria-label",
            `${t.columnType} ${columnName(col)}`,
          );
          for (const type of [
            "auto",
            "text",
            "number",
            "date",
            "boolean",
            "percent",
          ] as const) {
            const option = element(
              "option",
              "",
              t[
                type === "auto"
                  ? "autoType"
                  : type === "boolean"
                    ? "booleanType"
                    : type
              ],
            );
            option.value = type;
            select.append(option);
          }
          th.append(select);
          headings.append(th);
          return select;
        });
        table.append(headings);
        for (const row of rows.slice(0, 10)) {
          const tr = element("tr");
          for (const value of row) tr.append(element("td", "", value));
          table.append(tr);
        }
        preview.append(table);
      } catch (error) {
        preview.append(element("p", "", (error as Error).message));
        types = [];
      }
    };
    encoding.onchange = delimiter.onchange = render;
    render();
    return () => {
      const rows = parseCsv(
        new TextDecoder(encoding.value, { fatal: true }).decode(bytes),
        delimiter.value,
      );
      const columns = types.map((input) => input.value as ColumnType);
      if (rows.length && columns.length !== rows[0].length)
        throw new Error(t.importPreview);
      rows.forEach((row, index) => {
        if (!(header.checked && index === 0))
          row.forEach((value, col) => parseInputValue(value, columns[col]));
      });
      return {
        text,
        delimiter: delimiter.value,
        columns,
        header: header.checked,
      };
    };
  });
}

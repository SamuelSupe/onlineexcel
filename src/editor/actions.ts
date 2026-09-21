import {
  readClipboard,
  writeClipboard,
  type ClipboardPayload,
} from "./clipboard";
import { officeAction } from "./office-actions";
import { csvImportDialog } from "./import-dialog";
import type { Workbook } from "../runtime/client";
import type { CellStyle, SheetMeta, Rect, InputValue } from "../core/types";
import { address } from "../core/address";
import { sortDialog, filterDialog } from "./data-dialogs";
import { autoFit } from "./autofit";
import { element, download, type Field } from "./dom";
import { labels } from "./locale";
export interface FindState {
  search: string;
  matchCase: boolean;
  entireCell: boolean;
  sheetId?: string;
}
export interface ActionContext {
  dialogParent: HTMLElement;
  signal: AbortSignal;
  findState: FindState;
  workbook: Workbook;
  sheet: SheetMeta;
  sheets: SheetMeta[];
  sheetId: string;
  selection: Rect;
  selectedStyle: CellStyle;
  readOnly: boolean;
  t: ReturnType<typeof labels>;
  ask: (
    title: string,
    fields: Field[],
    message?: string,
    confirm?: string,
  ) => Promise<Record<string, string> | null>;
  activate: (id: string) => Promise<void>;
  select: (sheetId: string, range: string | Rect) => Promise<void>;
  parseInput: (value: string) => InputValue;
  notify: (message: string) => void;
  showDiagnostics: () => Promise<void>;
  makeClipboard: (cut: boolean) => ClipboardPayload;
  pasteText: (
    text: string,
    html?: string,
    mode?: "all" | "values" | "formulas" | "formats",
    transpose?: boolean,
  ) => Promise<void>;
  run: (fn: () => unknown) => Promise<void>;
  canvas: HTMLCanvasElement;
  runLong: <T>(work: (signal: AbortSignal) => Promise<T>) => Promise<T>;
}
export async function executeAction(
  context: ActionContext,
  name: string,
  value?: string,
) {
  const {
    dialogParent,
    signal,
    findState,
    workbook,
    sheet,
    sheets,
    sheetId,
    selection,
    selectedStyle,
    readOnly,
    t,
    ask,
    activate,
    select,
    notify,
    makeClipboard,
    showDiagnostics,
    pasteText,
    run,
    canvas,
    runLong,
  } = context;
  if (
    [
      "duplicateSheet",
      "hideSheet",
      "showSheet",
      "lockCells",
      "unlockCells",
      "protectSheet",
      "unprotectSheet",
      "nameManager",
      "validation",
      "conditionalFormat",
    ].includes(name)
  ) {
    await officeAction(context, name);
    return;
  }
  const style = (s: CellStyle) => workbook.setStyle(sheetId, selection, s);
  if (
    name === "bold" ||
    name === "italic" ||
    name === "underline" ||
    name === "wrap"
  ) {
    await style({ [name]: !selectedStyle[name] });
    return;
  }
  if (["left", "center", "right"].includes(name)) {
    await style({ align: name as CellStyle["align"] });
    return;
  }
  if (name === "font") {
    await style({ fontFamily: value });
    return;
  }
  if (name === "fontSize") {
    await style({ fontSize: Number(value) });
    return;
  }
  if (name === "color" || name === "fill") {
    await style({ [name === "color" ? "color" : "background"]: value });
    return;
  }
  if (name === "format") {
    await style({ numberFormat: value });
    return;
  }
  if (name === "verticalAlign") {
    await style({ verticalAlign: value as CellStyle["verticalAlign"] });
    return;
  }
  if (name === "customFormat") {
    const result = await ask(t.customFormat, [
      {
        key: "format",
        label: t.format,
        value: selectedStyle.numberFormat ?? "General",
      },
    ]);
    if (result) await style({ numberFormat: result.format || "General" });
    return;
  }
  if (name === "borderOptions") {
    const result = await ask(t.borderOptions, [
      {
        key: "kind",
        label: t.borderOptions,
        options: [
          { value: "all", label: t.border },
          { value: "outline", label: t.borderOutline },
          { value: "none", label: t.borderNone },
        ],
      },
      { key: "color", label: t.color, type: "color", value: "#aab9b0" },
    ]);
    if (result) {
      const empty = { top: "", right: "", bottom: "", left: "" };
      if (result.kind !== "outline")
        await style({
          border: Object.fromEntries(
            Object.keys(empty).map((side) => [
              side,
              result.kind === "none" ? "" : result.color,
            ]),
          ),
        });
      else
        await workbook.transaction([
          {
            type: "style",
            sheetId,
            range: selection,
            style: { border: empty },
          },
          ...(["top", "bottom", "left", "right"] as const).map((side) => ({
            type: "style" as const,
            sheetId,
            range: {
              ...selection,
              ...(side === "top"
                ? { r2: selection.r1 }
                : side === "bottom"
                  ? { r1: selection.r2 }
                  : side === "left"
                    ? { c2: selection.c1 }
                    : { c1: selection.c2 }),
            },
            style: { border: { [side]: result.color } },
          })),
        ]);
    }
    return;
  }
  if (name === "autoFitRows" || name === "autoFitColumns") {
    await runLong((signal) =>
      autoFit(
        workbook,
        sheet,
        selection,
        name === "autoFitRows" ? "row" : "column",
        signal,
      ),
    );
    return;
  }
  if (name === "border") {
    await style({
      border: {
        top: "#aab9b0",
        right: "#aab9b0",
        bottom: "#aab9b0",
        left: "#aab9b0",
      },
    });
    return;
  }
  if (name === "undo") await workbook.undo();
  else if (name === "redo") await workbook.redo();
  else if (name === "merge") await workbook.merge(sheetId, selection);
  else if (name === "unmerge") await workbook.unmerge(sheetId, selection);
  else if (name === "clear") await workbook.clear(sheetId, selection);
  else if (name === "freeze")
    await workbook.freeze(sheetId, selection.r1, selection.c1);
  else if (name === "unfreeze") await workbook.freeze(sheetId, 0, 0);
  else if (name === "addRow")
    await workbook.insertRows(
      sheetId,
      selection.r1,
      selection.r2 - selection.r1 + 1,
    );
  else if (name === "deleteRow")
    await workbook.deleteRows(
      sheetId,
      selection.r1,
      selection.r2 - selection.r1 + 1,
    );
  else if (name === "addColumn")
    await workbook.insertColumns(
      sheetId,
      selection.c1,
      selection.c2 - selection.c1 + 1,
    );
  else if (name === "deleteColumn")
    await workbook.deleteColumns(
      sheetId,
      selection.c1,
      selection.c2 - selection.c1 + 1,
    );
  else if (
    ["hideRows", "hideColumns", "showRows", "showColumns"].includes(name)
  ) {
    const isRow = name.endsWith("Rows"),
      show = name.startsWith("show"),
      indexes = show
        ? isRow
          ? sheet.hiddenRows
          : sheet.hiddenColumns
        : Array.from(
            {
              length: isRow
                ? selection.r2 - selection.r1 + 1
                : selection.c2 - selection.c1 + 1,
            },
            (_, i) => i + (isRow ? selection.r1 : selection.c1),
          );
    await workbook.setDimensions(sheetId, isRow ? "row" : "column", indexes, {
      hidden: !show,
    });
  } else if (name === "rowHeight" || name === "columnWidth") {
    const result = await ask(t[name], [
      {
        key: "size",
        label: t.value,
        type: "number",
        value: name === "rowHeight" ? "28" : "112",
      },
    ]);
    if (result)
      await workbook.setDimensions(
        sheetId,
        name === "rowHeight" ? "row" : "column",
        Array.from(
          {
            length:
              name === "rowHeight"
                ? selection.r2 - selection.r1 + 1
                : selection.c2 - selection.c1 + 1,
          },
          (_, i) => i + (name === "rowHeight" ? selection.r1 : selection.c1),
        ),
        { size: Number(result.size) },
      );
  } else if (name === "addSheet" || name === "rename") {
    const result = await ask(t[name], [
      {
        key: "name",
        label: t.name,
        value:
          name === "rename" ? sheet.name : `${t.sheet}${sheets.length + 1}`,
      },
    ]);
    if (result) {
      if (name === "rename") await workbook.renameSheet(sheetId, result.name);
      else await activate(await workbook.addSheet(result.name));
    }
  } else if (name === "deleteSheet") {
    if (await ask(t.deleteSheet, [], sheet.name))
      await workbook.deleteSheet(sheetId);
  } else if (name === "moveLeft" || name === "moveRight")
    await workbook.reorderSheet(
      sheetId,
      Math.max(
        0,
        Math.min(
          sheets.length - 1,
          sheets.findIndex((s) => s.id === sheetId) +
            (name === "moveLeft" ? -1 : 1),
        ),
      ),
    );
  else if (name === "namedRange") {
    const result = await ask(t.namedRange, [
      { key: "name", label: t.name, value: "MyRange" },
    ]);
    if (result) await workbook.defineName(result.name, sheetId, selection);
  } else if (name === "sortAsc" || name === "sortDesc" || name === "sort") {
    const result = await sortDialog(
      { parent: dialogParent, signal, t, workbook, sheetId, selection },
      name === "sortDesc" ? "desc" : "asc",
    );
    if (result) {
      await workbook.sort(sheetId, result.range, result.keys, result.header);
      await select(sheetId, result.range);
    }
  } else if (name === "filter") {
    const result = await filterDialog(
      { parent: dialogParent, signal, t, workbook, sheetId, selection },
      sheet.filter,
    );
    if (result) await workbook.filter(sheetId, result.range, result.rules);
  } else if (name === "clearFilter")
    await workbook.filter(sheetId, undefined, []);
  else if (name === "find" || name === "findNext") {
    const result =
      name === "findNext" && findState.search
        ? {
            search: findState.search,
            mode: "find",
            matchCase: String(findState.matchCase),
            entireCell: String(findState.entireCell),
            replacement: "",
          }
        : await ask(t.find, [
            { key: "search", label: t.search, value: findState.search },
            {
              key: "matchCase",
              label: t.matchCase,
              type: "checkbox",
              value: String(findState.matchCase),
            },
            {
              key: "entireCell",
              label: t.entireCell,
              type: "checkbox",
              value: String(findState.entireCell),
            },
            ...(!readOnly
              ? [
                  { key: "replacement", label: t.replacement },
                  {
                    key: "mode",
                    label: t.find,
                    options: [
                      { value: "find", label: t.findNext },
                      { value: "replace", label: t.replaceAll },
                    ],
                  },
                ]
              : []),
          ]);
    if (result) {
      const matchCase = result.matchCase === "true",
        entireCell = result.entireCell === "true";
      const continuing =
        findState.search === result.search &&
        findState.matchCase === matchCase &&
        findState.entireCell === entireCell &&
        findState.sheetId === sheetId;
      Object.assign(findState, {
        search: result.search,
        matchCase,
        entireCell,
        sheetId,
      });
      if (result.mode === "replace")
        await workbook.replace(sheetId, result.search, result.replacement, {
          matchCase,
          entireCell,
        });
      else {
        const found = await workbook.findNext(sheetId, result.search, {
          matchCase,
          entireCell,
          after: continuing ? address(selection.r1, selection.c1) : undefined,
        });
        if (found) {
          await select(sheetId, found.address);
          canvas.focus();
        } else notify(t.noMatches);
      }
    }
  } else if (name === "pasteSpecial") {
    const result = await ask(t.pasteSpecial, [
      {
        key: "mode",
        label: t.pasteSpecial,
        options: [
          { value: "all", label: t.pasteAll },
          { value: "values", label: t.pasteValues },
          { value: "formulas", label: t.pasteFormulas },
          { value: "formats", label: t.pasteFormats },
        ],
      },
      {
        key: "transpose",
        label: t.pasteTranspose,
        type: "checkbox",
        value: "false",
      },
    ]);
    if (result) {
      const data = await readClipboard();
      await pasteText(
        data.text,
        data.html,
        result.mode as "all",
        result.transpose === "true",
      );
    }
  } else if (name === "copy" || name === "cut") {
    const data = makeClipboard(name === "cut");
    await writeClipboard(data);
    notify(t.copied);
  } else if (name === "paste") {
    const data = await readClipboard();
    await pasteText(data.text, data.html);
  } else if (name === "open") {
    const picker = element("input");
    picker.type = "file";
    picker.accept = ".xlsx,.csv,.json";
    picker.setAttribute("aria-label", t.chooseFile);
    picker.onchange = () =>
      run(async () => {
        const file = picker.files?.[0];
        if (!file) return;
        if (file.name.toLowerCase().endsWith(".csv")) {
          const imported = await csvImportDialog(dialogParent, file, t, signal);
          if (imported)
            await runLong((signal) =>
              workbook.importCsv(imported.text, {
                sheetId,
                signal,
                start: address(selection.r1, selection.c1),
                ...imported,
              }),
            );
        } else if (file.name.endsWith(".json"))
          await runLong(async (signal) =>
            workbook.importJSON(JSON.parse(await file.text()), { signal }),
          );
        else await runLong((signal) => workbook.importXlsx(file, { signal }));
        const diagnostics = await workbook.getDiagnostics();
        if (diagnostics.length) await showDiagnostics();
      });
    picker.click();
  } else if (name === "saveXlsx") {
    const diagnostics = await workbook.getDiagnostics();
    let allowLossy = false;
    if (diagnostics.some((d) => d.lossy)) {
      if (
        !(await ask(
          t.compatibility,
          [],
          diagnostics
            .filter((d) => d.lossy)
            .map((d) => d.message)
            .join("\n"),
          t.exportAnyway,
        ))
      )
        return;
      allowLossy = true;
    }
    const result = await runLong((signal) =>
      workbook.exportXlsx({ allowLossy, signal }),
    );
    download(
      result.data.buffer as ArrayBuffer,
      "Workbook.xlsx",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
  } else if (name === "saveCsv") {
    const result = await runLong((signal) =>
      workbook.exportCsv(sheetId, { signal }),
    );
    download(result.text, sheet.name + ".csv", "text/csv;charset=utf-8");
  } else if (name === "saveJson")
    download(
      JSON.stringify(await workbook.exportJSON()),
      "Workbook.json",
      "application/json",
    );
  canvas.focus();
}

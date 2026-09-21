import type { WorkbookPersistence } from "../runtime/persistence";
import { persistenceUI } from "./persistence-ui";
import { showDiagnostics } from "./diagnostics";
import { paintCommand, type PaintSource } from "./format-painter";
import { productCss } from "./product-styles";
import { icon } from "./icons";
import { createClipboardPayload, parseHtmlTable } from "./clipboard";
import { manageEditor, type EditorView, type ViewState } from "./lifecycle";
import type { IconName } from "./icons";
import { parseInputValue } from "../core/input";
import { Workbook } from "../runtime/client";
import type {
  CellStyle,
  Command,
  InputValue,
  Rect,
  RegionCell,
  SheetMeta,
} from "../core/types";
import {
  address,
  contains,
  intersects,
  keyOf,
  parseRange,
  rangeName,
} from "../core/address";
import { formatValue } from "../core/format";
import { isError } from "../formula/values";
import { parseCsv, writeCsv } from "../io/csv";
import {
  Axis,
  COLUMN_HEADER,
  DEFAULT_COLUMN,
  DEFAULT_ROW,
  ROW_HEADER,
  validateZoom,
} from "./geometry";
import { cellBox, drawGrid, type RenderState } from "./renderer";
import { css } from "./styles";
import {
  labels,
  resolveLocale,
  errorMessage,
  type Label,
  type Locale,
} from "./locale";
export { supportedLocales } from "./locale";
export type { Locale } from "./locale";
import { dialog, element, type Field } from "./dom";
import { createToolbar } from "./toolbar";
import { executeAction, type FindState } from "./actions";
import { referenceInsertion, cycleReference } from "./formula-reference";
import { createFormulaHelp } from "./formula-help";
export interface EditorOptions {
  workbook: Workbook;
  /** Per-editor UI language. Defaults to zh-CN; workbook data is not translated. */
  locale?: Locale;
  readOnly?: boolean;
  toolbar?: boolean;
  sheetId?: string;
  zoom?: number;
  persistence?: WorkbookPersistence;
  styleNonce?: string;
  onError?: (error: Error) => void;
  toolbarItems?: readonly string[];
  actions?: readonly EditorAction[];
  onImport?: (context: EditorActionContext) => void | Promise<void>;
  onExport?: (
    format: "xlsx" | "csv" | "json",
    context: EditorActionContext,
  ) => void | Promise<void>;
}
export type EditorConfiguration = Partial<
  Omit<EditorOptions, "workbook" | "sheetId">
>;
export interface EditorActionContext {
  workbook: Workbook;
  sheetId: string;
  range: Rect;
  signal: AbortSignal;
}
export interface EditorAction {
  id: string;
  label: string;
  icon: IconName;
  allowReadOnly?: boolean;
  run(context: EditorActionContext): void | Promise<void>;
}
export interface EditState {
  editing: boolean;
  dirty: boolean;
  pending: boolean;
}
export interface Editor {
  readonly ready: Promise<void>;
  setOptions(options: EditorConfiguration): Promise<void>;
  select(sheetId: string, range: string | Rect): Promise<void>;
  getSelection(): { sheetId: string; range: Rect };
  resize(): void;
  getZoom(): number;
  setZoom(zoom: number): Promise<void>;
  getEditState(): EditState;
  commitEdit(): Promise<void>;
  cancelEdit(): void;
  destroy(): void;
  destroy(options: { commit: true }): Promise<void>;
}
interface ClipboardState {
  book: Workbook;
  sheetId: string;
  range: Rect;
  revision: number;
  cut: boolean;
  text: string;
  html: string;
  token: string;
  cells: RegionCell[];
  merges: Rect[];
  validations: NonNullable<SheetMeta["validations"]>;
  conditionalFormats: NonNullable<SheetMeta["conditionalFormats"]>;
  rows: number[];
  columns: number[];
}
let clipboard: ClipboardState | undefined;
export function mountEditor(
  container: HTMLElement,
  options: EditorOptions,
): Editor {
  return manageEditor(container, options, mountView);
}
function mountView(
  container: HTMLElement,
  options: EditorOptions,
  restored?: ViewState,
): EditorView {
  let zoom = validateZoom(restored?.zoom ?? options.zoom ?? 1);
  if (options.persistence && options.persistence.workbook !== options.workbook)
    throw new Error("Persistence belongs to a different workbook");
  const workbook = options.workbook.withOptions({ origin: "editor" }),
    locale = resolveLocale(options.locale),
    t = labels(locale),
    controller = new AbortController(),
    signal = controller.signal;
  const host = element("div");
  host.lang = locale;
  host.style.cssText = "height:100%;width:100%";
  const root = host.attachShadow({ mode: "open" });
  container.append(host);
  const style = element("style");
  if (options.styleNonce) style.nonce = options.styleNonce;
  style.textContent = css + productCss;
  root.append(style);
  const shell = element("div", "shell");
  shell.style.position = "relative";
  root.append(shell);
  const toolbar = createToolbar(
    t,
    (name, value) => run(() => action(name, value)),
    !!options.readOnly,
    signal,
    { items: options.toolbarItems, actions: options.actions },
  );
  if (options.toolbar !== false) shell.append(toolbar.menu, toolbar.ribbon);
  const formulaBar = element("div", "formula-bar"),
    addressInput = element("input", "address"),
    formulaInput = element("input", "formula");
  addressInput.setAttribute("aria-label", t.cellAddress);
  formulaInput.setAttribute("aria-label", t.formula);
  formulaInput.readOnly = !!options.readOnly;
  formulaInput.spellcheck = false;
  formulaBar.append(
    addressInput,
    element("span", "formula-symbol", "ƒx"),
    formulaInput,
  );
  shell.append(formulaBar);
  const viewport = element("div", "viewport"),
    gridLayer = element("div", "grid-layer"),
    scroll = element("div", "scroll"),
    spacer = element("div", "spacer"),
    canvas = element("canvas", "canvas"),
    input = element("textarea", "cell-input");
  input.hidden = true;
  input.spellcheck = false;
  input.setAttribute("aria-label", t.cellEditor);
  canvas.tabIndex = 0;
  canvas.setAttribute("role", "grid");
  canvas.setAttribute("aria-label", t.sheet);
  canvas.setAttribute("aria-readonly", String(!!options.readOnly));
  const accessible = element("div", "a11y");
  accessible.id = "cell-" + crypto.randomUUID();
  accessible.setAttribute("role", "gridcell");
  canvas.setAttribute("aria-activedescendant", accessible.id);
  canvas.setAttribute("aria-owns", accessible.id);
  accessible.setAttribute("aria-selected", "true");
  const toast = element("div", "toast");
  toast.hidden = true;
  toast.setAttribute("role", "status");
  scroll.append(spacer);
  gridLayer.append(scroll, canvas, input, accessible);
  viewport.append(gridLayer, toast);
  shell.append(viewport);
  const sheetbar = element("div", "sheetbar"),
    addSheet = element("button", "add-sheet", "+"),
    tabs = element("div", "tabs");
  addSheet.title = t.addSheet;
  addSheet.setAttribute("aria-label", t.addSheet);
  addSheet.disabled = !!options.readOnly;
  sheetbar.append(addSheet, tabs);
  shell.append(sheetbar);
  const statusbar = element("div", "statusbar"),
    status = element("span", "", t.loading),
    stats = element("span", "stats"),
    diagnosticsButton = element("button"),
    cancelOperation = element("button", "", t.cancel);
  cancelOperation.hidden = true;
  let operationController: AbortController | undefined;
  cancelOperation.onclick = () => operationController?.abort();
  statusbar.append(
    status,
    cancelOperation,
    diagnosticsButton,
    stats,
    zoomControl(),
  );
  shell.append(statusbar);
  let sheets: SheetMeta[] = [],
    sheetId = restored?.sheetId ?? options.sheetId ?? "",
    sheet: SheetMeta,
    rows: Axis,
    columns: Axis,
    revision = 0,
    dateSystem: 1900 | 1904 = 1900;
  let selection: Rect = restored?.selection ?? { r1: 0, c1: 0, r2: 0, c2: 0 },
    anchor = { row: 0, col: 0 },
    focus = { row: 0, col: 0 },
    cells = new Map<number, RegionCell>(),
    cachedRanges: Rect[] = [];
  let width = 0,
    height = 0,
    frame = 0,
    loadVersion = 0,
    selectedVersion = 0,
    isEditing = false,
    composing = false,
    destroyed = false,
    toastTimer = 0,
    selectionData: RegionCell[] = [],
    selectionReady = false;
  let editTarget:
    | {
        sheetId: string;
        row: number;
        column: number;
        initial: string;
        textFormat: boolean;
        scrollTop: number;
        scrollLeft: number;
      }
    | undefined;
  let pendingCommit: Promise<unknown> | undefined;
  let commitFailure: Error | undefined;
  let closing: Promise<void> | undefined;
  let referenceRange: Rect | undefined;
  let referenceDrag:
    | {
        source: HTMLInputElement | HTMLTextAreaElement;
        before: string;
        after: string;
        row: number;
        col: number;
      }
    | undefined;
  let selectionRequest: { key: string; promise: Promise<void> } | undefined;
  let drag:
      | { row: number; col: number; fill: boolean; source: Rect }
      | undefined,
    resizeDrag:
      | { axis: "row" | "column"; index: number; start: number; size: number }
      | undefined,
    contextMenu: HTMLElement | undefined;
  const findState: FindState = {
    search: "",
    matchCase: false,
    entireCell: false,
  };
  const unsubscribers: (() => void)[] = [];
  const on = <K extends keyof HTMLElementEventMap>(
    node: HTMLElement | ShadowRoot,
    type: K,
    listener: (event: HTMLElementEventMap[K]) => void,
  ) => node.addEventListener(type, listener as EventListener, { signal });
  function notify(message: string) {
    if (destroyed) return;
    toast.textContent = message;
    toast.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = window.setTimeout(() => {
      toast.hidden = true;
    }, 4500);
  }
  async function run(task: () => unknown | Promise<unknown>) {
    try {
      await task();
    } catch (e) {
      if (destroyed) return;
      const err = e as Error;
      notify(errorMessage(err, locale));
      options.onError?.(err);
    }
  }
  const state = (): RenderState => ({
    sheet,
    rows,
    columns,
    cells,
    selection,
    reference: isEditing ? referenceRange : undefined,
    zoom,
    scrollLeft: scroll.scrollLeft,
    scrollTop: scroll.scrollTop,
    width,
    height,
    dateSystem,
    accent:
      getComputedStyle(host).getPropertyValue("--oe-accent").trim() ||
      "#16734c",
  });
  function render() {
    frame = 0;
    if (destroyed || !sheet || !width || !height) return;
    drawGrid(canvas, state());
  }
  function schedule() {
    if (!frame && !destroyed) frame = requestAnimationFrame(render);
  }
  function resize() {
    gridLayer.style.transform = `scale(${zoom})`;
    gridLayer.style.width = `${100 / zoom}%`;
    gridLayer.style.height = `${100 / zoom}%`;
    width = Math.max(1, viewport.clientWidth / zoom - 14);
    height = Math.max(1, viewport.clientHeight / zoom - 14);
    schedule();
    formulaHelp.position();
    if (sheet) void run(() => fetchVisible());
  }
  function geometry() {
    rows = new Axis(
      sheet.rowCount,
      sheet.rowHeights,
      [...sheet.hiddenRows, ...(sheet.filteredRows ?? [])],
      DEFAULT_ROW,
    );
    columns = new Axis(
      sheet.columnCount,
      sheet.columnWidths,
      sheet.hiddenColumns,
      DEFAULT_COLUMN,
    );
    spacer.style.width = columns.total + ROW_HEADER + "px";
    spacer.style.height = rows.total + COLUMN_HEADER + "px";
    canvas.setAttribute("aria-rowcount", String(sheet.rowCount));
    canvas.setAttribute("aria-colcount", String(sheet.columnCount));
  }
  function renderTabs() {
    tabs.replaceChildren();
    for (const s of sheets.filter((s) => !s.hidden)) {
      const button = element(
        "button",
        "sheet-tab" + (s.id === sheetId ? " selected" : ""),
        s.name,
      );
      button.setAttribute("role", "tab");
      button.setAttribute("aria-selected", String(s.id === sheetId));
      button.onpointerdown = (event) => {
        if (referenceSource()) event.preventDefault();
      };
      button.onclick = () =>
        run(() =>
          referenceSource() ? activateReference(s.id) : activate(s.id),
        );
      button.ondblclick = () => run(() => action("rename"));
      button.oncontextmenu = (e) => {
        e.preventDefault();
        void run(async () => {
          await activate(s.id);
          showMenu(
            e.clientX - host.getBoundingClientRect().left,
            e.clientY - host.getBoundingClientRect().top,
            ["rename", "moveLeft", "moveRight", "deleteSheet"],
          );
        });
      };
      tabs.append(button);
    }
  }
  async function refresh() {
    const metadata = await workbook.getMetadata();
    if (destroyed) return;
    sheets = metadata.sheets;
    revision = metadata.revision;
    dateSystem = metadata.dateSystem;
    if (!sheets.some((s) => s.id === sheetId && !s.hidden)) {
      sheetId = sheets.find((s) => !s.hidden)!.id;
      selection = { r1: 0, c1: 0, r2: 0, c2: 0 };
    }
    sheet = sheets.find((s) => s.id === sheetId)!;
    selection = {
      r1: Math.min(selection.r1, sheet.rowCount - 1),
      r2: Math.min(selection.r2, sheet.rowCount - 1),
      c1: Math.min(selection.c1, sheet.columnCount - 1),
      c2: Math.min(selection.c2, sheet.columnCount - 1),
    };
    geometry();
    renderTabs();
    cachedRanges = [];
    cells.clear();
    resize();
    await fetchVisible(true);
    await updateSelection();
    for (const [action, enabled] of [
      ["undo", metadata.canUndo],
      ["redo", metadata.canRedo],
    ] as const)
      root
        .querySelectorAll<HTMLButtonElement>(`[data-action="${action}"]`)
        .forEach((b) => {
          b.disabled = !!options.readOnly || !enabled;
        });
  }
  const sheetViews = new Map<
    string,
    { selection: Rect; top: number; left: number }
  >(restored?.sheets);
  async function activate(id: string) {
    if (isEditing || pendingCommit) await commitEdit();
    if (sheetId)
      sheetViews.set(sheetId, {
        selection: { ...selection },
        top: scroll.scrollTop,
        left: scroll.scrollLeft,
      });
    sheetId = id;
    const view = sheetViews.get(id);
    selection = view?.selection ?? { r1: 0, c1: 0, r2: 0, c2: 0 };
    anchor = { row: selection.r1, col: selection.c1 };
    focus = { ...anchor };
    await refresh();
    scroll.scrollTop = view?.top ?? 0;
    scroll.scrollLeft = view?.left ?? 0;
    await fetchVisible();
    canvas.focus();
  }
  function visibleRanges(): Rect[] {
    const vr = rows.visible(
        scroll.scrollTop,
        height - COLUMN_HEADER,
        sheet.frozenRows,
      ),
      vc = columns.visible(
        scroll.scrollLeft,
        width - ROW_HEADER,
        sheet.frozenColumns,
      );
    const ranges: Rect[] = [];
    const spans = (items: { index: number }[]) => {
      const result: { start: number; end: number }[] = [];
      for (const { index } of items) {
        const last = result.at(-1);
        if (last && index === last.end + 1) last.end = index;
        else result.push({ start: index, end: index });
      }
      return result;
    };
    for (const fr of [false, true])
      for (const fc of [false, true]) {
        const rs = vr.filter((r) => r.frozen === fr),
          cs = vc.filter((c) => c.frozen === fc);
        if (!rs.length || !cs.length) continue;
        for (const r of spans(rs))
          for (const c of spans(cs))
            ranges.push({ r1: r.start, r2: r.end, c1: c.start, c2: c.end });
      }
    return ranges;
  }
  async function fetchVisible(force = false) {
    if (!sheet || destroyed) return;
    const visible = visibleRanges();
    if (
      !force &&
      visible.every((r) =>
        cachedRanges.some(
          (c) => c.r1 <= r.r1 && c.c1 <= r.c1 && c.r2 >= r.r2 && c.c2 >= r.c2,
        ),
      )
    )
      return;
    const version = ++loadVersion,
      sid = sheetId;
    const ranges = visible.map((r) => ({
      r1: Math.max(0, r.r1 - 16),
      r2: Math.min(sheet.rowCount - 1, r.r2 + 24),
      c1: Math.max(0, r.c1 - 2),
      c2: Math.min(sheet.columnCount - 1, r.c2 + 3),
    }));
    // Fetch merge anchors when the visible portion starts in the middle of a merged cell.
    for (const merge of sheet.merges)
      if (
        visible.some(
          (v) =>
            v.r1 <= merge.r2 &&
            v.r2 >= merge.r1 &&
            v.c1 <= merge.c2 &&
            v.c2 >= merge.c1,
        ) &&
        !ranges.some((r) => contains(r, merge.r1, merge.c1))
      )
        ranges.push({ r1: merge.r1, r2: merge.r1, c1: merge.c1, c2: merge.c1 });
    // Hidden rows can put adjacent visible cells thousands of rows apart.
    // Merge only overlapping overscan windows, then respect the Worker read limit.
    const requests: Rect[] = [];
    ranges.sort((a, b) => a.c1 - b.c1 || a.c2 - b.c2 || a.r1 - b.r1);
    for (const range of ranges) {
      const last = requests.at(-1);
      if (
        last &&
        last.c1 === range.c1 &&
        last.c2 === range.c2 &&
        range.r1 <= last.r2 + 1
      )
        last.r2 = Math.max(last.r2, range.r2);
      else requests.push({ ...range });
    }
    const bounded: Rect[] = [];
    for (const range of requests) {
      const step = Math.max(1, Math.floor(200000 / (range.c2 - range.c1 + 1)));
      for (let row = range.r1; row <= range.r2; row += step)
        bounded.push({
          ...range,
          r1: row,
          r2: Math.min(range.r2, row + step - 1),
        });
    }
    const data = await Promise.all(
      bounded.map((r) => workbook.getRegion(sid, r)),
    );
    if (destroyed || sid !== sheetId || version !== loadVersion) return;
    const next = new Map<number, RegionCell>();
    for (const result of data)
      for (const cell of result.cells)
        next.set(keyOf(cell.row, cell.column), cell);
    cells = next;
    cachedRanges = bounded;
    schedule();
    if (!isEditing && root.activeElement !== formulaInput) updateFormulaBar();
  }
  const validationChoice = element("select", "validation-choice");
  validationChoice.setAttribute("aria-label", t.chooseValue);
  validationChoice.hidden = true;
  viewport.append(validationChoice);
  validationChoice.onchange = () => {
    if (validationChoice.value)
      void run(async () => {
        await commitEdit();
        await workbook.setValues(
          sheetId,
          address(selection.r1, selection.c1),
          [[validationChoice.value]],
          { parseFormulas: false },
        );
      });
  };
  function updateFormulaBar() {
    if (root.activeElement !== addressInput)
      addressInput.value =
        selection.r1 === selection.r2 && selection.c1 === selection.c2
          ? address(selection.r1, selection.c1)
          : rangeName(selection);
    const cell = selectedCell();
    const validation = sheet?.validations?.find(
      (rule) =>
        rule.type === "list" &&
        contains(rule.range, selection.r1, selection.c1),
    );
    validationChoice.hidden = !validation || !!options.readOnly;
    validationChoice.replaceChildren(element("option", "", t.chooseValue));
    validationChoice.options[0].value = "";
    for (const value of validation?.values ?? []) {
      const option = element("option", "", value);
      option.value = value;
      validationChoice.append(option);
    }
    if (!isEditing) formulaInput.value = editText(cell);
    accessible.textContent = `${address(selection.r1, selection.c1)} ${cell ? formatValue(cell.value, cell.style, dateSystem) : ""}`;
    accessible.setAttribute("aria-rowindex", String(selection.r1 + 1));
    accessible.setAttribute("aria-colindex", String(selection.c1 + 1));
  }
  function updateSelection(): Promise<void> {
    const key = `${sheetId}:${revision}:${rangeName(selection)}`;
    if (selectionRequest?.key === key) return selectionRequest.promise;
    const promise = loadSelection().finally(() => {
      if (selectionRequest?.promise === promise) selectionRequest = undefined;
    });
    selectionRequest = { key, promise };
    return promise;
  }
  async function loadSelection() {
    const version = ++selectedVersion,
      sid = sheetId,
      range = { ...selection };
    selectionReady = false;
    updateFormulaBar();
    schedule();
    workbook.emit("selection", { sheetId, range });
    const count = (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1);
    if (count > 100000) {
      stats.textContent = `${count.toLocaleString(locale)} ${t.cells}`;
      selectionData = [];
      return;
    }
    const result = await workbook.getRegion(sid, range);
    if (destroyed || version !== selectedVersion || sid !== sheetId) return;
    selectionData = result.cells;
    selectionReady = true;
    toolbar.update(
      result.cells.map((c) => c.style),
      result.cells.length < count,
    );
    const hiddenRows = new Set([
      ...sheet.hiddenRows,
      ...(sheet.filteredRows ?? []),
    ]);
    const hiddenColumns = new Set(sheet.hiddenColumns);
    const values = result.cells.filter(
        (c) =>
          c.value !== null &&
          !hiddenRows.has(c.row) &&
          !hiddenColumns.has(c.column),
      ),
      numbers = values
        .map((c) => c.value)
        .filter((v): v is number => typeof v === "number"),
      sum = numbers.reduce((a, b) => a + b, 0);
    stats.textContent = values.length
      ? `${t.count} ${values.length}${numbers.length ? `　${t.average} ${(sum / numbers.length).toLocaleString(locale, { maximumFractionDigits: 2 })}　${t.sum} ${sum.toLocaleString(locale, { maximumFractionDigits: 2 })}` : ""}`
      : t.help;
    if (!isEditing && root.activeElement !== formulaInput) updateFormulaBar();
  }
  async function select(id: string, range: string | Rect) {
    if (isEditing || pendingCommit) await commitEdit();
    if (id !== sheetId) await activate(id);
    const next = parseRange(range);
    if (next.r2 >= sheet.rowCount || next.c2 >= sheet.columnCount)
      throw new Error(t.selectionBounds);
    selection = next;
    anchor = { row: next.r1, col: next.c1 };
    focus = { row: next.r2, col: next.c2 };
    reveal(next.r1, next.c1);
    await fetchVisible();
    await updateSelection();
  }
  function reveal(row: number, col: number) {
    if (row >= sheet.frozenRows) {
      const top = rows.offsets[row],
        bottom = rows.offsets[row + 1];
      if (top < scroll.scrollTop + rows.offsets[sheet.frozenRows])
        scroll.scrollTop = top - rows.offsets[sheet.frozenRows];
      else if (bottom > scroll.scrollTop + height - COLUMN_HEADER)
        scroll.scrollTop = bottom - height + COLUMN_HEADER;
    }
    if (col >= sheet.frozenColumns) {
      const left = columns.offsets[col],
        right = columns.offsets[col + 1];
      if (left < scroll.scrollLeft + columns.offsets[sheet.frozenColumns])
        scroll.scrollLeft = left - columns.offsets[sheet.frozenColumns];
      else if (right > scroll.scrollLeft + width - ROW_HEADER)
        scroll.scrollLeft = right - width + ROW_HEADER;
    }
  }
  function point(event: MouseEvent): {
    row: number;
    col: number;
    x: number;
    y: number;
  } {
    const box = canvas.getBoundingClientRect(),
      x = (event.clientX - box.left) / zoom,
      y = (event.clientY - box.top) / zoom;
    const rx = x - ROW_HEADER,
      ry = y - COLUMN_HEADER;
    return {
      row: rows.at(
        Math.max(
          0,
          ry + (ry < rows.offsets[sheet.frozenRows] ? 0 : scroll.scrollTop),
        ),
      ),
      col: columns.at(
        Math.max(
          0,
          rx +
            (rx < columns.offsets[sheet.frozenColumns] ? 0 : scroll.scrollLeft),
        ),
      ),
      x,
      y,
    };
  }
  function normalize(
    a: { row: number; col: number },
    b: { row: number; col: number },
  ): Rect {
    return {
      r1: Math.min(a.row, b.row),
      c1: Math.min(a.col, b.col),
      r2: Math.max(a.row, b.row),
      c2: Math.max(a.col, b.col),
    };
  }
  function selectedCell() {
    return (
      cells.get(keyOf(selection.r1, selection.c1)) ??
      selectionData.find(
        (c) => c.row === selection.r1 && c.column === selection.c1,
      )
    );
  }
  function editText(cell?: RegionCell): string {
    if (cell?.formula) return cell.formula;
    const value = cell?.value;
    if (typeof value === "string" && /^[=']/.test(value)) return "'" + value;
    return isError(value) ? value.error : String(value ?? "");
  }
  function beginDraft(): boolean {
    if (options.readOnly || !sheet || destroyed || closing) return false;
    if (isEditing) return true;
    const cell = selectedCell();
    if (cell?.spill) {
      notify(t.spillEdit);
      return false;
    }
    editTarget = {
      sheetId,
      row: selection.r1,
      column: selection.c1,
      initial: editText(cell),
      textFormat: cell?.style.numberFormat === "@",
      scrollTop: scroll.scrollTop,
      scrollLeft: scroll.scrollLeft,
    };
    input.value = editTarget.initial;
    isEditing = true;
    return true;
  }
  async function beginEdit(value?: string) {
    if (!beginDraft()) return;
    reveal(selection.r1, selection.c1);
    if (editTarget) {
      editTarget.scrollTop = scroll.scrollTop;
      editTarget.scrollLeft = scroll.scrollLeft;
    }
    const box = cellBox(selection.r1, selection.c1, state());
    input.style.left = box.x + "px";
    input.style.top = box.y + "px";
    input.style.width = Math.max(box.width, 130) + "px";
    input.style.height = Math.max(box.height, 30) + "px";
    if (value !== undefined) input.value = value;
    input.hidden = false;
    input.focus();
    if (value === undefined) input.select();
    else input.setSelectionRange(input.value.length, input.value.length);
    formulaInput.value = input.value;
    formulaHelp.update(input);
  }
  function parseInput(
    value: string,
    textFormat = selectedCell()?.style.numberFormat === "@",
  ): InputValue {
    return parseInputValue(value, textFormat ? "text" : "auto", dateSystem)
      .value;
  }
  async function commitEdit() {
    if (!isEditing || composing || options.readOnly || !editTarget)
      return pendingCommit;
    const target = editTarget;
    const value = input.value;
    endEdit();
    if (value === target.initial) return pendingCommit;
    selectionReady = false;
    const parsed = parseInputValue(
      value,
      target.textFormat ? "text" : "auto",
      dateSystem,
    );
    const range = {
      r1: target.row,
      r2: target.row,
      c1: target.column,
      c2: target.column,
    };
    const operation = workbook.transaction([
      {
        type: "setValues",
        sheetId: target.sheetId,
        range,
        values: [[parsed.value]],
        parseFormulas: !target.textFormat && !value.startsWith("'"),
      },
      ...(parsed.format && !sheet.protected
        ? [
            {
              type: "style" as const,
              sheetId: target.sheetId,
              range,
              style: { numberFormat: parsed.format },
            },
          ]
        : []),
    ]);
    pendingCommit = operation;
    try {
      await operation;
      commitFailure = undefined;
    } catch (error) {
      commitFailure = error as Error;
      if (!isEditing && !destroyed) {
        editTarget = target;
        isEditing = true;
        input.value = value;
        formulaInput.value = value;
      }
      throw error;
    } finally {
      if (pendingCommit === operation) pendingCommit = undefined;
    }
  }
  function endEdit() {
    isEditing = false;
    editTarget = undefined;
    referenceDrag = undefined;
    referenceRange = undefined;
    schedule();
    input.hidden = true;
    formulaHelp.hide();
  }
  function referenceSource() {
    if (!isEditing || composing || editTarget?.textFormat || options.readOnly)
      return;
    const source =
      root.activeElement === formulaInput
        ? formulaInput
        : root.activeElement === input
          ? input
          : undefined;
    if (!source) return;
    const span = referenceInsertion(
      source.value,
      source.selectionStart ?? 0,
      source.selectionEnd ?? 0,
    );
    if (span) return { source, span };
  }
  async function activateReference(id: string) {
    if (id === sheetId) return;
    const caret =
      referenceSource()?.source.selectionStart ?? input.value.length;
    sheetId = id;
    scroll.scrollTop = 0;
    scroll.scrollLeft = 0;
    selection = { r1: 0, c1: 0, r2: 0, c2: 0 };
    referenceRange = undefined;
    if (editTarget) {
      editTarget.scrollTop = scroll.scrollTop;
      editTarget.scrollLeft = scroll.scrollLeft;
    }
    input.hidden = true;
    formulaInput.focus();
    formulaInput.setSelectionRange(caret, caret);
    await refresh();
    if (editTarget) {
      editTarget.scrollTop = scroll.scrollTop;
      editTarget.scrollLeft = scroll.scrollLeft;
    }
  }
  function insertReference(row: number, col: number) {
    if (!referenceDrag || !editTarget) return;
    referenceRange = normalize(
      { row: referenceDrag.row, col: referenceDrag.col },
      { row, col },
    );
    const prefix =
      sheetId === editTarget.sheetId
        ? ""
        : "\'" + sheet.name.replace(/\'/g, "\'\'") + "\'!";
    const name =
      referenceRange.r1 === referenceRange.r2 &&
      referenceRange.c1 === referenceRange.c2
        ? address(row, col)
        : rangeName(referenceRange);
    const value = referenceDrag.before + prefix + name + referenceDrag.after;
    input.value = formulaInput.value = value;
    const caret = referenceDrag.before.length + prefix.length + name.length;
    referenceDrag.source.setSelectionRange(caret, caret);
    formulaHelp.hide();
    schedule();
  }
  async function flushEdit() {
    if (destroyed) throw new Error(t.editorDestroyed);
    if (composing) throw new Error(t.compositionPending);
    await commitEdit();
    if (commitFailure) throw commitFailure;
    await refresh();
  }
  function cancelEdit() {
    if (destroyed) return;
    endEdit();
    commitFailure = undefined;
    updateFormulaBar();
  }
  async function move(dr: number, dc: number, extend = false) {
    const nextVisible = (axis: Axis, current: number, delta: number) => {
      let next = Math.min(axis.count - 1, Math.max(0, current + delta));
      const step = Math.sign(delta);
      while (step && !axis.size(next) && next > 0 && next < axis.count - 1)
        next += step;
      return axis.size(next) ? next : current;
    };
    const row = nextVisible(rows, extend ? focus.row : selection.r1, dr);
    const col = nextVisible(columns, extend ? focus.col : selection.c1, dc);
    focus = { row, col };
    selection = extend
      ? normalize(anchor, { row, col })
      : { r1: row, r2: row, c1: col, c2: col };
    if (!extend) anchor = { row, col };
    reveal(row, col);
    schedule();
    await fetchVisible();
    await updateSelection();
  }
  async function ask(
    title: string,
    fields: Field[],
    message?: string,
    confirm?: string,
  ) {
    return dialog(shell, title, fields, t, { message, confirm, signal });
  }
  function closeMenu() {
    contextMenu?.remove();
    contextMenu = undefined;
  }
  function showMenu(x: number, y: number, items: string[]) {
    closeMenu();
    contextMenu = element("div", "context-menu");
    contextMenu.style.left = Math.min(x, host.clientWidth - 200) + "px";
    contextMenu.style.top =
      Math.max(0, Math.min(y, host.clientHeight - items.length * 32 - 20)) +
      "px";
    for (const name of items) {
      const b = element("button", "", t[name as Label] ?? name);
      b.disabled = !!options.readOnly && name !== "copy";
      b.onclick = () => {
        closeMenu();
        void run(() => action(name));
      };
      contextMenu.append(b);
    }
    shell.append(contextMenu);
  }
  function clipboardText(): string {
    const lookup = new Map(
      selectionData.map((c) => [keyOf(c.row, c.column), c]),
    );
    const data: InputValue[][] = [];
    const hiddenRows = new Set([
      ...sheet.hiddenRows,
      ...(sheet.filteredRows ?? []),
    ]);
    const hiddenColumns = new Set(sheet.hiddenColumns);
    for (let r = selection.r1; r <= selection.r2; r++) {
      if (hiddenRows.has(r)) continue;
      const row: InputValue[] = [];
      for (let c = selection.c1; c <= selection.c2; c++) {
        if (hiddenColumns.has(c)) continue;
        const cell = lookup.get(keyOf(r, c));
        row.push(cell ? formatValue(cell.value, cell.style, dateSystem) : "");
      }
      data.push(row);
    }
    return writeCsv(data, "\t").replace(/^\uFEFF/, "");
  }
  function captureClipboard(cut: boolean): ClipboardState {
    if (!selectionReady) throw new Error(t.clipboardLimit);
    return {
      book: options.workbook,
      sheetId,
      range: { ...selection },
      revision,
      cut,
      ...createClipboardPayload(clipboardText()),
      cells: structuredClone(selectionData),
      validations: structuredClone(
        (sheet.validations ?? []).filter((rule) =>
          intersects(rule.range, selection),
        ),
      ),
      conditionalFormats: structuredClone(
        (sheet.conditionalFormats ?? []).filter((rule) =>
          intersects(rule.range, selection),
        ),
      ),
      merges: sheet.merges
        .filter((merge) => intersects(merge, selection))
        .map((merge) => ({ ...merge })),
      rows: Array.from(
        { length: selection.r2 - selection.r1 + 1 },
        (_, i) => selection.r1 + i,
      ).filter(
        (r) =>
          !sheet.hiddenRows.includes(r) && !sheet.filteredRows?.includes(r),
      ),
      columns: Array.from(
        { length: selection.c2 - selection.c1 + 1 },
        (_, i) => selection.c1 + i,
      ).filter((c) => !sheet.hiddenColumns.includes(c)),
    };
  }
  function makeClipboard(cut: boolean): ClipboardState {
    return (clipboard = captureClipboard(cut));
  }
  let paintSource: PaintSource | undefined;
  function updatePainter() {
    root
      .querySelector('[data-action="formatPainter"]')
      ?.setAttribute("aria-pressed", String(!!paintSource));
    canvas.style.cursor = paintSource ? "crosshair" : "cell";
  }
  on(toolbar.menu, "click", updatePainter);
  async function applyPaint() {
    if (!paintSource || options.readOnly) return;
    const source = paintSource;
    const single =
      selection.r1 === selection.r2 && selection.c1 === selection.c2;
    const visible = (
      start: number,
      end: number,
      count: number | undefined,
      max: number,
      hidden: number[],
    ) => {
      const excluded = new Set(hidden),
        result: number[] = [];
      for (
        let i = start;
        i < max && (count !== undefined ? result.length < count : i <= end);
        i++
      )
        if (!excluded.has(i)) result.push(i);
      if (count !== undefined && result.length !== count)
        throw new Error(t.selectionBounds);
      return result;
    };
    const targetRows = visible(
      selection.r1,
      selection.r2,
      single ? source.rows.length : undefined,
      sheet.rowCount,
      [...sheet.hiddenRows, ...(sheet.filteredRows ?? [])],
    );
    const targetColumns = visible(
      selection.c1,
      selection.c2,
      single ? source.columns.length : undefined,
      sheet.columnCount,
      sheet.hiddenColumns,
    );
    await workbook.transaction([
      paintCommand(source, sheetId, targetRows, targetColumns),
    ]);
    paintSource = undefined;
    updatePainter();
  }
  function zoomControl() {
    const label = element("label", "zoom-control", t.zoom),
      control = element("select");
    control.setAttribute("aria-label", t.zoom);
    for (const value of [
      ...new Set([0.5, 0.75, 1, 1.25, 1.5, 1.75, 2, zoom]),
    ].sort((a, b) => a - b)) {
      const option = element("option", "", Math.round(value * 100) + "%");
      option.value = String(value);
      control.append(option);
    }
    control.value = String(zoom);
    control.onchange = () => void run(() => setZoom(Number(control.value)));
    label.append(control);
    return label;
  }
  async function setZoom(value: number) {
    validateZoom(value);
    await flushEdit();
    if (destroyed) throw new Error("Editor is destroyed");
    zoom = value;
    const control = root.querySelector<HTMLSelectElement>(
      ".zoom-control select",
    )!;
    if (
      ![...control.options].some((option) => Number(option.value) === value)
    ) {
      const option = element("option", "", Math.round(value * 100) + "%");
      option.value = String(value);
      control.append(option);
    }
    control.value = String(value);
    resize();
    if (sheet) {
      reveal(selection.r1, selection.c1);
      await fetchVisible(true);
    }
  }
  let navigation = Promise.resolve();
  async function jumpBoundary(dr: number, dc: number, extend: boolean) {
    await flushEdit();
    const id = sheetId,
      original = rangeName(selection),
      from = extend ? focus : { row: selection.r1, col: selection.c1 };
    const target = await workbook.getNavigationTarget(
      id,
      address(from.row, from.col),
      dr ? (dr < 0 ? "up" : "down") : dc < 0 ? "left" : "right",
    );
    if (destroyed || id !== sheetId || original !== rangeName(selection))
      return;
    focus = { row: target.row, col: target.column };
    selection = extend
      ? normalize(anchor, focus)
      : {
          r1: target.row,
          r2: target.row,
          c1: target.column,
          c2: target.column,
        };
    if (!extend) anchor = { ...focus };
    reveal(target.row, target.column);
    schedule();
    await fetchVisible();
    await updateSelection();
  }
  const dirtyDraft = () =>
    !!commitFailure ||
    !!pendingCommit ||
    !!(editTarget && input.value !== editTarget.initial);
  if (options.persistence) {
    unsubscribers.push(
      options.persistence.bindEditor({
        commit: flushEdit,
        dirty: dirtyDraft,
        cancel: cancelEdit,
      }),
    );
    unsubscribers.push(
      persistenceUI(shell, options.persistence, locale, run, dirtyDraft),
    );
    const panel = shell.querySelector(".persistence-bar")!;
    shell.insertBefore(panel, formulaBar);
  }
  async function pasteText(
    text: string,
    html?: string,
    mode: "all" | "values" | "formulas" | "formats" = "all",
    transpose = false,
  ) {
    if (options.readOnly) return;
    const table = html ? parseHtmlTable(html) : undefined;
    const copied =
      clipboard &&
      table?.token === clipboard.token &&
      clipboard.text.replace(/\r\n/g, "\n") === text.replace(/\r\n/g, "\n")
        ? clipboard
        : undefined;
    const visibleTargets = (
      start: number,
      count: number,
      max: number,
      hidden: number[],
    ) => {
      const indices: number[] = [],
        excluded = new Set(hidden);
      for (let i = start; i < max && indices.length < count; i++)
        if (!excluded.has(i)) indices.push(i);
      if (indices.length !== count) throw new Error(t.selectionBounds);
      return indices;
    };
    if (copied && !copied.cut) {
      await workbook.transaction([
        {
          type: "paste",
          sheetId,
          source: copied.range,
          cells: copied.cells,
          sourceRows: copied.rows,
          sourceColumns: copied.columns,
          targetRow: selection.r1,
          targetColumn: selection.c1,
          targetRows: visibleTargets(
            selection.r1,
            transpose ? copied.columns.length : copied.rows.length,
            sheet.rowCount,
            [...sheet.hiddenRows, ...(sheet.filteredRows ?? [])],
          ),
          targetColumns: visibleTargets(
            selection.c1,
            transpose ? copied.rows.length : copied.columns.length,
            sheet.columnCount,
            sheet.hiddenColumns,
          ),
          merges: copied.merges,
          validations: copied.validations,
          conditionalFormats: copied.conditionalFormats,
          mode,
          transpose,
        },
      ]);
      return;
    }
    if (copied?.cut) {
      const currentRevision = (await copied.book.getMetadata()).revision;
      if (
        copied.book !== options.workbook ||
        copied.revision !== currentRevision ||
        mode !== "all" ||
        transpose
      )
        throw new Error(
          "Cut contents changed or cannot be moved here; copy the selection again",
        );
      if (
        copied.rows.length !== copied.range.r2 - copied.range.r1 + 1 ||
        copied.columns.length !== copied.range.c2 - copied.range.c1 + 1
      )
        throw new Error("Cut requires a contiguous visible selection");
      const targetRows = visibleTargets(
        selection.r1,
        copied.rows.length,
        sheet.rowCount,
        [...sheet.hiddenRows, ...(sheet.filteredRows ?? [])],
      );
      const targetColumns = visibleTargets(
        selection.c1,
        copied.columns.length,
        sheet.columnCount,
        sheet.hiddenColumns,
      );
      if (
        targetRows.some((row, i) => row !== selection.r1 + i) ||
        targetColumns.some((column, i) => column !== selection.c1 + i)
      )
        throw new Error("Cut requires a contiguous visible selection");
      await workbook.copyRange(
        copied.sheetId,
        copied.range,
        sheetId,
        address(selection.r1, selection.c1),
        { cut: copied.cut },
      );
      if (copied.cut && clipboard === copied) clipboard = undefined;
      return;
    }
    let values: InputValue[][];
    const styles = table?.styles ?? new Map<number, CellStyle>();
    const textFormat = selectedCell()?.style.numberFormat === "@";
    const parseInput = (value: string, row: number, column: number) => {
      const parsed = parseInputValue(
        value,
        textFormat ? "text" : "auto",
        dateSystem,
      );
      if (parsed.format) {
        const key = keyOf(row, column);
        styles.set(key, { ...styles.get(key), numberFormat: parsed.format });
      }
      return parsed.value;
    };
    values = table
      ? table.values.map((row, r) =>
          row.map((value, c) =>
            value === null ? null : parseInput(value, r, c),
          ),
        )
      : parseCsv(text, "\t").map((row, r) =>
          row.map((value, c) => parseInput(value, r, c)),
        );
    if (!values.length) return;
    if (mode === "formats")
      throw new Error("Formats-only paste requires a spreadsheet clipboard");
    const source = {
      r1: 0,
      c1: 0,
      r2: values.length - 1,
      c2: values[0].length - 1,
    };
    const targetRows = visibleTargets(
      selection.r1,
      transpose ? values[0].length : values.length,
      sheet.rowCount,
      [...sheet.hiddenRows, ...(sheet.filteredRows ?? [])],
    );
    const targetColumns = visibleTargets(
      selection.c1,
      transpose ? values.length : values[0].length,
      sheet.columnCount,
      sheet.hiddenColumns,
    );
    const formats: Command[] = [];
    if (mode === "all" && !sheet.protected)
      values.forEach((row, r) =>
        row.forEach((_, c) => {
          const style = styles.get(keyOf(r, c));
          if (!style) return;
          const tr = targetRows[transpose ? c : r],
            tc = targetColumns[transpose ? r : c];
          formats.push({
            type: "style",
            sheetId,
            range: { r1: tr, r2: tr, c1: tc, c2: tc },
            style,
          });
        }),
      );
    if (mode === "all")
      for (const merge of table?.merges ?? []) {
        const range = {
          r1: targetRows[transpose ? merge.c1 : merge.r1],
          r2: targetRows[transpose ? merge.c2 : merge.r2],
          c1: targetColumns[transpose ? merge.r1 : merge.c1],
          c2: targetColumns[transpose ? merge.r2 : merge.c2],
        };
        if (
          range.r2 - range.r1 !==
            (transpose ? merge.c2 - merge.c1 : merge.r2 - merge.r1) ||
          range.c2 - range.c1 !==
            (transpose ? merge.r2 - merge.r1 : merge.c2 - merge.c1)
        )
          throw new Error("Merged cells require contiguous paste destinations");
        formats.push({ type: "merge", sheetId, range });
      }
    await workbook.transaction([
      {
        type: "paste",
        sheetId,
        source,
        cells: values.flatMap((row, r) =>
          row.map((value, c) => ({
            row: r,
            column: c,
            value,
            style: {},
          })),
        ),
        targetRow: selection.r1,
        targetColumn: selection.c1,
        targetRows,
        targetColumns,
        mode: mode === "all" ? "values" : mode,
        transpose,
      },
      ...formats,
    ]);
  }
  async function runLong<T>(
    work: (signal: AbortSignal) => Promise<T>,
  ): Promise<T> {
    if (operationController) throw new Error(t.fileBusy);
    const controller = new AbortController();
    operationController = controller;
    cancelOperation.hidden = false;
    try {
      return await work(controller.signal);
    } finally {
      operationController = undefined;
      cancelOperation.hidden = true;
    }
  }
  async function action(name: string, value?: string) {
    if (!sheet) return;
    if (name === "diagnostics") {
      await flushEdit();
      await showDiagnostics(shell, workbook, locale, signal, select);
      return;
    }
    if (name === "shortcuts") {
      await ask(t.shortcuts, [], t.shortcutGuide, t.close);
      return;
    }
    if (name === "saveNow") {
      await options.persistence?.save();
      return;
    }
    if (name === "formatPainter") {
      if (options.readOnly) return;
      await flushEdit();
      await updateSelection();
      paintSource = paintSource ? undefined : captureClipboard(false);
      updatePainter();
      if (paintSource) notify(t.paintHint);
      canvas.focus();
      return;
    }
    const custom = options.actions?.find(
      (item) => "custom:" + item.id === name,
    );
    if (
      options.readOnly &&
      !custom?.allowReadOnly &&
      !["saveXlsx", "saveCsv", "saveJson", "copy", "find", "findNext"].includes(
        name,
      )
    )
      return;
    if (isEditing || pendingCommit) {
      await commitEdit();
      await refresh();
    }
    const format = (
      { saveXlsx: "xlsx", saveCsv: "csv", saveJson: "json" } as const
    )[name as "saveXlsx" | "saveCsv" | "saveJson"];
    if (
      custom ||
      (name === "open" && options.onImport) ||
      (format && options.onExport)
    ) {
      await runLong(async (signal) => {
        const context = { workbook, sheetId, range: { ...selection }, signal };
        if (custom) await custom.run(context);
        else if (format) await options.onExport!(format, context);
        else await options.onImport!(context);
      });
      return;
    }
    await executeAction(
      {
        dialogParent: shell,
        signal,
        findState,
        workbook,
        sheet,
        sheets,
        sheetId,
        selection,
        selectedStyle:
          cells.get(keyOf(selection.r1, selection.c1))?.style ?? {},
        readOnly: !!options.readOnly,
        t,
        ask,
        activate,
        select,
        parseInput,
        notify,
        makeClipboard,
        showDiagnostics: () =>
          showDiagnostics(shell, workbook, locale, signal, select),
        pasteText,
        run,
        canvas,
        runLong,
      },
      name,
      value,
    );
  }
  on(scroll, "scroll", () => {
    if (
      isEditing &&
      input.value.startsWith("=") &&
      root.activeElement === input &&
      !referenceDrag
    ) {
      const caret = input.selectionStart;
      input.hidden = true;
      formulaInput.focus();
      formulaInput.setSelectionRange(caret, caret);
      formulaHelp.update(formulaInput);
    }
    if (
      isEditing &&
      !referenceDrag &&
      !input.value.startsWith("=") &&
      editTarget &&
      (scroll.scrollTop !== editTarget.scrollTop ||
        scroll.scrollLeft !== editTarget.scrollLeft)
    )
      void run(commitEdit);
    schedule();
    void run(() => fetchVisible());
  });
  canvas.addEventListener(
    "wheel",
    (event) => {
      event.preventDefault();
      if (event.shiftKey) scroll.scrollLeft += event.deltaY || event.deltaX;
      else {
        scroll.scrollTop += event.deltaY;
        scroll.scrollLeft += event.deltaX;
      }
    },
    { signal, passive: false },
  );
  on(canvas, "pointerdown", (event) => {
    if (!sheet || event.button !== 0) return;
    closeMenu();
    const reference = referenceSource();
    const p = point(event);
    if (reference && p.x >= ROW_HEADER && p.y >= COLUMN_HEADER) {
      event.preventDefault();
      referenceDrag = {
        source: reference.source,
        before: reference.source.value.slice(0, reference.span.start),
        after: reference.source.value.slice(reference.span.end),
        row: p.row,
        col: p.col,
      };
      canvas.setPointerCapture(event.pointerId);
      insertReference(p.row, p.col);
      return;
    }
    if (isEditing) void run(commitEdit);
    canvas.focus();
    canvas.setPointerCapture(event.pointerId);
    const box = cellBox(p.row, p.col, state());
    if (
      !options.readOnly &&
      sheet.filter &&
      p.y < COLUMN_HEADER &&
      p.col >= sheet.filter.range.c1 &&
      p.col <= sheet.filter.range.c2 &&
      p.x > box.x + box.width - 24 &&
      p.x < box.x + box.width - 5
    ) {
      selection = {
        r1: sheet.filter.range.r1,
        r2: sheet.filter.range.r1,
        c1: p.col,
        c2: p.col,
      };
      void run(() => action("filter"));
      return;
    }
    if (
      !options.readOnly &&
      p.y < COLUMN_HEADER &&
      Math.abs(p.x - box.x - box.width) < 5
    ) {
      resizeDrag = {
        axis: "column",
        index: p.col,
        start: p.x,
        size: box.width,
      };
      return;
    }
    if (
      !options.readOnly &&
      p.x < ROW_HEADER &&
      Math.abs(p.y - box.y - box.height) < 5
    ) {
      resizeDrag = { axis: "row", index: p.row, start: p.y, size: box.height };
      return;
    }
    const end = cellBox(selection.r2, selection.c2, state());
    const fill =
      !options.readOnly &&
      !paintSource &&
      Math.abs(p.x - end.x - end.width) < 7 &&
      Math.abs(p.y - end.y - end.height) < 7;
    drag = { row: p.row, col: p.col, fill, source: { ...selection } };
    if (fill) return;
    if (p.x < ROW_HEADER && p.y < COLUMN_HEADER)
      selection = {
        r1: 0,
        c1: 0,
        r2: sheet.rowCount - 1,
        c2: sheet.columnCount - 1,
      };
    else if (p.y < COLUMN_HEADER)
      selection = { r1: 0, r2: sheet.rowCount - 1, c1: p.col, c2: p.col };
    else if (p.x < ROW_HEADER)
      selection = { r1: p.row, r2: p.row, c1: 0, c2: sheet.columnCount - 1 };
    else {
      if (!event.shiftKey) anchor = { row: p.row, col: p.col };
      focus = { row: p.row, col: p.col };
      selection = normalize(anchor, p);
      const merge = sheet.merges.find((m) => contains(m, p.row, p.col));
      if (merge) selection = { ...merge };
    }
    schedule();
    void run(updateSelection);
  });
  on(canvas, "pointermove", (event) => {
    if (!sheet) return;
    const p = point(event);
    if (referenceDrag) {
      if (p.y > height - 16) scroll.scrollTop += 24;
      if (p.x > width - 16) scroll.scrollLeft += 24;
      insertReference(p.row, p.col);
      return;
    }
    if (resizeDrag) {
      canvas.style.cursor =
        resizeDrag.axis === "row" ? "row-resize" : "col-resize";
      return;
    }
    if (!drag) {
      const box = cellBox(p.row, p.col, state());
      canvas.style.cursor =
        p.y < COLUMN_HEADER && Math.abs(p.x - box.x - box.width) < 5
          ? "col-resize"
          : p.x < ROW_HEADER && Math.abs(p.y - box.y - box.height) < 5
            ? "row-resize"
            : paintSource
              ? "crosshair"
              : "cell";
      return;
    }
    if (p.y > height - 16) scroll.scrollTop += 24;
    if (p.x > width - 16) scroll.scrollLeft += 24;
    if (p.y < COLUMN_HEADER + 8 && scroll.scrollTop) scroll.scrollTop -= 24;
    if (p.x < ROW_HEADER + 8 && scroll.scrollLeft) scroll.scrollLeft -= 24;
    if (drag.fill)
      selection = {
        r1: Math.min(drag.source.r1, p.row),
        r2: Math.max(drag.source.r2, p.row),
        c1: Math.min(drag.source.c1, p.col),
        c2: Math.max(drag.source.c2, p.col),
      };
    else if (p.x >= ROW_HEADER && p.y >= COLUMN_HEADER)
      selection = normalize(anchor, p);
    schedule();
  });
  on(canvas, "pointerup", (event) => {
    if (!sheet) return;
    if (referenceDrag) {
      const source = referenceDrag.source;
      referenceDrag = undefined;
      if (editTarget) {
        editTarget.scrollTop = scroll.scrollTop;
        editTarget.scrollLeft = scroll.scrollLeft;
      }
      if (canvas.hasPointerCapture(event.pointerId))
        canvas.releasePointerCapture(event.pointerId);
      source.focus();
      return;
    }
    if (resizeDrag) {
      const p = point(event),
        d = resizeDrag;
      resizeDrag = undefined;
      void run(() =>
        workbook.setDimensions(sheetId, d.axis, [d.index], {
          size: Math.max(8, d.size + (d.axis === "row" ? p.y : p.x) - d.start),
        }),
      );
    }
    if (drag?.fill) {
      const source = drag.source;
      void run(() => workbook.fill(sheetId, source, selection));
    }
    if (drag && !drag.fill && paintSource) void run(applyPaint);
    drag = undefined;
    if (canvas.hasPointerCapture(event.pointerId))
      canvas.releasePointerCapture(event.pointerId);
    void run(updateSelection);
  });
  on(canvas, "dblclick", () => {
    if (!paintSource) void run(() => beginEdit());
  });
  on(canvas, "contextmenu", (event) => {
    event.preventDefault();
    const box = host.getBoundingClientRect();
    showMenu(event.clientX - box.left, event.clientY - box.top, [
      "copy",
      "cut",
      "paste",
      "clear",
      "addRow",
      "addColumn",
      "deleteRow",
      "deleteColumn",
      "merge",
      "unmerge",
    ]);
  });
  on(canvas, "keydown", (event) => {
    if (!sheet) return;
    if (event.key === "F1") {
      event.preventDefault();
      void run(() => action("shortcuts"));
      return;
    }
    if (event.key === "Escape" && paintSource) {
      paintSource = undefined;
      updatePainter();
      return;
    }
    if (event.key === "Process" || event.keyCode === 229) {
      void run(() => beginEdit(""));
      return;
    }
    if (event.isComposing) return;
    const ctrl = event.ctrlKey || event.metaKey;
    if (ctrl && ["c", "v", "x"].includes(event.key.toLowerCase())) return;
    if (ctrl && event.key.toLowerCase() === "z") {
      event.preventDefault();
      void run(() => action(event.shiftKey ? "redo" : "undo"));
      return;
    }
    if (ctrl && event.key.toLowerCase() === "y") {
      event.preventDefault();
      void run(() => action("redo"));
      return;
    }
    if (event.key === "F3") {
      event.preventDefault();
      void run(() => action("findNext"));
      return;
    }
    if (ctrl && event.key.toLowerCase() === "f") {
      event.preventDefault();
      void run(() => action("find"));
      return;
    }
    if (ctrl && event.key.toLowerCase() === "s") {
      event.preventDefault();
      void run(() => action(options.persistence ? "saveNow" : "saveXlsx"));
      return;
    }
    if (ctrl && event.key.toLowerCase() === "a") {
      event.preventDefault();
      selection = {
        r1: 0,
        c1: 0,
        r2: sheet.rowCount - 1,
        c2: sheet.columnCount - 1,
      };
      void run(updateSelection);
      return;
    }
    const directions: Record<string, [number, number]> = {
      ArrowDown: [1, 0],
      ArrowUp: [-1, 0],
      ArrowLeft: [0, -1],
      ArrowRight: [0, 1],
      Tab: [0, event.shiftKey ? -1 : 1],
      Enter: [event.shiftKey ? -1 : 1, 0],
      PageDown: [Math.max(1, Math.floor(height / DEFAULT_ROW) - 1), 0],
      PageUp: [-Math.max(1, Math.floor(height / DEFAULT_ROW) - 1), 0],
    };
    if (directions[event.key]) {
      event.preventDefault();
      const [dr, dc] = directions[event.key];
      if (ctrl && event.key.startsWith("Arrow")) {
        navigation = navigation.then(() =>
          run(() => jumpBoundary(dr, dc, event.shiftKey)),
        );
      } else
        void run(() =>
          move(
            dr,
            dc,
            event.shiftKey && event.key !== "Tab" && event.key !== "Enter",
          ),
        );
      return;
    }
    if (event.key === "Home") {
      event.preventDefault();
      void run(() => select(sheetId, address(ctrl ? 0 : selection.r1, 0)));
      return;
    }
    if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      void run(() => action("clear"));
      return;
    }
    if (event.key === "F2") {
      event.preventDefault();
      void run(() => beginEdit());
      return;
    }
    if (!ctrl && !event.altKey && event.key.length === 1) {
      event.preventDefault();
      void run(() => beginEdit(event.key));
    }
  });
  const formulaHelp = createFormulaHelp(
    shell,
    [input, formulaInput],
    locale,
    signal,
    (source) => {
      if (source === input) formulaInput.value = input.value;
      else if (isEditing) input.value = formulaInput.value;
    },
  );
  on(canvas, "compositionstart", () => {
    if (!options.readOnly) void run(() => beginEdit(""));
  });
  for (const source of [input, formulaInput]) {
    on(source, "compositionstart", () => {
      composing = true;
    });
    on(source, "compositionend", () => {
      composing = false;
    });
    on(source, "keydown", (event) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
        event.preventDefault();
        event.stopPropagation();
        void run(() => action(options.persistence ? "saveNow" : "saveXlsx"));
        return;
      }
      if (event.key === "F1") {
        event.preventDefault();
        void run(() => action("shortcuts"));
        return;
      }
      if (
        event.key !== "F4" ||
        composing ||
        event.isComposing ||
        options.readOnly
      )
        return;
      const next = cycleReference(source.value, source.selectionStart ?? 0);
      if (!next) return;
      event.preventDefault();
      input.value = formulaInput.value = next.value;
      source.setSelectionRange(next.caret, next.caret);
      formulaHelp.update(source);
    });
    on(source, "input", () => {
      referenceRange = undefined;
      schedule();
    });
    on(source, "blur", (event) => {
      if (referenceDrag) return;
      if (event.relatedTarget === input || event.relatedTarget === formulaInput)
        return;
      if (source === formulaInput && isEditing)
        input.value = formulaInput.value;
      void run(commitEdit);
    });
  }
  on(input, "input", () => {
    formulaInput.value = input.value;
  });
  on(input, "keydown", (event) => {
    if (event.isComposing || composing || event.keyCode === 229) return;
    if (formulaHelp.keydown(event)) return;
    if (event.key === "Escape") {
      event.preventDefault();
      endEdit();
      canvas.focus();
      updateFormulaBar();
    } else if (
      (event.key === "Enter" && !event.altKey) ||
      event.key === "Tab"
    ) {
      event.preventDefault();
      void run(async () => {
        const commit = commitEdit();
        const movement = move(
          event.key === "Enter" ? (event.shiftKey ? -1 : 1) : 0,
          event.key === "Tab" ? (event.shiftKey ? -1 : 1) : 0,
        );
        canvas.focus();
        await Promise.all([commit, movement]);
      });
    }
  });
  on(formulaInput, "focus", () => {
    if (beginDraft()) {
      formulaInput.value = input.value;
      formulaHelp.update(formulaInput);
    }
  });
  on(formulaInput, "input", () => {
    const value = formulaInput.value;
    if (beginDraft()) input.value = value;
  });
  on(formulaInput, "keydown", (event) => {
    if (
      options.readOnly ||
      event.isComposing ||
      composing ||
      event.keyCode === 229
    )
      return;
    if (formulaHelp.keydown(event)) return;
    if (event.key === "Enter") {
      event.preventDefault();
      if (isEditing) input.value = formulaInput.value;
      void run(async () => {
        const commit = commitEdit();
        canvas.focus();
        await commit;
      });
    }
    if (event.key === "Escape") {
      event.preventDefault();
      endEdit();
      canvas.focus();
      updateFormulaBar();
    }
  });
  on(addressInput, "blur", updateFormulaBar);
  on(addressInput, "keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      void run(async () => {
        await select(sheetId, addressInput.value);
        if (!destroyed && root.activeElement === addressInput) canvas.focus();
      });
    }
  });
  for (const type of ["copy", "cut"] as const)
    on(root, type, (event) => {
      if (
        root.activeElement?.matches("input,textarea,[contenteditable=true]") ||
        root.activeElement?.closest(".dialog")
      )
        return;
      if (type === "cut" && options.readOnly) return;
      event.preventDefault();
      void run(() => {
        const data = makeClipboard(type === "cut");
        event.clipboardData?.setData("text/plain", data.text);
        event.clipboardData?.setData("text/html", data.html);
      });
    });
  on(root, "paste", (event) => {
    if (
      root.activeElement?.matches("input,textarea,[contenteditable=true]") ||
      root.activeElement?.closest(".dialog")
    )
      return;
    event.preventDefault();
    void run(() =>
      pasteText(
        event.clipboardData?.getData("text/plain") ?? "",
        event.clipboardData?.getData("text/html"),
      ),
    );
  });
  addSheet.onclick = () => run(() => action("addSheet"));
  const updateDiagnostics = (count: number) =>
    diagnosticsButton.replaceChildren(
      icon("diagnostics"),
      element("span", "", t.diagnostics + (count ? " · " + count : "")),
    );
  updateDiagnostics(0);
  diagnosticsButton.className = "tool-button";
  diagnosticsButton.onclick = () => run(() => action("diagnostics"));
  unsubscribers.push(
    workbook.on("change", () => {
      void run(refresh);
    }),
    workbook.on("calculation", (event) => {
      status.textContent =
        event.status === "idle"
          ? options.readOnly
            ? t.readonly
            : t.ready
          : t.calculating;
    }),
    workbook.on("progress", (event) => {
      const stage: Record<string, string> = {
        apply: t.progressApply,
        calculate: t.progressCalculate,
        unzip: t.progressUnzip,
        worksheets: t.progressWorksheets,
      };
      status.textContent = `${stage[event.stage] ?? t.processing} ${Math.round(event.progress * 100)}%`;
    }),
    workbook.on("diagnostics", (items) => {
      updateDiagnostics(items.length);
    }),
  );
  const observer = new ResizeObserver(resize);
  observer.observe(viewport);
  const ready = Promise.all([
    refresh(),
    workbook.getDiagnostics().then((items) => {
      if (!destroyed) updateDiagnostics(items.length);
    }),
    workbook.getFunctions().then((catalog) => formulaHelp.setCatalog(catalog)),
  ])
    .then(async () => {
      if (destroyed) return;
      if (restored) {
        anchor = { row: selection.r1, col: selection.c1 };
        focus = { ...anchor };
        scroll.scrollTop = restored.top;
        scroll.scrollLeft = restored.left;
        await fetchVisible();
      }
      status.textContent = options.readOnly ? t.readonly : t.ready;
    })
    .catch((error) => {
      notify(errorMessage(error as Error, locale));
      options.onError?.(error as Error);
      throw error;
    });
  return {
    captureView: () => {
      if (operationController) throw new Error(t.fileBusy);
      return {
        zoom,
        sheetId,
        selection: { ...selection },
        top: scroll.scrollTop,
        left: scroll.scrollLeft,
        sheets: [...sheetViews],
      };
    },
    detach: () => cleanup(true),
    ready,
    select,
    getSelection: () => ({ sheetId, range: { ...selection } }),
    resize,
    getZoom: () => zoom,
    setZoom,
    getEditState: () => ({
      editing: isEditing,
      dirty:
        !!commitFailure || !!(editTarget && input.value !== editTarget.initial),
      pending: !!pendingCommit,
    }),
    commitEdit: flushEdit,
    cancelEdit,
    destroy,
  };
  function destroy(): void;
  function destroy(options: { commit: true }): Promise<void>;
  function destroy(options?: { commit: true }): void | Promise<void> {
    if (!options?.commit) {
      cleanup();
      return;
    }
    if (destroyed) return Promise.resolve();
    if (closing) return closing;
    const commit = flushEdit();
    shell.inert = true;
    closing = commit
      .then(() => cleanup())
      .finally(() => {
        closing = undefined;
        shell.inert = false;
      });
    return closing;
  }
  function cleanup(preserveClipboard = false) {
    if (destroyed) return;
    destroyed = true;
    controller.abort();
    operationController?.abort();
    observer.disconnect();
    unsubscribers.forEach((fn) => fn());
    cancelAnimationFrame(frame);
    clearTimeout(toastTimer);
    host.remove();
    cells.clear();
    selectionData = [];
    if (!preserveClipboard && clipboard?.book === options.workbook)
      clipboard = undefined;
  }
}

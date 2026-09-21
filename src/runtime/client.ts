import {
  writeRecords,
  readRecords,
  readChunks,
  writeChunks,
  type RecordColumn,
  type RecordOptions,
  type ReadChunkOptions,
  type WriteChunksResult,
} from "./records";
import { WorkbookError } from "./errors";
import type { listFunctions } from "../formula/functions";
import { keyOf } from "../core/address";
import { parseCell, parseRange } from "./arguments";
import type {
  CellStyle,
  ChangeEvent,
  Command,
  Diagnostic,
  FilterRule,
  InputValue,
  OperationOptions,
  ProgressEvent,
  Rect,
  Region,
  SortKey,
  WorkbookSnapshot,
} from "../core/types";
import type { WorkbookModel } from "../core/model";
export interface WorkbookEvents {
  change: ChangeEvent;
  selection: { sheetId: string; range: Rect };
  calculation: {
    status: "calculating" | "idle";
    operationId?: string;
    origin?: string;
  };
  progress: ProgressEvent;
  diagnostics: Diagnostic[];
  error: Error;
}
type Metadata = ReturnType<WorkbookModel["metadata"]>;
interface Pending {
  resolve(value: unknown): void;
  reject(reason: unknown): void;
  cleanup(): void;
  operation: string;
  operationId: string;
}
export class Workbook {
  private nextId = 1;
  private pending = new Map<number, Pending>();
  private listeners = new Map<
    keyof WorkbookEvents,
    Set<(value: any) => void>
  >();
  private disposed = false;
  constructor(
    private worker: Worker,
    private defaults: { requestTimeout?: number } = {},
  ) {
    worker.onmessage = ({ data }) => {
      if (!data || typeof data !== "object")
        return this.fail("PROTOCOL_MISMATCH", "Malformed Worker response");
      if (data.event) {
        if (!validEvent(data.event, data.data))
          return this.fail(
            "PROTOCOL_MISMATCH",
            "Unknown or malformed Worker event",
          );
        this.emit(data.event, data.data);
        return;
      }
      if (
        !Number.isSafeInteger(data.id) ||
        (!("result" in data) && !("error" in data))
      )
        return this.fail("PROTOCOL_MISMATCH", "Malformed Worker response");
      const pending = this.pending.get(data.id);
      if (!pending) return;
      this.pending.delete(data.id);
      pending.cleanup();
      if (data.error) {
        if (
          typeof data.error.message !== "string" ||
          typeof data.error.code !== "string"
        ) {
          pending.reject(
            new WorkbookError("PROTOCOL_MISMATCH", "Malformed Worker error", {
              operation: pending.operation,
              operationId: pending.operationId,
              outcome: "unknown",
            }),
          );
          return this.fail("PROTOCOL_MISMATCH", "Malformed Worker error");
        }
        pending.reject(
          Object.assign(
            new WorkbookError(data.error.code, data.error.message),
            data.error,
          ),
        );
      } else pending.resolve(data.result);
    };
    worker.onerror = (event) =>
      this.fail("WORKER_FAILED", event.message || "Workbook Worker failed");
    worker.onmessageerror = () =>
      this.fail("SERIALIZATION_FAILED", "Cannot deserialize Worker response");
  }
  private fail(
    code:
      | "WORKER_FAILED"
      | "PROTOCOL_MISMATCH"
      | "SERIALIZATION_FAILED"
      | "TIMEOUT",
    message: string,
  ): void {
    if (this.disposed) return;
    const error = new WorkbookError(code, message, { outcome: "unknown" });
    for (const item of this.pending.values()) {
      item.cleanup();
      item.reject(
        new WorkbookError(code, message, {
          operation: item.operation,
          operationId: item.operationId,
          outcome: "unknown",
        }),
      );
    }
    this.pending.clear();
    this.emit("error", error);
    void this.dispose();
  }
  /** Returns a facade sharing this workbook; options attach to its requests without mutable global context. */
  withOptions(options: OperationOptions): Workbook {
    const root = this;
    let facade: Workbook;
    facade = new Proxy(this, {
      get(target, key) {
        if (key === "request")
          return (
            operation: string,
            args?: unknown,
            local?: OperationOptions,
            transfer?: Transferable[],
          ) =>
            root.request(operation, args, { ...options, ...local }, transfer);
        if (key === "dispose") return root.dispose.bind(root);
        const value = Reflect.get(target, key);
        return typeof value === "function" ? value.bind(facade) : value;
      },
    });
    return facade;
  }
  on<K extends keyof WorkbookEvents>(
    event: K,
    callback: (value: WorkbookEvents[K]) => void,
  ): () => void {
    if (!this.listeners.has(event)) this.listeners.set(event, new Set());
    this.listeners.get(event)!.add(callback);
    return () => this.listeners.get(event)?.delete(callback);
  }
  /** Selection is a view event; it does not change workbook contents or history. */
  emit<K extends keyof WorkbookEvents>(
    event: K,
    value: WorkbookEvents[K],
  ): void {
    for (const callback of this.listeners.get(event) ?? []) {
      try {
        callback(value);
      } catch (error) {
        console.error("OnlineExcel event listener failed", error);
      }
    }
  }
  request<T>(
    operation: string,
    args: unknown = {},
    options: OperationOptions = {},
    transfer: Transferable[] = [],
  ): Promise<T> {
    const operationId = options.operationId ?? crypto.randomUUID();
    const context = {
      operation,
      operationId,
      outcome: "not-executed" as const,
    };
    if (this.disposed)
      return Promise.reject(
        new WorkbookError("DISPOSED", "Workbook is disposed", context),
      );
    if (options.signal?.aborted)
      return Promise.reject(
        new WorkbookError("CANCELLED", "Operation cancelled", context),
      );
    const timeout = options.timeout ?? this.defaults.requestTimeout ?? 0;
    if (!Number.isFinite(timeout) || timeout < 0)
      return Promise.reject(
        new WorkbookError("INVALID_ARGUMENT", "Invalid timeout", context),
      );
    const id = this.nextId++;
    return new Promise<T>((resolve, reject) => {
      let timer: ReturnType<typeof setTimeout> | undefined;
      const cancel = () => {
        try {
          this.worker.postMessage({ id, operation: "cancel" });
        } catch {
          this.fail("WORKER_FAILED", "Cannot cancel Worker operation");
        }
      };
      const cleanup = () => {
        options.signal?.removeEventListener("abort", cancel);
        clearTimeout(timer);
      };
      this.pending.set(id, {
        resolve: resolve as (value: unknown) => void,
        reject,
        cleanup,
        operation,
        operationId,
      });
      options.signal?.addEventListener("abort", cancel, { once: true });
      if (timeout)
        timer = setTimeout(
          () =>
            this.fail("TIMEOUT", "Worker request timed out; workbook closed"),
          timeout,
        );
      try {
        this.worker.postMessage(
          { id, operation, args, operationId, origin: options.origin ?? "api" },
          transfer,
        );
      } catch (error) {
        this.pending.delete(id);
        cleanup();
        reject(
          new WorkbookError("SERIALIZATION_FAILED", String(error), context),
        );
      }
    });
  }

  async getMetadata(): Promise<Metadata> {
    return this.request("metadata");
  }
  async getRegion(sheetId: string, range: string | Rect): Promise<Region> {
    return this.request("region", { sheetId, range: parseRange(range) });
  }
  async getValues(
    sheetId: string,
    range: string | Rect,
  ): Promise<import("../core/types").Scalar[][]> {
    const result = await this.getRegion(sheetId, range),
      r = result.range;
    const values = Array.from({ length: r.r2 - r.r1 + 1 }, () =>
      Array(r.c2 - r.c1 + 1).fill(null),
    );
    for (const cell of result.cells)
      values[cell.row - r.r1][cell.column - r.c1] = cell.value;
    return values;
  }
  async transaction(
    commands: Command[],
    options?: OperationOptions,
  ): Promise<ChangeEvent> {
    return this.request("commands", { commands }, options);
  }
  async setValues(
    sheetId: string,
    range: string | Rect,
    values: InputValue[][],
    options?: OperationOptions & { parseFormulas?: boolean },
  ): Promise<ChangeEvent> {
    return this.transaction(
      [
        {
          type: "setValues",
          sheetId,
          range: parseRange(range),
          values,
          parseFormulas: options?.parseFormulas,
        },
      ],
      options,
    );
  }
  async setFormula(
    sheetId: string,
    cell: string,
    formula: string,
  ): Promise<ChangeEvent> {
    const position = parseCell(cell);
    return this.transaction([
      { type: "setFormula", sheetId, ...position, formula },
    ]);
  }
  async setStyle(
    sheetId: string,
    range: string | Rect,
    style: CellStyle,
  ): Promise<ChangeEvent> {
    return this.transaction([
      { type: "style", sheetId, range: parseRange(range), style },
    ]);
  }
  async clear(
    sheetId: string,
    range: string | Rect,
    formats = false,
  ): Promise<ChangeEvent> {
    return this.transaction([
      { type: "clear", sheetId, range: parseRange(range), formats },
    ]);
  }
  async addSheet(
    name: string,
    options?: { rows?: number; columns?: number },
  ): Promise<string> {
    const id = crypto.randomUUID();
    await this.transaction([{ type: "addSheet", id, name, ...options }]);
    return id;
  }
  async deleteSheet(sheetId: string): Promise<ChangeEvent> {
    return this.transaction([{ type: "deleteSheet", sheetId }]);
  }
  async renameSheet(sheetId: string, name: string): Promise<ChangeEvent> {
    return this.transaction([{ type: "renameSheet", sheetId, name }]);
  }
  async reorderSheet(sheetId: string, index: number): Promise<ChangeEvent> {
    return this.transaction([{ type: "reorderSheet", sheetId, index }]);
  }
  async setDimensions(
    sheetId: string,
    axis: "row" | "column",
    indexes: number[],
    options: { size?: number; hidden?: boolean },
  ): Promise<ChangeEvent> {
    return this.transaction([
      { type: "dimensions", sheetId, axis, indexes, ...options },
    ]);
  }
  async insertRows(
    sheetId: string,
    index: number,
    count = 1,
  ): Promise<ChangeEvent> {
    return this.transaction([
      { type: "structure", sheetId, axis: "row", index, count },
    ]);
  }
  async deleteRows(
    sheetId: string,
    index: number,
    count = 1,
  ): Promise<ChangeEvent> {
    return this.transaction([
      { type: "structure", sheetId, axis: "row", index, count, delete: true },
    ]);
  }
  async insertColumns(
    sheetId: string,
    index: number,
    count = 1,
  ): Promise<ChangeEvent> {
    return this.transaction([
      { type: "structure", sheetId, axis: "column", index, count },
    ]);
  }
  async deleteColumns(
    sheetId: string,
    index: number,
    count = 1,
  ): Promise<ChangeEvent> {
    return this.transaction([
      {
        type: "structure",
        sheetId,
        axis: "column",
        index,
        count,
        delete: true,
      },
    ]);
  }
  async merge(sheetId: string, range: string | Rect): Promise<ChangeEvent> {
    return this.transaction([
      { type: "merge", sheetId, range: parseRange(range) },
    ]);
  }
  async unmerge(sheetId: string, range: string | Rect): Promise<ChangeEvent> {
    return this.transaction([
      { type: "merge", sheetId, range: parseRange(range), unmerge: true },
    ]);
  }
  async freeze(
    sheetId: string,
    rows: number,
    columns: number,
  ): Promise<ChangeEvent> {
    return this.transaction([{ type: "freeze", sheetId, rows, columns }]);
  }
  async sort(
    sheetId: string,
    range: string | Rect,
    keys: SortKey[],
    header = false,
  ): Promise<ChangeEvent> {
    return this.transaction([
      { type: "sort", sheetId, range: parseRange(range), keys, header },
    ]);
  }
  async filter(
    sheetId: string,
    range: string | Rect | undefined,
    rules: FilterRule[],
  ): Promise<ChangeEvent> {
    return this.transaction([
      {
        type: "filter",
        sheetId,
        range: range ? parseRange(range) : undefined,
        rules,
      },
    ]);
  }
  async copyRange(
    sheetId: string,
    range: string | Rect,
    targetSheetId: string,
    target: string,
    options?: { cut?: boolean; valuesOnly?: boolean },
  ): Promise<ChangeEvent> {
    const cell = parseCell(target);
    return this.transaction([
      {
        type: "copy",
        sheetId,
        range: parseRange(range),
        targetSheetId,
        targetRow: cell.row,
        targetColumn: cell.column,
        ...options,
      },
    ]);
  }
  async fill(
    sheetId: string,
    source: string | Rect,
    target: string | Rect,
  ): Promise<ChangeEvent> {
    return this.transaction([
      {
        type: "fill",
        sheetId,
        source: parseRange(source),
        target: parseRange(target),
      },
    ]);
  }
  async duplicateSheet(sheetId: string, name: string): Promise<string> {
    const id = crypto.randomUUID();
    await this.transaction([{ type: "duplicateSheet", sheetId, name, id }]);
    return id;
  }
  setSheetHidden(sheetId: string, hidden: boolean) {
    return this.transaction([{ type: "sheetVisibility", sheetId, hidden }]);
  }
  protectSheet(sheetId: string, enabled = true) {
    return this.transaction([{ type: "protect", sheetId, enabled }]);
  }
  async setValidation(
    sheetId: string,
    range: string | Rect,
    rule?: Omit<import("../core/types").ValidationRule, "range">,
  ) {
    return this.transaction([
      { type: "validation", sheetId, range: parseRange(range), rule },
    ]);
  }
  async setConditionalFormat(
    sheetId: string,
    range: string | Rect,
    rule?: Omit<import("../core/types").ConditionalRule, "range">,
  ) {
    return this.transaction([
      { type: "conditionalFormat", sheetId, range: parseRange(range), rule },
    ]);
  }
  renameName(name: string, newName: string) {
    return this.transaction([{ type: "renameName", name, newName }]);
  }
  deleteName(name: string) {
    return this.transaction([{ type: "deleteName", name }]);
  }
  async defineName(
    name: string,
    sheetId: string,
    range: string | Rect,
  ): Promise<ChangeEvent> {
    return this.transaction([
      { type: "defineName", name, sheetId, range: parseRange(range) },
    ]);
  }
  async getNavigationTarget(
    sheetId: string,
    cell: string,
    direction: "up" | "down" | "left" | "right",
  ): Promise<{ row: number; column: number }> {
    return this.request("navigationTarget", {
      sheetId,
      ...parseCell(cell),
      direction,
    });
  }
  async getDataRegion(sheetId: string, range: string | Rect): Promise<Rect> {
    return this.request("dataRegion", { sheetId, range: parseRange(range) });
  }
  async getDistinctValues(
    sheetId: string,
    range: string | Rect,
    column: number,
  ): Promise<{ values: string[]; truncated: boolean }> {
    return this.request("distinctValues", {
      sheetId,
      range: parseRange(range),
      column,
    });
  }
  async findNext(
    sheetId: string,
    search: string,
    options: { after?: string; matchCase?: boolean; entireCell?: boolean } = {},
  ): Promise<{
    address: string;
    value: import("../core/types").Scalar;
  } | null> {
    const position = options.after ? parseCell(options.after) : undefined;
    return this.request("findNext", {
      sheetId,
      search,
      ...options,
      after: position ? keyOf(position.row, position.column) : undefined,
    });
  }
  async find(
    sheetId: string,
    search: string,
    matchCase = false,
  ): Promise<{ address: string; value: unknown }[]> {
    return this.request("find", { sheetId, search, matchCase });
  }
  async replace(
    sheetId: string,
    search: string,
    replacement: string,
    options?: { matchCase?: boolean; entireCell?: boolean },
  ): Promise<ChangeEvent> {
    return this.transaction([
      { type: "replace", sheetId, search, replacement, ...options },
    ]);
  }
  async undo(): Promise<ChangeEvent | null> {
    return this.request("undo");
  }
  async redo(): Promise<ChangeEvent | null> {
    return this.request("redo");
  }
  setRecords(
    sheetId: string,
    records: readonly Record<string, InputValue>[],
    columns: readonly RecordColumn[],
    options?: RecordOptions,
  ): Promise<ChangeEvent | null> {
    return writeRecords(this, sheetId, records, columns, options);
  }
  readRecords(
    sheetId: string,
    range: string | Rect,
    columns: readonly RecordColumn[],
    options?: ReadChunkOptions,
  ) {
    return readRecords(this, sheetId, range, columns, options);
  }
  readChunks(
    sheetId: string,
    range: string | Rect,
    options?: ReadChunkOptions,
  ) {
    return readChunks(this, sheetId, range, options);
  }
  writeChunks(
    sheetId: string,
    start: string,
    source: AsyncIterable<InputValue[][]> | Iterable<InputValue[][]>,
    options?: OperationOptions,
  ): Promise<WriteChunksResult> {
    return writeChunks(this, sheetId, start, source, options);
  }
  async getFunctions(): Promise<ReturnType<typeof listFunctions>> {
    return this.request("functions");
  }
  /** Snapshot and revision captured together in the Worker queue. Persist both as one save point. */
  async createSavePoint(
    options?: OperationOptions,
  ): Promise<{ snapshot: WorkbookSnapshot; revision: number }> {
    return this.request("savePoint", {}, options);
  }
  async exportJSON(): Promise<WorkbookSnapshot> {
    return this.request("snapshot");
  }
  async importJSON(
    snapshot: WorkbookSnapshot,
    options?: OperationOptions,
  ): Promise<Metadata> {
    return this.request("importJSON", { snapshot }, options);
  }
  async importXlsx(
    input: Blob | ArrayBuffer | Uint8Array,
    options?: OperationOptions,
  ): Promise<Metadata & { diagnostics: Diagnostic[] }> {
    const data =
      input instanceof Blob
        ? new Uint8Array(await input.arrayBuffer())
        : input instanceof Uint8Array
          ? input.slice()
          : new Uint8Array(input.slice(0));
    return this.request("importXlsx", { data }, options, [data.buffer]);
  }
  async exportXlsx(
    options?: OperationOptions & { allowLossy?: boolean },
  ): Promise<{ data: Uint8Array; diagnostics: Diagnostic[] }> {
    return this.request(
      "exportXlsx",
      { allowLossy: options?.allowLossy },
      options,
    );
  }
  async importCsv(
    text: string,
    options?: OperationOptions & {
      sheetId?: string;
      start?: string;
      delimiter?: string;
      columns?: import("../core/input").ColumnType[];
      header?: boolean;
    },
  ): Promise<{ diagnostics: Diagnostic[] }> {
    return this.request(
      "importCsv",
      {
        text,
        sheetId: options?.sheetId,
        start: options?.start,
        delimiter: options?.delimiter,
        columns: options?.columns,
        header: options?.header,
      },
      options,
    );
  }
  async exportCsv(
    sheetId: string,
    options?: OperationOptions & { range?: string | Rect; delimiter?: string },
  ): Promise<{ text: string; diagnostics: Diagnostic[] }> {
    return this.request(
      "exportCsv",
      {
        sheetId,
        range: options?.range ? parseRange(options.range) : undefined,
        delimiter: options?.delimiter,
      },
      options,
    );
  }
  async getDiagnostics(): Promise<Diagnostic[]> {
    return this.request("diagnostics");
  }
  async dispose(): Promise<void> {
    if (this.disposed) return;
    this.disposed = true;
    this.worker.terminate();
    this.worker.onmessage =
      this.worker.onerror =
      this.worker.onmessageerror =
        null;
    for (const p of this.pending.values()) {
      p.cleanup();
      p.reject(
        new WorkbookError("DISPOSED", "Workbook is disposed", {
          operation: p.operation,
          operationId: p.operationId,
          outcome: "unknown",
        }),
      );
    }
    this.pending.clear();
    this.listeners.clear();
  }
}

function validEvent(event: unknown, data: any): boolean {
  switch (event) {
    case "change":
      return (
        !!data &&
        Number.isSafeInteger(data.revision) &&
        Array.isArray(data.changes) &&
        ["edit", "undo", "redo", "import"].includes(data.source)
      );
    case "calculation":
      return !!data && ["calculating", "idle"].includes(data.status);
    case "progress":
      return (
        !!data &&
        typeof data.operation === "string" &&
        typeof data.stage === "string" &&
        Number.isFinite(data.progress)
      );
    case "diagnostics":
      return Array.isArray(data);
    default:
      return false;
  }
}

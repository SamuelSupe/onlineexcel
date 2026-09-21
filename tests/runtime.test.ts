import { afterEach, describe, expect, it, vi } from "vitest";
import { Workbook } from "../src/runtime/client";
import { initializeWorkbook } from "../src/runtime/create";
import { WorkbookError } from "../src/runtime/errors";
import { PROTOCOL_VERSION, LIBRARY_VERSION } from "../src/runtime/protocol";
import { WorkbookModel } from "../src/core/model";
import { registerFunctions } from "../src/runtime/extensions";
import { functions } from "../src/formula/functions";
import { parseRange } from "../src/core/address";
import {
  createPersistence,
  type SavedWorkbook,
} from "../src/runtime/persistence";
class FakeWorker {
  onmessage: ((event: { data: any }) => void) | null = null;
  onerror: ((event: { message: string }) => void) | null = null;
  onmessageerror: (() => void) | null = null;
  terminated = false;
  sent: any[] = [];
  response?: (message: any) => unknown;
  postMessage(message: any) {
    this.sent.push(message);
    if (this.response)
      queueMicrotask(() => {
        try {
          this.onmessage?.({
            data: { id: message.id, result: this.response!(message) },
          });
        } catch (error) {
          this.onmessage?.({
            data: {
              id: message.id,
              error: {
                code: "INVALID_ARGUMENT",
                message: (error as Error).message,
              },
            },
          });
        }
      });
  }
  terminate() {
    this.terminated = true;
  }
  asWorker() {
    return this as unknown as Worker;
  }
}
afterEach(() => {
  vi.useRealTimers();
  functions.delete("SDK.DOUBLE");
});
describe("Optional persistence", () => {
  function workbook() {
    let model = new WorkbookModel({
      sheets: [{ name: "Draft", rows: 100, columns: 10 }],
    });
    const worker = new FakeWorker(),
      book = new Workbook(worker.asWorker());
    worker.response = ({ operation, args }) => {
      if (operation === "metadata") return model.metadata();
      if (operation === "savePoint")
        return { snapshot: model.snapshot(), revision: model.revision };
      if (operation === "commands") {
        const change = model.execute(args.commands);
        worker.onmessage?.({ data: { event: "change", data: change } });
        return change;
      }
      if (operation === "importJSON") {
        const next = new WorkbookModel({ snapshot: args.snapshot });
        next.revision = model.revision + 1;
        model = next;
        worker.onmessage?.({
          data: {
            event: "change",
            data: { revision: model.revision, changes: [], source: "import" },
          },
        });
        return model.metadata();
      }
      throw new Error("Unexpected operation " + operation);
    };
    return {
      book,
      worker,
      id: model.sheets[0].meta.id,
      snapshot: () => model.snapshot(),
    };
  }
  it("serializes saves and persists edits made while an earlier write is pending", async () => {
    vi.useFakeTimers();
    const { book, id } = workbook();
    let release!: () => void;
    const stored: SavedWorkbook[] = [];
    const save = vi.fn(async (_key: string, value: SavedWorkbook) => {
      stored.push(value);
      if (stored.length === 1)
        await new Promise<void>((resolve) => {
          release = resolve;
        });
    });
    const persistence = createPersistence(book, {
      key: "doc",
      debounceMs: 20,
      storage: { load: async () => undefined, save },
    });
    try {
      await persistence.ready;
      await book.setValues(id, "A1", [[1]]);
      await vi.advanceTimersByTimeAsync(20);
      expect(save).toHaveBeenCalledTimes(1);
      await book.setValues(id, "A1", [[2]]);
      await vi.advanceTimersByTimeAsync(100);
      expect(save).toHaveBeenCalledTimes(1);
      release();
      await vi.advanceTimersByTimeAsync(25);
      expect(save).toHaveBeenCalledTimes(2);
      expect(
        new WorkbookModel({ snapshot: stored[0].snapshot }).region(
          id,
          parseRange("A1"),
        ).cells[0].value,
      ).toBe(1);
      expect(
        new WorkbookModel({ snapshot: stored[1].snapshot }).region(
          id,
          parseRange("A1"),
        ).cells[0].value,
      ).toBe(2);
      expect(persistence.hasUnsavedChanges()).toBe(false);
    } finally {
      persistence.dispose();
      await book.dispose();
    }
  });
  it("keeps the stored draft until explicit recovery and cancels bound drafts only after valid import", async () => {
    vi.useFakeTimers();
    const { book, id, snapshot } = workbook();
    await book.setValues(id, "A1", [[42]]);
    const record: SavedWorkbook = {
      version: 1,
      savedAt: 1,
      snapshot: snapshot(),
    };
    await book.setValues(id, "A1", [[99]]);
    const save = vi.fn(async () => {}),
      cancel = vi.fn();
    const persistence = createPersistence(book, {
      key: "doc",
      storage: { load: async () => record, save },
    });
    persistence.bindEditor({
      commit: async () => {},
      dirty: () => false,
      cancel,
    });
    try {
      await persistence.ready;
      await vi.advanceTimersByTimeAsync(5000);
      expect(save).not.toHaveBeenCalled();
      await expect(persistence.save()).rejects.toThrow(/recovery/);
      const valid = record.snapshot;
      record.snapshot = {} as typeof valid;
      await expect(persistence.restore()).rejects.toThrow();
      expect(cancel).not.toHaveBeenCalled();
      expect(snapshot()).not.toEqual(valid);
      record.snapshot = valid;
      await persistence.restore();
      expect(cancel).toHaveBeenCalledOnce();
      expect(snapshot()).toEqual(valid);
      expect(persistence.getState().status).toBe("saved");
    } finally {
      persistence.dispose();
      await book.dispose();
    }
  });
  it("bounds retries, keeps failures unsaved and allows a manual retry with the latest data", async () => {
    vi.useFakeTimers();
    const { book, id } = workbook();
    const save = vi.fn(async (): Promise<void> => {
      throw new Error("offline");
    });
    const persistence = createPersistence(book, {
      key: "doc",
      debounceMs: 5,
      retryDelayMs: 5,
      maxRetries: 2,
      storage: { load: async () => undefined, save },
    });
    try {
      await persistence.ready;
      await vi.advanceTimersByTimeAsync(100);
      expect(save).toHaveBeenCalledTimes(3);
      expect(persistence.getState().status).toBe("error");
      await book.setValues(id, "A1", [[8]]);
      await vi.advanceTimersByTimeAsync(100);
      expect(save).toHaveBeenCalledTimes(3);
      expect(persistence.hasUnsavedChanges()).toBe(true);
      save.mockImplementation(async () => {});
      await persistence.retry();
      expect(persistence.hasUnsavedChanges()).toBe(false);
    } finally {
      persistence.dispose();
      await book.dispose();
    }
  });
  it("does not overwrite after a load failure and aborts pending storage when disposed", async () => {
    vi.useFakeTimers();
    const { book } = workbook();
    const load = vi.fn(async (): Promise<SavedWorkbook | undefined> => {
      throw new Error("unavailable");
    });
    let signal!: AbortSignal;
    const save = vi.fn(
      async (_key: string, _record: SavedWorkbook, current: AbortSignal) => {
        signal = current;
        await new Promise<void>((_resolve, reject) =>
          current.addEventListener(
            "abort",
            () => reject(new DOMException("Cancelled", "AbortError")),
            { once: true },
          ),
        );
      },
    );
    const persistence = createPersistence(book, {
      key: "doc",
      storage: { load, save },
    });
    try {
      await expect(persistence.ready).rejects.toThrow("unavailable");
      await expect(persistence.save()).rejects.toThrow("unavailable");
      expect(save).not.toHaveBeenCalled();
      load.mockResolvedValue(undefined);
      await persistence.retry();
      const pending = expect(persistence.save()).rejects.toMatchObject({
        name: "AbortError",
      });
      await vi.advanceTimersByTimeAsync(0);
      persistence.dispose();
      await pending;
      expect(signal.aborted).toBe(true);
    } finally {
      persistence.dispose();
      await book.dispose();
    }
  });
});
describe("SDK Worker boundary", () => {
  it("uses the factory, negotiates versions and rejects stale Worker assets", async () => {
    const worker = new FakeWorker();
    worker.response = () => ({
      protocolVersion: PROTOCOL_VERSION,
      libraryVersion: LIBRARY_VERSION,
    });
    const book = await initializeWorkbook({
      workerFactory: () => worker.asWorker(),
    });
    expect(worker.sent[0].args.workerFactory).toBeUndefined();
    await book.dispose();
    const stale = new FakeWorker();
    stale.response = () => ({ protocolVersion: 0 });
    await expect(
      initializeWorkbook({ workerFactory: () => stale.asWorker() }),
    ).rejects.toMatchObject({ code: "PROTOCOL_MISMATCH" });
    expect(stale.terminated).toBe(true);
  });
  it("times out initialization and closes a silent Worker", async () => {
    vi.useFakeTimers();
    const worker = new FakeWorker();
    const pending = expect(
      initializeWorkbook({
        workerFactory: () => worker.asWorker(),
        initializationTimeout: 20,
      }),
    ).rejects.toMatchObject({
      code: "TIMEOUT",
      operation: "init",
      outcome: "unknown",
    });
    await vi.advanceTimersByTimeAsync(21);
    await pending;
    expect(worker.terminated).toBe(true);
  });
  it("does not retry an uncertain mutation and rejects all queued work on timeout", async () => {
    vi.useFakeTimers();
    const worker = new FakeWorker(),
      book = new Workbook(worker.asWorker());
    const mutation = expect(
      book.setValues("s", "A1", [[1]], { timeout: 10, operationId: "save-1" }),
    ).rejects.toMatchObject({
      code: "TIMEOUT",
      operationId: "save-1",
      outcome: "unknown",
    });
    const read = expect(book.getMetadata()).rejects.toMatchObject({
      code: "TIMEOUT",
    });
    await vi.advanceTimersByTimeAsync(11);
    await Promise.all([mutation, read]);
    expect(worker.sent).toHaveLength(2);
    await expect(book.getMetadata()).rejects.toMatchObject({
      code: "DISPOSED",
      outcome: "not-executed",
    });
  });
  it("isolates scoped request metadata under concurrent calls", async () => {
    const worker = new FakeWorker();
    worker.response = (message) => message;
    const book = new Workbook(worker.asWorker());
    await Promise.all([
      book.withOptions({ origin: "editor" }).setFormula("s", "A1", "=1"),
      book
        .withOptions({ origin: "autosave", operationId: "host-1" })
        .createSavePoint(),
    ]);
    expect(worker.sent.map((item) => item.origin)).toEqual([
      "editor",
      "autosave",
    ]);
    expect(worker.sent[1].operationId).toBe("host-1");
    expect(new Set(worker.sent.map((item) => item.id)).size).toBe(2);
    await book.dispose();
  });
  it.each(["malformed", "messageerror", "error"])(
    "settles outstanding requests on %s",
    async (kind) => {
      const worker = new FakeWorker(),
        book = new Workbook(worker.asWorker());
      const pending = expect(book.getMetadata()).rejects.toBeInstanceOf(
        WorkbookError,
      );
      if (kind === "malformed") worker.onmessage?.({ data: null });
      else if (kind === "messageerror") worker.onmessageerror?.();
      else worker.onerror?.({ message: "CSP failure" });
      await pending;
      expect(worker.terminated).toBe(true);
    },
  );
  it("rejects invalid addresses asynchronously without posting a mutation", async () => {
    const worker = new FakeWorker(),
      book = new Workbook(worker.asWorker());
    await expect(book.setValues("s", "nonsense", [[1]])).rejects.toMatchObject({
      code: "INVALID_ARGUMENT",
      outcome: "not-executed",
    });
    expect(worker.sent).toHaveLength(0);
    await book.dispose();
  });
  it("reports committed chunks on source failure and protects record keys", async () => {
    const worker = new FakeWorker();
    worker.response = () => ({ revision: 1 });
    const book = new Workbook(worker.asWorker());
    async function* source() {
      yield [[1]];
      throw new Error("network stopped");
    }
    await expect(book.writeChunks("s", "A1", source())).rejects.toMatchObject({
      details: { committed: { rows: 1, chunks: 1, revision: 1 } },
    });
    await expect(
      book.setRecords("s", [], [{ key: "id" }, { key: "id" }]),
    ).rejects.toMatchObject({ code: "INVALID_ARGUMENT" });
    await book.setRecords("s", [{}], [{ key: "toString" }]);
    expect(worker.sent.at(-1).args.commands[0].values).toEqual([[null]]);
    await book.dispose();
  });
});
describe("Worker extensions", () => {
  it("calculates custom arrays, dependency updates and error results", () => {
    registerFunctions([
      {
        name: "SDK.DOUBLE",
        minArgs: 1,
        maxArgs: 1,
        signature: "value",
        evaluate: (args) => [[Number(args[0]) * 2, 5]],
      },
    ]);
    const model = new WorkbookModel(),
      sheetId = model.sheets[0].meta.id;
    model.execute([
      {
        type: "setValues",
        sheetId,
        range: parseRange("A1:B1"),
        values: [[3, "=SDK.DOUBLE(A1)"]],
      },
    ]);
    expect(
      model
        .region(sheetId, parseRange("B1:C1"))
        .cells.map((cell) => cell.value),
    ).toEqual([6, 5]);
    model.execute([
      { type: "setValues", sheetId, range: parseRange("A1"), values: [[4]] },
    ]);
    expect(model.region(sheetId, parseRange("B1")).cells[0].value).toBe(8);
    model.execute([
      {
        type: "setValues",
        sheetId,
        range: parseRange("A1"),
        values: [["invalid"]],
      },
    ]);
    expect(model.region(sheetId, parseRange("B1")).cells[0].value).toEqual({
      error: "#VALUE!",
    });
    expect(() =>
      registerFunctions([
        {
          name: "SUM",
          minArgs: 1,
          maxArgs: 1,
          signature: "v",
          evaluate: () => 0,
        },
      ]),
    ).toThrow(WorkbookError);
  });
  it("rolls back calculated business-validation failures without changing history", async () => {
    const model = new WorkbookModel(),
      sheetId = model.sheets[0].meta.id;
    model.execute([
      {
        type: "setValues",
        sheetId,
        range: parseRange("A1:B1"),
        values: [[2, "=A1*3"]],
      },
    ]);
    const before = model.snapshot(),
      revision = model.revision;
    await expect(
      model.executeAsync(
        [
          {
            type: "setValues",
            sheetId,
            range: parseRange("A1"),
            values: [[10]],
          },
        ],
        async () => {},
        () => {
          expect(model.region(sheetId, parseRange("B1")).cells[0].value).toBe(
            30,
          );
          throw new WorkbookError("VALIDATION_FAILED", "Limit exceeded");
        },
      ),
    ).rejects.toMatchObject({ code: "VALIDATION_FAILED" });
    expect(model.snapshot()).toEqual(before);
    expect(model.revision).toBe(revision);
    model.undo();
    expect(model.region(sheetId, parseRange("A1")).cells).toHaveLength(0);
  });
});

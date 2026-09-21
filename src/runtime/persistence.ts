import type { Workbook } from "./client";
import type { WorkbookSnapshot } from "../core/types";

export interface SavedWorkbook {
  version: 1;
  savedAt: number;
  snapshot: WorkbookSnapshot;
}
export interface PersistenceStorage {
  load(key: string, signal: AbortSignal): Promise<SavedWorkbook | undefined>;
  /** Resolve only after durable storage succeeds. Saving the same key must replace its previous record. */
  save(key: string, value: SavedWorkbook, signal: AbortSignal): Promise<void>;
}
export interface PersistenceState {
  status:
    | "loading"
    | "recovery"
    | "saved"
    | "dirty"
    | "saving"
    | "retrying"
    | "error"
    | "disposed";
  revision: number;
  savedRevision: number;
  savedAt?: number;
  recoverySavedAt?: number;
  attempt: number;
  error?: Error;
}
export interface PersistenceOptions {
  key: string;
  storage: PersistenceStorage;
  debounceMs?: number;
  maxRetries?: number;
  retryDelayMs?: number;
  warnBeforeUnload?: boolean;
  /** Called for explicit save(), not background saves, so autosave cannot interrupt IME input. */
  beforeSave?: () => void | Promise<void>;
  hasDraft?: () => boolean;
}
export interface WorkbookPersistence {
  readonly workbook: Workbook;
  readonly ready: Promise<void>;
  getState(): PersistenceState;
  subscribe(listener: (state: PersistenceState) => void): () => void;
  hasUnsavedChanges(): boolean;
  bindEditor(editor: {
    commit: () => Promise<void>;
    dirty: () => boolean;
    cancel?: () => void;
  }): () => void;
  save(): Promise<void>;
  retry(): Promise<void>;
  /** Replaces the workbook only after its snapshot has passed normal import validation. */
  restore(): Promise<void>;
  /** Explicitly replaces the pending stored draft with the current workbook on the next save. */
  discardRecovery(): void;
  dispose(): void;
}

export function createPersistence(
  workbook: Workbook,
  options: PersistenceOptions,
): WorkbookPersistence {
  const debounceMs = options.debounceMs ?? 800;
  const maxRetries = options.maxRetries ?? 2;
  const retryDelayMs = options.retryDelayMs ?? 1000;
  if (
    !options.key ||
    ![debounceMs, maxRetries, retryDelayMs].every(Number.isFinite) ||
    debounceMs < 0 ||
    retryDelayMs < 0 ||
    !Number.isInteger(maxRetries) ||
    maxRetries < 0 ||
    maxRetries > 10
  )
    throw new Error("Invalid persistence options");
  const controller = new AbortController();
  const editors = new Set<{
    commit: () => Promise<void>;
    dirty: () => boolean;
    cancel?: () => void;
  }>();
  const listeners = new Set<(state: PersistenceState) => void>();
  let state: PersistenceState = {
    status: "loading",
    revision: -1,
    savedRevision: -1,
    attempt: 0,
  };
  let loaded = false,
    disposed = false,
    restoring = false;
  let initialRevision = -1;
  let recovery: SavedWorkbook | undefined;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let operation: Promise<void> | undefined;
  let fatalError: Error | undefined;
  const guard = () => {
    if (disposed) throw new DOMException("Persistence disposed", "AbortError");
    if (fatalError) throw fatalError;
  };
  const emit = (patch: Partial<PersistenceState>) => {
    if (disposed) return;
    state = { ...state, ...patch };
    for (const listener of listeners) {
      try {
        listener({ ...state });
      } catch (error) {
        console.error("OnlineExcel persistence listener failed", error);
      }
    }
  };
  const dirty = () => state.revision !== state.savedRevision;
  const cancelTimer = () => {
    clearTimeout(timer);
    timer = undefined;
  };
  function schedule() {
    cancelTimer();
    if (
      !loaded ||
      recovery ||
      disposed ||
      restoring ||
      operation ||
      !dirty() ||
      state.status === "error"
    )
      return;
    timer = setTimeout(() => {
      timer = undefined;
      void persist(false).catch(() => {});
    }, debounceMs);
  }
  const off = workbook.on("change", (event) => {
    emit({ revision: event.revision });
    if (!restoring && loaded && !recovery) {
      if (!operation && state.status !== "error") emit({ status: "dirty" });
      schedule();
    }
  });
  const offError = workbook.on("error", (error) => {
    fatalError = error;
    cancelTimer();
    controller.abort();
    emit({ status: "error", error });
  });
  async function load() {
    guard();
    emit({ status: "loading", error: undefined });
    try {
      const [stored, metadata] = await Promise.all([
        options.storage.load(options.key, controller.signal),
        workbook.getMetadata(),
      ]);
      guard();
      if (
        stored &&
        (stored.version !== 1 ||
          !Number.isFinite(stored.savedAt) ||
          !stored.snapshot)
      )
        throw new Error("Unsupported saved workbook record");
      initialRevision = metadata.revision;
      recovery = stored;
      loaded = true;
      emit({
        revision: Math.max(state.revision, metadata.revision),
        status: stored ? "recovery" : "dirty",
        recoverySavedAt: stored?.savedAt,
        error: undefined,
      });
      schedule();
    } catch (error) {
      emit({ status: "error", error: error as Error });
      throw error;
    }
  }
  function delay(ms: number) {
    return new Promise<void>((resolve, reject) => {
      const abort = () => {
        clearTimeout(wait);
        reject(new DOMException("Persistence disposed", "AbortError"));
      };
      const wait = setTimeout(() => {
        controller.signal.removeEventListener("abort", abort);
        resolve();
      }, ms);
      controller.signal.addEventListener("abort", abort, { once: true });
    });
  }
  function persist(explicit: boolean): Promise<void> {
    guard();
    if (operation)
      return operation.then(() => (explicit ? persist(true) : undefined));
    const task = async () => {
      if (!loaded) await ready;
      guard();
      if (!loaded || recovery)
        throw new Error("Resolve saved draft recovery before saving");
      cancelTimer();
      try {
        if (explicit) {
          await options.beforeSave?.();
          for (const editor of editors) await editor.commit();
        }
        guard();
        if (!dirty()) return;
        emit({ status: "saving", error: undefined, attempt: 0 });
        const point = await workbook.createSavePoint();
        guard();
        const record: SavedWorkbook = {
          version: 1,
          savedAt: Date.now(),
          snapshot: point.snapshot,
        };
        for (let attempt = 0; ; attempt++) {
          try {
            await options.storage.save(options.key, record, controller.signal);
            guard();
            emit({
              savedRevision: point.revision,
              savedAt: record.savedAt,
              error: undefined,
              attempt: 0,
            });
            emit({ status: dirty() ? "dirty" : "saved" });
            break;
          } catch (error) {
            guard();
            if (attempt >= maxRetries) throw error;
            emit({
              status: "retrying",
              attempt: attempt + 1,
              error: error as Error,
            });
            await delay(retryDelayMs * 2 ** attempt);
            guard();
            emit({ status: "saving" });
          }
        }
      } catch (error) {
        emit({ status: "error", error: error as Error });
        throw error;
      }
    };
    operation = task().finally(() => {
      operation = undefined;
      schedule();
    });
    return operation;
  }
  const ready = load();
  // Callers may observe a load error through the status UI before awaiting ready.
  void ready.catch(() => {});
  const hasUnsavedChanges = () =>
    !disposed &&
    (!!options.hasDraft?.() ||
      [...editors].some((editor) => editor.dirty()) ||
      (recovery || !loaded ? state.revision > initialRevision : dirty()));
  const beforeUnload = (event: BeforeUnloadEvent) => {
    if (hasUnsavedChanges()) {
      event.preventDefault();
      event.returnValue = "";
    }
  };
  if (options.warnBeforeUnload !== false && typeof window !== "undefined")
    window.addEventListener("beforeunload", beforeUnload);
  async function restore() {
    guard();
    if (operation || restoring)
      throw new Error("Persistence operation in progress");
    if (!recovery) return;
    cancelTimer();
    restoring = true;
    const saved = recovery;
    emit({ status: "loading", error: undefined });
    try {
      const change = await workbook.importJSON(saved.snapshot, {
        signal: controller.signal,
      });
      guard();
      for (const editor of editors) editor.cancel?.();
      recovery = undefined;
      emit({
        status: "saved",
        savedRevision: change.revision,
        revision: Math.max(state.revision, change.revision),
        savedAt: saved.savedAt,
        recoverySavedAt: undefined,
      });
      if (dirty()) emit({ status: "dirty" });
    } catch (error) {
      emit({ status: "error", error: error as Error });
      throw error;
    } finally {
      restoring = false;
      schedule();
    }
  }
  return {
    workbook,
    ready,
    getState: () => ({ ...state }),
    subscribe(listener) {
      listeners.add(listener);
      listener({ ...state });
      return () => {
        listeners.delete(listener);
      };
    },
    hasUnsavedChanges,
    bindEditor(editor) {
      guard();
      editors.add(editor);
      return () => {
        editors.delete(editor);
      };
    },
    save: () => persist(true),
    async retry() {
      if (!loaded) await load();
      else if (recovery) await restore();
      else await persist(true);
    },
    restore,
    discardRecovery() {
      guard();
      if (!recovery || restoring) return;
      recovery = undefined;
      emit({
        status: "dirty",
        recoverySavedAt: undefined,
        error: undefined,
        savedRevision: -1,
      });
      schedule();
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      cancelTimer();
      controller.abort();
      off();
      offError();
      if (typeof window !== "undefined")
        window.removeEventListener("beforeunload", beforeUnload);
      state = { ...state, status: "disposed" };
      for (const listener of listeners) {
        try {
          listener({ ...state });
        } catch (error) {
          console.error("OnlineExcel persistence listener failed", error);
        }
      }
      listeners.clear();
      editors.clear();
      recovery = undefined;
    },
  };
}

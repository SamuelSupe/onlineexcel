import type { PersistenceStorage, SavedWorkbook } from "./persistence";

/** An optional local store. Use a distinct key for each document; the host owns the key and storage choice. */
export function createIndexedDBStorage(
  database = "onlineexcel-drafts",
): PersistenceStorage {
  function access<T>(
    key: string,
    value: SavedWorkbook | undefined,
    signal: AbortSignal,
  ): Promise<T> {
    return new Promise((resolve, reject) => {
      if (signal.aborted) {
        reject(new DOMException("Storage cancelled", "AbortError"));
        return;
      }
      const open = indexedDB.open(database, 1);
      let stopped = false;
      const stop = (error: unknown) => {
        stopped = true;
        signal.removeEventListener("abort", cancelOpen);
        reject(error);
      };
      const cancelOpen = () =>
        stop(new DOMException("Storage cancelled", "AbortError"));
      signal.addEventListener("abort", cancelOpen, { once: true });
      open.onupgradeneeded = () => open.result.createObjectStore("workbooks");
      open.onerror = () => stop(open.error);
      open.onblocked = () =>
        stop(new Error("Draft storage is blocked by another browser tab"));
      open.onsuccess = () => {
        const db = open.result;
        signal.removeEventListener("abort", cancelOpen);
        if (stopped || signal.aborted) {
          db.close();
          return;
        }
        let transaction: IDBTransaction;
        const abort = () => {
          try {
            transaction.abort();
          } catch {
            /* The transaction may already have committed. */
          }
        };
        const cleanup = () => {
          signal.removeEventListener("abort", abort);
          db.close();
        };
        try {
          transaction = db.transaction(
            "workbooks",
            value ? "readwrite" : "readonly",
          );
          signal.addEventListener("abort", abort, { once: true });
          const store = transaction.objectStore("workbooks");
          const request = value ? store.put(value, key) : store.get(key);
          transaction.oncomplete = () => {
            cleanup();
            resolve(request.result as T);
          };
          transaction.onabort = transaction.onerror = () => {
            cleanup();
            reject(
              transaction.error ??
                new DOMException("Storage cancelled", "AbortError"),
            );
          };
        } catch (error) {
          cleanup();
          reject(error);
        }
      };
    });
  }
  return {
    load: (key, signal) =>
      access<SavedWorkbook | undefined>(key, undefined, signal),
    save: (key, value, signal) => access<void>(key, value, signal),
  };
}

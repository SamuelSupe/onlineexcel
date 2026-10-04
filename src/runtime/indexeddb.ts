import type { PersistenceStorage, SavedWorkbook } from "./persistence";

interface EncodedWorkbook {
  encoding: "onlineexcel-json-chunks-v1";
  record: SavedWorkbook;
  sheets: Blob[];
}
const checkCancelled = (signal: AbortSignal) => {
  if (signal.aborted) throw new DOMException("Storage cancelled", "AbortError");
};
function checkpoint(signal: AbortSignal) {
  return new Promise<void>((resolve, reject) =>
    setTimeout(() => {
      try {
        checkCancelled(signal);
        resolve();
      } catch (error) {
        reject(error);
      }
    }, 0),
  );
}
async function encode(
  value: SavedWorkbook,
  signal: AbortSignal,
): Promise<EncodedWorkbook> {
  const encoded: EncodedWorkbook = {
    encoding: "onlineexcel-json-chunks-v1",
    record: {
      ...value,
      snapshot: {
        ...value.snapshot,
        sheets: value.snapshot.sheets.map((sheet) => ({ ...sheet, cells: [] })),
      },
    },
    sheets: [],
  };
  let yielded = performance.now();
  for (const sheet of value.snapshot.sheets) {
    const parts: string[] = [];
    for (let offset = 0; offset < sheet.cells.length; offset += 4096) {
      checkCancelled(signal);
      parts.push(
        JSON.stringify(sheet.cells.slice(offset, offset + 4096)) + "\n",
      );
      if (performance.now() - yielded >= 8) {
        await checkpoint(signal);
        yielded = performance.now();
      }
    }
    encoded.sheets.push(new Blob(parts, { type: "application/x-ndjson" }));
  }
  checkCancelled(signal);
  return encoded;
}
async function decode(
  value: SavedWorkbook | EncodedWorkbook | undefined,
  signal: AbortSignal,
): Promise<SavedWorkbook | undefined> {
  if (!value || !("encoding" in value)) return value;
  if (
    value.encoding !== "onlineexcel-json-chunks-v1" ||
    value.sheets.length !== value.record.snapshot.sheets.length
  )
    throw new Error("Unsupported saved workbook encoding");
  let yielded = performance.now();
  for (let index = 0; index < value.sheets.length; index++) {
    const text = await value.sheets[index].text();
    const cells = value.record.snapshot.sheets[index].cells;
    for (
      let start = 0, end = text.indexOf("\n");
      end !== -1;
      start = end + 1, end = text.indexOf("\n", start)
    ) {
      checkCancelled(signal);
      cells.push(...JSON.parse(text.slice(start, end)));
      if (performance.now() - yielded >= 8) {
        await checkpoint(signal);
        yielded = performance.now();
      }
    }
  }
  checkCancelled(signal);
  return value.record;
}

/** An optional local store. Use a distinct key for each document; the host owns the key and storage choice. */
export function createIndexedDBStorage(
  database = "onlineexcel-drafts",
): PersistenceStorage {
  function access<T>(
    key: string,
    value: EncodedWorkbook | undefined,
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
    load: async (key, signal) =>
      decode(
        await access<SavedWorkbook | EncodedWorkbook | undefined>(
          key,
          undefined,
          signal,
        ),
        signal,
      ),
    // Blob records avoid one synchronous structured clone of every cell in IDB.put.
    save: async (key, value, signal) =>
      access<void>(key, await encode(value, signal), signal),
  };
}

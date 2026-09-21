import type * as Library from "../src/index";

export async function productChecks(
  library: typeof Library,
  equal: (actual: unknown, expected: unknown, label: string) => void,
  waitUntil: (predicate: () => boolean) => Promise<void>,
) {
  const book = await library.createWorkbook({
    sheets: [{ name: "Product", rows: 100, columns: 20 }],
  });
  const id = (await book.getMetadata()).sheets[0].id;
  const host = document.createElement("div");
  host.style.height = "640px";
  document.body.append(host);
  const errors: Error[] = [];
  const database = "onlineexcel-qa-" + crypto.randomUUID();
  const localStorage = library.createIndexedDBStorage(database);
  let storageUnavailable = false;
  const storage: Library.PersistenceStorage = {
    load: (key, signal) => localStorage.load(key, signal),
    async save(key, record, signal) {
      if (storageUnavailable)
        throw new Error("Storage temporarily unavailable");
      await localStorage.save(key, record, signal);
    },
  };
  let persistence = library.createPersistence(book, {
    key: "doc",
    storage,
    debounceMs: 30,
    maxRetries: 0,
  });
  const editor = library.mountEditor(host, {
    workbook: book,
    persistence,
    locale: "en-US",
    onError: (error) => errors.push(error),
  });
  const root = () => host.firstElementChild!.shadowRoot!;
  const grid = () => root().querySelector<HTMLCanvasElement>("canvas")!;
  const key = (value: string, modifiers: KeyboardEventInit = {}) =>
    grid().dispatchEvent(
      new KeyboardEvent("keydown", {
        key: value,
        bubbles: true,
        cancelable: true,
        ...modifiers,
      }),
    );
  try {
    await editor.ready;
    await persistence.ready;
    const existingDatabase = database + "-existing";
    await new Promise<void>((resolve, reject) => {
      const request = indexedDB.open(existingDatabase, 1);
      request.onsuccess = () => {
        request.result.close();
        resolve();
      };
      request.onerror = () => reject(request.error);
    });
    let storageError = "";
    try {
      await library
        .createIndexedDBStorage(existingDatabase)
        .load("doc", new AbortController().signal);
    } catch (error) {
      storageError = (error as Error).name;
    } finally {
      await new Promise<void>((resolve, reject) => {
        const request = indexedDB.deleteDatabase(existingDatabase);
        request.onsuccess = () => resolve();
        request.onerror = () => reject(request.error);
      });
    }
    equal(
      storageError,
      "NotFoundError",
      "An incompatible existing IndexedDB rejects cleanly instead of hanging",
    );
    await book.setValues(id, "A1:A5", [[1], [2], [3], [null], [5]]);
    await waitUntil(() => persistence.getState().status === "saved");
    const stored = await storage.load("doc", new AbortController().signal);
    equal(
      stored?.snapshot.sheets[0].cells.length,
      4,
      "IndexedDB autosave persists committed data",
    );
    await editor.select(id, "A1");
    key("ArrowDown", { ctrlKey: true, shiftKey: true });
    await waitUntil(() => editor.getSelection().range.r2 === 2);
    equal(
      editor.getSelection().range,
      library.parseRange("A1:A3"),
      "Ctrl+Shift+arrow extends to the data boundary",
    );
    key("ArrowDown", { metaKey: true, shiftKey: true });
    await waitUntil(() => editor.getSelection().range.r2 === 4);
    equal(
      editor.getSelection().range,
      library.parseRange("A1:A5"),
      "Command+Shift+arrow crosses a blank gap",
    );
    await editor.setZoom(1.5);
    equal(
      root().querySelector<HTMLElement>(".grid-layer")!.style.transform,
      "scale(1.5)",
      "Zoom scales the actual grid",
    );
    await editor.setOptions({ locale: "ja-JP" });
    equal(
      editor.getZoom(),
      1.5,
      "Runtime zoom survives editor reconfiguration",
    );
    await editor.setOptions({ locale: "en-US" });
    key("F1");
    await waitUntil(() => !!root().querySelector(".dialog"));
    equal(
      root()
        .querySelector(".dialog")!
        .textContent!.includes("Keyboard shortcuts"),
      true,
      "F1 opens shortcut help",
    );
    root().querySelector<HTMLFormElement>(".dialog")!.requestSubmit();
    await waitUntil(() => !root().querySelector(".dialog"));
    await editor.select(id, "B1");
    key("F2");
    await waitUntil(() => editor.getEditState().editing);
    let input = root().querySelector<HTMLTextAreaElement>(".cell-input")!;
    input.value = "saved draft";
    input.dispatchEvent(new Event("input", { bubbles: true }));
    const unload = new Event("beforeunload", { cancelable: true });
    window.dispatchEvent(unload);
    equal(
      unload.defaultPrevented,
      true,
      "Uncommitted drafts trigger the unload guard",
    );
    input.dispatchEvent(
      new KeyboardEvent("keydown", {
        key: "s",
        ctrlKey: true,
        bubbles: true,
        cancelable: true,
      }),
    );
    await waitUntil(
      () =>
        !editor.getEditState().editing &&
        persistence.getState().status === "saved",
    );
    equal(
      await book.getValues(id, "B1"),
      [["saved draft"]],
      "Ctrl+S commits the current input and saves it",
    );
    const savedRevision = (await book.getMetadata()).revision;
    await waitUntil(
      () =>
        persistence.getState().status === "saved" &&
        persistence.getState().savedRevision === savedRevision,
    );
    persistence.dispose();
    await book.setValues(id, "B1", [["current"]]);
    persistence = library.createPersistence(book, {
      key: "doc",
      storage,
      debounceMs: 30,
      maxRetries: 0,
    });
    await persistence.ready;
    await editor.setOptions({ persistence });
    await editor.select(id, "B1");
    key("F2");
    await waitUntil(() => editor.getEditState().editing);
    input = root().querySelector<HTMLTextAreaElement>(".cell-input")!;
    input.value = "pending input";
    input.dispatchEvent(new Event("input", { bubbles: true }));
    const restore = [
      ...root().querySelectorAll<HTMLButtonElement>(".persistence-bar button"),
    ].find((button) => button.textContent === "Restore draft")!;
    equal(
      restore.hidden,
      false,
      "Stored draft requires an explicit recovery choice",
    );
    restore.click();
    await waitUntil(() => persistence.getState().status === "saved");
    equal(
      editor.getEditState().dirty,
      false,
      "Successful recovery cancels the old editor draft",
    );
    equal(
      await book.getValues(id, "B1"),
      [["saved draft"]],
      "Recovery restores the stored workbook",
    );
    const second = await book.addSheet("Issues");
    await book.setValues(second, "A1:C2", [
      ["=1/0", "=SEQUENCE(2)", "=UNKNOWN()"],
      [null, "blocked", null],
    ]);
    await editor.select(id, "A1");
    root().querySelector<HTMLButtonElement>(".statusbar .tool-button")!.click();
    await waitUntil(() => !!root().querySelector(".diagnostics-dialog"));
    const filters = root().querySelectorAll<HTMLSelectElement>(
      ".issue-filters select",
    );
    filters[0].value = "error";
    filters[0].dispatchEvent(new Event("change"));
    filters[1].value = second;
    filters[1].dispatchEvent(new Event("change"));
    equal(
      root().querySelectorAll(".issue-card").length,
      2,
      "Diagnostics filters distinguish formula errors from compatibility warnings",
    );
    const spill = [...root().querySelectorAll<HTMLElement>(".issue-card")].find(
      (card) => card.textContent!.includes("SPILL_CONFLICT"),
    )!;
    spill.querySelector<HTMLButtonElement>(".issue-locate")!.click();
    await waitUntil(
      () =>
        editor.getSelection().sheetId === second &&
        editor.getSelection().range.c1 === 1,
    );
    equal(
      editor.getSelection().range,
      library.parseRange("B1"),
      "Diagnostics locate the affected cell on another sheet",
    );
    storageUnavailable = true;
    await book.setValues(second, "D1", [[9]]);
    await waitUntil(() => persistence.getState().status === "error");
    equal(
      persistence.hasUnsavedChanges(),
      true,
      "Failed background saves retain unsaved state",
    );
    const retry = [
      ...root().querySelectorAll<HTMLButtonElement>(".persistence-bar button"),
    ].find((button) => button.textContent === "Retry")!;
    equal(retry.hidden, false, "A failed save exposes the retry action");
    storageUnavailable = false;
    retry.click();
    await waitUntil(() => persistence.getState().status === "saved");
    equal(
      persistence.hasUnsavedChanges(),
      false,
      "Retry saves the latest workbook and clears the dirty state",
    );
    equal(
      errors.map((error) => error.message),
      [],
      "Product workflows report no editor errors",
    );
  } finally {
    editor.destroy();
    persistence.dispose();
    await book.dispose();
    host.remove();
    await new Promise<void>((resolve, reject) => {
      const request = indexedDB.deleteDatabase(database);
      request.onsuccess = () => resolve();
      request.onerror = () => reject(request.error);
    });
  }
}

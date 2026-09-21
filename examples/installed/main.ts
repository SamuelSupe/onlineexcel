import {
  createWorkbook,
  mountEditor,
  supportedLocales,
  WorkbookError,
  type ChangeEvent,
  type Editor,
  type EditorConfiguration,
  type Workbook,
} from "onlineexcel";
import {
  OnlineExcel as ReactExcel,
  type EditorHandle,
} from "onlineexcel/react";
import { OnlineExcel as VueExcel } from "onlineexcel/vue";
import { createElement, StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { createApp, h, shallowRef, nextTick } from "vue";
const asset = (name: string) =>
  new URL("./onlineexcel/" + name, document.baseURI).href;
const workbookOptions = {
  workerFactory: () => new Worker(asset("worker.js"), { type: "module" }),
  workerModules: [new URL("./business.js", document.baseURI).href],
};
const output = document.querySelector<HTMLPreElement>("#results")!;
const container = document.querySelector<HTMLElement>("#editor")!;
const run = document.querySelector<HTMLButtonElement>("#run")!;
let book: Workbook | undefined, editor: Editor | undefined;
const faults: string[] = [];
window.addEventListener("error", (event) => faults.push(event.message));
window.addEventListener("unhandledrejection", (event) =>
  faults.push(String(event.reason)),
);
function check(ok: unknown, name: string) {
  if (!ok) throw new Error(name);
  output.textContent += "\nPASS " + name;
}
async function rejected(work: Promise<unknown>, code: string) {
  try {
    await work;
  } catch (error) {
    check(
      error instanceof WorkbookError && error.code === code,
      "structured error " + code,
    );
    return error as WorkbookError;
  }
  throw new Error("Expected " + code);
}
async function until(predicate: () => unknown) {
  const deadline = performance.now() + 8000;
  while (!predicate()) {
    if (performance.now() > deadline) throw new Error("UI update timed out");
    await new Promise((resolve) => setTimeout(resolve, 16));
  }
}
const root = () => container.firstElementChild!.shadowRoot!;
run.onclick = async () => {
  run.disabled = true;
  faults.length = 0;
  output.textContent =
    "Production Vite build from installed npm archive; strict CSP";
  editor?.destroy();
  await book?.dispose();
  try {
    book = await createWorkbook(workbookOptions);
    const sheet = (await book.getMetadata()).sheets[0].id;
    const changes: ChangeEvent[] = [];
    book.on("change", (event) => changes.push(event));
    await book
      .withOptions({ origin: "host", operationId: "seed" })
      .setValues(sheet, "A1:B1", [[4, "=ACME.DOUBLE(A1)"]]);
    check(
      JSON.stringify(await book.getValues(sheet, "B1:C1")) === "[[8,4]]",
      "custom function dynamic array in real Worker",
    );
    check(
      changes[0].origin === "host" && changes[0].operationId === "seed",
      "change origin and correlation ID",
    );
    check(
      (await book.getFunctions()).some(
        (item) => item.name === "ACME.DOUBLE" && item.signature === "value",
      ),
      "Worker function catalog and signature",
    );
    const save = await book.createSavePoint();
    const failure = await rejected(
      book.setValues(sheet, "A1", [[-1]]),
      "VALIDATION_FAILED",
    );
    check(
      failure.sheetId === sheet && failure.range?.r1 === 0 && !!failure.details,
      "validation error location and details",
    );
    check(
      (await book.createSavePoint()).revision === save.revision &&
        (await book.getValues(sheet, "A1"))[0][0] === 4,
      "rejected transaction leaves values and revision intact",
    );
    const bad = structuredClone(save.snapshot);
    bad.sheets[0].cells.find(([key]) => key === 0)![1].value = -2;
    await rejected(book.importJSON(bad), "VALIDATION_FAILED");
    check(
      (await book.createSavePoint()).revision === save.revision,
      "rejected import preserves current workbook",
    );
    await book.setRecords(
      sheet,
      [
        { sku: "001", amount: 12 },
        { sku: "=literal", amount: 24 },
      ],
      [{ key: "sku", title: "SKU" }, { key: "amount" }],
      { start: "A5", header: true },
    );
    const records = [];
    for await (const rows of book.readRecords(
      sheet,
      "A6:B7",
      [{ key: "sku" }, { key: "amount" }],
      { rowsPerChunk: 1 },
    ))
      records.push(...rows);
    check(
      records[1].sku === "=literal" && records[0].sku === "001",
      "records preserve identifiers and literal formulas",
    );
    const chunks = book.readChunks(sheet, "A5:B7", { rowsPerChunk: 1 });
    await chunks.next();
    await book.setValues(sheet, "F1", [[1]]);
    await rejected(chunks.next(), "REVISION_CONFLICT");
    const written = await book.writeChunks(sheet, "H1", [[[1, 2]], [[3, 4]]]);
    check(
      written.rows === 2 && written.chunks === 2,
      "chunked writes report committed progress",
    );
    await book.undo();
    check(
      (await book.getValues(sheet, "H2"))[0][0] === null,
      "chunk writes have independent undo transactions",
    );
    const abort = new AbortController();
    abort.abort();
    await rejected(book.createSavePoint({ signal: abort.signal }), "CANCELLED");
    let exported = 0,
      imported = 0,
      custom = 0;
    editor = mountEditor(container, {
      workbook: book,
      locale: "en-US",
      styleNonce: "sdk-example",
      toolbarItems: ["bold", "open", "saveXlsx"],
      actions: [
        {
          id: "save-host",
          label: "Host action",
          icon: "saveJson",
          allowReadOnly: true,
          run: async (context) => {
            await context.workbook.setValues(sheet, "J1", [[++custom]]);
          },
        },
      ],
      onImport: () => {
        imported++;
      },
      onExport: async (_format, context) => {
        await context.workbook.createSavePoint();
        exported++;
      },
    });
    await editor.ready;
    await editor.select(sheet, "A1");
    const priorView = container.firstElementChild;
    await editor.setOptions({ locale: "en-US" });
    check(
      container.firstElementChild === priorView,
      "unchanged options preserve the existing view",
    );
    const draft = root().querySelector<HTMLInputElement>(".formula")!;
    draft.focus();
    draft.value = "-1";
    draft.dispatchEvent(new Event("input", { bubbles: true }));
    await rejected(editor.setOptions({ locale: "ja-JP" }), "VALIDATION_FAILED");
    check(
      container.firstElementChild === priorView && editor.getEditState().dirty,
      "failed option update preserves draft and view",
    );
    editor.cancelEdit();
    await editor.select(sheet, "K120");
    const before = editor.getSelection(),
      scroll = root().querySelector<HTMLElement>(".scroll")!;
    const position = [scroll.scrollLeft, scroll.scrollTop];
    await editor.setOptions({ locale: "ja-JP", readOnly: true });
    check(
      editor.getSelection().sheetId === before.sheetId &&
        Object.entries(before.range).every(
          ([key, value]) =>
            editor!.getSelection().range[key as keyof typeof before.range] ===
            value,
        ),
      "runtime options preserve selection",
    );
    const afterScroll = root().querySelector<HTMLElement>(".scroll")!;
    check(
      JSON.stringify([afterScroll.scrollLeft, afterScroll.scrollTop]) ===
        JSON.stringify(position),
      "runtime options preserve scroll",
    );
    check(
      root().querySelector("canvas")?.getAttribute("aria-readonly") ===
        "true" && container.firstElementChild?.getAttribute("lang") === "ja-JP",
      "language and readonly update",
    );
    check(
      !root().querySelector('[data-action="italic"]') &&
        !!root().querySelector('[data-action="bold"]'),
      "toolbar whitelist",
    );
    (
      root().querySelector('[data-action="saveXlsx"]') as HTMLButtonElement
    ).click();
    await until(() => exported === 1);
    const host = [...root().querySelectorAll("button")].find(
      (button) => button.textContent === "Host action",
    )!;
    check(
      !!host.querySelector("svg") && !!host.querySelector(".tool-text"),
      "custom toolbar button has icon and text",
    );
    host.click();
    await until(() => custom === 1 && exported === 1);
    await until(() => changes.some((event) => event.origin === "editor"));
    check(
      changes.at(-1)?.origin === "editor",
      "editor and API changes remain distinguishable",
    );
    await editor.setOptions({ readOnly: false, locale: "en-US" });
    (root().querySelector('[data-action="open"]') as HTMLButtonElement).click();
    await until(() => imported === 1);
    check(
      imported === 1 && exported === 1,
      "host file callbacks replace default file dialogs",
    );
    await editor.select(sheet, "L1");
    const formula = root().querySelector<HTMLInputElement>(".formula")!;
    formula.focus();
    formula.value = "=ACME.D";
    formula.dispatchEvent(new Event("input", { bubbles: true }));
    check(
      root()
        .querySelector(".formula-help")
        ?.textContent?.includes("ACME.DOUBLE"),
      "custom formula autocomplete and signature",
    );
    editor.cancelEdit();
    await editor.setOptions({ toolbar: false });
    check(!root().querySelector(".menubar"), "runtime toolbar visibility");
    await editor.setOptions({ toolbar: true });
    const isolated = await createWorkbook({
      workerUrl: asset("worker.js"),
      snapshot: save.snapshot,
    });
    check(
      (await isolated.getValues(sheet, "B1"))[0][0] &&
        (await isolated.getDiagnostics()).some(
          (item) => item.code === "UNSUPPORTED_FORMULA",
        ),
      "custom functions isolated between Workers",
    );
    await isolated.dispose();
    check(supportedLocales.length === 5, "five locale exports");
    await adapters(book);
    check(!faults.length, "no window errors, rejected promises or CSP faults");
    output.textContent += "\nALL SDK INTEGRATION CHECKS PASSED";
  } catch (error) {
    output.textContent += "\nFAIL " + (error as Error).stack;
  } finally {
    run.disabled = false;
  }
};
async function adapters(external: Workbook) {
  const parent = document.querySelector("#frameworks")!;
  const reactHost = document.createElement("div");
  reactHost.className = "framework";
  parent.append(reactHost);
  const react = createRoot(reactHost);
  let reactHandle: EditorHandle | undefined;
  const renderReact = (locale: "en-US" | "ko-KR") =>
    react.render(
      createElement(
        StrictMode,
        {},
        createElement(ReactExcel, {
          workbookOptions,
          options: { locale, styleNonce: "sdk-example" },
          onReady: (handle) => (reactHandle = handle),
          onError: (error) => faults.push(error.message),
        }),
      ),
    );
  renderReact("en-US");
  await until(() => reactHandle);
  const owned = reactHandle!.workbook;
  const id = (await owned.getMetadata()).sheets[0].id;
  await owned.setValues(id, "A1", [[10]]);
  const originalView = reactHost.querySelector('[lang="en-US"]');
  renderReact("en-US");
  await new Promise(requestAnimationFrame);
  check(
    reactHost.querySelector('[lang="en-US"]') === originalView,
    "React rerender with equivalent options keeps the view",
  );
  renderReact("ko-KR");
  await until(() => reactHost.querySelector('[lang="ko-KR"]'));
  check(
    (await owned.getValues(id, "A1"))[0][0] === 10,
    "React StrictMode initializes once per live mount and retains data on option changes",
  );
  react.unmount();
  await rejected(owned.getMetadata(), "DISPOSED");
  reactHost.remove();
  const vueHost = document.createElement("div");
  vueHost.className = "framework";
  parent.append(vueHost);
  const options = shallowRef<EditorConfiguration>({
    locale: "en-US",
    readOnly: true,
    styleNonce: "sdk-example",
  });
  let vueHandle: EditorHandle | undefined;
  const app = createApp({
    render: () =>
      h(VueExcel, {
        workbook: external,
        options: options.value,
        onReady: (handle) => (vueHandle = handle),
        onError: (error) => faults.push(error.message),
      }),
  });
  app.mount(vueHost);
  await until(() => vueHandle);
  options.value = { locale: "zh-TW", styleNonce: "sdk-example" };
  await nextTick();
  await until(() => vueHost.querySelector('[lang="zh-TW"]'));
  check(
    vueHost
      .querySelector('[lang="zh-TW"]')
      ?.shadowRoot?.querySelector("canvas")
      ?.getAttribute("aria-readonly") === "false",
    "removing adapter readonly option restores the default",
  );
  check(
    vueHandle!.workbook === external,
    "Vue uses external workbook and updates options",
  );
  app.unmount();
  await external.getMetadata();
  check(
    !vueHost.children.length,
    "Vue unmount removes view and preserves host-owned workbook",
  );
  vueHost.remove();
  const rapid = document.createElement("div");
  parent.append(rapid);
  const temporary = createRoot(rapid);
  let started = 0,
    stopped = 0,
    readyAfterUnmount = 0;
  temporary.render(
    createElement(ReactExcel, {
      workbookOptions: {
        workerModules: [new URL("./deferred-module.js", document.baseURI).href],
        workerFactory: () => {
          const worker = new Worker(asset("worker.js"), { type: "module" });
          const terminate = worker.terminate.bind(worker);
          worker.terminate = () => {
            stopped++;
            terminate();
          };
          started++;
          return worker;
        },
      },
      onReady: () => readyAfterUnmount++,
      onError: (error) => faults.push(error.message),
    }),
  );
  await until(() => started === 1);
  temporary.unmount();
  await until(() => stopped === 1);
  check(
    !readyAfterUnmount && !rapid.children.length,
    "unmount during initialization terminates the late Worker without mounting a view",
  );
  rapid.remove();
}

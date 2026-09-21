import type { Workbook, Editor, WorkbookSnapshot } from "../src/index";
import { keyOf } from "../src/core/address";
import { officeChecks } from "./office-qa";
import { productChecks } from "./product-qa";
const builtUrl = "/dist/index.js";
const library = (await import(
  /* @vite-ignore */ builtUrl
)) as typeof import("../src/index");
const output = document.querySelector<HTMLPreElement>("#results")!,
  container = document.querySelector<HTMLElement>("#editor")!;
let currentBook: Workbook | undefined, currentEditor: Editor | undefined;
const failures: string[] = [];
window.addEventListener("error", (event) => {
  failures.push(event.message);
});
window.addEventListener("unhandledrejection", (event) => {
  failures.push(String(event.reason));
});
const frame = () =>
  new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
async function waitUntil(predicate: () => boolean) {
  const deadline = performance.now() + 5000;
  while (!predicate()) {
    if (performance.now() > deadline)
      throw new Error("Editor update timed out");
    await frame();
  }
}
function log(message: string) {
  output.textContent += "\n" + message;
}
function equal(actual: unknown, expected: unknown, label: string) {
  if (JSON.stringify(actual) !== JSON.stringify(expected))
    throw new Error(
      `${label}: ${JSON.stringify(actual)} != ${JSON.stringify(expected)}`,
    );
  log("PASS " + label);
}
async function cleanup() {
  currentEditor?.destroy();
  await currentBook?.dispose();
  currentBook = undefined;
  currentEditor = undefined;
  container.replaceChildren();
}
const run = document.querySelector<HTMLButtonElement>("#run")!,
  perf = document.querySelector<HTMLButtonElement>("#perf")!;
run.onclick = async () => {
  run.disabled = true;
  output.textContent = "Testing built ESM package and Workers";
  try {
    await cleanup();
    const book = await library.createWorkbook({
      sheets: [{ name: "Test", rows: 10000, columns: 100 }],
    });
    currentBook = book;
    const id = (await book.getMetadata()).sheets[0].id;
    await book.setValues(id, "A1:B2", [
      [2, 3],
      ["=SUM(A1:B1)", "=A2*2"],
    ]);
    equal(await book.getValues(id, "A2:B2"), [[5, 10]], "Worker formulas");
    const rangeBook = await library.createWorkbook({
      sheets: [{ name: "Ranges", rows: 100, columns: 10 }],
    });
    try {
      const rangeId = (await rangeBook.getMetadata()).sheets[0].id;
      await rangeBook.setValues(rangeId, "A1:A3", [[1], [2], [3]]);
      await rangeBook.setValues(rangeId, "B5:C5", [
        ["=SUM(A3:A1)", "=SUM(Ranges!$1:$1)"],
      ]);
      await rangeBook.copyRange(rangeId, "A1:A3", rangeId, "A2", {
        valuesOnly: true,
      });
      equal(
        await rangeBook.getValues(rangeId, "A1:A4"),
        [[1], [1], [2], [3]],
        "Overlapping values-only paste through Worker",
      );
      await rangeBook.undo();
      equal(
        await rangeBook.getValues(rangeId, "A1:A4"),
        [[1], [2], [3], [null]],
        "Undo restores overlapping paste",
      );
      await rangeBook.redo();
      equal(
        await rangeBook.getValues(rangeId, "A1:A4"),
        [[1], [1], [2], [3]],
        "Redo restores overlapping paste",
      );
      await rangeBook.undo();
      await rangeBook.copyRange(rangeId, "A1:A3", rangeId, "C1", {
        cut: true,
        valuesOnly: true,
      });
      equal(
        await rangeBook.getValues(rangeId, "C1:C3"),
        [[1], [2], [3]],
        "Values-only cut preserves source values",
      );
      await rangeBook.undo();
      await rangeBook.deleteRows(rangeId, 1);
      equal(
        await rangeBook.getValues(rangeId, "B4:C4"),
        [[4, 1]],
        "Reversed and absolute whole-row references after deletion",
      );
      await rangeBook.setFormula(
        rangeId,
        "D10",
        "=IFERROR({#N/A;#N/A},{10;20})",
      );
      await rangeBook.setFormula(
        rangeId,
        "G10",
        "=IF({TRUE;FALSE},{1,2},{3,4})",
      );
      await rangeBook.setFormula(rangeId, "D15", '=MATCH(">5",{"abc";">5"},0)');
      equal(
        await rangeBook.getValues(rangeId, "D10:D11"),
        [[10], [20]],
        "Element-wise IFERROR fallback through Worker",
      );
      equal(
        await rangeBook.getValues(rangeId, "G10:H11"),
        [
          [1, 2],
          [3, 4],
        ],
        "IF broadcasts row and column arrays through Worker",
      );
      equal(
        await rangeBook.getValues(rangeId, "D15"),
        [[2]],
        "Exact lookup keeps comparison characters literal",
      );
      await rangeBook.defineName("Amount", rangeId, "A1");
      await rangeBook.setFormula(rangeId, "D20", "=SUM(Amount)");
      await rangeBook.transaction([
        {
          type: "defineName",
          sheetId: rangeId,
          name: "Amount",
          range: library.parseRange("A2"),
        },
        {
          type: "copy",
          sheetId: rangeId,
          range: library.parseRange("D20"),
          targetSheetId: rangeId,
          targetRow: 19,
          targetColumn: 4,
          valuesOnly: true,
        },
      ]);
      equal(
        await rangeBook.getValues(rangeId, "D20:E20"),
        [[3, 3]],
        "Named-range changes are visible to copies within one transaction",
      );
      const file = await rangeBook.exportXlsx();
      await rangeBook.importXlsx(file.data);
      const restoredId = (await rangeBook.getMetadata()).sheets[0].id;
      equal(
        await rangeBook.getValues(restoredId, "B4:C4"),
        [[4, 1]],
        "Adjusted references survive XLSX roundtrip",
      );
      equal(
        await rangeBook.getValues(restoredId, "D10:D11"),
        [[10], [20]],
        "Conditional array recalculates after XLSX roundtrip",
      );
    } finally {
      await rangeBook.dispose();
    }
    currentEditor = library.mountEditor(container, {
      workbook: book,
      onError: (error) => log("Editor error: " + error.message),
    });
    await currentEditor.ready;
    await currentEditor.select(id, "B2");
    const editorRoot = container.firstElementChild!.shadowRoot!;
    equal(
      (editorRoot.querySelector(".formula") as HTMLInputElement).value,
      "=A2*2",
      "Editor formula bar",
    );
    await currentEditor.select(id, "D4");
    const grid = editorRoot.querySelector("canvas")!;
    grid.dispatchEvent(
      new KeyboardEvent("keydown", {
        key: "ArrowUp",
        shiftKey: true,
        bubbles: true,
      }),
    );
    grid.dispatchEvent(
      new KeyboardEvent("keydown", {
        key: "ArrowUp",
        shiftKey: true,
        bubbles: true,
      }),
    );
    equal(
      currentEditor.getSelection().range,
      { r1: 1, c1: 3, r2: 3, c2: 3 },
      "Backward keyboard range extension",
    );
    await currentEditor.select(id, "D4");
    grid.dispatchEvent(
      new KeyboardEvent("keydown", { key: "F2", bubbles: true }),
    );
    const cellInput =
      editorRoot.querySelector<HTMLTextAreaElement>(".cell-input")!;
    cellInput.dispatchEvent(
      new CompositionEvent("compositionstart", { bubbles: true }),
    );
    cellInput.value = "中文组合输入";
    cellInput.dispatchEvent(
      new KeyboardEvent("keydown", {
        key: "Enter",
        isComposing: true,
        bubbles: true,
      }),
    );
    equal(
      await book.getValues(id, "D4"),
      [[null]],
      "Composition Enter does not commit",
    );
    cellInput.dispatchEvent(
      new CompositionEvent("compositionend", {
        data: "中文组合输入",
        bubbles: true,
      }),
    );
    cellInput.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
    );
    await waitUntil(() => cellInput.hidden);
    equal(
      await book.getValues(id, "D4"),
      [["中文组合输入"]],
      "Composition completes without losing text",
    );
    await currentEditor.select(id, "D5");
    grid.dispatchEvent(
      new KeyboardEvent("keydown", { key: "=", bubbles: true }),
    );
    const help = editorRoot.querySelector<HTMLElement>(".formula-help")!;
    const formulaInput =
      editorRoot.querySelector<HTMLInputElement>(".formula")!;
    equal(help.hidden, false, "Equals opens formula assistance");
    cellInput.value = "=SU";
    cellInput.setSelectionRange(3, 3);
    cellInput.dispatchEvent(new InputEvent("input", { bubbles: true }));
    cellInput.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Tab", bubbles: true }),
    );
    equal(
      [cellInput.value, formulaInput.value],
      ["=SUM(", "=SUM("],
      "Completion preserves the edit and synchronizes the formula bar",
    );
    cellInput.value = "=SUM(1,2)";
    cellInput.dispatchEvent(new InputEvent("input", { bubbles: true }));
    cellInput.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
    );
    equal(
      await book.getValues(id, "D5"),
      [[3]],
      "Completed formula calculates through Worker",
    );
    await currentEditor.select(id, "D6");
    formulaInput.focus();
    formulaInput.value = "=AV";
    formulaInput.setSelectionRange(3, 3);
    formulaInput.dispatchEvent(new InputEvent("input", { bubbles: true }));
    formulaInput.dispatchEvent(
      new CompositionEvent("compositionstart", { bubbles: true }),
    );
    equal(help.hidden, true, "Formula assistance yields to IME composition");
    formulaInput.dispatchEvent(
      new CompositionEvent("compositionend", { bubbles: true }),
    );
    formulaInput.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Tab", bubbles: true }),
    );
    equal(
      formulaInput.value,
      "=AVERAGE(",
      "Formula bar supports completion after composition",
    );
    formulaInput.value = "'=literal";
    formulaInput.dispatchEvent(new InputEvent("input", { bubbles: true }));
    equal(
      help.hidden,
      true,
      "Literal equals text does not open formula assistance",
    );
    formulaInput.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Escape", bubbles: true }),
    );
    equal(
      await book.getValues(id, "D6"),
      [[null]],
      "Escape cancels formula bar edits",
    );
    const editingChecks: [unknown, unknown, string][] = [];
    await currentEditor.select(id, "J12");
    grid.dispatchEvent(
      new KeyboardEvent("keydown", { key: "1", bubbles: true }),
    );
    cellInput.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
    );
    editingChecks.push([
      [
        currentEditor.getSelection().range.r1,
        editorRoot.activeElement === grid,
      ],
      [12, true],
      "Enter immediately advances selection and restores keyboard focus",
    ]);
    (editorRoot.activeElement ?? document.body).dispatchEvent(
      new KeyboardEvent("keydown", { key: "2", bubbles: true }),
    );
    (editorRoot.activeElement ?? document.body).dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
    );
    editingChecks.push([
      await book.getValues(id, "J12:J13"),
      [[1], [2]],
      "Rapid consecutive edits target separate cells",
    ]);
    await currentEditor.select(id, "J15");
    formulaInput.focus();
    formulaInput.value = "5";
    formulaInput.dispatchEvent(new InputEvent("input", { bubbles: true }));
    formulaInput.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
    );
    const addressBox = editorRoot.querySelector<HTMLInputElement>(".address")!;
    addressBox.focus();
    await book.getValues(id, "J15");
    await frame();
    editingChecks.push([
      editorRoot.activeElement === addressBox,
      true,
      "Completed formula-bar commit does not steal newer focus",
    ]);
    addressBox.focus();
    addressBox.value = "J17";
    addressBox.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
    );
    formulaInput.focus();
    formulaInput.value = "=1+2";
    formulaInput.dispatchEvent(new InputEvent("input", { bubbles: true }));
    await book.getMetadata();
    await frame();
    editingChecks.push([
      editorRoot.activeElement === formulaInput,
      true,
      "Address navigation does not steal a newer formula draft focus",
    ]);
    editingChecks.push([
      await book.getValues(id, "J17"),
      [[null]],
      "Address navigation leaves the new formula draft uncommitted",
    ]);
    formulaInput.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Escape", bubbles: true }),
    );
    await currentEditor.select(id, "J18");
    grid.dispatchEvent(
      new KeyboardEvent("keydown", { key: "7", bubbles: true }),
    );
    const gridScroll = editorRoot.querySelector<HTMLElement>(".scroll")!;
    gridScroll.scrollTop += 28;
    gridScroll.dispatchEvent(new Event("scroll"));
    editingChecks.push([
      await book.getValues(id, "J18"),
      [[7]],
      "Actual scroll movement still commits the active draft",
    ]);

    addressBox.focus();
    addressBox.value = "J19";
    await book.setValues(id, "J30", [[9]]);
    await book.getMetadata();
    await frame();
    editingChecks.push([
      addressBox.value,
      "J19",
      "Workbook refresh preserves the address being typed",
    ]);
    addressBox.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
    );
    await book.getMetadata();
    await frame();
    const typedDestination = currentEditor.getSelection().range;
    editingChecks.push([
      [typedDestination.r1, typedDestination.c1],
      [18, 9],
      "Address Enter navigates to the typed destination after refresh",
    ]);

    await book.setValues(id, "L1:M1", [[2, "=L1*2"]]);
    await book.setStyle(id, "M1", { bold: true, numberFormat: "0.00" });
    await currentEditor.select(id, "M1");
    grid.focus();
    const transfer = new DataTransfer();
    grid.dispatchEvent(
      new ClipboardEvent("copy", {
        clipboardData: transfer,
        bubbles: true,
        composed: true,
      }),
    );
    editingChecks.push([
      transfer.getData("text/plain"),
      "4.00",
      "Copy is ready immediately after API selection",
    ]);
    await currentEditor.select(id, "M2");
    grid.dispatchEvent(
      new ClipboardEvent("paste", {
        clipboardData: transfer,
        bubbles: true,
        composed: true,
      }),
    );
    await book.getMetadata();
    await currentEditor.select(id, "M3");
    grid.dispatchEvent(
      new ClipboardEvent("paste", {
        clipboardData: transfer,
        bubbles: true,
        composed: true,
      }),
    );
    const repeated = (await book.getRegion(id, "M3")).cells[0];
    editingChecks.push([
      [repeated?.formula, repeated?.style.bold],
      ["=(L3*2)", true],
      "Repeated paste retains relative formulas and styles",
    ]);
    await currentEditor.select(id, "J5");
    formulaInput.focus();
    formulaInput.value = "=SUM(10,20)";
    formulaInput.dispatchEvent(new InputEvent("input", { bubbles: true }));
    await currentEditor.select(id, "J6");
    editingChecks.push([
      await book.getValues(id, "J5"),
      [[30]],
      "Formula-bar draft survives selection change",
    ]);

    await currentEditor.select(id, "J7");
    grid.dispatchEvent(
      new KeyboardEvent("keydown", { key: "F2", bubbles: true }),
    );
    cellInput.value = "42";
    cellInput.dispatchEvent(new InputEvent("input", { bubbles: true }));
    await currentEditor.select(id, "J8");
    if (!cellInput.hidden)
      cellInput.dispatchEvent(
        new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
      );
    editingChecks.push([
      await book.getValues(id, "J7:J8"),
      [[42], [null]],
      "API selection commits to the original edit target",
    ]);

    await book.setValues(id, "J9", [["=1+1"]], { parseFormulas: false });
    await currentEditor.select(id, "J9");
    await waitUntil(() => formulaInput.value.endsWith("=1+1"));
    grid.dispatchEvent(
      new KeyboardEvent("keydown", { key: "F2", bubbles: true }),
    );
    cellInput.dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
    );
    editingChecks.push([
      await book.getValues(id, "J9"),
      [["=1+1"]],
      "Unchanged literal text is never reinterpreted as a formula",
    ]);

    const readonlyHost = document.createElement("div");
    readonlyHost.style.height = "360px";
    document.body.append(readonlyHost);
    await book.setValues(id, "J10", [["=1+1"]], { parseFormulas: false });
    const readonlyEditor = library.mountEditor(readonlyHost, {
      workbook: book,
      readOnly: true,
    });
    try {
      await readonlyEditor.ready;
      await readonlyEditor.select(id, "J10");
      const beforeRevision = (await book.getMetadata()).revision;
      readonlyHost
        .firstElementChild!.shadowRoot!.querySelector(".formula")!
        .dispatchEvent(
          new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
        );
      editingChecks.push([
        (await book.getMetadata()).revision,
        beforeRevision,
        "Read-only formula bar Enter cannot mutate the workbook",
      ]);
    } finally {
      readonlyEditor.destroy();
      readonlyHost.remove();
    }

    const filtered = await library.createWorkbook({
      sheets: [{ name: "Sparse", rows: 100000, columns: 100 }],
    });
    const filteredHost = document.createElement("div");
    filteredHost.style.height = "420px";
    document.body.append(filteredHost);
    let filteredEditor: Editor | undefined;
    try {
      const sid = (await filtered.getMetadata()).sheets[0].id;
      await filtered.setValues(sid, "A1", [["Header"]]);
      await filtered.setValues(sid, "A50001", [["keep"]]);
      await filtered.setValues(sid, "A100000", [["keep"]]);
      await filtered.filter(sid, "A1:A100000", [
        { column: 0, operator: "eq", value: "keep" },
      ]);
      filteredEditor = library.mountEditor(filteredHost, {
        workbook: filtered,
      });
      let readyError = "";
      try {
        await filteredEditor.ready;
      } catch (error) {
        readyError = (error as Error).message;
      }
      editingChecks.push([
        readyError,
        "",
        "Sparse filtered rows stay within the viewport read limit",
      ]);
    } finally {
      filteredEditor?.destroy();
      filteredHost.remove();
      await filtered.dispose();
    }
    for (const [actual, , name] of editingChecks)
      log(`${name}: ${JSON.stringify(actual)}`);
    for (const [actual, expected, name] of editingChecks)
      equal(actual, expected, name);
    currentEditor.destroy();
    currentEditor = library.mountEditor(container, {
      workbook: book,
      locale: "en-US",
    });
    await currentEditor.ready;
    equal(container.children.length, 1, "Unmount and remount");
    await officeChecks(library, equal, waitUntil);
    await productChecks(library, equal, waitUntil);
    const localized: {
      editor: Editor;
      host: HTMLElement;
      root: ShadowRoot;
      home: string;
    }[] = [];
    try {
      const languages = [
        ["zh-CN", "开始", "查找 / 替换", "公式模式"],
        ["zh-TW", "常用", "尋找 / 取代", "公式模式"],
        ["en-US", "Home", "Find / replace", "Formula mode"],
        ["ja-JP", "ホーム", "検索 / 置換", "数式モード"],
        ["ko-KR", "홈", "찾기 / 바꾸기", "수식 모드"],
      ] as const;
      for (const [locale, home, find, formulaMode] of languages) {
        const host = document.createElement("div");
        host.style.height = "420px";
        document.body.append(host);
        const editor = library.mountEditor(host, { workbook: book, locale });
        const root = host.firstElementChild!.shadowRoot!;
        localized.push({ editor, host, root, home });
        await editor.ready;
        equal(
          host.firstElementChild!.getAttribute("lang"),
          locale,
          locale + " language inherited inside Shadow DOM",
        );
        equal(
          root.querySelector('[data-tab="home"]')!.textContent,
          home,
          locale + " toolbar",
        );
        root.querySelector<HTMLButtonElement>('[data-action="find"]')!.click();
        equal(
          root.querySelector(".dialog h2")!.textContent,
          find,
          locale + " dialog",
        );
        root
          .querySelector(".dialog")!
          .dispatchEvent(
            new KeyboardEvent("keydown", { key: "Escape", bubbles: true }),
          );
        await editor.select(id, "H8");
        const formula = root.querySelector<HTMLInputElement>(".formula")!;
        formula.focus();
        formula.value = "=SU";
        formula.setSelectionRange(3, 3);
        formula.dispatchEvent(new InputEvent("input", { bubbles: true }));
        equal(
          root.querySelector(".formula-help-title")!.textContent,
          "ƒx  " + formulaMode,
          locale + " formula guidance",
        );
        formula.dispatchEvent(
          new KeyboardEvent("keydown", { key: "Tab", bubbles: true }),
        );
        equal(
          formula.value,
          "=SUM(",
          locale + " canonical function completion",
        );
        formula.value = "=SUM(1,2)";
        formula.dispatchEvent(
          new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
        );
        equal(
          await book.getValues(id, "H8"),
          [[3]],
          locale + " formula calculation",
        );
      }
      for (const view of localized)
        equal(
          view.root.querySelector('[data-tab="home"]')!.textContent,
          view.home,
          "Concurrent language isolation: " + view.home,
        );
      equal(
        (await book.getMetadata()).sheets[0].name,
        "Test",
        "UI locale leaves workbook content unchanged",
      );
    } finally {
      for (const view of localized) {
        view.editor.destroy();
        view.host.remove();
      }
    }
    const other = await library.createWorkbook();
    try {
      const otherId = (await other.getMetadata()).sheets[0].id;
      await other.setValues(otherId, "A1", [[99]]);
      const secondContainer = document.createElement("div");
      secondContainer.style.height = "360px";
      document.body.append(secondContainer);
      const secondEditor = library.mountEditor(secondContainer, {
        workbook: other,
        readOnly: true,
        toolbar: false,
      });
      await secondEditor.ready;
      const secondRoot = secondContainer.firstElementChild!.shadowRoot!;
      const secondCanvas = secondRoot.querySelector("canvas")!;
      secondCanvas.dispatchEvent(
        new KeyboardEvent("keydown", { key: "Delete", bubbles: true }),
      );
      secondCanvas.dispatchEvent(
        new KeyboardEvent("keydown", { key: "x", bubbles: true }),
      );
      equal(
        await other.getValues(otherId, "A1"),
        [[99]],
        "Read-only editor rejects keyboard writes",
      );
      equal(
        secondRoot.querySelector(".menubar"),
        null,
        "Toolbar can be hidden",
      );
      equal(container.children.length, 1, "Two editors mount simultaneously");
      secondEditor.destroy();
      secondContainer.remove();
      equal(
        await book.getValues(id, "A1"),
        [[2]],
        "Independent workbook instances",
      );
    } finally {
      await other.dispose();
    }
    const data = await book.exportXlsx();
    await book.importXlsx(data.data);
    const restored = (await book.getMetadata()).sheets[0].id;
    equal(
      await book.getValues(restored, "A2:B2"),
      [[5, 10]],
      "XLSX Worker roundtrip",
    );
    const saved = await book.exportJSON();
    await book.setValues(restored, "A1", [[7]]);
    await book.importJSON(saved);
    equal(await book.getValues(restored, "A1"), [[2]], "JSON restore");
    const cancelled = new AbortController();
    const importing = book.importXlsx(data.data, { signal: cancelled.signal });
    cancelled.abort();
    let aborted = false;
    try {
      await importing;
    } catch (error) {
      aborted = (error as Error).name === "AbortError";
    }
    equal(aborted, true, "Import cancellation");
    equal(
      await book.getValues(restored, "A1"),
      [[2]],
      "Cancelled import leaves workbook intact",
    );
    let rejected = false;
    try {
      await book.transaction([
        {
          type: "setValues",
          sheetId: restored,
          range: { r1: 0, c1: 0, r2: 0, c2: 0 },
          values: [[88]],
        },
        { type: "freeze", sheetId: restored, rows: -1, columns: 0 },
      ]);
    } catch {
      rejected = true;
    }
    equal(rejected, true, "Failed transaction rejects");
    equal(
      await book.getValues(restored, "A1"),
      [[2]],
      "Failed transaction rolls back",
    );
    await book.setFormula(restored, "G1", "=SEQUENCE(2,2)");
    const csv = await book.exportCsv(restored);
    equal(
      csv.text.includes("3,4"),
      true,
      "CSV includes full dynamic array extent",
    );
    const noHistory = await library.createWorkbook({ historyLimit: 0 });
    try {
      await noHistory.importJSON(saved);
      await noHistory.setValues(restored, "A1", [[10]]);
      equal(
        await noHistory.undo(),
        null,
        "History configuration survives import",
      );
    } finally {
      await noHistory.dispose();
    }
    equal(failures, [], "No window errors");
    log("ALL INTEGRATION CHECKS PASSED");
  } catch (error) {
    log("FAIL " + (error as Error).stack);
  } finally {
    run.disabled = false;
  }
};
perf.onclick = async () => {
  perf.disabled = true;
  output.textContent =
    "Building deterministic 100,000 × 100 sheet; 1,000,000 nonempty cells / 50,000 formulas";
  try {
    await cleanup();
    const book = await library.createWorkbook();
    currentBook = book;
    const cells: WorkbookSnapshot["sheets"][number]["cells"] = [];
    for (let r = 0; r < 100000; r++)
      for (let c = 0; c < 10; c++)
        cells.push([
          keyOf(r, c),
          c === 9 && r < 50000
            ? { formula: `=SUM(A${r + 1}:I${r + 1})` }
            : {
                value:
                  c === 0 ? `ID-${r + 1}` : ((r * 31 + c * 17) % 10000) / 100,
              },
        ]);
    const snapshot: WorkbookSnapshot = {
      version: 1,
      dateSystem: 1900,
      styles: [{}],
      names: {},
      sheets: [
        {
          id: "bench",
          name: "Benchmark",
          rowCount: 100000,
          columnCount: 100,
          cells,
          rowHeights: {},
          columnWidths: {},
          hiddenRows: [],
          hiddenColumns: [],
          frozenRows: 1,
          frozenColumns: 1,
          merges: [],
        },
      ],
    };
    let start = performance.now();
    await book.importJSON(snapshot);
    const loadMs = performance.now() - start;
    log(`Worker load ${loadMs.toFixed(1)} ms`);
    cells.length = 0;
    currentEditor = library.mountEditor(container, {
      workbook: book,
      locale: "en-US",
    });
    await currentEditor.ready;
    const root = container.firstElementChild!.shadowRoot!,
      scroll = root.querySelector<HTMLElement>(".scroll")!,
      canvas = root.querySelector<HTMLCanvasElement>("canvas")!;
    const editing: number[] = [];
    for (let i = 0; i < 30; i++) {
      const cell = `B${(i % 10) + 1}`;
      await currentEditor.select("bench", cell);
      await frame();
      const start = performance.now();
      await book.setValues("bench", cell, [[i + 100]]);
      // Wait for the editor's visible selection to reflect the committed revision, then a paint.
      await waitUntil(
        () => root.querySelector(".a11y")?.textContent === `${cell} ${i + 100}`,
      );
      await frame();
      editing.push(performance.now() - start);
    }
    editing.sort((a, b) => a - b);
    const editP95Ms = editing[Math.floor(editing.length * 0.95)];
    log(`Editing feedback P95 ${editP95Ms.toFixed(1)} ms`);
    const times: number[] = [];
    let previous = performance.now();
    for (let i = 0; i < 180; i++) {
      scroll.scrollTop = i * 32;
      await frame();
      const now = performance.now();
      times.push(now - previous);
      previous = now;
    }
    const fps = 1000 / (times.reduce((a, b) => a + b, 0) / times.length);
    times.sort((a, b) => a - b);
    const info = {
      userAgent: navigator.userAgent,
      viewport: {
        width: canvas.clientWidth,
        height: canvas.clientHeight,
        dpr: devicePixelRatio,
      },
      loadMs,
      editP50Ms: editing[Math.floor(editing.length * 0.5)],
      editP95Ms,
      scrollFps: fps,
      frameP95Ms: times[Math.floor(times.length * 0.95)],
      windowErrors: failures,
    };
    log(JSON.stringify(info, null, 2));
    log(
      (fps >= 50 && editP95Ms <= 100 ? "PASS" : "TARGET NOT MET") +
        " browser performance gates",
    );
  } catch (error) {
    log("FAIL " + (error as Error).stack);
  } finally {
    perf.disabled = false;
  }
};

import type * as Library from "../src/index";

export async function officeChecks(
  library: typeof Library,
  equal: (actual: unknown, expected: unknown, label: string) => void,
  waitUntil: (predicate: () => boolean) => Promise<void>,
) {
  const book = await library.createWorkbook({
      sheets: [{ name: "Office", rows: 200, columns: 20 }],
    }),
    id = (await book.getMetadata()).sheets[0].id;
  const host = document.createElement("div");
  host.style.height = "640px";
  document.body.append(host);
  const errors: Error[] = [];
  const editor = library.mountEditor(host, {
    workbook: book,
    locale: "en-US",
    onError: (error) => errors.push(error),
  });
  try {
    await editor.ready;
    const root = host.firstElementChild!.shadowRoot!;
    const action = (name: string) =>
      root.querySelector<HTMLButtonElement>(`[data-action="${name}"]`)!.click();
    const tab = (name: string) =>
      root.querySelector<HTMLButtonElement>(`[data-tab="${name}"]`)!.click();
    const form = () => root.querySelector<HTMLFormElement>(".dialog")!;
    const open = async (name: string) => {
      action(name);
      await waitUntil(() => !!form());
    };
    const submit = async () => {
      const revision = (await book.getMetadata()).revision;
      form().requestSubmit();
      await waitUntil(() => !form());
      for (let i = 0; i < 100; i++) {
        if ((await book.getMetadata()).revision > revision) return;
        await new Promise(requestAnimationFrame);
      }
      throw new Error("Dialog did not commit a change");
    };
    await book.setValues(id, "A1:C4", [
      ["Name", "Team", "Score"],
      ["Zoe", "B", 20],
      ["Amy", "A", 30],
      ["Amy", "A", 10],
    ]);
    await editor.select(id, "A1");
    tab("data");
    await open("sort");
    equal(
      form().querySelector<HTMLInputElement>('input[type="checkbox"]')!.checked,
      true,
      "Sort protects the header by default",
    );
    const add = [...form().querySelectorAll("button")].find((b) =>
      b.textContent?.includes("Add sort level"),
    )!;
    add.click();
    const levels = form().querySelectorAll("fieldset");
    levels[1].querySelector("select")!.value = "2";
    levels[1].querySelectorAll("select")[1].value = "desc";
    await submit();
    equal(
      await book.getValues(id, "A1:C4"),
      [
        ["Name", "Team", "Score"],
        ["Amy", "A", 30],
        ["Amy", "A", 10],
        ["Zoe", "B", 20],
      ],
      "Multi-key sorting expands the selected cell and keeps rows together",
    );
    await book.undo();
    equal(
      (await book.getValues(id, "A2"))[0][0],
      "Zoe",
      "Sorting is one undo transaction",
    );
    await editor.select(id, "A1:C4");
    await open("filter");
    const filterColumn = form().querySelector("select")!;
    await waitUntil(() => !filterColumn.disabled);
    const mode = form().querySelectorAll("select")[1];
    mode.value = "values";
    mode.dispatchEvent(new Event("change"));
    [...form().querySelectorAll<HTMLLabelElement>(".filter-value")]
      .find((l) => l.textContent === "Zoe")!
      .querySelector<HTMLInputElement>("input")!
      .click();
    filterColumn.value = "2";
    filterColumn.dispatchEvent(new Event("change"));
    await waitUntil(() => !filterColumn.disabled);
    mode.value = "conditions";
    mode.dispatchEvent(new Event("change"));
    form().querySelectorAll("select")[2].value = "gte";
    const condition = [
      ...form().querySelectorAll<HTMLLabelElement>("label"),
    ].find((l) => l.textContent === "Value 1")!;
    (root.getElementById(condition.htmlFor) as HTMLInputElement).value = "20";
    await submit();
    equal(
      (await book.getMetadata()).sheets[0].filteredRows,
      [1, 3],
      "Value-list and second-column condition filters combine with AND",
    );
    action("clearFilter");
    await book.getMetadata();
    await editor.select(id, "A1");
    await open("find");
    (form().elements.namedItem("search") as HTMLInputElement).value = "Amy";
    form().requestSubmit();
    await waitUntil(() => editor.getSelection().range.r1 === 2);
    root
      .querySelector("canvas")!
      .dispatchEvent(
        new KeyboardEvent("keydown", { key: "F3", bubbles: true }),
      );
    await waitUntil(() => editor.getSelection().range.r1 === 3);
    equal(
      editor.getSelection().range.r1,
      3,
      "Find next traverses subsequent matches with F3",
    );
    root
      .querySelector("canvas")!
      .dispatchEvent(
        new KeyboardEvent("keydown", { key: "F3", bubbles: true }),
      );
    await waitUntil(() => editor.getSelection().range.r1 === 2);
    equal(
      editor.getSelection().range.r1,
      2,
      "Find next wraps to the first match",
    );

    await editor.select(id, "A2:C4");
    tab("view");
    await open("rowHeight");
    (form().elements.namedItem("size") as HTMLInputElement).value = "42";
    await submit();
    equal(
      (await book.getMetadata()).sheets[0].rowHeights,
      { 1: 42, 2: 42, 3: 42 },
      "Row height applies to every selected row",
    );
    await open("columnWidth");
    (form().elements.namedItem("size") as HTMLInputElement).value = "160";
    await submit();
    equal(
      (await book.getMetadata()).sheets[0].columnWidths,
      { 0: 160, 1: 160, 2: 160 },
      "Column width applies to every selected column",
    );
    await book.setValues(id, "A6", [
      ["A long label that needs a wider column"],
    ]);
    await editor.select(id, "A6");
    action("autoFitColumns");
    let autoWidth = 0;
    for (let i = 0; i < 100; i++) {
      autoWidth = (await book.getMetadata()).sheets[0].columnWidths[0];
      if (autoWidth > 160) break;
      await new Promise(requestAnimationFrame);
    }
    equal(
      autoWidth > 160,
      true,
      "Auto-fit measures rendered content across the selected column",
    );
    await book.undo();
    equal(
      (await book.getMetadata()).sheets[0].columnWidths[0],
      160,
      "Auto-fit is undoable",
    );
    await book.setStyle(id, "A6", { wrap: true });
    action("autoFitRows");
    let autoHeight = 0;
    for (let i = 0; i < 100; i++) {
      autoHeight = (await book.getMetadata()).sheets[0].rowHeights[5] ?? 0;
      if (autoHeight > 28) break;
      await new Promise(requestAnimationFrame);
    }
    equal(
      autoHeight > 28,
      true,
      "Auto-fit row height accounts for wrapped text",
    );
    tab("home");
    await editor.select(id, "C2:C4");
    await open("customFormat");
    (form().elements.namedItem("format") as HTMLInputElement).value = "0.000";
    await submit();
    equal(
      (await book.getRegion(id, "C2:C4")).cells.map(
        (c) => c.style.numberFormat,
      ),
      ["0.000", "0.000", "0.000"],
      "Custom formats apply to the complete selection",
    );
    await book.setStyle(id, "C2:C4", {
      border: { top: "#112233", bottom: "#112233" },
    });
    await open("borderOptions");
    (form().elements.namedItem("kind") as HTMLSelectElement).value = "none";
    await submit();
    equal(
      Object.values(
        (await book.getRegion(id, "C2")).cells[0].style.border!,
      ).every((value) => value === ""),
      true,
      "Border settings can remove borders without changing contents",
    );

    await book.setValues(id, "H1:I4", [
      ["Label", "Amount"],
      ["Keep", 10],
      ["Drop", 20],
      ["Keep", 30],
    ]);
    await book.setStyle(id, "H2", {
      fontSize: 24,
      bold: true,
      numberFormat: "0.00%",
    });
    await editor.select(id, "H2");
    tab("data");
    tab("home");
    equal(
      root.querySelector<HTMLSelectElement>('[data-style="fontSize"]')!.value,
      "24",
      "Toolbar restores selected font size after tab switches",
    );
    equal(
      root.querySelector('[data-action="bold"]')!.getAttribute("aria-pressed"),
      "true",
      "Toolbar indicates active bold state",
    );
    await editor.select(id, "H2:H3");
    equal(
      root.querySelector<HTMLSelectElement>('[data-style="fontSize"]')!.value,
      "—",
      "Mixed style selection is represented explicitly",
    );
    await book.filter(id, "H1:I4", [
      { column: 7, operator: "in", values: ["Keep"] },
    ]);
    await editor.select(id, "H1:I4");
    const grid = root.querySelector<HTMLCanvasElement>("canvas")!;
    grid.focus();
    const transfer = new DataTransfer();
    grid.dispatchEvent(
      new ClipboardEvent("copy", {
        bubbles: true,
        composed: true,
        clipboardData: transfer,
      }),
    );
    equal(
      transfer.getData("text/plain").includes("Drop"),
      false,
      "Filtered clipboard excludes invisible rows",
    );
    equal(
      root.querySelector(".stats")?.textContent?.includes("40") ??
        root.textContent!.includes("Sum 40"),
      true,
      "Filtered statistics use visible values",
    );
    await editor.select(id, "K10");
    grid.dispatchEvent(
      new ClipboardEvent("paste", {
        bubbles: true,
        composed: true,
        clipboardData: transfer,
      }),
    );
    for (
      let i = 0;
      i < 100 && (await book.getValues(id, "L12"))[0][0] !== 30;
      i++
    )
      await new Promise(requestAnimationFrame);
    equal(
      await book.getValues(id, "K10:L12"),
      [
        ["Label", "Amount"],
        ["Keep", 10],
        ["Keep", 30],
      ],
      "Internal paste compacts visible copied rows",
    );
    await book.filter(id, undefined, []);
    await book.importCsv("Code,Rate,Date\n00123,25%,2026-09-20", {
      sheetId: id,
      start: "H10",
      columns: ["text", "percent", "date"],
      header: true,
    });
    equal(
      (await book.getValues(id, "H11:I11"))[0],
      ["00123", 0.25],
      "Typed CSV preserves identifiers and converts percentages",
    );
    await book.setValidation(id, "H15:H16", {
      type: "list",
      values: ["Yes", "No"],
    });
    await editor.select(id, "H15");
    const choice = root.querySelector<HTMLSelectElement>(".validation-choice")!;
    equal(choice.hidden, false, "List validation exposes a dropdown");
    choice.value = "Yes";
    choice.dispatchEvent(new Event("change"));
    for (
      let i = 0;
      i < 100 && (await book.getValues(id, "H15"))[0][0] !== "Yes";
      i++
    )
      await new Promise(requestAnimationFrame);
    equal(
      await book.getValues(id, "H15"),
      [["Yes"]],
      "Dropdown writes through workbook validation",
    );
    const paste = async (data: DataTransfer) => {
      const revision = (await book.getMetadata()).revision;
      grid.focus();
      grid.dispatchEvent(
        new ClipboardEvent("paste", {
          bubbles: true,
          composed: true,
          clipboardData: data,
        }),
      );
      for (let i = 0; i < 100; i++) {
        if ((await book.getMetadata()).revision > revision) return;
        await new Promise(requestAnimationFrame);
      }
      throw new Error("Clipboard did not commit a change");
    };
    await book.setConditionalFormat(id, "H15:H16", {
      operator: "eq",
      value: "Yes",
      style: { italic: true, background: "#ff0000" },
    });
    await editor.select(id, "H15:H16");
    const template = new DataTransfer();
    grid.focus();
    grid.dispatchEvent(
      new ClipboardEvent("copy", {
        bubbles: true,
        composed: true,
        clipboardData: template,
      }),
    );
    await editor.select(id, "K15");
    await paste(template);
    await editor.select(id, "K15");
    equal(
      choice.hidden,
      false,
      "Copied templates retain their validation dropdown",
    );
    equal(
      (await book.getRegion(id, "K15")).cells[0].displayStyle?.italic,
      true,
      "Copied templates retain conditional formatting",
    );
    let invalidCopy = false;
    try {
      await book.setValues(id, "K15", [["Invalid"]]);
    } catch {
      invalidCopy = true;
    }
    equal(invalidCopy, true, "Copied validation rejects invalid values");
    await book.undo();
    await editor.select(id, "K15");
    equal(
      choice.hidden,
      true,
      "Undo removes pasted values and their rules together",
    );
    await book.redo();

    await editor.select(id, "M20");
    const external = new DataTransfer();
    external.setData("text/plain", "Bold\t25%\t2026-09-20");
    external.setData(
      "text/html",
      '<table><tr><td style="font-weight:bold;font-style:italic;text-align:center">Bold</td><td>25%</td><td>2026-09-20</td></tr></table>',
    );
    await paste(external);
    const externalCells = (await book.getRegion(id, "M20:O20")).cells;
    equal(
      externalCells.map((c) => c.value),
      ["Bold", 0.25, 46285],
      "External paste parses percentages and dates",
    );
    equal(
      externalCells.map((c) => c.style.numberFormat),
      [undefined, "0.00%", "yyyy-mm-dd"],
      "External paste keeps inferred number formats",
    );
    equal(
      [
        externalCells[0].style.bold,
        externalCells[0].style.italic,
        externalCells[0].style.align,
      ],
      [true, true, "center"],
      "External HTML paste preserves supported inline styles",
    );
    await book.undo();
    equal(
      (await book.getRegion(id, "M20:O20")).cells,
      [],
      "External values and formats undo as one transaction",
    );
    external.clearData("text/html");
    await paste(external);
    equal(
      (await book.getRegion(id, "N20:O20")).cells.map(
        (c) => c.style.numberFormat,
      ),
      ["0.00%", "yyyy-mm-dd"],
      "Plain TSV paste also retains inferred formats",
    );

    await editor.select(id, "M30");
    const spanning = new DataTransfer();
    spanning.setData("text/plain", "Heading\t\tRight\n\t\tBelow");
    spanning.setData(
      "text/html",
      '<table><tr><td rowspan="2" colspan="2" style="font-weight:bold">Heading<br>line 2</td><td>Right</td></tr><tr><td>Below</td></tr></table>',
    );
    await paste(spanning);
    equal(
      await book.getValues(id, "M30:O31"),
      [
        ["Heading\nline 2", null, "Right"],
        [null, null, "Below"],
      ],
      "HTML spans keep adjacent cells and line breaks in their correct positions",
    );
    equal(
      (await book.getMetadata()).sheets[0].merges.some(
        (m) => m.r1 === 29 && m.r2 === 30 && m.c1 === 12 && m.c2 === 13,
      ),
      true,
      "HTML spans create real merged cells",
    );
    equal(
      (await book.getRegion(id, "M30")).cells[0].style.bold,
      true,
      "HTML merged cells retain supported formatting",
    );
    await book.undo();
    equal(
      await book.getValues(id, "M30:O31"),
      [
        [null, null, null],
        [null, null, null],
      ],
      "HTML merged paste undoes atomically",
    );
    equal(
      (await book.getMetadata()).sheets[0].merges.some((m) => m.r1 === 29),
      false,
      "Undo removes the HTML merge",
    );
    await book.redo();

    await book.setValues(id, "M35", [["=1+1"]]);
    await book.setStyle(id, "M35", { bold: true });
    await editor.select(id, "M35");
    const formulaCopy = new DataTransfer();
    grid.focus();
    grid.dispatchEvent(
      new ClipboardEvent("copy", {
        bubbles: true,
        composed: true,
        clipboardData: formulaCopy,
      }),
    );
    await editor.select(id, "O35");
    const sameText = new DataTransfer();
    sameText.setData("text/plain", formulaCopy.getData("text/plain"));
    await paste(sameText);
    const externalSame = (await book.getRegion(id, "O35")).cells[0];
    equal(
      [externalSame.value, externalSame.formula, externalSame.style.bold],
      [2, undefined, undefined],
      "External matching text cannot replay internal formulas or styles",
    );
    await editor.select(id, "Q35");
    await paste(formulaCopy);
    const internalSame = (await book.getRegion(id, "Q35")).cells[0];
    equal(
      [internalSame.value, !!internalSame.formula, internalSame.style.bold],
      [2, true, true],
      "Identified internal clipboard retains formulas and styles",
    );

    await book.setValues(id, "O41", [["hidden destination"]]);
    await book.setDimensions(id, "row", [40], { hidden: true });
    await editor.select(id, "M38:M39");
    await book.setValues(id, "M38:M39", [["move"], ["both"]]);
    await editor.select(id, "M38:M39");
    const cut = new DataTransfer();
    grid.focus();
    grid.dispatchEvent(
      new ClipboardEvent("cut", {
        bubbles: true,
        composed: true,
        clipboardData: cut,
      }),
    );
    await editor.select(id, "O40");
    const beforeCut = await book.exportJSON();
    const errorCount = errors.length;
    grid.focus();
    grid.dispatchEvent(
      new ClipboardEvent("paste", {
        bubbles: true,
        composed: true,
        clipboardData: cut,
      }),
    );
    await waitUntil(() => errors.length > errorCount);
    equal(
      await book.exportJSON(),
      beforeCut,
      "Cut into hidden destination rows is rejected without changing source or destination",
    );
    await book.setDimensions(id, "row", [40], { hidden: false });

    await editor.select(id, "M25");
    tab("data");
    await open("validation");
    (form().elements.namedItem("type") as HTMLSelectElement).value = "date";
    (form().elements.namedItem("minimum") as HTMLInputElement).value =
      "2026-09-01";
    (form().elements.namedItem("maximum") as HTMLInputElement).value =
      "2026-09-30";
    await submit();
    await open("validation");
    equal(
      (form().elements.namedItem("minimum") as HTMLInputElement).value,
      "2026-09-01",
      "Date validation reopens with a readable date",
    );
    await submit();
    await open("validation");
    (form().elements.namedItem("minimum") as HTMLInputElement).value =
      "2026-09-10";
    await submit();
    equal(
      (await book.getMetadata()).sheets[0].validations!.find(
        (r) => r.range.r1 === 24,
      )?.minimum,
      46275,
      "Date validation can be confirmed unchanged and then edited",
    );
    const second = await book.duplicateSheet(id, "Office copy");
    await editor.select(second, "J18");
    await editor.select(id, "H15");
    root.querySelectorAll<HTMLButtonElement>('[role="tab"]')[1].click();
    await waitUntil(() => editor.getSelection().sheetId === second);
    await waitUntil(
      () => root.querySelector<HTMLInputElement>(".address")?.value === "J18",
    );
    equal(
      editor.getSelection().range.r1,
      17,
      "Sheet tabs restore the previous selection",
    );
    await editor.select(id, "A1");

    await editor.select(id, "H20");
    const draft = root.querySelector<HTMLInputElement>(".formula")!;
    draft.focus();
    draft.value = "=SUM(";
    draft.dispatchEvent(new InputEvent("input", { bubbles: true }));
    const scroll = root.querySelector<HTMLElement>(".scroll")!;
    scroll.scrollTop += 100;
    scroll.dispatchEvent(new Event("scroll"));
    await book.getMetadata();
    equal(
      editor.getEditState().editing,
      true,
      "Formula draft survives scrolling to offscreen references",
    );
    equal(
      await book.getValues(id, "H20"),
      [[null]],
      "Scrolling does not store incomplete formulas",
    );
    editor.cancelEdit();
    await editor.select(id, "F1");
    const formula = root.querySelector<HTMLInputElement>(".formula")!;
    formula.focus();
    formula.value = "=SUM(C2:C4)";
    formula.dispatchEvent(new InputEvent("input", { bubbles: true }));
    equal(
      editor.getEditState().dirty,
      true,
      "Host can detect an uncommitted draft",
    );
    await editor.commitEdit();
    equal(
      await book.getValues(id, "F1"),
      [[60]],
      "Host commit waits for calculation before export",
    );
    equal(
      editor.getEditState(),
      { editing: false, dirty: false, pending: false },
      "Successful host commit leaves a clean editor",
    );
    formula.focus();
    formula.value = "999";
    formula.dispatchEvent(new InputEvent("input", { bubbles: true }));
    editor.cancelEdit();
    equal(
      await book.getValues(id, "F1"),
      [[60]],
      "Host can discard a draft without changing the workbook",
    );
    formula.blur();
    formula.focus();
    formula.value = "128";
    formula.dispatchEvent(new InputEvent("input", { bubbles: true }));
    formula.dispatchEvent(
      new CompositionEvent("compositionstart", { bubbles: true }),
    );
    let rejected = false;
    try {
      await editor.commitEdit();
    } catch {
      rejected = true;
    }
    equal(rejected, true, "Host commit rejects unfinished IME composition");
    rejected = false;
    try {
      await editor.destroy({ commit: true });
    } catch {
      rejected = true;
    }
    equal(
      [rejected, host.children.length],
      [true, 1],
      "Failed commit-and-destroy preserves the mounted view",
    );
    formula.dispatchEvent(
      new CompositionEvent("compositionend", { bubbles: true }),
    );
    await editor.destroy({ commit: true });
    equal(
      await book.getValues(id, "F1"),
      [[128]],
      "Commit-and-destroy preserves the final draft",
    );
    equal(host.children.length, 0, "Commit-and-destroy releases the view");
  } finally {
    editor.destroy();
    host.remove();
    await book.dispose();
  }
}

import { describe, it, expect } from "vitest";
import { WorkbookModel } from "../src/core/model";
import { keyOf, parseRange } from "../src/core/address";
import {
  dataRegion,
  distinctValues,
  findNext,
  navigationTarget,
} from "../src/core/queries";
import { paintCommand } from "../src/editor/format-painter";

it("navigates visible data boundaries, blank gaps, empty formulas and spilled cells", () => {
  const { model, id } = setup();
  set(model, id, "A1:A7", [[1], [2], [3], [null], [5], [null], ['=""']]);
  expect(navigationTarget(model, id, 0, 0, "down")).toEqual({
    row: 2,
    column: 0,
  });
  expect(navigationTarget(model, id, 2, 0, "down")).toEqual({
    row: 4,
    column: 0,
  });
  expect(navigationTarget(model, id, 4, 0, "down")).toEqual({
    row: 6,
    column: 0,
  });
  expect(navigationTarget(model, id, 6, 0, "down")).toEqual({
    row: 99,
    column: 0,
  });
  model.sheet(id).meta.hiddenRows = [3];
  expect(navigationTarget(model, id, 0, 0, "down")).toEqual({
    row: 4,
    column: 0,
  });
  model.sheet(id).meta.filteredRows = [4];
  expect(navigationTarget(model, id, 2, 0, "down")).toEqual({
    row: 6,
    column: 0,
  });
  set(model, id, "C1", [["=SEQUENCE(1,4)"]]);
  expect(navigationTarget(model, id, 0, 2, "right")).toEqual({
    row: 0,
    column: 5,
  });
  expect(navigationTarget(model, id, 0, 5, "left")).toEqual({
    row: 0,
    column: 2,
  });
});

it("paints repeating styles and conditional formats without changing values or validation, in one undo", () => {
  const { model, id } = setup();
  set(model, id, "A1:B1", [[1, 2]]);
  set(model, id, "D3:F3", [[7, 8, "=D3+E3"]]);
  model.execute([
    {
      type: "style",
      sheetId: id,
      range: parseRange("A1"),
      style: { bold: true },
    },
    {
      type: "style",
      sheetId: id,
      range: parseRange("D3:F3"),
      style: { italic: true },
    },
    {
      type: "validation",
      sheetId: id,
      range: parseRange("D3:F3"),
      rule: { type: "whole", minimum: 0, maximum: 20 },
    },
    {
      type: "conditionalFormat",
      sheetId: id,
      range: parseRange("A1"),
      rule: { operator: "gt", value: 5, style: { background: "#ff0000" } },
    },
  ]);
  const before = model.snapshot();
  model.execute([
    paintCommand(
      {
        rows: [0],
        columns: [0, 1],
        cells: model.region(id, parseRange("A1:B1")).cells,
        conditionalFormats: model.sheet(id).meta.conditionalFormats ?? [],
      },
      id,
      [2],
      [3, 4, 5],
    ),
  ]);
  const cells = model.region(id, parseRange("D3:F3")).cells;
  expect(cells.map((cell) => cell.value)).toEqual([7, 8, 15]);
  expect(cells.map((cell) => !!cell.style.bold)).toEqual([true, false, true]);
  expect(cells.every((cell) => !cell.style.italic)).toBe(true);
  expect(cells[2].displayStyle?.background).toBe("#ff0000");
  expect(() => set(model, id, "E3", [[21]])).toThrow(/Validation/);
  model.undo();
  expect(model.snapshot().sheets).toEqual(before.sheets);
});

it("expands a data region without crossing a blank separator and finds matches in row order", () => {
  const { model, id } = setup();
  set(model, id, "A3:B3", [["same", 3]]);
  set(model, id, "A1:B2", [
    ["Name", "Value"],
    ["same", 2],
  ]);
  set(model, id, "A5:B5", [["separate", 99]]);
  expect(dataRegion(model, id, parseRange("B2"))).toEqual(parseRange("A1:B3"));
  set(model, id, "C2", [["=Z99"]]);
  expect(dataRegion(model, id, parseRange("B2"))).toEqual(parseRange("A1:C3"));
  model.undo();
  expect(dataRegion(model, id, parseRange("B1:B100"))).toEqual(
    parseRange("A1:B5"),
  );
  expect(findNext(model, id, "same")?.address).toBe("A2");
  expect(findNext(model, id, "same", keyOf(1, 0))?.address).toBe("A3");
  expect(findNext(model, id, "same", keyOf(2, 0))?.address).toBe("A2");
  set(model, id, "D1", [["=SEQUENCE(3,1,10)"]]);
  expect(findNext(model, id, "12", undefined, false, true)?.address).toBe("D3");
});

it("combines value-list filters with conditions, preserves blanks and undoes the whole filter", () => {
  const { model, id } = setup();
  set(model, id, "A1:B5", [
    ["Name", "Score"],
    ["Amy", 10],
    ["amy", 20],
    [null, 30],
    ["Zoe", 40],
  ]);
  const range = parseRange("A1:B5");
  expect(distinctValues(model, id, range, 0).values).toEqual([
    "",
    "amy",
    "Zoe",
  ]);
  model.execute([
    {
      type: "filter",
      sheetId: id,
      range,
      rules: [
        { column: 0, operator: "in", values: ["Amy", ""] },
        { column: 1, operator: "gte", value: 20 },
      ],
    },
  ]);
  expect(model.sheet(id).meta.filteredRows).toEqual([1, 4]);
  model.undo();
  expect(model.sheet(id).meta.filteredRows).toBeUndefined();
  model.redo();
  expect(model.sheet(id).meta.filteredRows).toEqual([1, 4]);
  expect(() =>
    model.execute([
      {
        type: "filter",
        sheetId: id,
        range,
        rules: [
          { column: 0, operator: "in", values: [] },
          { column: 0, operator: "eq", value: "Amy" },
        ],
      },
    ]),
  ).toThrow();
  expect(model.sheet(id).meta.filteredRows).toEqual([1, 4]);
});
function setup() {
  const model = new WorkbookModel({
    sheets: [
      { name: "Data", rows: 100, columns: 30 },
      { name: "Summary", rows: 100, columns: 30 },
    ],
  });
  const id = model.sheets[0].meta.id;
  return { model, id, second: model.sheets[1].meta.id };
}
const set = (
  model: WorkbookModel,
  id: string,
  range: string,
  values: (string | number | boolean | null)[][],
) =>
  model.execute([
    { type: "setValues", sheetId: id, range: parseRange(range), values },
  ]);
const get = (model: WorkbookModel, id: string, cell: string) => {
  const r = parseRange(cell);
  return model.engine.get(id, keyOf(r.r1, r.c1));
};
it("pastes immutable region snapshots repeatedly with formulas, literals and exact styles", () => {
  const { model, id } = setup();
  set(model, id, "A1:B2", [
    [1.23456789, "=A1+$A$1"],
    [null, "=1/0"],
  ]);
  model.execute([
    {
      type: "setValues",
      sheetId: id,
      range: parseRange("C1"),
      values: [["=literal"]],
      parseFormulas: false,
    },
    {
      type: "style",
      sheetId: id,
      range: parseRange("B1"),
      style: { bold: true, numberFormat: "0.00" },
    },
  ]);
  const source = parseRange("A1:C2");
  const cells = model.region(id, source).cells;
  const before = structuredClone(cells);
  const paste = (targetRow: number, targetColumn: number) =>
    model.execute([
      {
        type: "paste",
        sheetId: id,
        source,
        cells,
        targetRow,
        targetColumn,
      },
    ]);
  paste(1, 0);
  expect(get(model, id, "A2")).toBe(1.23456789);
  expect(get(model, id, "C2")).toBe("=literal");
  set(model, id, "A1", [[9]]);
  model.execute([
    {
      type: "style",
      sheetId: id,
      range: parseRange("G1:I2"),
      style: { italic: true },
    },
  ]);
  paste(0, 6);
  expect(get(model, id, "G1")).toBe(1.23456789);
  expect(get(model, id, "H1")).toBe(10.23456789);
  expect(get(model, id, "H2")).toEqual({ error: "#DIV/0!" });
  const pasted = model.region(id, parseRange("H1:I1")).cells;
  expect(pasted[0].formula).toBe("=(G1+$A$1)");
  expect(pasted[0].style).toEqual({ bold: true, numberFormat: "0.00" });
  expect(pasted[1].formula).toBeUndefined();
  expect(get(model, id, "I1")).toBe("=literal");
  expect(model.sheet(id).cells.has(keyOf(1, 8))).toBe(false);
  expect(cells).toEqual(before);
  model.undo();
  expect(
    model
      .region(id, parseRange("G1:I2"))
      .cells.every((cell) => cell.style.italic && cell.value === null),
  ).toBe(true);
  model.redo();
  expect(get(model, id, "G1")).toBe(1.23456789);
});

it("rolls back a snapshot paste rejected for partial spilled results", () => {
  const { model, id } = setup();
  set(model, id, "A1", [["=SEQUENCE(2)"]]);
  set(model, id, "C1:C2", [[8], [9]]);
  const source = parseRange("A2");
  const before = model.snapshot();
  expect(() =>
    model.execute([
      {
        type: "paste",
        sheetId: id,
        source,
        cells: model.region(id, source).cells,
        targetRow: 0,
        targetColumn: 2,
      },
    ]),
  ).toThrow(/spill/i);
  expect(model.snapshot()).toEqual(before);
});
describe("workbook operations", () => {
  it("recalculates transitive cross-sheet references and supports atomic undo", () => {
    const { model, id, second } = setup();
    set(model, id, "A1:B2", [
      [2, 3],
      ["=A1+B1", "=A2*2"],
    ]);
    set(model, second, "A1", [["=Data!B2"]]);
    expect(get(model, second, "A1")).toBe(10);
    set(model, id, "A1", [[7]]);
    expect(get(model, second, "A1")).toBe(20);
    model.undo();
    expect(get(model, second, "A1")).toBe(10);
    model.redo();
    expect(get(model, second, "A1")).toBe(20);
  });
  it("rolls back all commands if a later command fails", () => {
    const { model, id } = setup();
    set(model, id, "A1", [[4]]);
    const revision = model.revision;
    const before = model.snapshot();
    expect(() =>
      model.execute([
        {
          type: "setValues",
          sheetId: id,
          range: parseRange("A1"),
          values: [[99]],
        },
        {
          type: "style",
          sheetId: id,
          range: parseRange("A1"),
          style: { bold: true, color: "#abc" },
        },
        { type: "merge", sheetId: id, range: parseRange("A1:ZZ1000") },
      ]),
    ).toThrow();
    expect(get(model, id, "A1")).toBe(4);
    expect(model.revision).toBe(revision);
    expect(model.snapshot()).toEqual(before);
    model.undo();
    expect(get(model, id, "A1")).toBe(null);
  });
  it("adjusts relative and absolute references during copy, fill and structural insert", () => {
    const { model, id } = setup();
    set(model, id, "A1:B3", [
      [2, "=A1+$A$1"],
      [3, null],
      [4, null],
    ]);
    model.execute([
      {
        type: "copy",
        sheetId: id,
        range: parseRange("B1"),
        targetSheetId: id,
        targetRow: 1,
        targetColumn: 1,
      },
    ]);
    expect(get(model, id, "B2")).toBe(5);
    model.execute([
      { type: "structure", sheetId: id, axis: "row", index: 0, count: 1 },
    ]);
    expect(get(model, id, "B3")).toBe(5);
    model.undo();
    expect(get(model, id, "B2")).toBe(5);
  });
  it("renames references and preserves text literals", () => {
    const { model, id, second } = setup();
    set(model, id, "A1", [[10]]);
    set(model, second, "A1", [[`=IF(Data!A1=10,"Data!A1","")`]]);
    model.execute([{ type: "renameSheet", sheetId: id, name: "My Data" }]);
    expect(get(model, second, "A1")).toBe("Data!A1");
    expect(model.sheets[1].cells.get(0)?.formula).toContain("'My Data'!A1");
  });
  it("materializes dynamic arrays, rejects partial edits and clears stale spill cells", () => {
    const { model, id } = setup();
    set(model, id, "A1", [["=SEQUENCE(3,2,10,2)"]]);
    expect(get(model, id, "B3")).toBe(20);
    expect(() => set(model, id, "B2", [[5]])).toThrow(/spill/i);
    set(model, id, "A1", [["=SEQUENCE(1,2)"]]);
    expect(get(model, id, "B1")).toBe(2);
    expect(get(model, id, "A3")).toBe(null);
  });
  it("recovers a blocked spill and recalculates spill references", () => {
    const { model, id } = setup();
    set(model, id, "A2", [[99]]);
    set(model, id, "A1", [["=SEQUENCE(3)"]]);
    expect(get(model, id, "A1")).toEqual({ error: "#SPILL!" });
    model.execute([{ type: "clear", sheetId: id, range: parseRange("A2") }]);
    expect(get(model, id, "A3")).toBe(3);
    set(model, id, "C1", [["=SUM(A1#)"]]);
    expect(get(model, id, "C1")).toBe(6);
    set(model, id, "A1", [["=SEQUENCE(4)"]]);
    expect(get(model, id, "C1")).toBe(10);
  });
  it("keeps filters derived while undo restores the edits that caused them", () => {
    const { model, id } = setup();
    set(model, id, "A1:B4", [
      ["name", "total"],
      ["c", 3],
      ["a", 1],
      ["b", 2],
    ]);
    model.execute([
      {
        type: "sort",
        sheetId: id,
        range: parseRange("A1:B4"),
        keys: [{ column: 1, direction: "asc" }],
        header: true,
      },
      {
        type: "filter",
        sheetId: id,
        range: parseRange("A1:B4"),
        rules: [{ column: 1, operator: "gt", value: 1 }],
      },
    ]);
    expect(get(model, id, "A2")).toBe("a");
    expect(model.sheet(id).meta.filteredRows).toEqual([1]);
    model.undo();
    expect(get(model, id, "A2")).toBe("c");
    expect(model.sheet(id).meta.filter).toBeUndefined();
  });
  it("reports unsupported functions and cycles without using imported caches", () => {
    const { model, id } = setup();
    set(model, id, "A1:B1", [["=FUTURE_FUNCTION(3)", "=A1"]]);
    expect(get(model, id, "B1")).toEqual({ error: "#NAME?" });
    set(model, id, "A1:B1", [["=B1", "=A1"]]);
    expect(get(model, id, "A1")).toEqual({ error: "#CYCLE!" });
    expect(
      model.getDiagnostics().some((d) => d.code === "CIRCULAR_REFERENCE"),
    ).toBe(true);
  });
  it("updates ranges when deleting a referenced first row", () => {
    const { model, id } = setup();
    set(model, id, "A1:A4", [[1], [2], [3], [4]]);
    set(model, id, "C8", [["=SUM(A1:A4)"]]);
    model.execute([
      {
        type: "structure",
        sheetId: id,
        axis: "row",
        index: 0,
        count: 1,
        delete: true,
      },
    ]);
    expect(get(model, id, "C7")).toBe(9);
  });
  it("preserves dates, names, styles and values through snapshots", () => {
    const { model, id } = setup();
    set(model, id, "A1", [[45678]]);
    model.execute([
      {
        type: "style",
        sheetId: id,
        range: parseRange("A1"),
        style: { numberFormat: "yyyy-mm-dd", bold: true },
      },
      {
        type: "defineName",
        name: "StartDate",
        sheetId: id,
        range: parseRange("A1"),
      },
    ]);
    set(model, id, "B1", [["=YEAR(StartDate)"]]);
    const restored = new WorkbookModel({ snapshot: model.snapshot() });
    expect(get(restored, id, "B1")).toBe(2025);
    expect(restored.region(id, parseRange("A1")).cells[0].style.bold).toBe(
      true,
    );
  });
  it("rolls back an aborted large paste", async () => {
    const model = new WorkbookModel({
        sheets: [{ name: "Data", rows: 5000, columns: 10 }],
      }),
      id = model.sheets[0].meta.id;
    await expect(
      model.executeAsync(
        [
          {
            type: "setValues",
            sheetId: id,
            range: parseRange("A1:A4000"),
            values: Array.from({ length: 4000 }, () => [7]),
          },
        ],
        async () => {
          throw new Error("cancelled");
        },
      ),
    ).rejects.toThrow("cancelled");
    expect(model.sheet(id).cells.size).toBe(0);
  });
});
it("recalculates a dependent created before its spill anchor", () => {
  const { model, id } = setup();
  set(model, id, "D1", [["=B2*2"]]);
  set(model, id, "A1", [["=SEQUENCE(2,2)"]]);
  expect(get(model, id, "D1")).toBe(8);
});
it("moves formulas across sheets without retargeting unrelated local references", () => {
  const { model, id, second } = setup();
  set(model, id, "A1:B1", [[8, "=A1*2"]]);
  set(model, second, "A1", [[100]]);
  model.execute([
    {
      type: "copy",
      sheetId: id,
      range: parseRange("B1"),
      targetSheetId: second,
      targetRow: 0,
      targetColumn: 1,
      cut: true,
    },
  ]);
  expect(get(model, second, "B1")).toBe(16);
  model.undo();
  expect(get(model, id, "B1")).toBe(16);
});
it("does not spill through merged cells", () => {
  const { model, id } = setup();
  model.execute([{ type: "merge", sheetId: id, range: parseRange("A2:B2") }]);
  set(model, id, "A1", [["=SEQUENCE(3)"]]);
  expect(get(model, id, "A1")).toEqual({ error: "#SPILL!" });
  model.execute([
    { type: "merge", sheetId: id, range: parseRange("A2:B2"), unmerge: true },
  ]);
  expect(get(model, id, "A3")).toBe(3);
  model.execute([{ type: "merge", sheetId: id, range: parseRange("A1:B3") }]);
  expect(get(model, id, "A1")).toEqual({ error: "#SPILL!" });
});
it("preserves whole-column references across column insertion and deletion", () => {
  const { model, id } = setup();
  set(model, id, "B1:B2", [[2], [3]]);
  set(model, id, "D1", [["=SUM(B:B)"]]);
  model.execute([
    {
      type: "structure",
      sheetId: id,
      axis: "column",
      index: 0,
      count: 1,
      delete: true,
    },
  ]);
  expect(get(model, id, "C1")).toBe(5);
});
it("keeps deleted-sheet references broken when a new sheet reuses the old name", () => {
  const { model, id, second } = setup();
  set(model, id, "A1", [[8]]);
  set(model, second, "A1", [["=Data!A1"]]);
  model.execute([
    { type: "deleteSheet", sheetId: id },
    { type: "addSheet", name: "Data" },
  ]);
  expect(get(model, second, "A1")).toEqual({ error: "#REF!" });
  model.undo();
  expect(get(model, second, "A1")).toBe(8);
});

it("sorts by formulas updated earlier in the same atomic transaction", () => {
  const model = new WorkbookModel();
  const id = model.sheets[0].meta.id;
  model.execute([
    {
      type: "setValues",
      sheetId: id,
      range: parseRange("A1:B2"),
      values: [
        [2, "=A1*2"],
        [1, "=A2*2"],
      ],
    },
  ]);
  model.execute([
    { type: "setValues", sheetId: id, range: parseRange("A1"), values: [[0]] },
    {
      type: "sort",
      sheetId: id,
      range: parseRange("A1:B2"),
      keys: [{ column: 1, direction: "desc" }],
    },
  ]);
  expect(model.engine.get(id, keyOf(0, 0))).toBe(1);
  expect(model.engine.get(id, keyOf(0, 1))).toBe(2);
  model.undo();
  expect(model.engine.get(id, keyOf(0, 0))).toBe(2);
  expect(model.engine.get(id, keyOf(0, 1))).toBe(4);
});
it("reads updated named ranges before dependent value copies and sorts in the same transaction", () => {
  const { model, id } = setup();
  set(model, id, "A1:B2", [
    [2, "=SUM(Amount)"],
    [3, "=5-SUM(Amount)"],
  ]);
  model.execute([
    {
      type: "defineName",
      sheetId: id,
      name: "Amount",
      range: parseRange("A1"),
    },
  ]);
  model.execute([
    {
      type: "defineName",
      sheetId: id,
      name: "Amount",
      range: parseRange("A2"),
    },
    {
      type: "copy",
      sheetId: id,
      range: parseRange("B1"),
      targetSheetId: id,
      targetRow: 0,
      targetColumn: 2,
      valuesOnly: true,
    },
    {
      type: "sort",
      sheetId: id,
      range: parseRange("B1:B2"),
      keys: [{ column: 1, direction: "asc" }],
    },
  ]);
  expect(get(model, id, "C1")).toBe(3);
  expect(get(model, id, "B1")).toBe(2);
  expect(get(model, id, "B2")).toBe(3);
  model.undo();
  expect(get(model, id, "C1")).toBe(null);
  expect(get(model, id, "B1")).toBe(2);
  expect(get(model, id, "B2")).toBe(3);
  expect(model.names.AMOUNT.range).toEqual(parseRange("A1"));
  model.redo();
  expect(get(model, id, "C1")).toBe(3);
  expect(model.names.AMOUNT.range).toEqual(parseRange("A2"));
});

it.each([false, true])(
  "snapshots overlapping values-only copies before mutation (cut=%s)",
  (cut) => {
    const { model, id } = setup();
    set(model, id, "A1:A3", [[1], ["=A1+1"], [3]]);
    model.execute([
      {
        type: "style",
        sheetId: id,
        range: parseRange("A2"),
        style: { bold: true },
      },
    ]);
    const targetStyle = model.sheet(id).cells.get(keyOf(1, 0))?.style;
    model.execute([
      {
        type: "copy",
        sheetId: id,
        range: parseRange("A1:A3"),
        targetSheetId: id,
        targetRow: 1,
        targetColumn: 0,
        valuesOnly: true,
        cut,
      },
    ]);
    const values = () =>
      ["A1", "A2", "A3", "A4"].map((cell) => get(model, id, cell));
    const expected = [cut ? null : 1, 1, 2, 3];
    expect(values()).toEqual(expected);
    expect(model.sheet(id).cells.get(keyOf(1, 0))?.style).toBe(targetStyle);
    expect(model.sheet(id).cells.get(keyOf(2, 0))?.formula).toBeUndefined();
    model.undo();
    expect(values()).toEqual([1, 2, 3, null]);
    expect(model.sheet(id).cells.get(keyOf(1, 0))?.formula).toBe("=A1+1");
    model.redo();
    expect(values()).toEqual(expected);
  },
);

it("adjusts reversed ranges on deletion while preserving endpoint anchors", () => {
  const { model, id } = setup();
  set(model, id, "A1:C3", [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9],
  ]);
  set(model, id, "E5", [["=SUM($C3:A$1)"]]);
  expect(get(model, id, "E5")).toBe(45);
  model.execute([
    {
      type: "structure",
      sheetId: id,
      axis: "row",
      index: 1,
      count: 1,
      delete: true,
    },
  ]);
  expect(get(model, id, "E4")).toBe(30);
  expect(model.sheet(id).cells.get(keyOf(3, 4))?.formula).toBe("=SUM($C2:A$1)");
  model.execute([
    {
      type: "structure",
      sheetId: id,
      axis: "column",
      index: 1,
      count: 1,
      delete: true,
    },
  ]);
  expect(get(model, id, "D4")).toBe(20);
  expect(model.sheet(id).cells.get(keyOf(3, 3))?.formula).toBe("=SUM($B2:A$1)");
  model.undo();
  expect(get(model, id, "E4")).toBe(30);
  model.redo();
  expect(get(model, id, "D4")).toBe(20);
});

it("calculates and copies sheet-qualified whole-row ranges with absolute anchors", () => {
  const { model, id, second } = setup();
  set(model, id, "A1:B3", [
    [1, 2],
    [3, 4],
    [5, 6],
  ]);
  set(model, second, "A5:C5", [
    ["=SUM(Data!$1:$1)", "=SUM(Data!1:1)", "=SUM(Data!$1:2)"],
  ]);
  expect(["A5", "B5", "C5"].map((cell) => get(model, second, cell))).toEqual([
    3, 3, 10,
  ]);
  model.execute([
    {
      type: "copy",
      sheetId: second,
      range: parseRange("A5:C5"),
      targetSheetId: second,
      targetRow: 5,
      targetColumn: 0,
    },
  ]);
  expect(["A6", "B6", "C6"].map((cell) => get(model, second, cell))).toEqual([
    3, 7, 21,
  ]);
  expect(model.sheet(second).cells.get(keyOf(5, 0))?.formula).toBe(
    "=SUM('Data'!$1:$1)",
  );
  model.execute([{ type: "renameSheet", sheetId: id, name: "My Data" }]);
  expect(get(model, second, "C6")).toBe(21);
  expect(model.sheet(second).cells.get(keyOf(5, 2))?.formula).toBe(
    "=SUM('My Data'!$1:3)",
  );
  model.execute([
    { type: "structure", sheetId: id, axis: "row", index: 0, count: 1 },
  ]);
  expect(get(model, second, "A6")).toBe(3);
  expect(model.sheet(second).cells.get(keyOf(5, 0))?.formula).toBe(
    "=SUM('My Data'!$2:$2)",
  );
});

it("validates complete transactions and enforces locked cells across direct edits and copies", () => {
  const { model, id } = setup();
  model.execute([
    {
      type: "validation",
      sheetId: id,
      range: parseRange("A1:A3"),
      rule: { type: "whole", minimum: 0, maximum: 10, allowBlank: false },
    },
  ]);
  expect(() => set(model, id, "A1:A2", [[3], [11]])).toThrow(/Validation/);
  expect(model.region(id, parseRange("A1:A2")).cells).toEqual([]);
  set(model, id, "A1:A2", [[3], [4]]);
  model.execute([
    {
      type: "style",
      sheetId: id,
      range: parseRange("A1"),
      style: { locked: false },
    },
    { type: "protect", sheetId: id, enabled: true },
  ]);
  set(model, id, "A1", [[5]]);
  expect(() => set(model, id, "A2", [[6]])).toThrow(/Protected/);
  expect(() =>
    model.execute([
      {
        type: "copy",
        sheetId: id,
        range: parseRange("A1"),
        targetSheetId: id,
        targetRow: 1,
        targetColumn: 0,
        valuesOnly: true,
      },
    ]),
  ).toThrow(/Protected/);
  expect(() =>
    model.execute([
      { type: "setFormula", sheetId: id, row: 0, column: 0, formula: "=99" },
    ]),
  ).toThrow(/Validation/);
  expect(model.engine.get(id, keyOf(0, 0))).toBe(5);
  model.undo();
  expect(model.engine.get(id, keyOf(0, 0))).toBe(3);
  model.redo();
  expect(model.engine.get(id, keyOf(0, 0))).toBe(5);
});

it("pastes visible rows, transposes formulas, preserves merged ranges and rolls back a blocked destination", () => {
  const { model, id } = setup();
  set(model, id, "A1:B3", [
    [1, "=A1*2"],
    [2, "=A2*2"],
    [3, "=A3*2"],
  ]);
  const source = parseRange("A1:B3"),
    cells = model.region(id, source).cells;
  model.execute([
    {
      type: "paste",
      sheetId: id,
      source,
      cells,
      sourceRows: [0, 2],
      targetRow: 5,
      targetColumn: 0,
      mode: "values",
    },
  ]);
  expect(model.engine.get(id, keyOf(6, 1))).toBe(6);
  model.execute([
    {
      type: "paste",
      sheetId: id,
      source: parseRange("A1:B1"),
      cells: cells.filter((c) => c.row === 0),
      targetRow: 9,
      targetColumn: 2,
      transpose: true,
    },
  ]);
  expect(model.region(id, parseRange("C11")).cells[0].formula).toBe("=(B11*2)");
  model.execute([
    {
      type: "paste",
      sheetId: id,
      source: parseRange("A1:B1"),
      cells: cells.filter((c) => c.row === 0 && c.column === 0),
      merges: [parseRange("A1:B1")],
      targetRow: 12,
      targetColumn: 0,
    },
  ]);
  expect(model.sheet(id).meta.merges).toContainEqual(parseRange("A13:B13"));
  model.undo();
  expect(model.sheet(id).meta.merges).toEqual([]);
  model.execute([
    {
      type: "setFormula",
      sheetId: id,
      row: 20,
      column: 0,
      formula: "=SEQUENCE(2)",
    },
  ]);
  expect(() =>
    model.execute([
      {
        type: "paste",
        sheetId: id,
        source: parseRange("A1"),
        cells: cells.filter((c) => c.row === 0 && c.column === 0),
        targetRow: 21,
        targetColumn: 0,
      },
    ]),
  ).toThrow(/spilled/);
});

it("manages sheet visibility, duplicate formulas, named ranges and conditional formats with undo", () => {
  const { model, id } = setup();
  set(model, id, "A1:B1", [[5, "=A1*2"]]);
  model.execute([
    { type: "duplicateSheet", sheetId: id, id: "copy", name: "Copy" },
    { type: "sheetVisibility", sheetId: id, hidden: true },
  ]);
  expect(model.engine.get("copy", keyOf(0, 1))).toBe(10);
  for (const other of model.sheets.filter(
    (s) => s.meta.id !== "copy" && !s.meta.hidden,
  ))
    model.execute([
      { type: "sheetVisibility", sheetId: other.meta.id, hidden: true },
    ]);
  expect(() =>
    model.execute([{ type: "sheetVisibility", sheetId: "copy", hidden: true }]),
  ).toThrow(/last visible/);
  model.execute([
    { type: "defineName", sheetId: id, range: parseRange("A1"), name: "Input" },
    {
      type: "setFormula",
      sheetId: "copy",
      row: 1,
      column: 0,
      formula: '=Input+LEN("Input")',
    },
  ]);
  model.execute([{ type: "renameName", name: "Input", newName: "Amount" }]);
  expect(model.sheet("copy").cells.get(keyOf(1, 0))?.formula).toContain(
    'LEN("Input")',
  );
  expect(model.engine.get("copy", keyOf(1, 0))).toBe(10);
  model.execute([{ type: "deleteName", name: "Amount" }]);
  expect(model.engine.get("copy", keyOf(1, 0))).toEqual({ error: "#REF!" });
  model.undo();
  expect(model.engine.get("copy", keyOf(1, 0))).toBe(10);
  model.execute([
    {
      type: "conditionalFormat",
      sheetId: "copy",
      range: parseRange("A1:B3"),
      rule: { operator: "gt", value: 7, style: { background: "#ff0000" } },
    },
  ]);
  expect(
    model.region("copy", parseRange("B1")).cells[0].displayStyle?.background,
  ).toBe("#ff0000");
  set(model, "copy", "A1", [[2]]);
  expect(
    model.region("copy", parseRange("B1")).cells[0].displayStyle,
  ).toBeUndefined();
});

it("preserves merges in skipped paste rows and columns and rejects partially covered merges atomically", () => {
  const { model, id } = setup();
  set(model, id, "A1:B2", [
    [1, 2],
    [3, 4],
  ]);
  model.execute([
    { type: "merge", sheetId: id, range: parseRange("D11:E11") },
    { type: "merge", sheetId: id, range: parseRange("I10:I12") },
    {
      type: "dimensions",
      sheetId: id,
      axis: "row",
      indexes: [10],
      hidden: true,
    },
    {
      type: "dimensions",
      sheetId: id,
      axis: "column",
      indexes: [8],
      hidden: true,
    },
  ]);
  const merges = structuredClone(model.sheet(id).meta.merges);
  const cells = model.region(id, parseRange("A1:B2")).cells;
  model.execute([
    {
      type: "paste",
      sheetId: id,
      source: parseRange("A1:B2"),
      cells,
      targetRow: 9,
      targetColumn: 3,
      targetRows: [9, 11],
    },
    {
      type: "paste",
      sheetId: id,
      source: parseRange("A1:B2"),
      cells,
      targetRow: 9,
      targetColumn: 7,
      targetColumns: [7, 9],
    },
  ]);
  expect(model.sheet(id).meta.merges).toEqual(merges);
  expect(get(model, id, "E12")).toBe(4);
  model.undo();
  expect(get(model, id, "D10")).toBe(null);
  expect(model.sheet(id).meta.merges).toEqual(merges);
  model.execute([{ type: "merge", sheetId: id, range: parseRange("D15:D17") }]);
  const before = model.snapshot();
  expect(() =>
    model.execute([
      {
        type: "paste",
        sheetId: id,
        source: parseRange("A1:A2"),
        cells,
        targetRow: 14,
        targetColumn: 3,
        targetRows: [14, 16],
      },
    ]),
  ).toThrow(/part of a merged/);
  expect(model.snapshot()).toEqual(before);
});

it("copies template rules across visible transposed cells, API copies and cuts with undo and validation rollback", () => {
  const { model, id } = setup();
  set(model, id, "A1:A3", [[2], [4], [8]]);
  model.execute([
    {
      type: "validation",
      sheetId: id,
      range: parseRange("A1:A3"),
      rule: { type: "whole", minimum: 1, maximum: 10 },
    },
    {
      type: "conditionalFormat",
      sheetId: id,
      range: parseRange("A1:A3"),
      rule: { operator: "gt", value: 5, style: { background: "#ff0000" } },
    },
    {
      type: "validation",
      sheetId: id,
      range: parseRange("D10:F10"),
      rule: { type: "textLength", minimum: 5 },
    },
  ]);
  const rules = structuredClone(model.sheet(id).meta);
  const command = {
    type: "paste" as const,
    sheetId: id,
    source: parseRange("A1:A3"),
    cells: model.region(id, parseRange("A1:A3")).cells,
    validations: rules.validations,
    conditionalFormats: rules.conditionalFormats,
    sourceRows: [0, 2],
    targetRow: 9,
    targetColumn: 3,
    targetColumns: [3, 5],
    transpose: true,
  };
  model.execute([command]);
  expect(get(model, id, "F10")).toBe(8);
  expect(
    model.region(id, parseRange("F10")).cells[0].displayStyle?.background,
  ).toBe("#ff0000");
  expect(() => set(model, id, "F10", [[999]])).toThrow(/Validation/);
  expect(() => set(model, id, "E10", [["x"]])).toThrow(/Validation/);
  model.undo();
  expect(model.sheet(id).meta.validations).toEqual(rules.validations);
  expect(get(model, id, "F10")).toBe(null);
  model.redo();
  const before = model.snapshot();
  expect(() =>
    model.execute([
      {
        ...command,
        cells: command.cells.map((cell) => ({ ...cell, value: 3 })),
        validations: [
          { type: "whole", range: parseRange("A1:A3"), minimum: 50 },
        ],
      },
    ]),
  ).toThrow(/Validation/);
  expect(model.snapshot()).toEqual(before);
  model.execute([
    {
      type: "copy",
      sheetId: id,
      range: parseRange("A1:A3"),
      targetSheetId: id,
      targetRow: 19,
      targetColumn: 0,
    },
  ]);
  expect(() => set(model, id, "A20", [[99]])).toThrow(/Validation/);
  model.execute([
    {
      type: "copy",
      sheetId: id,
      range: parseRange("A20:A22"),
      targetSheetId: id,
      targetRow: 19,
      targetColumn: 1,
      cut: true,
    },
  ]);
  expect(() => set(model, id, "B22", [[99]])).toThrow(/Validation/);
  expect(
    model.region(id, parseRange("B22")).cells[0].displayStyle?.background,
  ).toBe("#ff0000");
  model.undo();
  expect(get(model, id, "A22")).toBe(8);
  expect(
    model.sheet(id).meta.validations?.some((rule) => rule.range.c1 === 1),
  ).toBe(false);
});

it("rejects hidden content in merged cells without partially applying a write", () => {
  const { model, id } = setup();
  set(model, id, "A1", [["title"]]);
  model.execute([{ type: "merge", sheetId: id, range: parseRange("A1:B2") }]);
  const before = model.snapshot();
  expect(() => set(model, id, "A1:B1", [["changed", "invisible"]])).toThrow(
    /merged/i,
  );
  expect(model.snapshot()).toEqual(before);
  expect(() =>
    model.execute([
      { type: "setFormula", sheetId: id, row: 1, column: 1, formula: "=42" },
    ]),
  ).toThrow(/merged/i);
  const paste = {
    type: "paste" as const,
    sheetId: id,
    source: parseRange("A1:B2"),
    cells: [
      { row: 0, column: 0, value: "changed", style: {} },
      { row: 0, column: 1, value: "invisible", style: {} },
    ],
    targetRow: 0,
    targetColumn: 0,
  };
  expect(() => model.execute([{ ...paste, mode: "values" }])).toThrow(
    /merged/i,
  );
  expect(model.snapshot()).toEqual(before);
  expect(() =>
    model.execute([{ ...paste, merges: [parseRange("A1:B2")] }]),
  ).toThrow(/merged/i);
  expect(model.snapshot()).toEqual(before);
  model.execute([paste]);
  expect(model.sheet(id).meta.merges).toEqual([]);
  expect(get(model, id, "B1")).toBe("invisible");
  model.undo();
  expect(model.snapshot()).toEqual(before);
  set(model, id, "A1", [["visible"]]);
  expect(get(model, id, "A1")).toBe("visible");
});

it("fills backwards and irregular patterns without rewriting the source", () => {
  const { model, id } = setup();
  set(model, id, "A5:A7", [[1], [4], [10]]);
  model.execute([
    {
      type: "fill",
      sheetId: id,
      source: parseRange("A5:A7"),
      target: parseRange("A1:A10"),
    },
  ]);
  expect(
    model.region(id, parseRange("A1:A10")).cells.map((cell) => cell.value),
  ).toEqual([10, 1, 4, 10, 1, 4, 10, 1, 4, 10]);
  model.undo();
  expect(get(model, id, "A6")).toBe(4);
  expect(get(model, id, "A1")).toBe(null);
  set(model, id, "B5:B6", [[10], [20]]);
  model.execute([
    {
      type: "fill",
      sheetId: id,
      source: parseRange("B5:B6"),
      target: parseRange("B2:B8"),
    },
  ]);
  expect(
    model.region(id, parseRange("B2:B8")).cells.map((cell) => cell.value),
  ).toEqual([-20, -10, 0, 10, 20, 30, 40]);
  set(model, id, "C4:C5", [["a"], ["b"]]);
  model.execute([
    {
      type: "fill",
      sheetId: id,
      source: parseRange("C4:C5"),
      target: parseRange("C1:C7"),
    },
  ]);
  expect(
    model.region(id, parseRange("C1:C7")).cells.map((cell) => cell.value),
  ).toEqual(["b", "a", "b", "a", "b", "a", "b"]);
});

it("replaces a large spilled array atomically across async write checkpoints", async () => {
  const model = new WorkbookModel({
    sheets: [{ name: "Large", rows: 4000, columns: 5 }],
  });
  const id = model.sheets[0].meta.id;
  set(model, id, "A1", [["=SEQUENCE(3001)"]]);
  const original = model.snapshot(),
    revision = model.revision;
  const command = {
    type: "setValues" as const,
    sheetId: id,
    range: parseRange("A1:A3001"),
    values: Array.from({ length: 3001 }, () => [9]),
  };
  await model.executeAsync([command], async () => {});
  expect(get(model, id, "A3001")).toBe(9);
  model.undo();
  expect(get(model, id, "A3001")).toBe(3001);
  expect(model.snapshot()).toEqual(original);
  let checkpoints = 0;
  await expect(
    model.executeAsync([command], async () => {
      if (++checkpoints === 2) throw new Error("cancelled");
    }),
  ).rejects.toThrow("cancelled");
  expect(model.snapshot()).toEqual(original);
  expect(model.revision).toBe(revision + 2);
});

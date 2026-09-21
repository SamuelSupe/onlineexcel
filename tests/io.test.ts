import { parseInputValue } from "../src/core/input";
import { formatColor } from "../src/core/number-format";
import { it, expect } from "vitest";
import { formatValue } from "../src/core/format";
import { readXlsx, writeXlsx } from "../src/io/xlsx";
import { WorkbookModel } from "../src/core/model";
import { parseRange, keyOf } from "../src/core/address";
import { parseCsv, writeCsv } from "../src/io/csv";
import { zipSync, strToU8, unzipSync } from "fflate";
it("roundtrips value-list filters, blank selection and other column conditions", async () => {
  const model = new WorkbookModel();
  const sheetId = model.sheets[0].meta.id;
  model.execute([
    {
      type: "setValues",
      sheetId,
      range: parseRange("A1:B4"),
      values: [
        ["Name", "Value"],
        ["Amy", 10],
        [null, 20],
        ["Zoe", 30],
      ],
    },
    {
      type: "filter",
      sheetId,
      range: parseRange("A1:B4"),
      rules: [
        { column: 0, operator: "in", values: ["Amy", ""] },
        { column: 1, operator: "gte", value: 15 },
      ],
    },
  ]);
  const output = await writeXlsx(model.snapshot());
  const imported = await readXlsx(output.data);
  const restored = new WorkbookModel({ snapshot: imported.snapshot });
  expect(restored.sheets[0].meta.filteredRows).toEqual([1, 3]);
  expect(restored.sheets[0].meta.filter?.rules).toEqual([
    { column: 0, operator: "in", values: ["", "Amy"] },
    { column: 1, operator: "gte", value: 15 },
  ]);
  expect(imported.diagnostics.some((d) => d.code === "FILTER_VARIANT")).toBe(
    false,
  );
});
it("renders imported high-precision formats without crashing the grid", async () => {
  const model = new WorkbookModel();
  const sheetId = model.sheets[0].meta.id;
  const numberFormat = "0." + "0".repeat(30);
  model.execute([
    { type: "setValues", sheetId, range: parseRange("A1"), values: [[1.5]] },
    {
      type: "style",
      sheetId,
      range: parseRange("A1"),
      style: { numberFormat },
    },
  ]);
  const result = await readXlsx((await writeXlsx(model.snapshot())).data);
  const restored = new WorkbookModel({ snapshot: result.snapshot });
  const cell = restored.region(restored.sheets[0].meta.id, parseRange("A1"))
    .cells[0];
  expect(cell.style.numberFormat).toBe(numberFormat);
  expect(formatValue(cell.value, cell.style)).toBe("1.5" + "0".repeat(19));
  expect(
    formatValue(1.5, { numberFormat: "0." + "0".repeat(101) + "E+00" }),
  ).toBe("1.5" + "0".repeat(19) + "E+00");
});
it("handles multiline CSV fields, quotes and formula-looking text as data", () => {
  const rows = [["a,b", "line\nbreak", '"quoted"', "=SUM(A1:A2)"]];
  expect(parseCsv(writeCsv(rows))).toEqual(rows);
  expect(() => parseCsv('"unclosed')).toThrow();
});
it("preserves single-column blank records through CSV and clipboard TSV roundtrips", () => {
  for (const delimiter of [",", "\t"])
    for (const rows of [[[""]], [["a"], [""]], [["a"], [""], [""]]])
      expect(parseCsv(writeCsv(rows, delimiter), delimiter)).toEqual(rows);
  expect(parseCsv("a\r\n")).toEqual([["a"]]);
});
it("roundtrips styles, dimensions, formulas, names and dynamic arrays", async () => {
  const model = new WorkbookModel({
    sheets: [{ name: "Sales", rows: 100, columns: 20 }],
  });
  const id = model.sheets[0].meta.id;
  model.execute([
    {
      type: "setValues",
      sheetId: id,
      range: parseRange("A1:B3"),
      values: [
        ["Label", "Value"],
        ["a", 12],
        ["b", 34],
      ],
    },
    {
      type: "setFormula",
      sheetId: id,
      row: 0,
      column: 3,
      formula: "=SUM(B2:B3)",
    },
    {
      type: "setFormula",
      sheetId: id,
      row: 3,
      column: 3,
      formula: "=SEQUENCE(3,2)",
    },
    {
      type: "setFormula",
      sheetId: id,
      row: 7,
      column: 3,
      formula: "=SUM(D4#)",
    },
    {
      type: "style",
      sheetId: id,
      range: parseRange("A1:B1"),
      style: { bold: true, background: "#abcdef" },
    },
    { type: "freeze", sheetId: id, rows: 1, columns: 1 },
    {
      type: "defineName",
      name: "Amounts",
      sheetId: id,
      range: parseRange("B2:B3"),
    },
  ]);
  const exported = await writeXlsx(model.snapshot());
  const imported = await readXlsx(exported.data);
  const next = new WorkbookModel({ snapshot: imported.snapshot }),
    sid = next.sheets[0].meta.id;
  expect(next.engine.get(sid, keyOf(0, 3))).toBe(46);
  expect(next.engine.get(sid, keyOf(5, 4))).toBe(6);
  expect(next.engine.get(sid, keyOf(7, 3))).toBe(21);
  expect(
    next.region(sid, parseRange("A1")).cells[0].style.background?.toLowerCase(),
  ).toBe("#abcdef");
  expect(next.sheets[0].meta.frozenRows).toBe(1);
  expect(next.names.AMOUNTS.range).toEqual(parseRange("B2:B3"));
});
it("requires explicit consent for known file losses", async () => {
  const model = new WorkbookModel();
  model.diagnostics = [
    {
      code: "CHART",
      severity: "warning",
      lossy: true,
      message: "Chart not retained",
    },
  ];
  await expect(writeXlsx(model.snapshot())).rejects.toThrow(/discard/);
  expect(
    (await writeXlsx(model.snapshot(), { allowLossy: true })).data.length,
  ).toBeGreaterThan(0);
});
it("imports independent hand-authored shared string and shared formula parts", async () => {
  const files: Record<string, Uint8Array> = {};
  const put = (name: string, xml: string) => (files[name] = strToU8(xml));
  put(
    "_rels/.rels",
    '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  );
  put(
    "xl/workbook.xml",
    '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Source" sheetId="1" r:id="rId1"/></sheets></workbook>',
  );
  put(
    "xl/_rels/workbook.xml.rels",
    '<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="s" Type="x/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
  );
  put("xl/sharedStrings.xml", "<sst><si><t>中文 &amp; text</t></si></sst>");
  put(
    "xl/worksheets/sheet1.xml",
    '<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>5</v></c><c r="C1"><f t="shared" si="0" ref="C1:C2">B1*2</f><v>999</v></c><c r="D1"><f>_xlfn.FutureFunction  ( B1 )</f><v>999</v></c></row><row r="2"><c r="B2"><v>7</v></c><c r="C2"><f t="shared" si="0"/><v>999</v></c></row></sheetData></worksheet>',
  );
  const result = await readXlsx(zipSync(files));
  const model = new WorkbookModel({ snapshot: result.snapshot }),
    id = model.sheets[0].meta.id;
  expect(model.engine.get(id, 0)).toBe("中文 & text");
  expect(model.engine.get(id, keyOf(1, 2))).toBe(14);
  expect(model.engine.get(id, keyOf(0, 3))).toEqual({ error: "#NAME?" });
  const roundtrip = await readXlsx((await writeXlsx(model.snapshot())).data);
  expect(
    roundtrip.snapshot.sheets[0].cells.find(([key]) => key === keyOf(0, 3))?.[1]
      .formula,
  ).toBe("=_xlfn.FutureFunction  ( B1 )");
});
it("rejects malformed or hostile XML without fetching entities", async () => {
  const model = new WorkbookModel(),
    output = await writeXlsx(model.snapshot());
  const parts = unzipSync(output.data);
  parts["xl/workbook.xml"] = strToU8(
    '<!DOCTYPE workbook [<!ENTITY x SYSTEM "file:///etc/passwd">]><workbook>&x;</workbook>',
  );
  await expect(readXlsx(zipSync(parts))).rejects.toThrow(/DOCTYPE/);
});
it("cancels imports before returning a workbook", async () => {
  const output = await writeXlsx(new WorkbookModel().snapshot());
  const controller = new AbortController();
  controller.abort();
  await expect(
    readXlsx(output.data, { signal: controller.signal }),
  ).rejects.toThrow();
});

it("restores contains filters and allows clearing recalculated hidden rows", async () => {
  const model = new WorkbookModel();
  const id = model.sheets[0].meta.id;
  model.execute([
    {
      type: "setValues",
      sheetId: id,
      range: parseRange("A1:A4"),
      values: [["Header"], ["alpha"], ["beta"], ["alphabet"]],
    },
    {
      type: "filter",
      sheetId: id,
      range: parseRange("A1:A4"),
      rules: [{ column: 0, operator: "contains", value: "alpha" }],
    },
  ]);
  const result = await readXlsx((await writeXlsx(model.snapshot())).data);
  const next = new WorkbookModel({ snapshot: result.snapshot });
  const sheet = next.sheets[0];
  expect(sheet.meta.filteredRows).toEqual([2]);
  expect(sheet.meta.filter?.rules[0]).toMatchObject({
    operator: "contains",
    value: "alpha",
  });
  next.execute([{ type: "filter", sheetId: sheet.meta.id, rules: [] }]);
  expect(sheet.meta.hiddenRows).toEqual([]);
  expect(sheet.meta.filteredRows).toBeUndefined();
});

it("roundtrips validation, locked cells, hidden sheets and conditional formatting", async () => {
  const model = new WorkbookModel({
      sheets: [{ name: "Input" }, { name: "Hidden" }],
    }),
    id = model.sheets[0].meta.id;
  model.execute([
    {
      type: "setValues",
      sheetId: id,
      range: parseRange("A1:B1"),
      values: [["Yes", 8]],
    },
    {
      type: "validation",
      sheetId: id,
      range: parseRange("A1:A3"),
      rule: { type: "list", values: ["Yes", "No"], allowBlank: false },
    },
    {
      type: "validation",
      sheetId: id,
      range: parseRange("B1:B3"),
      rule: { type: "whole", minimum: 1, maximum: 10 },
    },
    {
      type: "conditionalFormat",
      sheetId: id,
      range: parseRange("B1:B3"),
      rule: {
        operator: "gt",
        value: 5,
        style: {
          background: "#ff0000",
          color: "#123456",
          fontFamily: "Arial",
          fontSize: 14,
          bold: false,
          italic: true,
          underline: false,
          numberFormat: "0.00%",
          align: "right",
          verticalAlign: "top",
          wrap: false,
          border: { top: "#112233", bottom: "" },
          locked: false,
        },
      },
    },
    {
      type: "style",
      sheetId: id,
      range: parseRange("A1:B3"),
      style: { locked: false },
    },
    { type: "protect", sheetId: id, enabled: true },
    { type: "sheetVisibility", sheetId: model.sheets[1].meta.id, hidden: true },
  ]);
  const imported = await readXlsx((await writeXlsx(model.snapshot())).data);
  expect(imported.diagnostics.filter((d) => d.lossy)).toEqual([]);
  expect(imported.snapshot.sheets[0].conditionalFormats?.[0].style).toEqual({
    ...model.sheet(id).meta.conditionalFormats![0].style,
    background: "#FF0000",
  });
  const next = new WorkbookModel({ snapshot: imported.snapshot }),
    sid = next.sheets[0].meta.id;
  expect(next.sheets[0].meta.protected).toBe(true);
  expect(next.sheets[1].meta.hidden).toBe(true);
  expect(
    next
      .region(sid, parseRange("B1"))
      .cells[0].displayStyle?.background?.toLowerCase(),
  ).toBe("#ff0000");
  expect(() =>
    next.execute([
      {
        type: "setValues",
        sheetId: sid,
        range: parseRange("A2"),
        values: [["Invalid"]],
      },
    ]),
  ).toThrow(/Validation/);
  next.execute([
    {
      type: "setValues",
      sheetId: sid,
      range: parseRange("A2"),
      values: [["No"]],
    },
  ]);
  expect(() =>
    next.execute([
      {
        type: "setValues",
        sheetId: sid,
        range: parseRange("C1"),
        values: [[3]],
      },
    ]),
  ).toThrow(/Protected/);
});

it("reports unsupported differential styling before a potentially lossy XLSX export", async () => {
  const model = new WorkbookModel(),
    id = model.sheets[0].meta.id;
  model.execute([
    {
      type: "conditionalFormat",
      sheetId: id,
      range: parseRange("A1"),
      rule: { operator: "gt", value: 5, style: { italic: true } },
    },
  ]);
  const files = unzipSync((await writeXlsx(model.snapshot())).data);
  const xml = new TextDecoder().decode(files["xl/styles.xml"]);
  files["xl/styles.xml"] = strToU8(
    xml.replace('<i val="1"/>', '<i val="1"/><strike/>'),
  );
  const imported = await readXlsx(zipSync(files));
  expect(
    imported.diagnostics.some((d) => d.code === "RULE_VARIANT" && d.lossy),
  ).toBe(true);
  await expect(writeXlsx(imported.snapshot)).rejects.toThrow(/lossy/i);
  expect(
    (await writeXlsx(imported.snapshot, { allowLossy: true })).data.length,
  ).toBeGreaterThan(0);
});

it("parses office input without destroying identifiers and renders conditional numeric sections", () => {
  expect(parseInputValue("00123").value).toBe("00123");
  expect(parseInputValue("1234567890123456").value).toBe("1234567890123456");
  expect(parseInputValue("1,234.50").value).toBe(1234.5);
  expect(parseInputValue("25%").value).toBe(0.25);
  expect(parseInputValue("2026-09-20").format).toBe("yyyy-mm-dd");
  expect(() => parseInputValue("2026-02-30", "date")).toThrow();
  expect(parseInputValue("00123", "text").value).toBe("00123");
  expect(formatValue(-12, { numberFormat: '0;[Red](0);"zero"' })).toBe("(12)");
  expect(formatValue(0, { numberFormat: '0;[Red](0);"zero"' })).toBe("zero");
  expect(formatValue(5, { numberFormat: "[>=10]0.00;[Red]0.0" })).toBe("5.0");
  expect(formatColor(-12, "0;[Red](0)")).toBe("#ff0000");
  expect(formatValue(1.25, { numberFormat: "# ?/?" })).toBe("1 1/4");
  expect(formatValue(2.5, { numberFormat: "[h]:mm:ss" })).toBe("60:00:00");
});

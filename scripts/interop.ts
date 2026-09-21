import { readFile, writeFile } from "node:fs/promises";
import { readXlsx } from "../src/io/xlsx-read";
import { writeModelXlsx } from "../src/io/xlsx-write";
import { WorkbookModel } from "../src/core/model";
import { keyOf, parseRange } from "../src/core/address";
const input = await readXlsx(await readFile(".cache/external.xlsx"));
const model = new WorkbookModel({ snapshot: input.snapshot });
const id = model.sheets[0].meta.id;
if (model.engine.get(id, keyOf(1, 3)) !== 120)
  throw new Error("Independent formula did not calculate");
if (model.dateSystem !== 1904)
  throw new Error("1904 date system not preserved");
if (model.sheets[0].meta.filter?.rules[0]?.operator !== "in")
  throw new Error("Independent value-list filter was not imported");
model.execute([
  {
    type: "setFormula",
    sheetId: id,
    row: 7,
    column: 0,
    formula: "=SEQUENCE(2,2)",
  },
  { type: "setValues", sheetId: id, range: parseRange("C2"), values: [[4]] },
]);
const output = await writeModelXlsx(model, { allowLossy: true });
await writeFile(".cache/interop-output.xlsx", output.data);
console.log(
  JSON.stringify({
    importedSheets: model.sheets.length,
    totalAfterEdit: model.engine.get(id, keyOf(1, 3)),
    diagnostics: input.diagnostics,
  }),
);

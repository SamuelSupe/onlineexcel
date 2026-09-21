import { performance } from "node:perf_hooks";
import { writeFile, mkdir } from "node:fs/promises";
import { WorkbookModel } from "../src/core/model";
import { keyOf, parseRange } from "../src/core/address";
import { writeModelXlsx } from "../src/io/xlsx-write";
import { readXlsx } from "../src/io/xlsx-read";
import type { WorkbookSnapshot } from "../src/core/types";
const rowCount = Number(process.env.BENCH_ROWS ?? 100000);
const cells: WorkbookSnapshot["sheets"][number]["cells"] = [];
for (let r = 0; r < rowCount; r++)
  for (let c = 0; c < 10; c++)
    cells.push([
      keyOf(r, c),
      c === 9 && r < rowCount / 2
        ? { formula: `=SUM(A${r + 1}:I${r + 1})` }
        : {
            value: c === 0 ? `ID-${r + 1}` : ((r * 31 + c * 17) % 10000) / 100,
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
      rowCount,
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
const started = performance.now();
const model = new WorkbookModel({ snapshot, historyLimit: 5 });
const loadMs = performance.now() - started;
const edits: number[] = [];
for (let i = 0; i < 50; i++) {
  const start = performance.now();
  model.execute([
    {
      type: "setValues",
      sheetId: "bench",
      range: parseRange(`B${i + 1}`),
      values: [[i]],
    },
  ]);
  edits.push(performance.now() - start);
}
edits.sort((a, b) => a - b);
console.log(
  JSON.stringify({
    stage: "load-edit",
    rows: rowCount,
    populatedCells: cells.length,
    formulaCells: rowCount / 2,
    loadMs,
    editP95Ms: edits[Math.floor(edits.length * 0.95)],
    memory: process.memoryUsage(),
  }),
);
let start = performance.now();
model.engine.rebuild();
const dependencyRebuildMs = performance.now() - start;
start = performance.now();
model.engine.recalculate();
const fullRecalculateMs = performance.now() - start;
start = performance.now();
model.execute([
  {
    type: "sort",
    sheetId: "bench",
    range: { r1: 0, c1: 0, r2: rowCount - 1, c2: 9 },
    keys: [{ column: 1, direction: "desc" }],
  },
]);
const sortMs = performance.now() - start;
start = performance.now();
const exported = await writeModelXlsx(model);
const exportMs = performance.now() - start;
start = performance.now();
const imported = await readXlsx(exported.data);
const importMs = performance.now() - start;
start = performance.now();
const importedModel = new WorkbookModel({
  snapshot: imported.snapshot,
  historyLimit: 0,
});
const importCalculateMs = performance.now() - start;
if (importedModel.sheets[0].cells.size !== cells.length)
  throw new Error("Imported cell count differs");
const result = {
  timestamp: new Date().toISOString(),
  environment: {
    node: process.version,
    platform: process.platform,
    arch: process.arch,
  },
  rows: rowCount,
  columns: 100,
  populatedCells: cells.length,
  formulaCells: rowCount / 2,
  loadMs,
  editP50Ms: edits[Math.floor(edits.length * 0.5)],
  editP95Ms: edits[Math.floor(edits.length * 0.95)],
  dependencyRebuildMs,
  fullRecalculateMs,
  sortMs,
  exportMs,
  importMs,
  importCalculateMs,
  xlsxBytes: exported.data.length,
  importedCells: imported.snapshot.sheets[0].cells.length,
  peakRssMiB: process.resourceUsage().maxRSS / 1024,
  memory: process.memoryUsage(),
};
await mkdir(".cache", { recursive: true });
await writeFile(".cache/benchmark.json", JSON.stringify(result, null, 2));
console.log(JSON.stringify(result, null, 2));

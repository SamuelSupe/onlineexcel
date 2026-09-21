import { writeFile } from "node:fs/promises";
import { listFunctions } from "../src/formula/functions";
import { LIBRARY_VERSION } from "../src/runtime/protocol";
const functions = listFunctions();
await writeFile(
  "docs/functions.json",
  JSON.stringify(
    { schemaVersion: 1, libraryVersion: LIBRARY_VERSION, functions },
    null,
    2,
  ) + "\n",
);
let markdown =
  "# Formula capabilities\n\n" +
  functions.length +
  " distinct functions. Optional parameters appear in brackets. Functions use invariant English names and comma separators. This list describes implemented functions, not certification against every Excel edge case.\n\n";
for (const category of [...new Set(functions.map((f) => f.category))].sort()) {
  markdown +=
    "## " +
    category +
    "\n\n| Function | Parameters | Notes |\n| --- | --- | --- |\n";
  for (const fn of functions.filter((f) => f.category === category))
    markdown += `| ${fn.name} | ${fn.signature} | ${fn.notes ?? ""} |\n`;
  markdown += "\n";
}
await writeFile("docs/functions.md", markdown);
console.log(`Documented ${functions.length} functions`);

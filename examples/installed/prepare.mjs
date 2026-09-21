import { copyFile, writeFile } from "node:fs/promises";
await copyFile("business.js", "public/business.js");
await writeFile(
  "public/deferred-module.js",
  "await new Promise(resolve => setTimeout(resolve, 100));\n",
);

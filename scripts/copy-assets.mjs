#!/usr/bin/env node
import { cp, mkdir, readdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { resolve, join } from "node:path";
const destination = process.argv[2];
if (!destination || process.argv.length !== 3) {
  console.error("Usage: onlineexcel-copy-assets <public/onlineexcel>");
  process.exitCode = 1;
} else {
  const source = fileURLToPath(new URL("../dist/", import.meta.url));
  const target = resolve(destination);
  if (target === resolve(source))
    throw new Error("Destination must differ from package dist directory");
  await mkdir(target, { recursive: true });
  for (const file of await readdir(source)) {
    if (
      file === "worker.js" ||
      file === "onlineexcel.js" ||
      /^chunk-.*\.js$/.test(file)
    )
      await cp(join(source, file), join(target, file));
  }
  console.log(`OnlineExcel browser and Worker assets copied to ${target}`);
}

import { createHash } from "node:crypto";
import { cp, mkdir, readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { execFileSync } from "node:child_process";
const destination = resolve(".cache/sdk-example");
await mkdir(destination, { recursive: true });
await cp("examples/installed", destination, { recursive: true });
const pkg = JSON.parse(await readFile(destination + "/package.json", "utf8"));
const { version } = JSON.parse(await readFile("package.json", "utf8"));
const archive = await readFile(`artifacts/onlineexcel-${version}.tgz`);
const archiveName =
  "onlineexcel-" +
  createHash("sha256").update(archive).digest("hex").slice(0, 16) +
  ".tgz";
await writeFile(destination + "/" + archiveName, archive);
pkg.dependencies.onlineexcel = "file:" + destination + "/" + archiveName;
await writeFile(destination + "/package.json", JSON.stringify(pkg, null, 2));
execFileSync(
  "npm",
  [
    "install",
    "--ignore-scripts",
    "--no-audit",
    "--no-fund",
    "--prefer-offline",
  ],
  { cwd: destination, stdio: "inherit" },
);
execFileSync(
  resolve("node_modules/.bin/tsc"),
  ["-p", destination + "/tsconfig.json"],
  { stdio: "inherit" },
);
execFileSync("npm", ["run", "build"], { cwd: destination, stdio: "inherit" });
console.log(
  "Installed-package production example: /.cache/sdk-example/dist/index.html",
);

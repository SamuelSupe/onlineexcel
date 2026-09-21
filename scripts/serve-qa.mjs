import { createServer } from "node:http";
import { readFile, stat } from "node:fs/promises";
import { resolve, extname, sep } from "node:path";
import { build } from "esbuild";
const root = process.cwd();
const built = await build({
  entryPoints: ["demo/qa.ts"],
  bundle: true,
  write: false,
  platform: "browser",
  format: "esm",
  target: "es2022",
});
const html = (await readFile("demo/qa.html", "utf8")).replace(
  "./qa.ts",
  "/qa.js",
);
const mime = {
  ".js": "text/javascript",
  ".json": "application/json",
  ".html": "text/html",
  ".css": "text/css",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
};
const server = createServer(async (request, response) => {
  try {
    const pathname = decodeURIComponent(
      new URL(request.url, "http://localhost").pathname,
    );
    if (pathname === "/qa") {
      response.setHeader("Content-Type", "text/html");
      response.end(html);
      return;
    }
    if (pathname === "/qa.js") {
      response.setHeader("Content-Type", "text/javascript");
      response.end(built.outputFiles[0].contents);
      return;
    }
    const path = resolve(root, "." + pathname);
    if (!path.startsWith(root + sep)) {
      response.writeHead(403);
      response.end();
      return;
    }
    const info = await stat(path);
    if (!info.isFile()) {
      response.writeHead(404);
      response.end();
      return;
    }
    response.setHeader(
      "Content-Type",
      mime[extname(path)] ?? "application/octet-stream",
    );
    response.setHeader("Cache-Control", "no-store");
    response.end(await readFile(path));
  } catch {
    response.writeHead(404);
    response.end("Not found");
  }
});
server.listen(5174, "127.0.0.1", () =>
  console.log("Built-package QA: http://127.0.0.1:5174/qa"),
);

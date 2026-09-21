import { build } from "esbuild";
import { rm, mkdir } from "node:fs/promises";
await rm("dist", { recursive: true, force: true });
await mkdir("dist", { recursive: true });
const common = {
  bundle: true,
  platform: "browser",
  target: "es2022",
  sourcemap: true,
  logLevel: "info",
};
await build({
  ...common,
  entryPoints: {
    index: "src/index.ts",
    core: "src/core/index.ts",
    xlsx: "src/io/xlsx.ts",
    editor: "src/editor/index.ts",
    worker: "src/runtime/worker.ts",
    react: "src/adapters/react.tsx",
    vue: "src/adapters/vue.ts",
  },
  outdir: "dist",
  format: "esm",
  splitting: true,
  external: ["react", "react/jsx-runtime", "vue"],
});
await build({
  ...common,
  entryPoints: ["src/browser.ts"],
  outfile: "dist/onlineexcel.js",
  format: "iife",
  globalName: "OnlineExcel",
  define: { __ONLINEEXCEL_IIFE__: "true" },
});

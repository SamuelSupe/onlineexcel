import type { Workbook } from "./runtime/client";
import { initializeWorkbook } from "./runtime/create";
export { WorkbookError } from "./runtime/errors";
import type { WorkbookOptions } from "./core/types";
export { Workbook } from "./runtime/client";
export { mountEditor, supportedLocales } from "./editor/index";
export { listFunctions } from "./formula/functions";
export { parseRange, parseCell, address, columnName } from "./core/address";
const script = document.currentScript as HTMLScriptElement | null;
const defaultWorkerUrl = script?.src
  ? new URL("./worker.js", script.src)
  : undefined;
export async function createWorkbook(
  options: WorkbookOptions = {},
): Promise<Workbook> {
  return initializeWorkbook(options, defaultWorkerUrl);
}

export { createPersistence } from "./runtime/persistence";
export { createIndexedDBStorage } from "./runtime/indexeddb";

import type { Workbook } from "./runtime/client";
import { initializeWorkbook } from "./runtime/create";
import type { WorkbookOptions } from "./core/types";
export { Workbook } from "./runtime/client";
export { mountEditor, supportedLocales } from "./editor/index";
export { listFunctions } from "./formula/functions";
export { parseRange, parseCell, address, columnName } from "./core/address";
export type * from "./core/types";
export type {
  Editor,
  EditorOptions,
  EditorConfiguration,
  EditorAction,
  EditorActionContext,
  EditState,
  Locale,
} from "./editor/index";
export type { WorkbookEvents } from "./runtime/client";
/** Creates an isolated browser Worker. Supply workerUrl when assets are served separately. */
export async function createWorkbook(
  options: WorkbookOptions = {},
): Promise<Workbook> {
  return initializeWorkbook(options, new URL("./worker.js", import.meta.url));
}

export type { ColumnType } from "./core/input";

export { WorkbookError } from "./runtime/errors";
export type { WorkbookErrorCode, ErrorContext } from "./runtime/errors";
export type {
  CustomFunction,
  FunctionValue,
  WorkbookModule,
  ValidationContext,
} from "./runtime/extensions";
export type {
  RecordColumn,
  RecordOptions,
  ReadChunkOptions,
  WriteChunksResult,
} from "./runtime/records";
export type { IconName } from "./editor/icons";

export { createPersistence } from "./runtime/persistence";
export { createIndexedDBStorage } from "./runtime/indexeddb";
export type {
  WorkbookPersistence,
  PersistenceOptions,
  PersistenceState,
  PersistenceStorage,
  SavedWorkbook,
} from "./runtime/persistence";

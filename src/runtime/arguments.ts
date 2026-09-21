import {
  parseCell as readCell,
  parseRange as readRange,
} from "../core/address";
import type { Rect } from "../core/types";
import { WorkbookError, asWorkbookError } from "./errors";
export function parseRange(range: string | Rect): Rect {
  try {
    return readRange(range);
  } catch (error) {
    throw new WorkbookError(
      "INVALID_ARGUMENT",
      asWorkbookError(error).message,
      { outcome: "not-executed" },
    );
  }
}
export function parseCell(cell: string) {
  try {
    return readCell(cell);
  } catch (error) {
    throw new WorkbookError(
      "INVALID_ARGUMENT",
      asWorkbookError(error).message,
      { outcome: "not-executed" },
    );
  }
}

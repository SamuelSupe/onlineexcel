import type { Workbook } from "./client";
import type {
  ChangeEvent,
  InputValue,
  OperationOptions,
  Rect,
  Region,
  Scalar,
} from "../core/types";
import { parseRange } from "./arguments";
import { WorkbookError } from "./errors";
export interface RecordColumn {
  key: string;
  title?: string;
}
export interface RecordOptions extends OperationOptions {
  start?: string;
  header?: boolean;
}
export async function writeRecords(
  workbook: Workbook,
  sheetId: string,
  records: readonly Record<string, InputValue>[],
  columns: readonly RecordColumn[],
  options: RecordOptions = {},
): Promise<ChangeEvent | null> {
  checkColumns(columns);
  const values: InputValue[][] = records.map((record) =>
    columns.map((column) =>
      Object.hasOwn(record, column.key) ? (record[column.key] ?? null) : null,
    ),
  );
  if (options.header)
    values.unshift(columns.map((column) => column.title ?? column.key));
  if (!values.length) return null;
  const start = parseRange(options.start ?? "A1");
  return workbook.setValues(
    sheetId,
    {
      ...start,
      r2: start.r1 + values.length - 1,
      c2: start.c1 + columns.length - 1,
    },
    values,
    { ...options, parseFormulas: false },
  );
}
export interface ReadChunkOptions extends OperationOptions {
  rowsPerChunk?: number;
  /** true (default) rejects concurrent changes instead of returning mixed revisions. */
  consistent?: boolean;
}
export async function* readChunks(
  workbook: Workbook,
  sheetId: string,
  range: string | Rect,
  options: ReadChunkOptions = {},
): AsyncGenerator<Region> {
  const area = parseRange(range),
    width = area.c2 - area.c1 + 1;
  const size =
    options.rowsPerChunk ??
    Math.max(1, Math.min(1000, Math.floor(200000 / width)));
  if (!Number.isInteger(size) || size < 1 || size * width > 200000)
    throw new WorkbookError(
      "INVALID_ARGUMENT",
      "Chunk size must fit the 200000-cell region limit",
    );
  let revision: number | undefined;
  for (let row = area.r1; row <= area.r2; row += size) {
    const chunk = await workbook.request<Region>(
      "region",
      {
        sheetId,
        range: { ...area, r1: row, r2: Math.min(area.r2, row + size - 1) },
        expectedRevision: options.consistent === false ? undefined : revision,
      },
      options,
    );
    revision = chunk.revision;
    yield chunk;
  }
}
export async function* readRecords(
  workbook: Workbook,
  sheetId: string,
  range: string | Rect,
  columns: readonly RecordColumn[],
  options: ReadChunkOptions = {},
): AsyncGenerator<Record<string, Scalar>[]> {
  checkColumns(columns);
  const area = parseRange(range);
  if (area.c2 - area.c1 + 1 !== columns.length)
    throw new WorkbookError(
      "INVALID_ARGUMENT",
      "Columns must match range width",
    );
  for await (const chunk of readChunks(workbook, sheetId, area, options)) {
    const values = Array.from(
      { length: chunk.range.r2 - chunk.range.r1 + 1 },
      () => Array<Scalar>(columns.length).fill(null),
    );
    for (const cell of chunk.cells)
      values[cell.row - chunk.range.r1][cell.column - chunk.range.c1] =
        cell.value;
    yield values.map((row) =>
      Object.fromEntries(
        columns.map((column, index) => [column.key, row[index]]),
      ),
    );
  }
}
export interface WriteChunksResult {
  rows: number;
  chunks: number;
  revision?: number;
}
/** Each yielded chunk is an atomic, undoable transaction. Earlier chunks remain committed on failure. */
export async function writeChunks(
  workbook: Workbook,
  sheetId: string,
  start: string,
  source: AsyncIterable<InputValue[][]> | Iterable<InputValue[][]>,
  options: OperationOptions = {},
): Promise<WriteChunksResult> {
  const area = parseRange(start),
    result: WriteChunksResult = { rows: 0, chunks: 0 };
  try {
    for await (const values of source) {
      if (options.signal?.aborted)
        throw new WorkbookError("CANCELLED", "Operation cancelled");
      if (!values.length) continue;
      const change = await workbook.setValues(
        sheetId,
        {
          r1: area.r1 + result.rows,
          r2: area.r1 + result.rows + values.length - 1,
          c1: area.c1,
          c2: area.c1 + values[0].length - 1,
        },
        values,
        { ...options, parseFormulas: false },
      );
      result.rows += values.length;
      result.chunks++;
      result.revision = change.revision;
    }
    return result;
  } catch (error) {
    const value =
      error instanceof WorkbookError
        ? error
        : new WorkbookError("OPERATION_FAILED", String(error));
    value.details = { cause: value.details, committed: { ...result } };
    throw value;
  }
}
function checkColumns(columns: readonly RecordColumn[]): void {
  if (
    !columns.length ||
    columns.some((column) => typeof column.key !== "string" || !column.key) ||
    new Set(columns.map((column) => column.key)).size !== columns.length
  )
    throw new WorkbookError(
      "INVALID_ARGUMENT",
      "Columns require unique nonempty keys",
    );
}

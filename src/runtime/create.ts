import type { WorkbookOptions } from "../core/types";
import { Workbook } from "./client";
import { WorkbookError, asWorkbookError } from "./errors";
import { PROTOCOL_VERSION, LIBRARY_VERSION } from "./protocol";
export async function initializeWorkbook(
  options: WorkbookOptions,
  defaultUrl?: URL,
): Promise<Workbook> {
  const {
    workerUrl,
    workerFactory,
    initializationTimeout = 15000,
    requestTimeout = 0,
    workerModules,
    ...data
  } = options;
  for (const timeout of [initializationTimeout, requestTimeout]) {
    if (!Number.isFinite(timeout) || timeout < 0)
      throw new WorkbookError(
        "INVALID_ARGUMENT",
        "Timeout must be a nonnegative finite number",
        { outcome: "not-executed" },
      );
  }
  let workbook: Workbook | undefined;
  try {
    const url = workerUrl ?? defaultUrl;
    if (!workerFactory && !url)
      throw new WorkbookError(
        "INVALID_ARGUMENT",
        "Pass workerUrl or workerFactory",
      );
    const worker = workerFactory
      ? workerFactory()
      : new Worker(url!, { type: "module", name: "OnlineExcel" });
    workbook = new Workbook(worker, { requestTimeout });
    const result = await workbook.request<{
      protocolVersion: number;
      libraryVersion: string;
    }>(
      "init",
      {
        ...data,
        protocolVersion: PROTOCOL_VERSION,
        libraryVersion: LIBRARY_VERSION,
        workerModules: workerModules?.map(
          (url) => new URL(url, document.baseURI).href,
        ),
      },
      { timeout: initializationTimeout },
    );
    if (
      result?.protocolVersion !== PROTOCOL_VERSION ||
      result.libraryVersion !== LIBRARY_VERSION
    )
      throw new WorkbookError(
        "PROTOCOL_MISMATCH",
        "Library and Worker assets have different versions",
        { operation: "init", outcome: "not-executed" },
      );
    return workbook;
  } catch (error) {
    await workbook?.dispose();
    throw asWorkbookError(error, { operation: "init" });
  }
}

import type { Diagnostic, Rect } from "../core/types";
export type WorkbookErrorCode =
  | "INVALID_ARGUMENT"
  | "OPERATION_FAILED"
  | "VALIDATION_FAILED"
  | "CANCELLED"
  | "DISPOSED"
  | "WORKER_FAILED"
  | "PROTOCOL_MISMATCH"
  | "TIMEOUT"
  | "SERIALIZATION_FAILED"
  | "LOSSY_EXPORT"
  | "MODULE_FAILED"
  | "REVISION_CONFLICT";
export interface ErrorContext {
  operation?: string;
  operationId?: string;
  sheetId?: string;
  range?: Rect;
  /** Whether the caller can know if a mutation committed. Never retry an unknown outcome automatically. */
  outcome?: "not-executed" | "rolled-back" | "unknown";
  details?: unknown;
  diagnostics?: Diagnostic[];
}
export class WorkbookError extends Error implements ErrorContext {
  readonly code: WorkbookErrorCode;
  operation?: string;
  operationId?: string;
  sheetId?: string;
  range?: Rect;
  outcome?: ErrorContext["outcome"];
  details?: unknown;
  diagnostics?: Diagnostic[];
  constructor(
    code: WorkbookErrorCode,
    message: string,
    context: ErrorContext = {},
  ) {
    super(message);
    this.name =
      code === "CANCELLED"
        ? "AbortError"
        : code === "LOSSY_EXPORT"
          ? "CompatibilityError"
          : "WorkbookError";
    this.code = code;
    Object.assign(this, context);
  }
}
export function asWorkbookError(
  error: unknown,
  context: ErrorContext = {},
): WorkbookError {
  if (error instanceof WorkbookError) return Object.assign(error, context);
  const value = error as {
    name?: string;
    message?: string;
    diagnostics?: Diagnostic[];
  };
  return new WorkbookError(
    value?.name === "AbortError"
      ? "CANCELLED"
      : value?.name === "CompatibilityError"
        ? "LOSSY_EXPORT"
        : "OPERATION_FAILED",
    value?.message ?? String(error),
    { diagnostics: value?.diagnostics, ...context },
  );
}

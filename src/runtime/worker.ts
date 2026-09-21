import { PROTOCOL_VERSION, LIBRARY_VERSION } from "./protocol";
import { WorkbookError, asWorkbookError } from "./errors";
import { registerFunctions, type WorkbookModule } from "./extensions";
import { listFunctions } from "../formula/functions";
import { parseInputValue } from "../core/input";
import { WorkbookModel } from "../core/model";
import { parseRange, keyOf, rowOf, columnOf } from "../core/address";
import { readXlsx } from "../io/xlsx-read";
import { CompatibilityError, writeModelXlsx } from "../io/xlsx-write";
import { parseCsv, writeCsv } from "../io/csv";
import { isArray, isError } from "../formula/values";
import type { ChangeEvent, InputValue, ProgressEvent } from "../core/types";
import {
  dataRegion,
  distinctValues,
  findNext,
  navigationTarget,
} from "../core/queries";
const scope = self as unknown as DedicatedWorkerGlobalScope;
let model: WorkbookModel;
let historyLimit: number | undefined;
const cancelled = new Set<number>();
let queue = Promise.resolve();
const modules: WorkbookModule[] = [];
function emit(type: string, data: unknown): void {
  scope.postMessage({ event: type, data });
}
scope.onmessage = (event) => {
  const request = event.data;
  if (
    !request ||
    !Number.isSafeInteger(request.id) ||
    typeof request.operation !== "string"
  )
    return;
  if (request.operation === "cancel") {
    cancelled.add(request.id);
    return;
  }
  queue = queue.then(async () => {
    const { id, operation, args } = request;
    const context = {
      operationId: request.operationId,
      origin: request.origin ?? "api",
    };
    const validate = (candidate = model, commands = args.commands ?? []) => {
      for (const module of modules) {
        const issues = module.validate?.({
          ...context,
          operation,
          commands,
          sheets: candidate.metadata().sheets,
          getRegion: (sheetId, range) =>
            candidate.region(sheetId, parseRange(range)),
        });
        if (issues !== undefined && !Array.isArray(issues)) {
          void Promise.resolve(issues).catch(() => {});
          throw new WorkbookError(
            "MODULE_FAILED",
            "Business validation must synchronously return an array of diagnostics",
          );
        }
        if (issues?.length) {
          const details = issues.map((issue) => ({
            code: String(issue.code),
            message: String(issue.message),
            severity: issue.severity === "warning" ? "warning" : "error",
            sheetId: issue.sheetId == null ? undefined : String(issue.sheetId),
            range: issue.range == null ? undefined : String(issue.range),
          }));
          throw new WorkbookError(
            "VALIDATION_FAILED",
            "Business validation rejected this operation",
            { details, outcome: "rolled-back" },
          );
        }
      }
    };
    const checkpoint = async () => {
      await new Promise((resolve) => setTimeout(resolve, 0));
      if (cancelled.has(id))
        throw new DOMException("Operation cancelled", "AbortError");
    };
    const progress = (value: { stage: string; progress: number }) =>
      emit("progress", {
        operation,
        ...context,
        ...value,
      } satisfies ProgressEvent);
    let change: ChangeEvent | null = null;
    try {
      if (cancelled.has(id))
        throw new DOMException("Operation cancelled", "AbortError");
      let result: unknown;
      switch (operation) {
        case "init":
          if (
            model ||
            args.protocolVersion !== PROTOCOL_VERSION ||
            args.libraryVersion !== LIBRARY_VERSION
          )
            throw new WorkbookError(
              "PROTOCOL_MISMATCH",
              "Library and Worker assets have different versions",
            );
          for (const url of args.workerModules ?? []) {
            try {
              const module: WorkbookModule = await import(
                /* @vite-ignore */ url
              );
              registerFunctions(module.functions ?? []);
              if (
                module.validate !== undefined &&
                typeof module.validate !== "function"
              )
                throw new Error("validate must be a function");
              modules.push(module);
            } catch (error) {
              throw new WorkbookError(
                "MODULE_FAILED",
                "Cannot load workbook module",
                { details: { url, message: String(error) } },
              );
            }
            await checkpoint();
          }
          historyLimit = args.historyLimit;
          model = new WorkbookModel(args);
          validate();
          result = {
            protocolVersion: PROTOCOL_VERSION,
            libraryVersion: LIBRARY_VERSION,
          };
          break;
        case "functions":
          result = listFunctions();
          break;
        case "metadata":
          result = model.metadata();
          break;
        case "region":
          if (
            args.expectedRevision !== undefined &&
            args.expectedRevision !== model.revision
          )
            throw new WorkbookError(
              "REVISION_CONFLICT",
              "Workbook changed during chunked read",
            );
          result = model.region(args.sheetId, args.range);
          break;
        case "commands":
          emit("calculation", { ...context, status: "calculating" });
          change = await model.executeAsync(
            args.commands,
            async (value) => {
              progress({ stage: "apply", progress: value });
              await checkpoint();
            },
            () => validate(),
          );
          result = change;
          break;
        case "undo":
          change = model.undo();
          result = change;
          break;
        case "redo":
          change = model.redo();
          result = change;
          break;
        case "navigationTarget":
          result = navigationTarget(
            model,
            args.sheetId,
            args.row,
            args.column,
            args.direction,
          );
          break;
        case "dataRegion":
          result = dataRegion(model, args.sheetId, args.range);
          break;
        case "distinctValues":
          result = distinctValues(model, args.sheetId, args.range, args.column);
          break;
        case "findNext":
          result = findNext(
            model,
            args.sheetId,
            args.search,
            args.after,
            args.matchCase,
            args.entireCell,
          );
          break;
        case "find":
          result = model.find(args.sheetId, args.search, args.matchCase);
          break;
        case "savePoint":
          result = { snapshot: model.snapshot(), revision: model.revision };
          break;
        case "snapshot":
          result = model.snapshot();
          break;
        case "diagnostics":
          result = model.getDiagnostics();
          break;
        case "importJSON": {
          const next = new WorkbookModel({
            snapshot: args.snapshot,
            historyLimit,
          });
          validate(next, []);
          await checkpoint();
          next.revision = model.revision + 1;
          model = next;
          change = { revision: model.revision, changes: [], source: "import" };
          result = model.metadata();
          break;
        }
        case "importXlsx": {
          const imported = await readXlsx(args.data, {
            checkpoint,
            onProgress: progress,
          });
          progress({ stage: "calculate", progress: 0.95 });
          const next = new WorkbookModel({
            snapshot: imported.snapshot,
            historyLimit,
          });
          validate(next, []);
          await checkpoint();
          next.revision = model.revision + 1;
          model = next;
          change = { revision: model.revision, changes: [], source: "import" };
          result = { ...model.metadata(), diagnostics: model.getDiagnostics() };
          break;
        }
        case "exportXlsx":
          result = await writeModelXlsx(model, {
            allowLossy: args.allowLossy,
            checkpoint,
            onProgress: progress,
          });
          break;
        case "importCsv": {
          const raw = parseCsv(args.text, args.delimiter),
            start = parseRange(args.start ?? "A1"),
            formats: import("../core/types").Command[] = [],
            values = raw.map((row, r) =>
              row.map((text, c) => {
                const parsed = parseInputValue(
                  text,
                  args.header && r === 0
                    ? "text"
                    : (args.columns?.[c] ?? "text"),
                  model.dateSystem,
                );
                if (parsed.format)
                  formats.push({
                    type: "style",
                    sheetId: args.sheetId ?? model.sheets[0].meta.id,
                    range: {
                      r1: start.r1 + r,
                      r2: start.r1 + r,
                      c1: start.c1 + c,
                      c2: start.c1 + c,
                    },
                    style: { numberFormat: parsed.format },
                  });
                return parsed.value;
              }),
            ),
            sheetId = args.sheetId ?? model.sheets[0].meta.id;
          if (values.length) {
            const start = parseRange(args.start ?? "A1");
            change = await model.executeAsync(
              [
                {
                  type: "setValues",
                  sheetId,
                  range: {
                    ...start,
                    r2: start.r1 + values.length - 1,
                    c2: start.c1 + values[0].length - 1,
                  },
                  values,
                  parseFormulas: false,
                },
                ...formats,
              ],
              async (value) => {
                progress({ stage: "apply", progress: value });
                await checkpoint();
              },
              () => validate(),
            );
          }
          result = { diagnostics: [] };
          break;
        }
        case "exportCsv": {
          const sheet = model.sheet(args.sheetId);
          let r2 = 0,
            c2 = 0;
          for (const key of sheet.cells.keys()) {
            r2 = Math.max(r2, rowOf(key));
            c2 = Math.max(c2, columnOf(key));
          }
          for (const [key, cell] of sheet.cells) {
            if (!cell.formula) continue;
            const spill = model.engine.arrayResult(args.sheetId, key);
            if (isArray(spill)) {
              r2 = Math.max(r2, rowOf(key) + spill.rows - 1);
              c2 = Math.max(c2, columnOf(key) + spill.columns - 1);
            }
          }
          const range = args.range ?? { r1: 0, c1: 0, r2, c2 },
            rows: InputValue[][] = [];
          model.validateArea(sheet, range);
          for (let r = range.r1; r <= range.r2; r++) {
            const row: InputValue[] = [];
            for (let c = range.c1; c <= range.c2; c++) {
              const value = model.engine.get(args.sheetId, keyOf(r, c));
              row.push(isError(value) ? value.error : value);
            }
            rows.push(row);
            if (r % 2000 === 0) await checkpoint();
          }
          result = {
            text: writeCsv(rows, args.delimiter),
            diagnostics: [
              {
                code: "CSV_VALUES_ONLY",
                severity: "warning",
                message:
                  "CSV includes displayed values from one sheet; styles and formulas are not retained.",
              },
            ],
          };
          break;
        }
        default:
          throw new Error(`Unknown operation: ${operation}`);
      }
      if (change) {
        Object.assign(change, context);
        emit("change", change);
        emit("diagnostics", model.getDiagnostics());
      }
      emit("calculation", { ...context, status: "idle" });
      if (operation === "exportXlsx") {
        const output = result as { data: Uint8Array };
        scope.postMessage({ id, result }, [output.data.buffer as ArrayBuffer]);
      } else scope.postMessage({ id, result });
    } catch (e) {
      emit("calculation", { ...context, status: "idle" });
      const first = args?.commands?.length === 1 ? args.commands[0] : undefined;
      const err = asWorkbookError(e, {
        operation,
        ...context,
        sheetId: args?.sheetId ?? first?.sheetId,
        range: args?.range ?? first?.range,
        outcome: "rolled-back",
      });
      scope.postMessage({
        id,
        error: {
          ...err,
          name: err.name,
          message: err.message,
          diagnostics:
            e instanceof CompatibilityError ? e.diagnostics : undefined,
        },
      });
    } finally {
      cancelled.delete(id);
    }
  });
};

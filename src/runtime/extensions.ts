import type {
  Command,
  Diagnostic,
  OperationContext,
  Region,
  Scalar,
} from "../core/types";
import type { FunctionContext } from "../formula/registry";
import { functions } from "../formula/functions";
import { isArray, matrix } from "../formula/values";
import { WorkbookError } from "./errors";
export type FunctionValue = Scalar | Scalar[][];
export interface CustomFunction {
  name: string;
  minArgs: number;
  maxArgs: number;
  signature: string;
  description?: string;
  /** Scalars or rectangular arrays. Functions must be synchronous and deterministic. */
  evaluate(args: FunctionValue[], context: FunctionContext): FunctionValue;
}
export interface ValidationContext extends OperationContext {
  operation: string;
  commands: readonly Command[];
  getRegion(sheetId: string, range: string): Region;
  sheets: readonly { id: string; name: string }[];
}
export interface WorkbookModule {
  functions?: CustomFunction[];
  /** Runs against tentative calculated state. Return issues to roll back the whole transaction. */
  validate?(context: ValidationContext): void | Diagnostic[];
}
const errorCodes = new Set([
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
  "#SPILL!",
  "#CALC!",
  "#CYCLE!",
]);
function validScalar(value: unknown): value is Scalar {
  return (
    value === null ||
    typeof value === "string" ||
    typeof value === "boolean" ||
    (typeof value === "number" && Number.isFinite(value)) ||
    (!!value &&
      typeof value === "object" &&
      errorCodes.has((value as { error: string }).error))
  );
}
export function registerFunctions(definitions: CustomFunction[]): void {
  for (const definition of definitions) {
    const { name, minArgs, maxArgs, signature, evaluate } = definition;
    if (
      !/^[A-Z_][A-Z0-9_.]*$/.test(name) ||
      functions.has(name) ||
      !Number.isInteger(minArgs) ||
      !Number.isInteger(maxArgs) ||
      minArgs < 0 ||
      maxArgs < minArgs ||
      maxArgs > 255 ||
      typeof signature !== "string" ||
      (definition.description !== undefined &&
        typeof definition.description !== "string") ||
      typeof evaluate !== "function"
    )
      throw new WorkbookError(
        "MODULE_FAILED",
        `Invalid or duplicate function: ${name}`,
      );
    functions.set(name, {
      name,
      minArgs,
      maxArgs,
      signature,
      category: "custom",
      elementwise: false,
      notes: definition.description,
      evaluate(args, context) {
        try {
          const inputs = args.map((value) => {
            if (!isArray(value)) return value;
            if (value.rows * value.columns > 200000)
              throw new Error("Custom function range exceeds 200000 cells");
            return Array.from({ length: value.rows }, (_, r) =>
              Array.from({ length: value.columns }, (_, c) => value.get(r, c)),
            );
          });
          const output = evaluate(inputs, context);
          if (Array.isArray(output)) {
            const width = output[0]?.length;
            if (
              !width ||
              output.length * width > 200000 ||
              output.some(
                (row) =>
                  !Array.isArray(row) ||
                  row.length !== width ||
                  !row.every(validScalar),
              )
            )
              return { error: "#VALUE!" };
            return matrix(output.map((row) => row.map(cleanScalar)));
          }
          if (
            output &&
            typeof (output as unknown as { then?: unknown }).then === "function"
          )
            void Promise.resolve(output).catch(() => {});
          return validScalar(output)
            ? cleanScalar(output)
            : { error: "#VALUE!" };
        } catch {
          return { error: "#VALUE!" };
        }
      },
    });
  }
}

function cleanScalar(value: Scalar): Scalar {
  return value && typeof value === "object" ? { error: value.error } : value;
}

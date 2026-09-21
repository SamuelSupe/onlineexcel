import type { Value } from "./values";
export interface FunctionContext {
  row: number;
  column: number;
  dateSystem: 1900 | 1904;
  now: Date;
}
export interface FormulaFunction {
  name: string;
  minArgs: number;
  maxArgs: number;
  category: string;
  evaluate(args: Value[], context: FunctionContext): Value;
  elementwise: boolean;
  notes?: string;
  signature?: string;
}
export const functions = new Map<string, FormulaFunction>();
export function define(
  name: string,
  minArgs: number,
  maxArgs: number,
  category: string,
  evaluate: FormulaFunction["evaluate"],
  elementwise = false,
  notes?: string,
): void {
  functions.set(name, {
    name,
    minArgs,
    maxArgs,
    category,
    evaluate,
    elementwise,
    notes,
  });
}

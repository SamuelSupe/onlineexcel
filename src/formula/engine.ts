import { address, columnOf, contains, keyOf, rowOf } from "../core/address";
import type { Cell, Diagnostic, NamedRange, Rect, Scalar } from "../core/types";
import { parseFormula, type AST } from "./parser";
import { functions } from "./functions";
import {
  bool,
  compare,
  error,
  finite,
  isArray,
  isError,
  matrix,
  number,
  scalar,
  text,
  type ArrayValue,
  type Value,
} from "./values";
export interface FormulaHost {
  dateSystem: 1900 | 1904;
  cell(sheetId: string, key: number): Cell | undefined;
  sheetId(name: string): string | undefined;
  extent(sheetId: string): { rows: number; columns: number };
  canSpill(sheetId: string, range: Rect): boolean;
  named(name: string): NamedRange | undefined;
  formulas(): Iterable<{ sheetId: string; key: number; formula: string }>;
}
interface Dependency {
  sheetId: string;
  range: Rect;
}
interface Compiled {
  ast?: AST;
  formula: string;
  dependencies: Dependency[];
  volatile: boolean;
  custom?: boolean;
}
interface Spill {
  owner: string;
  value: Scalar;
}
export class FormulaEngine {
  private compiled = new Map<string, Compiled>();
  private cache = new Map<string, Value>();
  private astCache = new Map<string, AST>();
  private dependents = new Map<string, Set<string>>();
  private rangeDependents = new Map<string, Map<number, Map<string, Rect[]>>>();
  private visiting = new Set<string>();
  private spillCells = new Map<string, Spill>();
  private spillKeys = new Map<string, string[]>();
  private now = new Date();
  private newSpillCells = new Set<string>();
  private dirty = new Set<string>();
  private volatileCells = new Set<string>();
  private blockedSpills = new Set<string>();
  private diagnosticByCell = new Map<string, Diagnostic>();
  constructor(private host: FormulaHost) {}
  private id(sheetId: string, key: number): string {
    return `${sheetId}:${key}`;
  }
  private split(id: string): [string, number] {
    const i = id.lastIndexOf(":");
    return [id.slice(0, i), Number(id.slice(i + 1))];
  }
  rebuild(): void {
    this.compiled.clear();
    this.cache.clear();
    this.dependents.clear();
    this.rangeDependents.clear();
    this.spillCells.clear();
    this.spillKeys.clear();
    this.visiting.clear();
    this.astCache.clear();
    this.newSpillCells.clear();
    this.dirty.clear();
    this.volatileCells.clear();
    this.blockedSpills.clear();
    this.diagnosticByCell.clear();
    for (const cell of this.host.formulas())
      this.updateFormula(cell.sheetId, cell.key, cell.formula);
  }
  private dependency(
    ast: AST,
    sheetId: string,
    dependencies: Dependency[],
  ): void {
    if (ast.type === "ref") {
      const id = ast.sheet ? this.host.sheetId(ast.sheet) : sheetId;
      if (id)
        dependencies.push({
          sheetId: id,
          range: { r1: ast.row, c1: ast.column, r2: ast.row, c2: ast.column },
        });
    } else if (ast.type === "range") {
      const id = ast.start.sheet ? this.host.sheetId(ast.start.sheet) : sheetId;
      if (id)
        dependencies.push({
          sheetId: id,
          range: {
            r1: Math.min(ast.start.row, ast.end.row),
            c1: Math.min(ast.start.column, ast.end.column),
            r2: Math.max(ast.start.row, ast.end.row),
            c2: Math.max(ast.start.column, ast.end.column),
          },
        });
    } else if (ast.type === "name") {
      const named = this.host.named(ast.name);
      if (named)
        dependencies.push({ sheetId: named.sheetId, range: named.range });
    } else if (ast.type === "binary") {
      this.dependency(ast.left, sheetId, dependencies);
      this.dependency(ast.right, sheetId, dependencies);
    } else if (ast.type === "unary")
      this.dependency(ast.value, sheetId, dependencies);
    else if (ast.type === "call")
      ast.args.forEach((a) => this.dependency(a, sheetId, dependencies));
    else if (ast.type === "array")
      ast.rows.flat().forEach((a) => this.dependency(a, sheetId, dependencies));
  }
  private buckets(range: Rect): number[] {
    const start = Math.floor(range.r1 / 256),
      end = Math.floor(range.r2 / 256);
    return end - start > 128
      ? [-1]
      : Array.from({ length: end - start + 1 }, (_, i) => start + i);
  }
  updateFormula(sheetId: string, key: number, formula?: string): void {
    const owner = this.id(sheetId, key),
      previous = this.compiled.get(owner);
    if (!previous && !formula) return;
    if (previous?.formula === formula) return;
    for (const { sheetId: sid, range } of previous?.dependencies ?? []) {
      if (range.r1 === range.r2 && range.c1 === range.c2) {
        const id = this.id(sid, keyOf(range.r1, range.c1)),
          owners = this.dependents.get(id);
        owners?.delete(owner);
        if (!owners?.size) this.dependents.delete(id);
      } else
        for (const bucket of this.buckets(range))
          this.rangeDependents.get(sid)?.get(bucket)?.delete(owner);
    }
    this.compiled.delete(owner);
    this.cache.delete(owner);
    this.dirty.delete(owner);
    this.volatileCells.delete(owner);
    this.blockedSpills.delete(owner);
    this.diagnosticByCell.delete(owner);
    if (!formula) return;
    let ast: AST | undefined;
    try {
      ast = this.astCache.get(formula) ?? parseFormula(formula);
      if (this.astCache.size < 100_000) this.astCache.set(formula, ast);
    } catch {
      /* The original formula remains editable even when parsing fails. */
    }
    const dependencies: Dependency[] = [];
    if (ast) this.dependency(ast, sheetId, dependencies);
    this.compiled.set(owner, {
      ast,
      custom: ast ? usesCustomFunction(ast) : false,
      formula,
      dependencies,
      volatile: /\b(?:NOW|TODAY|RAND|RANDBETWEEN)\s*\(/i.test(formula),
    });
    this.dirty.add(owner);
    if (/\b(?:NOW|TODAY|RAND|RANDBETWEEN)\s*\(/i.test(formula))
      this.volatileCells.add(owner);
    for (const { sheetId: sid, range } of dependencies) {
      if (range.r1 === range.r2 && range.c1 === range.c2) {
        const id = this.id(sid, keyOf(range.r1, range.c1));
        if (!this.dependents.has(id)) this.dependents.set(id, new Set());
        this.dependents.get(id)!.add(owner);
      } else {
        if (!this.rangeDependents.has(sid))
          this.rangeDependents.set(sid, new Map());
        for (const bucket of this.buckets(range)) {
          const index = this.rangeDependents.get(sid)!;
          if (!index.has(bucket)) index.set(bucket, new Map());
          const owners = index.get(bucket)!;
          owners.set(owner, [...(owners.get(owner) ?? []), range]);
        }
      }
    }
  }
  invalidate(
    changes: Iterable<{ sheetId: string; key: number }>,
    includeVolatile = true,
  ): void {
    const queue: string[] = [],
      seen = new Set<string>();
    const add = (id: string) => {
      if (!seen.has(id)) {
        seen.add(id);
        queue.push(id);
      }
    };
    for (const change of changes) add(this.id(change.sheetId, change.key));
    if (includeVolatile) for (const id of this.volatileCells) add(id);
    for (const id of this.blockedSpills) add(id);
    for (let i = 0; i < queue.length; i++) {
      const id = queue[i];
      this.cache.delete(id);
      if (this.compiled.has(id)) this.dirty.add(id);
      for (const spill of this.spillKeys.get(id) ?? []) {
        this.spillCells.delete(spill);
        add(spill);
      }
      this.spillKeys.delete(id);
      for (const owner of this.dependents.get(id) ?? []) add(owner);
      const [sid, key] = this.split(id),
        row = rowOf(key),
        col = columnOf(key);
      for (const bucket of [Math.floor(row / 256), -1])
        for (const [owner, ranges] of this.rangeDependents
          .get(sid)
          ?.get(bucket) ?? [])
          if (ranges.some((range) => contains(range, row, col))) add(owner);
    }
    this.now = new Date();
  }
  recalculate(): void {
    for (let pass = 0; pass < 32; pass++) {
      for (const id of this.dirty) {
        const [sid, key] = this.split(id);
        this.get(sid, key);
      }
      if (!this.newSpillCells.size) return;
      const changed = [...this.newSpillCells].map((id) => {
        const [sheetId, key] = this.split(id);
        return { sheetId, key };
      });
      this.newSpillCells.clear();
      this.invalidate(changed, false);
    }
    for (const id of this.compiled.keys())
      if (!this.cache.has(id)) this.cache.set(id, error("#CALC!"));
  }
  async recalculateAsync(yieldControl: () => Promise<void>): Promise<void> {
    let count = 0;
    for (const id of this.compiled.keys()) {
      const [sid, key] = this.split(id);
      this.get(sid, key);
      if (++count % 1000 === 0) await yieldControl();
    }
  }
  spillOwner(
    sheetId: string,
    key: number,
  ): { row: number; column: number } | undefined {
    const spill = this.spillCells.get(this.id(sheetId, key));
    if (!spill) return;
    const [, anchor] = this.split(spill.owner);
    return { row: rowOf(anchor), column: columnOf(anchor) };
  }
  get(sheetId: string, key: number): Scalar {
    const id = this.id(sheetId, key),
      raw = this.host.cell(sheetId, key);
    if (!raw?.formula)
      return this.spillCells.get(id)?.value ?? raw?.value ?? null;
    if (this.visiting.has(id) || this.visiting.size > 1000)
      return error("#CYCLE!");
    if (this.cache.has(id)) return scalar(this.cache.get(id)!);
    if (!this.compiled.has(id)) this.updateFormula(sheetId, key, raw.formula);
    const compiled = this.compiled.get(id)!;
    this.visiting.add(id);
    let result: Value;
    try {
      result = compiled.ast
        ? this.evaluate(compiled.ast, sheetId, rowOf(key), columnOf(key))
        : error("#VALUE!");
    } catch (e) {
      result = isError(e) ? e : error("#VALUE!");
    }
    if (isArray(result)) {
      const source = result,
        row = rowOf(key),
        col = columnOf(key),
        extent = this.host.extent(sheetId);
      if (source.rows < 1 || source.columns < 1) result = error("#CALC!");
      else if (
        source.rows * source.columns > 1_000_000 ||
        row + source.rows > extent.rows ||
        col + source.columns > extent.columns ||
        (source.rows * source.columns > 1 &&
          !this.host.canSpill(sheetId, {
            r1: row,
            c1: col,
            r2: row + source.rows - 1,
            c2: col + source.columns - 1,
          }))
      )
        result = error("#SPILL!");
      else {
        const values: Scalar[][] = [],
          keys: string[] = [];
        let blocked = false;
        for (let r = 0; r < source.rows && !blocked; r++) {
          const line: Scalar[] = [];
          values.push(line);
          for (let c = 0; c < source.columns; c++) {
            const target = keyOf(row + r, col + c),
              targetId = this.id(sheetId, target),
              cell = this.host.cell(sheetId, target),
              existing = this.spillCells.get(targetId);
            if (
              (r || c) &&
              ((cell &&
                ((cell.value !== undefined && cell.value !== null) ||
                  cell.formula)) ||
                (existing && existing.owner !== id))
            ) {
              blocked = true;
              break;
            }
            try {
              line.push(source.get(r, c));
            } catch (e) {
              line.push(isError(e) ? e : error("#VALUE!"));
            }
            if (r || c) keys.push(targetId);
          }
        }
        if (blocked) result = error("#SPILL!");
        else {
          result = matrix(values);
          this.spillKeys.set(id, keys);
          let i = 0;
          for (let r = 0; r < source.rows; r++)
            for (let c = 0; c < source.columns; c++)
              if (r || c) {
                const target = keys[i++];
                const previous = this.spillCells.get(target);
                if (
                  !previous ||
                  JSON.stringify(previous.value) !==
                    JSON.stringify(values[r][c])
                )
                  this.newSpillCells.add(target);
                this.spillCells.set(target, { owner: id, value: values[r][c] });
              }
        }
      }
    }
    this.visiting.delete(id);
    this.cache.set(id, result);
    this.dirty.delete(id);
    const value = scalar(result);
    if (isError(value) && value.error === "#SPILL!") this.blockedSpills.add(id);
    else this.blockedSpills.delete(id);
    if (
      !compiled.ast ||
      (isError(value) && ["#NAME?", "#CYCLE!"].includes(value.error))
    ) {
      this.diagnosticByCell.set(id, {
        code: !compiled.ast
          ? "INVALID_FORMULA"
          : isError(value) && value.error === "#CYCLE!"
            ? "CIRCULAR_REFERENCE"
            : "UNSUPPORTED_FORMULA",
        severity: "warning",
        message: !compiled.ast
          ? "Formula could not be parsed; original text is preserved."
          : isError(value) && value.error === "#CYCLE!"
            ? "Circular or excessively deep formula dependency."
            : "Unknown function or name; cached Excel values are not used.",
        sheetId,
        range: address(rowOf(key), columnOf(key)),
      });
    } else if (isError(value)) {
      this.diagnosticByCell.set(id, {
        code: value.error === "#SPILL!" ? "SPILL_CONFLICT" : "FORMULA_ERROR",
        severity: "error",
        message: value.error,
        sheetId,
        range: address(rowOf(key), columnOf(key)),
      });
    } else if (compiled.custom) {
      this.diagnosticByCell.set(id, {
        code: "CUSTOM_FUNCTION_REQUIRES_MODULE",
        severity: "warning",
        message:
          "Formula requires a host Worker module. Excel and hosts without that module cannot recalculate it.",
        sheetId,
        range: address(rowOf(key), columnOf(key)),
      });
    } else this.diagnosticByCell.delete(id);
    return value;
  }
  arrayResult(sheetId: string, key: number): Value | undefined {
    this.get(sheetId, key);
    return this.cache.get(this.id(sheetId, key));
  }
  private range(sheetId: string, rect: Rect): ArrayValue {
    return {
      rows: rect.r2 - rect.r1 + 1,
      columns: rect.c2 - rect.c1 + 1,
      get: (r, c) => this.get(sheetId, keyOf(rect.r1 + r, rect.c1 + c)),
      reference: { sheetId, row: rect.r1, column: rect.c1 },
    };
  }
  private broadcast(
    values: Value[],
    evaluate: (args: Value[]) => Value,
  ): Value {
    const rows = Math.max(...values.map((v) => (isArray(v) ? v.rows : 1))),
      columns = Math.max(...values.map((v) => (isArray(v) ? v.columns : 1)));
    if (!values.some(isArray)) return evaluate(values);
    if (rows * columns > 1_000_000) return error("#NUM!");
    return {
      rows,
      columns,
      get: (r, c) => {
        try {
          return scalar(
            evaluate(values.map((v) => this.arrayElement(v, r, c))),
          );
        } catch (e) {
          return isError(e) ? e : error("#VALUE!");
        }
      },
    };
  }
  private arrayElement(value: Value, row: number, column: number): Scalar {
    if (!isArray(value)) return value;
    if (
      (value.rows !== 1 && row >= value.rows) ||
      (value.columns !== 1 && column >= value.columns)
    )
      return error("#N/A");
    return value.get(
      value.rows === 1 ? 0 : row,
      value.columns === 1 ? 0 : column,
    );
  }
  private selectArray(
    test: ArrayValue,
    whenTrue: () => Value,
    whenFalse: () => Value,
  ): Value {
    if (test.rows * test.columns > 1_000_000) return error("#NUM!");
    let needsTrue = false,
      needsFalse = false;
    for (let r = 0; r < test.rows && !(needsTrue && needsFalse); r++)
      for (let c = 0; c < test.columns && !(needsTrue && needsFalse); c++) {
        try {
          if (bool(test.get(r, c))) needsTrue = true;
          else needsFalse = true;
        } catch {
          // Invalid conditions propagate at their output position, not across the array.
        }
      }
    const yes = needsTrue ? whenTrue() : false,
      no = needsFalse ? whenFalse() : false,
      values = [test, yes, no],
      rows = Math.max(
        ...values.map((value) => (isArray(value) ? value.rows : 1)),
      ),
      columns = Math.max(
        ...values.map((value) => (isArray(value) ? value.columns : 1)),
      );
    if (rows * columns > 1_000_000) return error("#NUM!");
    return {
      rows,
      columns,
      get: (r, c) => {
        const condition = this.arrayElement(test, r, c);
        if (isError(condition)) return condition;
        return this.arrayElement(bool(condition) ? yes : no, r, c);
      },
    };
  }
  private evaluate(
    ast: AST,
    sheetId: string,
    row: number,
    column: number,
  ): Value {
    const ev = (a: AST): Value => {
      try {
        return this.evaluate(a, sheetId, row, column);
      } catch (e) {
        return isError(e) ? e : error("#VALUE!");
      }
    };
    switch (ast.type) {
      case "literal":
        return ast.value;
      case "name": {
        const named = this.host.named(ast.name);
        return named ? this.range(named.sheetId, named.range) : error("#NAME?");
      }
      case "ref": {
        const id = ast.sheet ? this.host.sheetId(ast.sheet) : sheetId;
        return id
          ? this.range(id, {
              r1: ast.row,
              r2: ast.row,
              c1: ast.column,
              c2: ast.column,
            })
          : error("#REF!");
      }
      case "range": {
        const id = ast.start.sheet
          ? this.host.sheetId(ast.start.sheet)
          : sheetId;
        if (!id || (ast.end.sheet && this.host.sheetId(ast.end.sheet) !== id))
          return error("#REF!");
        const extent = this.host.extent(id);
        return this.range(id, {
          r1: Math.min(ast.start.row, ast.end.row),
          c1: Math.min(ast.start.column, ast.end.column),
          r2:
            ast.whole === "column"
              ? extent.rows - 1
              : Math.max(ast.start.row, ast.end.row),
          c2:
            ast.whole === "row"
              ? extent.columns - 1
              : Math.max(ast.start.column, ast.end.column),
        });
      }
      case "array":
        return matrix(ast.rows.map((r) => r.map((a) => scalar(ev(a)))));
      case "unary": {
        if (ast.op === "#") {
          if (ast.value.type !== "ref") return error("#REF!");
          const ref = ast.value,
            sid = ref.sheet ? this.host.sheetId(ref.sheet) : sheetId;
          if (!sid) return error("#REF!");
          this.get(sid, keyOf(ref.row, ref.column));
          return (
            this.cache.get(this.id(sid, keyOf(ref.row, ref.column))) ??
            error("#REF!")
          );
        }
        if (ast.op === "@") return scalar(ev(ast.value));
        return this.broadcast([ev(ast.value)], (a) => {
          const n = number(a[0]);
          return ast.op === "%" ? n / 100 : ast.op === "-" ? -n : n;
        });
      }
      case "binary":
        return this.broadcast([ev(ast.left), ev(ast.right)], (a) => {
          const left = scalar(a[0]),
            right = scalar(a[1]);
          if (isError(left)) return left;
          if (isError(right)) return right;
          if (ast.op === "&") return text(left) + text(right);
          if (["=", "<>", "<", ">", "<=", ">="].includes(ast.op)) {
            const order = compare(left, right);
            return ast.op === "="
              ? order === 0
              : ast.op === "<>"
                ? order !== 0
                : ast.op === "<"
                  ? order < 0
                  : ast.op === ">"
                    ? order > 0
                    : ast.op === "<="
                      ? order <= 0
                      : order >= 0;
          }
          const x = number(left),
            y = number(right);
          return ast.op === "+"
            ? finite(x + y)
            : ast.op === "-"
              ? finite(x - y)
              : ast.op === "*"
                ? finite(x * y)
                : ast.op === "/"
                  ? y
                    ? finite(x / y)
                    : error("#DIV/0!")
                  : finite(x ** y);
        });
      case "call": {
        const fn = functions.get(ast.name);
        if (!fn) return error("#NAME?");
        if (ast.args.length < fn.minArgs || ast.args.length > fn.maxArgs)
          return error("#VALUE!");
        const args = ast.args;
        if (ast.name === "IF") {
          const test = ev(args[0]);
          if (!isArray(test) || (test.rows === 1 && test.columns === 1))
            return bool(test) ? ev(args[1]) : args[2] ? ev(args[2]) : false;
          return this.selectArray(
            test,
            () => ev(args[1]),
            () => (args[2] ? ev(args[2]) : false),
          );
        }
        if (ast.name === "IFERROR" || ast.name === "IFNA") {
          const v = ev(args[0]);
          const replace = (x: Scalar) =>
            isError(x) && (ast.name === "IFERROR" || x.error === "#N/A");
          if (!isArray(v)) return replace(v) ? ev(args[1]) : v;
          return this.selectArray(
            {
              rows: v.rows,
              columns: v.columns,
              get: (r, c) => replace(v.get(r, c)),
            },
            () => ev(args[1]),
            () => v,
          );
        }
        if (ast.name === "CHOOSE") {
          const index = Math.trunc(number(ev(args[0])));
          return index < 1 || index >= args.length
            ? error("#VALUE!")
            : ev(args[index]);
        }
        if (ast.name === "IFS") {
          if (args.length % 2) return error("#VALUE!");
          for (let i = 0; i < args.length; i += 2)
            if (bool(ev(args[i]))) return ev(args[i + 1]);
          return error("#N/A");
        }
        if (ast.name === "SWITCH") {
          const value = scalar(ev(args[0]));
          for (let i = 1; i + 1 < args.length; i += 2)
            if (compare(value, scalar(ev(args[i]))) === 0)
              return ev(args[i + 1]);
          return args.length % 2 === 0
            ? ev(args[args.length - 1])
            : error("#N/A");
        }
        const values = args.map(ev),
          ctx = {
            row,
            column,
            dateSystem: this.host.dateSystem,
            now: this.now,
          };
        return fn.elementwise
          ? this.broadcast(values, (a) => fn.evaluate(a, ctx))
          : fn.evaluate(values, ctx);
      }
    }
  }
  diagnostics(): Diagnostic[] {
    return [...this.diagnosticByCell.values()];
  }
}

function usesCustomFunction(ast: AST): boolean {
  switch (ast.type) {
    case "call":
      return (
        functions.get(ast.name)?.category === "custom" ||
        ast.args.some(usesCustomFunction)
      );
    case "binary":
      return usesCustomFunction(ast.left) || usesCustomFunction(ast.right);
    case "unary":
      return usesCustomFunction(ast.value);
    case "array":
      return ast.rows.some((row) => row.some(usesCustomFunction));
    default:
      return false;
  }
}

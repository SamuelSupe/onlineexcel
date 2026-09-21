import {
  columnIndex,
  columnName,
  MAX_COLUMNS,
  MAX_ROWS,
  parseCell,
} from "../core/address";
import type { Scalar } from "../core/types";
export type Ref = {
  type: "ref";
  sheet?: string;
  row: number;
  column: number;
  absoluteRow: boolean;
  absoluteColumn: boolean;
};
export type AST =
  | { type: "literal"; value: Scalar }
  | Ref
  | { type: "name"; name: string }
  | { type: "range"; start: Ref; end: Ref; whole?: "row" | "column" }
  | { type: "binary"; op: string; left: AST; right: AST }
  | { type: "unary"; op: string; value: AST }
  | { type: "call"; name: string; args: AST[]; originalName?: string }
  | { type: "array"; rows: AST[][] };
interface Token {
  kind: string;
  text: string;
}
function tokenize(input: string): Token[] {
  const tokens: Token[] = [];
  let i = input.startsWith("=") ? 1 : 0;
  while (i < input.length) {
    const rest = input.slice(i);
    if (/^\s/.test(rest)) {
      i++;
      continue;
    }
    const quoted = /^("(?:[^"]|"")*"|'(?:[^']|'')*')/.exec(rest);
    const error =
      /^#(?:DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|CYCLE!)/.exec(
        rest,
      );
    const number = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/.exec(rest);
    const word =
      /^(?:\$?[A-Za-z_\u0080-\uffff][A-Za-z0-9_.$\u0080-\uffff]*|\$[1-9]\d*)/.exec(
        rest,
      );
    const op = /^(?:<>|<=|>=|[+\-*/^&=<>(),;:{}!%#@])/.exec(rest);
    const match = quoted ?? error ?? number ?? word ?? op;
    if (!match)
      throw new Error(`Unexpected formula character at ${i}: ${rest[0]}`);
    tokens.push({
      kind: quoted
        ? rest[0] === '"'
          ? "string"
          : "sheet"
        : error
          ? "error"
          : number
            ? "number"
            : word
              ? "word"
              : "op",
      text: match[0],
    });
    i += match[0].length;
  }
  tokens.push({ kind: "eof", text: "" });
  return tokens;
}
const precedence: Record<string, number> = {
  "=": 1,
  "<>": 1,
  "<": 1,
  ">": 1,
  "<=": 1,
  ">=": 1,
  "&": 2,
  "+": 3,
  "-": 3,
  "*": 4,
  "/": 4,
  "^": 5,
};
export function parseFormula(input: string): AST {
  const tokens = tokenize(input);
  let at = 0;
  const peek = () => tokens[at];
  const take = () => tokens[at++];
  const expect = (text: string) => {
    if (take().text !== text) throw new Error(`Expected ${text}`);
  };
  const ref = (text: string, sheet?: string): AST => {
    try {
      return { type: "ref", ...parseCell(text), sheet };
    } catch {
      return { type: "name", name: sheet ? `${sheet}!${text}` : text };
    }
  };
  function atom(): AST {
    const token = take();
    let result: AST;
    if (token.text === "+" || token.text === "-" || token.text === "@")
      result = { type: "unary", op: token.text, value: expression(6) };
    else if (token.text === "(") {
      result = expression(0);
      expect(")");
    } else if (token.text === "{") {
      const rows: AST[][] = [[]];
      while (peek().text !== "}") {
        rows[rows.length - 1].push(expression(0));
        if (peek().text === ";") {
          take();
          rows.push([]);
        } else if (peek().text === ",") take();
        else break;
      }
      expect("}");
      if (rows.some((r) => r.length !== rows[0].length))
        throw new Error("Unequal array columns");
      result = { type: "array", rows };
    } else if (
      (token.kind === "word" ||
        token.kind === "sheet" ||
        token.kind === "number") &&
      peek().text === "!"
    ) {
      take();
      const sheet =
        token.kind === "sheet"
          ? token.text.slice(1, -1).replace(/''/g, "'")
          : token.text;
      result = ref(take().text, sheet);
    } else if (token.kind === "number")
      result = { type: "literal", value: Number(token.text) };
    else if (token.kind === "string")
      result = {
        type: "literal",
        value: token.text.slice(1, -1).replace(/""/g, '"'),
      };
    else if (token.kind === "error")
      result = { type: "literal", value: { error: token.text as "#REF!" } };
    else if (token.kind === "word") {
      if (peek().text === "(") {
        take();
        const args: AST[] = [];
        while (peek().text !== ")") {
          args.push(
            peek().text === ","
              ? { type: "literal", value: null }
              : expression(0),
          );
          if (peek().text !== ",") break;
          take();
          if (peek().text === ")") args.push({ type: "literal", value: null });
        }
        expect(")");
        const name = token.text
          .toUpperCase()
          .replace(/^_XLFN\./, "")
          .replace(/^_XLWS\./, "");
        result =
          ["ANCHORARRAY", "SINGLE"].includes(name) && args.length === 1
            ? {
                type: "unary",
                op: name === "ANCHORARRAY" ? "#" : "@",
                value: args[0],
              }
            : { type: "call", name, args, originalName: token.text };
      } else if (/^(TRUE|FALSE)$/i.test(token.text))
        result = {
          type: "literal",
          value: token.text.toUpperCase() === "TRUE",
        };
      else result = ref(token.text);
    } else throw new Error("Expected expression");
    if (peek().text === ":") {
      take();
      const end = atom();
      if (result.type === "ref" && end.type === "ref")
        result = {
          type: "range",
          start: result,
          end: { ...end, sheet: end.sheet ?? result.sheet },
        };
      else {
        const endpoint = (node: AST) =>
          node.type === "name"
            ? node.name
            : node.type === "literal" && typeof node.value === "number"
              ? String(node.value)
              : "";
        const pattern = /^(?:(.+)!)?(\$?)([A-Za-z]{1,3}|[1-9]\d*)$/;
        const a = pattern.exec(endpoint(result)),
          b = pattern.exec(endpoint(end));
        if (!a || !b) throw new Error("Invalid range");
        const whole = /^\d+$/.test(a[3]) ? "row" : "column";
        if ((/^\d+$/.test(b[3]) ? "row" : "column") !== whole)
          throw new Error("Invalid range");
        const first = whole === "row" ? Number(a[3]) - 1 : columnIndex(a[3]),
          last = whole === "row" ? Number(b[3]) - 1 : columnIndex(b[3]),
          limit = whole === "row" ? MAX_ROWS : MAX_COLUMNS;
        if (first >= limit || last >= limit) throw new Error("Invalid range");
        result = {
          type: "range",
          whole,
          start: {
            type: "ref",
            sheet: a[1],
            row: whole === "row" ? first : 0,
            column: whole === "column" ? first : 0,
            absoluteRow: whole === "column" || !!a[2],
            absoluteColumn: whole === "row" || !!a[2],
          },
          end: {
            type: "ref",
            sheet: b[1] ?? a[1],
            row: whole === "row" ? last : MAX_ROWS - 1,
            column: whole === "column" ? last : MAX_COLUMNS - 1,
            absoluteRow: whole === "column" || !!b[2],
            absoluteColumn: whole === "row" || !!b[2],
          },
        };
      }
    }
    while (peek().text === "%" || peek().text === "#")
      result = { type: "unary", op: take().text, value: result };
    return result;
  }
  function expression(min: number): AST {
    let left = atom();
    while (
      precedence[peek().text] !== undefined &&
      precedence[peek().text] >= min
    ) {
      const op = take().text,
        p = precedence[op];
      left = { type: "binary", op, left, right: expression(p + 1) };
    }
    return left;
  }
  const ast = expression(0);
  if (peek().kind !== "eof") throw new Error(`Unexpected token ${peek().text}`);
  return ast;
}
const quoteSheet = (name: string): string => `'${name.replace(/'/g, "''")}'!`;
export function printFormula(ast: AST): string {
  function print(node: AST): string {
    switch (node.type) {
      case "literal":
        return typeof node.value === "string"
          ? `"${node.value.replace(/"/g, '""')}"`
          : typeof node.value === "boolean"
            ? String(node.value).toUpperCase()
            : node.value && typeof node.value === "object"
              ? node.value.error
              : String(node.value ?? "");
      case "ref":
        return (
          (node.sheet ? quoteSheet(node.sheet) : "") +
          (node.absoluteColumn ? "$" : "") +
          columnName(node.column) +
          (node.absoluteRow ? "$" : "") +
          (node.row + 1)
        );
      case "name":
        return node.name;
      case "range": {
        const prefix = node.start.sheet ? quoteSheet(node.start.sheet) : "";
        if (node.whole === "column")
          return (
            prefix +
            (node.start.absoluteColumn ? "$" : "") +
            columnName(node.start.column) +
            ":" +
            (node.end.absoluteColumn ? "$" : "") +
            columnName(node.end.column)
          );
        if (node.whole === "row")
          return (
            prefix +
            (node.start.absoluteRow ? "$" : "") +
            (node.start.row + 1) +
            ":" +
            (node.end.absoluteRow ? "$" : "") +
            (node.end.row + 1)
          );
        return (
          print(node.start) +
          ":" +
          print({
            ...node.end,
            sheet:
              node.end.sheet === node.start.sheet ? undefined : node.end.sheet,
          })
        );
      }
      case "binary":
        return `(${print(node.left)}${node.op}${print(node.right)})`;
      case "unary":
        return node.op === "%" || node.op === "#"
          ? print(node.value) + node.op
          : node.op + print(node.value);
      case "call":
        return `${node.originalName ?? node.name}(${node.args.map(print).join(",")})`;
      case "array":
        return (
          "{" + node.rows.map((row) => row.map(print).join(",")).join(";") + "}"
        );
    }
  }
  return "=" + print(ast);
}
export function mapReferences(
  ast: AST,
  map: (ref: Ref) => AST,
  mapRange?: (range: Extract<AST, { type: "range" }>) => AST,
): AST {
  switch (ast.type) {
    case "ref":
      return map(ast);
    case "range": {
      if (mapRange) return mapRange(ast);
      const start = map(ast.start),
        end = map(ast.end);
      return start.type === "ref" && end.type === "ref"
        ? { ...ast, start, end }
        : { type: "literal", value: { error: "#REF!" } };
    }
    case "binary":
      return {
        ...ast,
        left: mapReferences(ast.left, map, mapRange),
        right: mapReferences(ast.right, map, mapRange),
      };
    case "unary":
      return { ...ast, value: mapReferences(ast.value, map, mapRange) };
    case "call":
      return {
        ...ast,
        args: ast.args.map((a) => mapReferences(a, map, mapRange)),
      };
    case "array":
      return {
        ...ast,
        rows: ast.rows.map((row) =>
          row.map((a) => mapReferences(a, map, mapRange)),
        ),
      };
    default:
      return ast;
  }
}
export function shiftFormula(formula: string, dr: number, dc: number): string {
  try {
    return printFormula(
      mapReferences(parseFormula(formula), (ref) => {
        const row = ref.row + (ref.absoluteRow ? 0 : dr),
          column = ref.column + (ref.absoluteColumn ? 0 : dc);
        return row < 0 || column < 0 || row >= MAX_ROWS || column >= MAX_COLUMNS
          ? { type: "literal", value: { error: "#REF!" } }
          : { ...ref, row, column };
      }),
    );
  } catch {
    return formula;
  }
}

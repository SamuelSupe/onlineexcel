import type { InputValue } from "../core/types";
export function parseCsv(input: string, delimiter = ","): string[][] {
  if (delimiter.length !== 1 || /[\r\n"]/.test(delimiter))
    throw new Error("Invalid CSV delimiter");
  input = input.replace(/^\uFEFF/, "");
  const rows: string[][] = [];
  let row: string[] = [],
    value = "",
    quoted = false,
    closed = false;
  for (let i = 0; i < input.length; i++) {
    const char = input[i];
    if (quoted) {
      if (char === '"') {
        if (input[i + 1] === '"') {
          value += '"';
          i++;
        } else {
          quoted = false;
          closed = true;
        }
      } else value += char;
    } else if (char === '"' && value === "" && !closed) quoted = true;
    else if (char === delimiter) {
      row.push(value);
      value = "";
      closed = false;
    } else if (char === "\r" || char === "\n") {
      if (char === "\r" && input[i + 1] === "\n") i++;
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
      closed = false;
    } else {
      if (closed || char === '"') throw new Error("Malformed CSV quoting");
      value += char;
    }
  }
  if (quoted) throw new Error("Unclosed CSV quoted field");
  if (value !== "" || row.length || (input.length && !/[\r\n]$/.test(input))) {
    row.push(value);
    rows.push(row);
  }
  if (!rows.length) return [];
  const width = rows.reduce((n, row) => Math.max(n, row.length), 0);
  return rows.map((row) => [...row, ...Array(width - row.length).fill("")]);
}
export function writeCsv(rows: InputValue[][], delimiter = ","): string {
  if (delimiter.length !== 1 || /[\r\n"]/.test(delimiter))
    throw new Error("Invalid CSV delimiter");
  const escape = (value: InputValue) => {
    const text = value === null ? "" : String(value);
    return text.includes(delimiter) || /[\r\n"]/.test(text)
      ? '"' + text.replace(/"/g, '""') + '"'
      : text;
  };
  return (
    "\uFEFF" +
    rows
      .map((row) => {
        const line = row.map(escape).join(delimiter);
        return line === "" && row.length === 1 ? '""' : line;
      })
      .join("\r\n")
  );
}

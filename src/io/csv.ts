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
export class CsvSizeError extends Error {}
export function writeCsvRow(
  row: InputValue[],
  delimiter = ",",
  maxLength = Infinity,
): string {
  if (delimiter.length !== 1 || /[\r\n"]/.test(delimiter))
    throw new Error("Invalid CSV delimiter");
  let length = Math.max(0, row.length - 1);
  const escape = (value: InputValue) => {
    const text = value === null ? "" : String(value);
    const quoted = text.includes(delimiter) || /[\r\n"]/.test(text);
    length += text.length + (quoted ? 2 : 0);
    if (length > maxLength)
      throw new CsvSizeError("CSV output exceeds its size limit");
    if (quoted && Number.isFinite(maxLength))
      for (let i = 0; i < text.length; i++)
        if (text[i] === '"' && ++length > maxLength)
          throw new CsvSizeError("CSV output exceeds its size limit");
    return quoted ? '"' + text.replace(/"/g, '""') + '"' : text;
  };
  const line = row.map(escape).join(delimiter);
  const result = line === "" && row.length === 1 ? '""' : line;
  if (result.length > maxLength)
    throw new CsvSizeError("CSV output exceeds its size limit");
  return result;
}
export function writeCsv(rows: InputValue[][], delimiter = ","): string {
  // Validate even when the input contains no rows.
  writeCsvRow([], delimiter);
  return "\uFEFF" + rows.map((row) => writeCsvRow(row, delimiter)).join("\r\n");
}

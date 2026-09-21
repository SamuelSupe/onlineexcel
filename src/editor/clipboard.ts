import type { CellStyle, Rect } from "../core/types";
import { keyOf } from "../core/address";
import { parseCsv } from "../io/csv";

export interface ClipboardPayload {
  text: string;
  html: string;
}

export function createClipboardPayload(text: string) {
  const token = crypto.randomUUID();
  const table = document.createElement("table");
  table.dataset.onlineexcelClipboard = token;
  for (const values of parseCsv(text, "\t")) {
    const row = table.insertRow();
    for (const value of values) row.insertCell().textContent = value;
  }
  return { text, html: table.outerHTML, token };
}

export async function writeClipboard(data: ClipboardPayload): Promise<void> {
  if (navigator.clipboard.write && typeof ClipboardItem !== "undefined")
    await navigator.clipboard.write([
      new ClipboardItem({
        "text/plain": new Blob([data.text], { type: "text/plain" }),
        "text/html": new Blob([data.html], { type: "text/html" }),
      }),
    ]);
  else await navigator.clipboard.writeText(data.text);
}

export async function readClipboard(): Promise<ClipboardPayload> {
  if (!navigator.clipboard.read)
    return { text: await navigator.clipboard.readText(), html: "" };
  const items = await navigator.clipboard.read();
  const item =
    items.find((entry) => entry.types.includes("text/html")) ??
    items.find((entry) => entry.types.includes("text/plain"));
  const read = async (type: string) =>
    item?.types.includes(type) ? (await item.getType(type)).text() : "";
  const [text, html] = await Promise.all([
    read("text/plain"),
    read("text/html"),
  ]);
  return { text, html };
}

export function parseHtmlTable(html: string) {
  const table = new DOMParser()
    .parseFromString(html, "text/html")
    .querySelector("table");
  if (!table) return undefined;
  const rows = [...table.rows].filter((row) => row.closest("table") === table);
  if (!rows.length) return undefined;
  if (rows.length > 200000) throw new Error("Invalid clipboard dimensions");
  const groupEnds = new Array<number>(rows.length);
  for (let r = rows.length - 1; r >= 0; r--)
    groupEnds[r] =
      rows[r + 1]?.parentElement === rows[r].parentElement
        ? groupEnds[r + 1]
        : r + 1;
  const values: (string | null)[][] = rows.map(() => []);
  const styles = new Map<number, CellStyle>();
  const merges: Rect[] = [];
  const occupied = new Set<number>();
  let width = 0;
  for (let r = 0; r < rows.length; r++) {
    let c = 0;
    for (const cell of rows[r].cells) {
      while (occupied.has(keyOf(r, c))) c++;
      const groupEnd = groupEnds[r];
      const height = Math.min(cell.rowSpan || groupEnd - r, groupEnd - r);
      const columns = cell.colSpan;
      width = Math.max(width, c + columns);
      if (width > 16384 || rows.length * width > 200000)
        throw new Error("Invalid clipboard dimensions");
      for (let dr = 0; dr < height; dr++)
        for (let dc = 0; dc < columns; dc++) {
          const key = keyOf(r + dr, c + dc);
          if (occupied.has(key))
            throw new Error("Overlapping cells in clipboard table");
          occupied.add(key);
          values[r + dr][c + dc] = null;
        }
      const content = cell.cloneNode(true) as HTMLElement;
      content.querySelectorAll("br").forEach((br) => br.replaceWith("\n"));
      values[r][c] = content.textContent ?? "";
      const style: CellStyle = {};
      if (
        cell.style.fontWeight === "bold" ||
        Number(cell.style.fontWeight) >= 600
      )
        style.bold = true;
      if (cell.style.fontStyle === "italic") style.italic = true;
      if (["left", "right", "center"].includes(cell.style.textAlign))
        style.align = cell.style.textAlign as CellStyle["align"];
      if (Object.keys(style).length) styles.set(keyOf(r, c), style);
      if (height > 1 || columns > 1)
        merges.push({ r1: r, r2: r + height - 1, c1: c, c2: c + columns - 1 });
      c += columns;
    }
  }
  if (!width) return undefined;
  return {
    values: values.map((row) =>
      Array.from({ length: width }, (_, c) => row[c] ?? null),
    ),
    styles,
    merges,
    token: table.dataset.onlineexcelClipboard,
  };
}

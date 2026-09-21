import { SaxesParser } from "saxes";
import { Unzip, UnzipInflate } from "fflate";
export type XmlAttributes = Record<string, string>;
export interface XmlNode {
  name: string;
  attributes: XmlAttributes;
  children: XmlNode[];
  text: string;
}
export const xmlEscape = (text: unknown): string =>
  String(text ?? "")
    .replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
export const local = (name: string): string => name.split(":").pop()!;
export function xmlParser(handlers: {
  open?: (name: string, attrs: XmlAttributes) => void;
  close?: (name: string) => void;
  text?: (text: string) => void;
}): SaxesParser {
  const parser = new SaxesParser();
  parser.on("doctype", () => {
    throw new Error("DOCTYPE is not permitted in XLSX XML");
  });
  parser.on("opentag", (tag) =>
    handlers.open?.(local(tag.name), tag.attributes as XmlAttributes),
  );
  parser.on("closetag", (tag) => handlers.close?.(local(tag.name)));
  parser.on("text", (text) => handlers.text?.(text));
  parser.on("cdata", (text) => handlers.text?.(text));
  return parser;
}
export function xmlTree(bytes: Uint8Array | undefined): XmlNode {
  if (!bytes) throw new Error("Missing XML part");
  const root: XmlNode = { name: "", attributes: {}, children: [], text: "" },
    stack = [root];
  let count = 0;
  const parser = xmlParser({
    open: (name, attributes) => {
      if (++count > 200_000 || stack.length > 100)
        throw new Error("XML part too complex");
      const child: XmlNode = { name, attributes, children: [], text: "" };
      stack[stack.length - 1].children.push(child);
      stack.push(child);
    },
    text: (text) => {
      stack[stack.length - 1].text += text;
    },
    close: () => {
      stack.pop();
    },
  });
  parser.write(new TextDecoder().decode(bytes)).close();
  return root.children[0];
}
export const child = (
  node: XmlNode | undefined,
  name: string,
): XmlNode | undefined => node?.children.find((n) => n.name === name);
export const children = (node: XmlNode | undefined, name: string): XmlNode[] =>
  node?.children.filter((n) => n.name === name) ?? [];
export async function parseXmlChunks(
  bytes: Uint8Array,
  parser: SaxesParser,
  checkpoint: () => Promise<void>,
): Promise<void> {
  const decoder = new TextDecoder();
  for (let offset = 0; offset < bytes.length; offset += 65536) {
    parser.write(
      decoder.decode(bytes.subarray(offset, offset + 65536), { stream: true }),
    );
    if (offset % 524288 === 0) await checkpoint();
  }
  parser.write(decoder.decode()).close();
}
export async function readArchive(
  data: Uint8Array,
  checkpoint: (progress: number) => Promise<void>,
): Promise<Map<string, Uint8Array>> {
  if (data[0] !== 0x50 || data[1] !== 0x4b)
    throw new Error("Only unencrypted .xlsx ZIP packages are supported");
  if (data.length > 128 * 1024 * 1024)
    throw new Error("Compressed workbook exceeds the 128 MiB import limit");
  const files = new Map<string, Uint8Array>();
  let total = 0,
    failure: Error | undefined,
    entries = 0;
  const unzip = new Unzip((file) => {
    if (
      ++entries > 10000 ||
      /(^|\/)\.\.(\/|$)|^\//.test(file.name) ||
      files.has(file.name)
    )
      throw new Error("Invalid archive path or entry count");
    const parts: Uint8Array[] = [];
    let size = 0;
    file.ondata = (err, data, final) => {
      if (err) {
        failure = err;
        return;
      }
      size += data.length;
      total += data.length;
      if (size > 256 * 1024 * 1024 || total > 512 * 1024 * 1024)
        throw new Error("Uncompressed workbook exceeds import limits");
      parts.push(data);
      if (final) {
        const bytes = new Uint8Array(size);
        let offset = 0;
        for (const part of parts) {
          bytes.set(part, offset);
          offset += part.length;
        }
        files.set(file.name, bytes);
      }
    };
    file.start();
  });
  unzip.register(UnzipInflate);
  for (let offset = 0; offset < data.length; offset += 65536) {
    unzip.push(
      data.subarray(offset, offset + 65536),
      offset + 65536 >= data.length,
    );
    if (failure) throw failure;
    await checkpoint(Math.min(1, (offset + 65536) / data.length));
  }
  return files;
}

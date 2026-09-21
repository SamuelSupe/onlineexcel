import type { PreservedObjects } from "../core/types";
import { children, xmlTree, type XmlNode } from "./xml";
import { serializeNode } from "./xlsx-rules";

const relFile = (path: string) => path.replace(/([^/]+)$/, "_rels/$1.rels");
function resolve(base: string, target: string): string {
  const out: string[] = [];
  for (const part of (target.startsWith("/")
    ? target.slice(1)
    : base.slice(0, base.lastIndexOf("/") + 1) + target
  ).split("/")) {
    if (part === "..") out.pop();
    else if (part && part !== ".") out.push(part);
  }
  return out.join("/");
}
function encode(bytes: Uint8Array) {
  let binary = "";
  for (let i = 0; i < bytes.length; i += 32768)
    binary += String.fromCharCode(...bytes.subarray(i, i + 32768));
  return btoa(binary);
}
const decode = (text: string) =>
  Uint8Array.from(atob(text), (c) => c.charCodeAt(0));
export function preserveObjects(
  files: Map<string, Uint8Array>,
  sheetPath: string,
  nodes: XmlNode[],
): { objects?: PreservedObjects; retained: Set<string> } {
  const path = relFile(sheetPath),
    retained = new Set<string>(),
    relationships = files.has(path) ? xmlTree(files.get(path)) : undefined;
  const selected = children(relationships, "Relationship").filter(
    (rel) =>
      /\/(drawing|vmlDrawing|comments|table)$/.test(rel.attributes.Type) &&
      rel.attributes.TargetMode !== "External",
  );
  if (!selected.length) return { retained };
  const parts: Record<string, string> = {};
  const visit = (file: string) => {
    if (retained.has(file)) return;
    const bytes = files.get(file);
    if (!bytes) throw new Error("Missing related object part");
    retained.add(file);
    parts[file] = encode(bytes);
    const rels = relFile(file);
    if (files.has(rels)) {
      retained.add(rels);
      parts[rels] = encode(files.get(rels)!);
      for (const rel of children(xmlTree(files.get(rels)), "Relationship"))
        if (rel.attributes.TargetMode !== "External") {
          if (
            !/\/(image|chart|chartStyle|chartColorStyle|drawing|vmlDrawing|comments|table|package)$/.test(
              rel.attributes.Type,
            )
          )
            throw new Error("Unsupported object relationship");
          visit(resolve(file, rel.attributes.Target));
        }
    }
  };
  for (const rel of selected) {
    const target = resolve(sheetPath, rel.attributes.Target);
    visit(target);
    rel.attributes.Target = "/" + target;
  }
  const types = xmlTree(files.get("[Content_Types].xml"));
  return {
    retained,
    objects: {
      elements: nodes
        .filter((node) =>
          ["drawing", "legacyDrawing", "tableParts"].includes(node.name),
        )
        .map(serializeNode),
      relationships: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${selected.map(serializeNode).join("")}</Relationships>`,
      parts,
      contentTypes: types.children
        .filter(
          (node) =>
            node.name === "Default" ||
            retained.has((node.attributes.PartName ?? "").replace(/^\//, "")),
        )
        .map(serializeNode),
    },
  };
}
export function exportObjects(objects: PreservedObjects, index: number) {
  const paths = Object.keys(objects.parts).filter(
    (path) => !path.endsWith(".rels"),
  );
  const mapping = new Map(
    paths.map((path, i) => [
      path,
      `xl/onlineexcel-objects/s${index}/p${i}.${path.split(".").at(-1)}`,
    ]),
  );
  const parts = new Map<string, Uint8Array>();
  function rewriteRelations(bytes: Uint8Array, base: string) {
    const root = xmlTree(bytes);
    for (const rel of children(root, "Relationship"))
      if (rel.attributes.TargetMode !== "External") {
        const target = mapping.get(resolve(base, rel.attributes.Target));
        if (!target) throw new Error("Missing preserved object target");
        rel.attributes.Target = "/" + target;
      }
    return new TextEncoder().encode(serializeNode(root));
  }
  for (const [old, path] of mapping) {
    if (!/^xl\//.test(old) || old.includes(".."))
      throw new Error("Invalid preserved object path");
    parts.set(path, decode(objects.parts[old]));
    if (objects.parts[relFile(old)])
      parts.set(
        relFile(path),
        rewriteRelations(decode(objects.parts[relFile(old)]), old),
      );
  }
  const types = objects.contentTypes.flatMap((text) => {
    const node = xmlTree(new TextEncoder().encode(text));
    if (node.name === "Default") return [serializeNode(node)];
    const target = mapping.get(node.attributes.PartName.replace(/^\//, ""));
    if (!target) return [];
    node.attributes.PartName = "/" + target;
    return [serializeNode(node)];
  });
  return {
    parts,
    types,
    relationships: rewriteRelations(
      new TextEncoder().encode(objects.relationships),
      "",
    ),
    elements: objects.elements.join(""),
  };
}

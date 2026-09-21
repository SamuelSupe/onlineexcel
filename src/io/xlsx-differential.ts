import type { CellStyle, ConditionalRule } from "../core/types";
import { builtinFormats } from "./xlsx-styles";
import { child, xmlEscape as x, type XmlNode } from "./xml";

export function readDifferential(
  node: XmlNode,
  report: (message: string) => void,
): CellStyle {
  const style: CellStyle = {};
  const unsupported = () =>
    report("Unsupported conditional-format styling was normalized or dropped.");
  const color = (n: XmlNode | undefined) => {
    if (!n) return;
    if (n.attributes.rgb) return "#" + n.attributes.rgb.slice(-6);
    unsupported();
  };
  const font = child(node, "font");
  for (const [tag, key] of [
    ["b", "bold"],
    ["i", "italic"],
    ["u", "underline"],
  ] as const) {
    const n = child(font, tag);
    if (n)
      style[key] = !["0", "false", "none"].includes(n.attributes.val ?? "1");
    if (
      tag === "u" &&
      n?.attributes.val &&
      !["single", "none"].includes(n.attributes.val)
    )
      unsupported();
  }
  if (child(font, "name"))
    style.fontFamily = child(font, "name")!.attributes.val;
  if (child(font, "sz"))
    style.fontSize = Number(child(font, "sz")!.attributes.val);
  const foreground = color(child(font, "color"));
  if (foreground !== undefined) style.color = foreground;
  if (
    font?.children.some(
      (n) =>
        ![
          "b",
          "i",
          "u",
          "name",
          "sz",
          "color",
          "family",
          "charset",
          "scheme",
        ].includes(n.name),
    )
  )
    unsupported();
  const fill = child(node, "fill"),
    pattern = child(fill, "patternFill");
  if (pattern?.attributes.patternType === "none") style.background = "";
  else if (pattern?.attributes.patternType === "solid") {
    const background = color(child(pattern, "fgColor"));
    if (background !== undefined) style.background = background;
  } else if (fill) unsupported();
  const format = child(node, "numFmt");
  if (format) {
    style.numberFormat =
      format.attributes.formatCode ??
      builtinFormats[Number(format.attributes.numFmtId)];
    if (style.numberFormat === undefined) unsupported();
  }
  const alignment = child(node, "alignment");
  if (alignment) {
    const { horizontal, vertical, wrapText, ...extra } = alignment.attributes;
    if (horizontal !== undefined) {
      if (["left", "center", "right"].includes(horizontal))
        style.align = horizontal as CellStyle["align"];
      else unsupported();
    }
    if (vertical !== undefined) {
      if (["top", "center", "bottom"].includes(vertical))
        style.verticalAlign = vertical as CellStyle["verticalAlign"];
      else unsupported();
    }
    if (wrapText !== undefined) style.wrap = ["1", "true"].includes(wrapText);
    if (Object.keys(extra).length) unsupported();
  }
  const border = child(node, "border");
  if (border) {
    style.border = {};
    for (const side of ["left", "right", "top", "bottom"] as const) {
      const edge = child(border, side);
      if (!edge) continue;
      style.border[side] = edge.attributes.style
        ? (color(child(edge, "color")) ?? "#000000")
        : "";
      if (edge.attributes.style && edge.attributes.style !== "thin")
        unsupported();
    }
    if (
      border.children.some(
        (n) =>
          !["left", "right", "top", "bottom"].includes(n.name) &&
          (n.attributes.style || n.children.length),
      )
    )
      unsupported();
  }
  const protection = child(node, "protection");
  if (protection?.attributes.locked !== undefined)
    style.locked = ["1", "true"].includes(protection.attributes.locked);
  if (protection?.attributes.hidden !== undefined) unsupported();
  if (
    node.children.some(
      (n) =>
        ![
          "font",
          "numFmt",
          "fill",
          "alignment",
          "border",
          "protection",
        ].includes(n.name),
    )
  )
    unsupported();
  return style;
}

export function writeDxfs(rules: ConditionalRule[], firstFormatId: number) {
  const rgb = (value: string) => {
    const hex = value.replace("#", "");
    return (
      "FF" +
      (hex.length === 3
        ? [...hex].map((c) => c + c).join("")
        : hex
      ).toUpperCase()
    );
  };
  const records = rules.map(({ style: s }, i) => {
    const font = [
      s.fontFamily === undefined ? "" : `<name val="${x(s.fontFamily)}"/>`,
      s.fontSize === undefined ? "" : `<sz val="${s.fontSize}"/>`,
      s.bold === undefined ? "" : `<b val="${+s.bold}"/>`,
      s.italic === undefined ? "" : `<i val="${+s.italic}"/>`,
      s.underline === undefined
        ? ""
        : `<u val="${s.underline ? "single" : "none"}"/>`,
      s.color ? `<color rgb="${rgb(s.color)}"/>` : "",
    ].join("");
    const alignment = [
      s.align === undefined ? "" : ` horizontal="${s.align}"`,
      s.verticalAlign === undefined ? "" : ` vertical="${s.verticalAlign}"`,
      s.wrap === undefined ? "" : ` wrapText="${+s.wrap}"`,
    ].join("");
    const border =
      s.border === undefined
        ? ""
        : `<border>${Object.entries(s.border)
            .map(([side, c]) =>
              c
                ? `<${side} style="thin"><color rgb="${rgb(c)}"/></${side}>`
                : `<${side}/>`,
            )
            .join("")}</border>`;
    const fill =
      s.background === undefined
        ? ""
        : `<fill><patternFill patternType="${s.background ? "solid" : "none"}">${s.background ? `<fgColor rgb="${rgb(s.background)}"/><bgColor indexed="64"/>` : ""}</patternFill></fill>`;
    return [
      "<dxf>",
      font ? `<font>${font}</font>` : "",
      s.numberFormat === undefined
        ? ""
        : `<numFmt numFmtId="${firstFormatId + i}" formatCode="${x(s.numberFormat)}"/>`,
      fill,
      alignment ? `<alignment${alignment}/>` : "",
      border,
      s.locked === undefined ? "" : `<protection locked="${+s.locked}"/>`,
      "</dxf>",
    ].join("");
  });
  return `<dxfs count="${records.length}">${records.join("")}</dxfs>`;
}

import type { CellStyle, Diagnostic } from "../core/types";
import { child, children, xmlEscape as x, xmlTree, type XmlNode } from "./xml";
export const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
export const REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
export const builtinFormats: Record<number, string> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  14: "mm-dd-yy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yy h:mm",
  37: "#,##0;(#,##0)",
  38: "#,##0;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  49: "@",
};
export function readStyles(
  files: Map<string, Uint8Array>,
  path: string,
  diagnostics: Diagnostic[],
): CellStyle[] {
  if (!files.has(path)) return [{}];
  const root = xmlTree(files.get(path));
  const report = (code: string, message: string) => {
    if (!diagnostics.some((d) => d.code === code))
      diagnostics.push({ code, message, severity: "warning", lossy: true });
  };
  const theme = [
    "#FFFFFF",
    "#000000",
    "#E7E6E6",
    "#44546A",
    "#4472C4",
    "#ED7D31",
    "#A5A5A5",
    "#FFC000",
    "#5B9BD5",
    "#70AD47",
    "#0563C1",
    "#954F72",
  ];
  const themeFile = [...files.keys()].find((n) =>
    /theme\/theme\d+\.xml$/.test(n),
  );
  if (themeFile) {
    const scheme = child(
      child(xmlTree(files.get(themeFile)), "themeElements"),
      "clrScheme",
    );
    const order = [
      "lt1",
      "dk1",
      "lt2",
      "dk2",
      "accent1",
      "accent2",
      "accent3",
      "accent4",
      "accent5",
      "accent6",
      "hlink",
      "folHlink",
    ];
    order.forEach((name, index) => {
      const color = child(scheme, name)?.children[0];
      const value = color?.attributes.lastClr ?? color?.attributes.val;
      if (value && /^[A-Fa-f0-9]{6}$/.test(value)) theme[index] = "#" + value;
    });
  }
  const color = (node: XmlNode | undefined): string | undefined => {
    if (!node) return;
    let value = node.attributes.rgb
      ? "#" + node.attributes.rgb.slice(-6)
      : node.attributes.theme
        ? theme[Number(node.attributes.theme)]
        : undefined;
    if (node.attributes.indexed) {
      const palette = [
        "#000000",
        "#FFFFFF",
        "#FF0000",
        "#00FF00",
        "#0000FF",
        "#FFFF00",
        "#FF00FF",
        "#00FFFF",
      ];
      const index = Number(node.attributes.indexed);
      if (index >= 16 && index < 64)
        report(
          "INDEXED_PALETTE",
          "Extended indexed palette colors are approximated.",
        );
      value = index < 64 ? palette[index % 8] : undefined;
    }
    const tint = Number(node.attributes.tint ?? 0);
    if (value && tint)
      value =
        "#" +
        [1, 3, 5]
          .map((offset) => {
            const n = parseInt(value!.slice(offset, offset + 2), 16);
            return Math.round(
              tint < 0 ? n * (1 + tint) : n * (1 - tint) + 255 * tint,
            )
              .toString(16)
              .padStart(2, "0");
          })
          .join("");
    return value;
  };
  const formats = { ...builtinFormats };
  for (const node of children(child(root, "numFmts"), "numFmt"))
    formats[Number(node.attributes.numFmtId)] = node.attributes.formatCode;
  const fonts = children(child(root, "fonts"), "font"),
    fills = children(child(root, "fills"), "fill"),
    borders = children(child(root, "borders"), "border");
  const enabled = (node: XmlNode | undefined) =>
    !!node && node.attributes.val !== "0" && node.attributes.val !== "false";
  return children(child(root, "cellXfs"), "xf").map((node) => {
    const font = fonts[Number(node.attributes.fontId ?? 0)],
      fill = child(fills[Number(node.attributes.fillId ?? 0)], "patternFill"),
      border = borders[Number(node.attributes.borderId ?? 0)],
      align = child(node, "alignment");
    if (
      child(font, "strike") ||
      child(font, "vertAlign") ||
      child(font, "outline") ||
      child(font, "shadow") ||
      (child(font, "u")?.attributes.val &&
        child(font, "u")?.attributes.val !== "single")
    )
      report(
        "ADVANCED_FONT",
        "Advanced font decorations are normalized to basic font styling.",
      );
    if (
      child(fills[Number(node.attributes.fillId ?? 0)], "gradientFill") ||
      (fill?.attributes.patternType &&
        !["solid", "none", "gray125"].includes(fill.attributes.patternType))
    )
      report("ADVANCED_FILL", "Gradient and patterned fills are not retained.");
    if (
      child(border, "diagonal")?.attributes.style ||
      child(border, "start")?.attributes.style ||
      child(border, "end")?.attributes.style
    )
      report(
        "ADVANCED_BORDER",
        "Diagonal and directional borders are not retained.",
      );
    if (formats[Number(node.attributes.numFmtId ?? 0)] === undefined)
      report(
        "NUMBER_FORMAT",
        "An unrecognized built-in number format is normalized to General.",
      );
    const style: CellStyle = {
      locked:
        child(node, "protection")?.attributes.locked === "0"
          ? false
          : undefined,
      fontFamily: child(font, "name")?.attributes.val,
      fontSize: child(font, "sz")
        ? Number(child(font, "sz")!.attributes.val)
        : undefined,
      bold: enabled(child(font, "b")),
      italic: enabled(child(font, "i")),
      underline: enabled(child(font, "u")),
      color: color(child(font, "color")),
      background:
        fill?.attributes.patternType === "solid"
          ? color(child(fill, "fgColor"))
          : undefined,
      align: ["left", "center", "right"].includes(
        align?.attributes.horizontal ?? "",
      )
        ? (align!.attributes.horizontal as CellStyle["align"])
        : undefined,
      verticalAlign: ["top", "center", "bottom"].includes(
        align?.attributes.vertical ?? "",
      )
        ? (align!.attributes.vertical as CellStyle["verticalAlign"])
        : undefined,
      wrap: align?.attributes.wrapText === "1",
      numberFormat: formats[Number(node.attributes.numFmtId ?? 0)] ?? "General",
    };
    for (const side of ["top", "right", "bottom", "left"] as const) {
      const edge = child(border, side);
      if (edge?.attributes.style) {
        style.border ??= {};
        style.border[side] = color(child(edge, "color")) ?? "#000000";
        if (
          edge.attributes.style !== "thin" &&
          !diagnostics.some((d) => d.code === "BORDER_STYLE")
        )
          diagnostics.push({
            code: "BORDER_STYLE",
            severity: "warning",
            lossy: true,
            message: "Non-thin borders are normalized to thin borders.",
          });
      }
    }
    if (align?.attributes.textRotation || align?.attributes.indent)
      if (!diagnostics.some((d) => d.code === "ADVANCED_ALIGNMENT"))
        diagnostics.push({
          code: "ADVANCED_ALIGNMENT",
          severity: "warning",
          lossy: true,
          message: "Text rotation and indentation are not retained.",
        });
    return style;
  });
}
export function writeStyles(styles: CellStyle[]): string {
  const color = (value: string | undefined, fallback: string) =>
    (value ?? fallback).replace("#", "").padStart(8, "F").toUpperCase();
  const font = (s: CellStyle) =>
    `<font><sz val="${s.fontSize ?? 11}"/><name val="${x(s.fontFamily ?? "Calibri")}"/><color rgb="${color(s.color, "#172B25")}"/>${s.bold ? "<b/>" : ""}${s.italic ? "<i/>" : ""}${s.underline ? "<u/>" : ""}</font>`;
  const border = (s: CellStyle) =>
    "<border>" +
    ["left", "right", "top", "bottom"]
      .map((side) => {
        const c = s.border?.[side as keyof NonNullable<CellStyle["border"]>];
        return c
          ? `<${side} style="thin"><color rgb="${color(c, "#000000")}"/></${side}>`
          : `<${side}/>`;
      })
      .join("") +
    "<diagonal/></border>";
  const formats = styles
    .map(
      (s, i) =>
        `<numFmt numFmtId="${164 + i}" formatCode="${x(s.numberFormat ?? "General")}"/>`,
    )
    .join("");
  const xfs = styles
    .map(
      (s, i) =>
        `<xf numFmtId="${164 + i}" fontId="${i}" fillId="${i + 2}" borderId="${i}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1"><alignment horizontal="${s.align ?? "general"}" vertical="${s.verticalAlign ?? "bottom"}" wrapText="${s.wrap ? 1 : 0}"/>${s.locked === undefined ? "" : `<protection locked="${s.locked ? 1 : 0}"/>`}</xf>`,
    )
    .join("");
  return (
    HEADER +
    `<styleSheet xmlns="${NS}"><numFmts count="${styles.length}">${formats}</numFmts><fonts count="${styles.length}">${styles.map(font).join("")}</fonts><fills count="${styles.length + 2}"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>${styles.map((s) => `<fill><patternFill patternType="${s.background ? "solid" : "none"}">${s.background ? `<fgColor rgb="${color(s.background, "#FFFFFF")}"/><bgColor indexed="64"/>` : ""}</patternFill></fill>`).join("")}</fills><borders count="${styles.length}">${styles.map(border).join("")}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${styles.length}">${xfs}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`
  );
}

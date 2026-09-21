const sheet = "M4 3h16v18H4z M4 9h16 M10 9v12";
const eye =
  "M2 12s3.5-6 10-6 10 6 10 6-3.5 6-10 6-10-6-10-6z M15 12a3 3 0 1 1-6 0 3 3 0 0 1 6 0";
const paths = {
  saveNow: "M4 3h16v18H4z M8 3v6h8V3 M8 21v-8h8v8",
  diagnostics: "M12 3L2 21h20L12 3z M12 9v5 M12 17v1",
  shortcuts:
    "M3 6h18v12H3z M6 9h2 M10 9h2 M14 9h2 M6 12h2 M10 12h2 M14 12h4 M7 15h10",
  formatPainter: "M4 3h16v7H4z M8 10v4h8v4h-4v4",
  pasteSpecial: sheet,
  duplicateSheet: sheet,
  hideSheet: sheet,
  showSheet: sheet,
  nameManager: sheet,
  validation: sheet,
  conditionalFormat: sheet,
  protectSheet: sheet,
  unprotectSheet: sheet,
  unlockCells: sheet,
  lockCells: sheet,

  sort: "M4 5h16 M4 12h12 M4 19h8",
  findNext: "M15 9a5 5 0 1 1-10 0 5 5 0 0 1 10 0 M13 13l7 7 M16 20h4v-4",
  customFormat: "M4 4h16v16H4z M8 9h8 M8 15h8 M10 7l-2 10 M16 7l-2 10",
  borderOptions: "M3 3h18v18H3z M8 12h8 M12 8v8",
  verticalAlign: "M3 4h18 M3 20h18 M7 9h10 M7 15h10",
  autoFitRows: "M3 3h18 M3 21h18 M12 5v14 M8 8l4-4 4 4 M8 16l4 4 4-4",
  autoFitColumns: "M3 3v18 M21 3v18 M5 12h14 M8 8l-4 4 4 4 M16 8l4 4-4 4",
  home: "M3 10l9-7 9 7 M5 9v12h5v-7h4v7h5V9",
  insert: "M4 4h16v16H4z M12 8v8 M8 12h8",
  data: "M4 3h16v18H4z M4 9h16 M4 15h16 M10 3v18",
  view: eye,
  open: "M4 14v7h16v-7 M12 3v12 M7 10l5 5 5-5",
  saveXlsx: "M4 14v7h16v-7 M12 15V3 M7 8l5-5 5 5",
  undo: "M8 4L3 9l5 5 M3 9h11a6 6 0 0 1 0 12",
  redo: "M16 4l5 5-5 5 M21 9H10a6 6 0 0 0 0 12",
  copy: "M9 9h12v12H9z M15 9V3H3v12h6",
  font: "M3 20L10 4l7 16 M6 14h8 M17 10h5 M19.5 10v10",
  fontSize: "M3 5h12 M9 5v15 M5 20h8 M18 7l3-3 3 3 M21 4v16 M18 17l3 3 3-3",
  bold: "M7 4h6a4 4 0 0 1 0 8H7z M7 12h7a4 4 0 0 1 0 8H7z",
  italic: "M11 4h9 M4 20h9 M15 4L9 20",
  underline: "M6 3v8a6 6 0 0 0 12 0V3 M4 21h16",
  color: "M6 17L12 3l6 14 M8 12h8 M3 21h18",
  fill: "M4 10l8-8 9 9-8 8-9-9z M4 10h16 M7 2l6 6 M20 16s-2 2-2 3a2 2 0 0 0 4 0c0-1-2-3-2-3",
  left: "M4 5h16 M4 10h10 M4 15h16 M4 20h10",
  center: "M4 5h16 M7 10h10 M4 15h16 M7 20h10",
  right: "M4 5h16 M10 10h10 M4 15h16 M10 20h10",
  wrap: "M3 5h18 M3 10h13a4 4 0 0 1 0 8h-5 M14 15l-3 3 3 3 M3 16h3",
  border: "M3 3h18v18H3z M3 9h18 M3 15h18 M9 3v18 M15 3v18",
  format: "M9 3L7 21 M17 3l-2 18 M3 9h18 M2 15h18",
  merge:
    "M9 3H3v18h6 M15 3h6v18h-6 M2 12h7 M6 9l3 3-3 3 M22 12h-7 M18 9l-3 3 3 3",
  unmerge:
    "M3 3h18v18H3z M12 3v18 M10 12H5 M8 9l-3 3 3 3 M14 12h5 M16 9l3 3-3 3",
  find: "M16 10a6 6 0 1 1-12 0 6 6 0 0 1 12 0 M15 15l6 6",
  addRow: sheet + " M6 5h6 M9 2v6",
  addColumn: sheet + " M15 12v6 M12 15h6",
  deleteRow: sheet + " M6 6h12",
  deleteColumn: sheet + " M15 12v6",
  addSheet: "M4 3h12l4 4v14H4z M16 3v5h4 M12 11v7 M8.5 14.5h7",
  rename: "M4 4h9 M4 4v17h16v-8 M10 14l1-4 8-8 3 3-8 8-4 1z",
  deleteSheet: "M4 3h12l4 4v14H4z M16 3v5h4 M9 12l6 6 M15 12l-6 6",
  namedRange: "M3 3h18v18H3z M7 16V8l10 8V8",
  sortAsc: "M5 4v16 M2 17l3 3 3-3 M12 5h3 M12 10h5 M12 15h7 M12 20h9",
  sortDesc: "M5 4v16 M2 17l3 3 3-3 M12 5h9 M12 10h7 M12 15h5 M12 20h3",
  filter: "M3 4h18l-7 8v7l-4 2v-9L3 4z",
  clearFilter: "M3 4h18l-7 8v7l-4 2v-9L3 4z M17 16l5 5 M22 16l-5 5",
  clear: "M3 15l10-12 9 8-8 10H8l-5-6z M8 9l9 9 M13 21h9",
  saveCsv: "M4 3h10l6 6v12H4z M14 3v6h6 M8 13h8 M8 17h8 M11 13v4",
  saveJson: "M8 3H6v6l-3 3 3 3v6h2 M16 3h2v6l3 3-3 3v6h-2",
  freeze: "M3 3h18v18H3z M3 9h18 M9 3v18",
  unfreeze: "M3 9h13 M9 3v13 M3 3v18h18V3H3 M3 3l18 18",
  rowHeight: "M3 4h18 M3 20h18 M12 5v14 M8 9l4-4 4 4 M8 15l4 4 4-4",
  columnWidth: "M4 3v18 M20 3v18 M5 12h14 M9 8l-4 4 4 4 M15 8l4 4-4 4",
  hideRows: "M3 4h18 M3 20h18 M7 8l5 4 5-4 M7 16l5-4 5 4",
  hideColumns: "M4 3v18 M20 3v18 M8 7l4 5-4 5 M16 7l-4 5 4 5",
  showRows: "M3 4h18 M3 20h18 M7 10l5-5 5 5 M7 14l5 5 5-5",
  showColumns: "M4 3v18 M20 3v18 M10 7l-5 5 5 5 M14 7l5 5-5 5",
  benchmark: "M3 20h18 M5 16v-5 M12 16V4 M19 16V8",
  code: "M8 5l-6 7 6 7 M16 5l6 7-6 7 M14 3l-4 18",
} as const;
export type IconName = keyof typeof paths;
export function icon(name: IconName): SVGSVGElement {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("width", "16");
  svg.setAttribute("height", "16");
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "1.7");
  svg.setAttribute("stroke-linecap", "round");
  svg.setAttribute("stroke-linejoin", "round");
  svg.setAttribute("aria-hidden", "true");
  svg.setAttribute("focusable", "false");
  svg.classList.add("tool-icon");
  const path = document.createElementNS(svg.namespaceURI, "path");
  path.setAttribute("d", paths[name]);
  svg.append(path);
  return svg;
}

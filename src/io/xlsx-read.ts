import { preserveObjects } from "./xlsx-objects";
import { readRules } from "./xlsx-rules";
import type { XmlNode } from "./xml";
import {
  keyOf,
  MAX_COLUMNS,
  parseCell,
  parseRange,
  rowOf,
  columnOf,
  contains,
} from "../core/address";
import type {
  Cell,
  Diagnostic,
  FilterRule,
  SheetSnapshot,
  WorkbookSnapshot,
} from "../core/types";
import { shiftFormula, parseFormula } from "../formula/parser";
import {
  readArchive,
  xmlTree,
  xmlParser,
  parseXmlChunks,
  children,
  child,
  type XmlAttributes,
} from "./xml";
import { readStyles } from "./xlsx-styles";
export interface XlsxProgress {
  stage: string;
  progress: number;
}
export interface XlsxReadOptions {
  signal?: AbortSignal;
  onProgress?: (progress: XlsxProgress) => void;
  checkpoint?: () => Promise<void>;
}
export async function readXlsx(
  input: Uint8Array | ArrayBuffer,
  options: XlsxReadOptions = {},
): Promise<{ snapshot: WorkbookSnapshot; diagnostics: Diagnostic[] }> {
  const checkpoint = async () => {
    options.signal?.throwIfAborted();
    if (options.checkpoint) await options.checkpoint();
    else await new Promise((resolve) => setTimeout(resolve, 0));
    options.signal?.throwIfAborted();
  };
  const files = await readArchive(
    input instanceof Uint8Array ? input : new Uint8Array(input),
    async (progress) => {
      options.onProgress?.({ stage: "unzip", progress: progress * 0.25 });
      await checkpoint();
    },
  );
  const diagnostics: Diagnostic[] = [];
  const retainedObjects = new Set<string>();
  const report = (code: string, message: string, sheetId?: string) => {
    if (!diagnostics.some((d) => d.code === code && d.sheetId === sheetId))
      diagnostics.push({
        code,
        severity: "warning",
        lossy: true,
        message,
        sheetId,
      });
  };
  if ([...files.keys()].some((n) => /vbaProject\.bin$/i.test(n)))
    throw new Error("Macro-enabled workbooks are not supported");
  const resolve = (base: string, target: string) => {
    const parts = (
        target.startsWith("/")
          ? target.slice(1)
          : base.slice(0, base.lastIndexOf("/") + 1) + target
      ).split("/"),
      normalized: string[] = [];
    for (const part of parts) {
      if (part === "..") normalized.pop();
      else if (part !== "." && part) normalized.push(part);
    }
    return normalized.join("/");
  };
  const rootRelations = xmlTree(files.get("_rels/.rels"));
  const workbookRel = children(rootRelations, "Relationship").find((n) =>
    n.attributes.Type.endsWith("/officeDocument"),
  );
  if (!workbookRel) throw new Error("No workbook part found");
  const workbookPath = resolve("", workbookRel.attributes.Target),
    workbook = xmlTree(files.get(workbookPath));
  const relPath = workbookPath.replace(/([^/]+)$/, "_rels/$1.rels"),
    relationships = new Map<string, { path: string; type: string }>();
  for (const rel of children(xmlTree(files.get(relPath)), "Relationship")) {
    if (rel.attributes.TargetMode === "External") {
      report(
        "EXTERNAL_RELATIONSHIP",
        "External relationships are not fetched or retained.",
      );
      continue;
    }
    if (
      !["worksheet", "styles", "sharedStrings", "theme", "sheetMetadata"].some(
        (type) => rel.attributes.Type.endsWith("/" + type),
      )
    )
      report(
        "UNSUPPORTED_RELATIONSHIP",
        "Workbook relationship is not retained: " + rel.attributes.Type,
      );
    relationships.set(rel.attributes.Id, {
      path: resolve(workbookPath, rel.attributes.Target),
      type: rel.attributes.Type,
    });
  }
  for (const path of files.keys()) {
    if (
      /^(customXml|docProps)\//.test(path) ||
      /^xl\/(embeddings|media|slicers|connections|queryTables|printerSettings)/.test(
        path,
      )
    )
      report(
        "UNSUPPORTED_PACKAGE_CONTENT",
        "Package metadata or object is not retained: " + path,
      );
    if (
      /^xl\/(charts|drawings|pivotTables|pivotCache|tables|externalLinks|comments|threadedComments)\//.test(
        path,
      ) ||
      /comments\d*\.xml$/.test(path)
    )
      report(
        "UNSUPPORTED_PART_" + path.split("/")[1].toUpperCase(),
        `Unsupported workbook content: ${path}`,
      );
  }
  const stylePath =
    [...relationships.values()].find((r) => r.type.endsWith("/styles"))?.path ??
    "xl/styles.xml";
  const styles = readStyles(files, stylePath, diagnostics),
    sharedPath = [...relationships.values()].find((r) =>
      r.type.endsWith("/sharedStrings"),
    )?.path;
  const strings: string[] = [];
  if (sharedPath && files.has(sharedPath)) {
    let insideText = false,
      value = "",
      rich = false;
    const parser = xmlParser({
      open: (name) => {
        if (name === "si") {
          value = "";
          rich = false;
        }
        if (name === "t") insideText = true;
        if (name === "rPr") rich = true;
      },
      text: (text) => {
        if (insideText) value += text;
      },
      close: (name) => {
        if (name === "t") insideText = false;
        if (name === "si") {
          strings.push(value);
          if (rich)
            report("RICH_TEXT", "Rich text runs are flattened to cell text.");
        }
      },
    });
    await parseXmlChunks(files.get(sharedPath)!, parser, checkpoint);
    files.delete(sharedPath);
  }
  const sheets: SheetSnapshot[] = [],
    nodes = children(child(workbook, "sheets"), "sheet");
  for (let index = 0; index < nodes.length; index++) {
    const info = nodes[index],
      id = `sheet-${index + 1}`,
      relation = relationships.get(info.attributes["r:id"]);
    if (
      !relation ||
      !files.has(relation.path) ||
      !relation.type.endsWith("/worksheet")
    )
      throw new Error(
        `Missing or unsupported worksheet: ${info.attributes.name}`,
      );
    const sheet: SheetSnapshot = {
      id,
      name: info.attributes.name,
      rowCount: 100000,
      columnCount: 100,
      rowHeights: {},
      columnWidths: {},
      hiddenRows: [],
      hiddenColumns: [],
      frozenRows: 0,
      frozenColumns: 0,
      merges: [],
      cells: [],
    };
    sheet.hidden =
      !!info.attributes.state && info.attributes.state !== "visible";
    const ruleNodes: XmlNode[] = [],
      ruleStack: XmlNode[] = [];
    let cellAttrs: XmlAttributes = {},
      formulaAttrs: XmlAttributes = {},
      value = "",
      formula = "",
      inline = "",
      field = "",
      cellAddress = "",
      row = 0,
      inferredColumn = 0;
    const shared = new Map<
        string,
        { row: number; column: number; formula: string }
      >(),
      pending: { cell: Cell; row: number; column: number; id: string }[] = [],
      arrays: { range: ReturnType<typeof parseRange>; anchor: number }[] = [];
    let currentFilterColumn = 0,
      filterGroupStart = 0,
      filterGroupAnd = false;
    const parser = xmlParser({
      open: (name, attrs) => {
        if (
          ruleStack.length ||
          [
            "dataValidations",
            "conditionalFormatting",
            "sheetProtection",
            "drawing",
            "legacyDrawing",
            "tableParts",
          ].includes(name)
        ) {
          const node: XmlNode = {
            name,
            attributes: attrs,
            text: "",
            children: [],
          };
          if (ruleStack.length) ruleStack.at(-1)!.children.push(node);
          else ruleNodes.push(node);
          ruleStack.push(node);
        }

        if (name === "dimension" && attrs.ref) {
          const range = parseRange(attrs.ref);
          sheet.rowCount = Math.max(sheet.rowCount, range.r2 + 1);
          sheet.columnCount = Math.max(sheet.columnCount, range.c2 + 1);
        }
        if (name === "row") {
          row = Number(attrs.r ?? row + 1) - 1;
          inferredColumn = 0;
          sheet.rowCount = Math.max(sheet.rowCount, row + 1);
          if (attrs.ht) sheet.rowHeights[row] = (Number(attrs.ht) * 4) / 3;
          if (attrs.hidden === "1") sheet.hiddenRows.push(row);
        }
        if (name === "col") {
          const first = Number(attrs.min) - 1,
            last = Number(attrs.max) - 1;
          if (first < 0 || last >= MAX_COLUMNS)
            throw new Error("Column outside Excel bounds");
          for (let col = first; col <= last; col++) {
            if (attrs.width)
              sheet.columnWidths[col] = Math.round(Number(attrs.width) * 7 + 5);
            if (attrs.hidden === "1") sheet.hiddenColumns.push(col);
          }
          sheet.columnCount = Math.max(sheet.columnCount, last + 1);
        }
        if (name === "pane") {
          if (attrs.state === "frozen" || attrs.state === "frozenSplit") {
            sheet.frozenRows = Number(attrs.ySplit ?? 0);
            sheet.frozenColumns = Number(attrs.xSplit ?? 0);
          } else
            report(
              "SPLIT_PANES",
              "Non-frozen split panes are not retained.",
              id,
            );
        }
        if (name === "mergeCell") {
          const range = parseRange(attrs.ref);
          sheet.merges.push(range);
          sheet.rowCount = Math.max(sheet.rowCount, range.r2 + 1);
          sheet.columnCount = Math.max(sheet.columnCount, range.c2 + 1);
        }
        if (name === "c") {
          cellAttrs = attrs;
          cellAddress = attrs.r ?? "";
          value = "";
          formula = "";
          inline = "";
          formulaAttrs = {};
        }
        if (name === "f") {
          field = "formula";
          formulaAttrs = attrs;
        }
        if (name === "v") field = "value";
        if (name === "t") field = "inline";
        if (name === "autoFilter")
          sheet.filter = { range: parseRange(attrs.ref), rules: [] };
        if (name === "filterColumn")
          currentFilterColumn =
            Number(attrs.colId) + (sheet.filter?.range.c1 ?? 0);
        if (name === "filters" && sheet.filter)
          sheet.filter.rules.push({
            column: currentFilterColumn,
            operator: "in",
            values: attrs.blank === "1" || attrs.blank === "true" ? [""] : [],
          });
        if (name === "dateGroupItem" && sheet.filter) {
          sheet.filter.rules = sheet.filter.rules.filter(
            (rule) =>
              rule.column !== currentFilterColumn || rule.operator !== "in",
          );
        }
        if (name === "filter" && sheet.filter) {
          const list = sheet.filter.rules.at(-1);
          if (list?.operator === "in") list.values.push(attrs.val ?? "");
        }
        if (name === "customFilters") {
          filterGroupStart = sheet.filter?.rules.length ?? 0;
          filterGroupAnd = attrs.and === "1" || attrs.and === "true";
        }
        if (name === "customFilter" && sheet.filter) {
          const operators: Record<
            string,
            Exclude<FilterRule["operator"], "in">
          > = {
            equal: "eq",
            notEqual: "neq",
            greaterThan: "gt",
            lessThan: "lt",
            greaterThanOrEqual: "gte",
            lessThanOrEqual: "lte",
          };
          const input = attrs.val,
            numeric = input !== "" && Number.isFinite(Number(input));
          let operator = operators[attrs.operator ?? "equal"] ?? "eq";
          let parsed: string | number = numeric ? Number(input) : input;
          if (!numeric && /[~*?]/.test(input)) {
            if (operator === "eq" && /^\*(?:~.|[^*?~])*\*$/.test(input)) {
              operator = "contains";
              parsed = input.slice(1, -1).replace(/~(.)/g, "$1");
            } else if (/^(?:~.|[^*?~])*$/.test(input))
              parsed = input.replace(/~(.)/g, "$1");
            else {
              report(
                "FILTER_WILDCARD",
                "This wildcard filter is not retained.",
                id,
              );
              return;
            }
          }
          sheet.filter.rules.push({
            column: currentFilterColumn,
            operator,
            value: parsed,
          });
        }
        if (
          [
            "dateGroupItem",
            "colorFilter",
            "dynamicFilter",
            "top10",
            "iconFilter",
          ].includes(name)
        )
          report(
            "FILTER_VARIANT",
            "This filter variant is not retained; custom comparison filters are supported.",
            id,
          );
        if (
          name === "sheetFormatPr" &&
          (attrs.defaultColWidth ||
            (attrs.defaultRowHeight && Number(attrs.defaultRowHeight) !== 21))
        )
          report(
            "DEFAULT_DIMENSIONS",
            "Default row/column dimensions are normalized; explicit dimensions are retained.",
            id,
          );
        if (
          [
            "hyperlinks",
            "pageSetup",
            "printOptions",
            "headerFooter",
            "sheetPr",
            "pageMargins",
            "rowBreaks",
            "colBreaks",
            "extLst",
            "drawing",
            "legacyDrawing",
            "oleObjects",
            "controls",
          ].includes(name)
        )
          report(
            name.toUpperCase(),
            `Worksheet feature ${name} is not retained.`,
            id,
          );
        if (name === "rPr")
          report("RICH_TEXT", "Rich text runs are flattened to cell text.", id);
      },
      text: (content) => {
        if (ruleStack.length) ruleStack.at(-1)!.text += content;
        if (field === "formula") formula += content;
        else if (field === "value") value += content;
        else if (field === "inline") inline += content;
      },
      close: (name) => {
        if (ruleStack.length) ruleStack.pop();
        if (
          name === "customFilters" &&
          sheet.filter &&
          !filterGroupAnd &&
          sheet.filter.rules.length - filterGroupStart > 1
        ) {
          sheet.filter.rules.splice(filterGroupStart);
          report(
            "FILTER_OR",
            "OR filters are not retained; comparison rules use AND.",
            id,
          );
        }
        if (["f", "v", "t"].includes(name)) field = "";
        if (name !== "c") return;
        const position = cellAddress
          ? parseCell(cellAddress)
          : { row, column: inferredColumn };
        inferredColumn = position.column + 1;
        sheet.rowCount = Math.max(sheet.rowCount, position.row + 1);
        sheet.columnCount = Math.max(sheet.columnCount, position.column + 1);
        const cell: Cell = { style: Number(cellAttrs.s ?? 0) };
        if (cell.style! >= styles.length)
          throw new Error("Invalid cell style index");
        if (formula || formulaAttrs.t) {
          if (formula) cell.formula = "=" + formula;
          if (formulaAttrs.t === "shared") {
            if (formula)
              shared.set(formulaAttrs.si, {
                ...position,
                formula: "=" + formula,
              });
            else pending.push({ cell, ...position, id: formulaAttrs.si });
          }
          if (formulaAttrs.t === "array" && formulaAttrs.ref) {
            if (!cellAttrs.cm)
              report(
                "LEGACY_ARRAY",
                "Legacy fixed array formulas are recalculated as dynamic arrays.",
                id,
              );
            arrays.push({
              range: parseRange(formulaAttrs.ref),
              anchor: keyOf(position.row, position.column),
            });
          }
        } else if (cellAttrs.t === "s") {
          const i = Number(value);
          if (!Number.isInteger(i) || i < 0 || i >= strings.length)
            throw new Error("Invalid shared string index");
          cell.value = strings[i];
        } else if (cellAttrs.t === "inlineStr") cell.value = inline;
        else if (cellAttrs.t === "str") cell.value = value;
        else if (cellAttrs.t === "b") cell.value = value === "1";
        else if (cellAttrs.t === "e")
          cell.value = { error: value as "#VALUE!" };
        else if (cellAttrs.t === "d") {
          cell.value = value;
          report("ISO_DATE_CELL", "ISO date cells are imported as text.", id);
        } else if (value !== "") {
          const n = Number(value);
          if (!Number.isFinite(n)) throw new Error("Invalid numeric cell");
          cell.value = n;
        }
        if (
          cell.value !== undefined ||
          cell.formula ||
          cell.style ||
          formulaAttrs.t
        )
          sheet.cells.push([keyOf(position.row, position.column), cell]);
      },
    });
    await parseXmlChunks(files.get(relation.path)!, parser, checkpoint);
    try {
      const preserved = preserveObjects(files, relation.path, ruleNodes);
      if (preserved.objects) sheet.objects = preserved.objects;
      for (const path of preserved.retained) retainedObjects.add(path);
    } catch {
      report(
        "OBJECT_PRESERVATION",
        "Some related objects could not be preserved.",
        id,
      );
    }
    if (sheet.objects)
      for (let i = diagnostics.length - 1; i >= 0; i--)
        if (
          diagnostics[i].sheetId === id &&
          /drawing/i.test(diagnostics[i].code)
        )
          diagnostics.splice(i, 1);
    readRules(
      ruleNodes,
      sheet,
      files.get(stylePath) ?? new TextEncoder().encode("<styleSheet/>"),
      diagnostics,
    );
    files.delete(relation.path);
    for (const item of pending) {
      const anchor = shared.get(item.id);
      if (!anchor) throw new Error("Missing shared formula anchor");
      item.cell.formula = shiftFormula(
        anchor.formula,
        item.row - anchor.row,
        item.column - anchor.column,
      );
    }
    if (arrays.length) {
      const byRow = new Map<number, typeof arrays>();
      for (const spill of arrays) {
        for (
          let bucket = Math.floor(spill.range.r1 / 256);
          bucket <= Math.floor(spill.range.r2 / 256);
          bucket++
        ) {
          const group = byRow.get(bucket) ?? [];
          group.push(spill);
          byRow.set(bucket, group);
        }
      }
      for (const [key, cell] of sheet.cells) {
        for (const spill of byRow.get(Math.floor(rowOf(key) / 256)) ?? []) {
          if (
            key !== spill.anchor &&
            contains(spill.range, rowOf(key), columnOf(key))
          ) {
            delete cell.value;
            delete cell.formula;
          }
        }
      }
    }
    if (sheet.filter?.rules.length) {
      const { r1, r2 } = sheet.filter.range;
      const outside = sheet.hiddenRows.filter((r) => r <= r1 || r > r2);
      if (outside.length !== sheet.hiddenRows.length) {
        sheet.hiddenRows = outside;
        report(
          "FILTER_HIDDEN_ROWS",
          "Hidden rows inside an active filter are recalculated; manually hidden rows within that range may need to be hidden again.",
          id,
        );
      }
    }
    sheets.push(sheet);
    options.onProgress?.({
      stage: "worksheets",
      progress: 0.25 + ((index + 1) / nodes.length) * 0.75,
    });
    await checkpoint();
  }
  for (let i = diagnostics.length - 1; i >= 0; i--) {
    const d = diagnostics[i];
    if ([...retainedObjects].some((path) => d.message.endsWith(": " + path)))
      diagnostics.splice(i, 1);
  }
  const names: WorkbookSnapshot["names"] = {};
  for (const name of children(child(workbook, "definedNames"), "definedName")) {
    if (name.attributes.name.startsWith("_xlnm.")) {
      report("PRINT_NAMES", "Print areas and print titles are not retained.");
      continue;
    }
    try {
      const ast = parseFormula("=" + name.text),
        start =
          ast.type === "range"
            ? ast.start
            : ast.type === "ref"
              ? ast
              : undefined;
      if (!start?.sheet || name.attributes.localSheetId !== undefined)
        throw new Error();
      const sheet = sheets.find(
        (s) => s.name.toUpperCase() === start.sheet!.toUpperCase(),
      );
      if (!sheet) throw new Error();
      const end = ast.type === "range" ? ast.end : start;
      names[name.attributes.name.toUpperCase()] = {
        sheetId: sheet.id,
        range: { r1: start.row, c1: start.column, r2: end.row, c2: end.column },
      };
    } catch {
      report(
        "DEFINED_NAME",
        `Unsupported defined name: ${name.attributes.name}`,
      );
    }
  }
  if (!sheets.length) throw new Error("Workbook has no worksheets");
  return {
    snapshot: {
      version: 1,
      sheets,
      styles: styles.length ? styles : [{}],
      names,
      dateSystem: ["1", "true"].includes(
        child(workbook, "workbookPr")?.attributes.date1904 ?? "0",
      )
        ? 1904
        : 1900,
      diagnostics,
    },
    diagnostics,
  };
}

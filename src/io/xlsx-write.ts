import { exportObjects } from "./xlsx-objects";
import { writeValidations, writeConditional } from "./xlsx-rules";
import { writeDxfs } from "./xlsx-differential";
import { Zip, ZipDeflate, strToU8 } from "fflate";
import { WorkbookModel, type SheetState } from "../core/model";
import type { Diagnostic, WorkbookSnapshot, Scalar } from "../core/types";
import { address, columnOf, keyOf, rangeName, rowOf } from "../core/address";
import { isArray, isError } from "../formula/values";
import { functions } from "../formula/functions";
import { parseFormula, printFormula, type AST } from "../formula/parser";
import { xmlEscape as x } from "./xml";
import { HEADER, NS, REL, writeStyles } from "./xlsx-styles";
import type { XlsxReadOptions } from "./xlsx-read";
export interface XlsxWriteOptions extends XlsxReadOptions {
  allowLossy?: boolean;
}
export class CompatibilityError extends Error {
  constructor(public diagnostics: Diagnostic[]) {
    super(
      "Export would discard unsupported content. Review diagnostics and pass allowLossy: true to proceed.",
    );
    this.name = "CompatibilityError";
  }
}
const futureFunctions = new Set([
  "XLOOKUP",
  "XMATCH",
  "IFS",
  "SWITCH",
  "CONCAT",
  "TEXTJOIN",
  "UNICHAR",
  "UNICODE",
  "NUMBERVALUE",
  "ISOWEEKNUM",
  "SORT",
  "SORTBY",
  "UNIQUE",
  "SEQUENCE",
  "FILTER",
  "TAKE",
  "DROP",
  "CHOOSECOLS",
  "CHOOSEROWS",
  "VSTACK",
  "HSTACK",
  "STDEV.S",
  "STDEV.P",
  "VAR.S",
  "VAR.P",
  "RANK.EQ",
  "MODE.SNGL",
  "PERCENTILE.INC",
  "QUARTILE.INC",
  "COVARIANCE.P",
  "COVARIANCE.S",
  "CEILING.MATH",
  "FLOOR.MATH",
]);
function exportFormula(formula: string): string {
  try {
    function visit(ast: AST): AST {
      if (ast.type === "call") {
        if (!functions.has(ast.name))
          throw new Error("Keep unknown formulas verbatim");
        return {
          ...ast,
          originalName: undefined,
          name: ast.name.startsWith("_")
            ? ast.name
            : futureFunctions.has(ast.name)
              ? "_xlfn." +
                (["FILTER", "SORT"].includes(ast.name) ? "_xlws." : "") +
                ast.name
              : ast.name,
          args: ast.args.map(visit),
        };
      }
      if (ast.type === "unary") {
        if (ast.op === "#")
          return {
            type: "call",
            name: "_xlfn.ANCHORARRAY",
            args: [visit(ast.value)],
          };
        return { ...ast, value: visit(ast.value) };
      }
      if (ast.type === "binary")
        return { ...ast, left: visit(ast.left), right: visit(ast.right) };
      if (ast.type === "array")
        return { ...ast, rows: ast.rows.map((row) => row.map(visit)) };
      return ast;
    }
    return printFormula(visit(parseFormula(formula))).slice(1);
  } catch {
    return formula.slice(1);
  }
}
function cellXml(
  ref: string,
  value: Scalar,
  style: number,
  formula?: string,
  spill?: string,
  dynamic = false,
): string {
  const type = isError(value)
    ? "e"
    : typeof value === "boolean"
      ? "b"
      : typeof value === "string"
        ? formula
          ? "str"
          : "inlineStr"
        : "n";
  const formulaXml = formula
    ? `<f${spill ? ` t="array" ref="${spill}" aca="1"` : ""}>${x(exportFormula(formula))}</f>`
    : "";
  const content =
    value === null
      ? ""
      : type === "inlineStr"
        ? `<is><t xml:space="preserve">${x(value)}</t></is>`
        : `<v>${x(isError(value) ? (value.error === "#CYCLE!" ? "#VALUE!" : value.error) : typeof value === "boolean" ? +value : value)}</v>`;
  return `<c r="${ref}" s="${style}" t="${type}"${dynamic ? ' cm="1"' : ""}>${formulaXml}${content}</c>`;
}
function* sheetXml(model: WorkbookModel, sheet: SheetState): Generator<string> {
  const meta = sheet.meta,
    hiddenRows = new Set([...meta.hiddenRows, ...(meta.filteredRows ?? [])]),
    hiddenColumns = new Set(meta.hiddenColumns);
  let maxRow = 0,
    maxColumn = 0;
  // Row XML is generated below; only coordinates of spilled values need an auxiliary index.
  const spillValues = new Map<number, { value: Scalar; style: number }>();
  for (const [key, cell] of sheet.cells) {
    maxRow = Math.max(maxRow, rowOf(key));
    maxColumn = Math.max(maxColumn, columnOf(key));
    if (cell.formula) {
      const result = model.engine.arrayResult(meta.id, key);
      if (isArray(result) && result.rows * result.columns > 1)
        for (let r = 0; r < result.rows; r++)
          for (let c = 0; c < result.columns; c++)
            if (r || c) {
              const target = keyOf(rowOf(key) + r, columnOf(key) + c);
              spillValues.set(target, {
                value: result.get(r, c),
                style: sheet.cells.get(target)?.style ?? cell.style ?? 0,
              });
              maxRow = Math.max(maxRow, rowOf(target));
              maxColumn = Math.max(maxColumn, columnOf(target));
            }
    }
  }
  const pane =
    meta.frozenRows || meta.frozenColumns
      ? `<pane xSplit="${meta.frozenColumns}" ySplit="${meta.frozenRows}" topLeftCell="${address(meta.frozenRows, meta.frozenColumns)}" activePane="${meta.frozenRows && meta.frozenColumns ? "bottomRight" : meta.frozenRows ? "bottomLeft" : "topRight"}" state="frozen"/>`
      : "";
  yield HEADER +
    `<worksheet xmlns="${NS}" xmlns:r="${REL}"><dimension ref="A1:${address(maxRow, maxColumn)}"/><sheetViews><sheetView workbookViewId="0">${pane}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="21"/><cols>`;
  for (let c = 0; c < meta.columnCount; c++)
    if (meta.columnWidths[c] !== undefined || hiddenColumns.has(c))
      yield `<col min="${c + 1}" max="${c + 1}" width="${((meta.columnWidths[c] ?? 112) - 5) / 7}" customWidth="1"${hiddenColumns.has(c) ? ' hidden="1"' : ""}/>`;
  yield "</cols><sheetData>";
  const keys = [
    ...new Set([...sheet.cells.keys(), ...spillValues.keys()]),
  ].sort((a, b) => a - b);
  const rowKeys = [
    ...new Set([
      ...keys.map(rowOf),
      ...Object.keys(meta.rowHeights).map(Number),
      ...hiddenRows,
    ]),
  ].sort((a, b) => a - b);
  let offset = 0;
  for (const r of rowKeys) {
    let xml = `<row r="${r + 1}"${meta.rowHeights[r] !== undefined ? ` ht="${(meta.rowHeights[r] * 3) / 4}" customHeight="1"` : ""}${hiddenRows.has(r) ? ' hidden="1"' : ""}>`;
    while (offset < keys.length && rowOf(keys[offset]) === r) {
      const key = keys[offset++],
        c = columnOf(key),
        cell = sheet.cells.get(key),
        spill = spillValues.get(key);
      const value = spill?.value ?? model.engine.get(meta.id, key);
      const result = cell?.formula
        ? model.engine.arrayResult(meta.id, key)
        : undefined;
      const dynamic = isArray(result) && result.rows * result.columns > 1;
      const spillRef = dynamic
        ? `${address(r, c)}:${address(r + result.rows - 1, c + result.columns - 1)}`
        : undefined;
      xml += cellXml(
        address(r, c),
        value,
        cell?.style ?? spill?.style ?? 0,
        cell?.formula,
        spillRef,
        dynamic,
      );
    }
    yield xml + "</row>";
  }
  yield "</sheetData>";
  if (meta.protected)
    yield `<sheetProtection sheet="1" objects="1" scenarios="1"/>`;
  if (meta.filter) {
    const operators: Record<string, string> = {
      eq: "equal",
      neq: "notEqual",
      gt: "greaterThan",
      lt: "lessThan",
      gte: "greaterThanOrEqual",
      lte: "lessThanOrEqual",
      contains: "equal",
    };
    const columns = new Map<number, typeof meta.filter.rules>();
    for (const rule of meta.filter.rules)
      columns.set(rule.column, [...(columns.get(rule.column) ?? []), rule]);
    yield `<autoFilter ref="${rangeName(meta.filter.range)}">`;
    for (const [column, rules] of columns) {
      const list = rules.find((rule) => rule.operator === "in");
      if (list?.operator === "in") {
        yield `<filterColumn colId="${column - meta.filter.range.c1}"><filters blank="${list.values.includes("") ? 1 : 0}">${list.values
          .filter((v) => v !== "")
          .map((v) => `<filter val="${x(v)}"/>`)
          .join("")}</filters></filterColumn>`;
        continue;
      }
      yield `<filterColumn colId="${column - meta.filter.range.c1}"><customFilters and="1">${rules
        .filter((rule) => rule.operator !== "in")
        .map(
          (rule) =>
            `<customFilter operator="${operators[rule.operator]}" val="${x(rule.operator === "contains" ? "*" + String(rule.value ?? "").replace(/([*?~])/g, "~$1") + "*" : typeof rule.value === "string" ? rule.value.replace(/([*?~])/g, "~$1") : rule.value)}"/>`,
        )
        .join("")}</customFilters></filterColumn>`;
    }
    yield "</autoFilter>";
  }
  if (meta.merges.length)
    yield `<mergeCells count="${meta.merges.length}">${meta.merges.map((m) => `<mergeCell ref="${rangeName(m)}"/>`).join("")}</mergeCells>`;
  const ruleOffset = model.sheets
    .slice(0, model.sheets.indexOf(sheet))
    .reduce((n, s) => n + (s.meta.conditionalFormats?.length ?? 0), 0);
  yield writeConditional(meta.conditionalFormats ?? [], ruleOffset);
  yield writeValidations(meta);
  if (meta.objects) yield meta.objects.elements.join("");
  yield "</worksheet>";
}
const metadataXml =
  HEADER +
  `<metadata xmlns="${NS}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{BDBB8CDC-FA1E-496E-A857-3C3F30C029C3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`;
export async function writeModelXlsx(
  model: WorkbookModel,
  options: XlsxWriteOptions = {},
): Promise<{ data: Uint8Array; diagnostics: Diagnostic[] }> {
  const diagnostics = model.getDiagnostics();
  if (diagnostics.some((d) => d.lossy) && !options.allowLossy)
    throw new CompatibilityError(diagnostics);
  const checkpoint = async () => {
    options.signal?.throwIfAborted();
    if (options.checkpoint) await options.checkpoint();
    else await new Promise((resolve) => setTimeout(resolve, 0));
    options.signal?.throwIfAborted();
  };
  const chunks: Uint8Array[] = [];
  let failure: Error | undefined;
  const zip = new Zip((err, data) => {
    if (err) failure = err;
    else chunks.push(data);
  });
  const add = async (path: string, content: Iterable<string>) => {
    const file = new ZipDeflate(path, { level: 3 });
    zip.add(file);
    let buffer = "",
      count = 0;
    for (const piece of content) {
      buffer += piece;
      if (buffer.length >= 65536) {
        file.push(strToU8(buffer));
        buffer = "";
        if (++count % 8 === 0) await checkpoint();
      }
    }
    file.push(strToU8(buffer), true);
    if (failure) throw failure;
    await checkpoint();
  };
  const sheets = model.sheets;
  const preserved = sheets.map((sheet, i) =>
    sheet.meta.objects ? exportObjects(sheet.meta.objects, i + 1) : undefined,
  );
  const objectTypes = [
    ...new Set(preserved.flatMap((item) => item?.types ?? [])),
  ]
    .filter((type) => !/<Default[^>]*Extension="(?:xml|rels)"/.test(type))
    .join("");
  await add("[Content_Types].xml", [
    HEADER +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}${objectTypes}</Types>`,
  ]);
  await add("_rels/.rels", [
    HEADER +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  ]);
  const defined = Object.entries(model.names)
    .map(([name, value]) => {
      const sheet = model.sheet(value.sheetId);
      return `<definedName name="${x(name)}">${x("'" + sheet.meta.name.replace(/'/g, "''") + "'!" + rangeName(value.range).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2"))}</definedName>`;
    })
    .join("");
  await add("xl/workbook.xml", [
    HEADER +
      `<workbook xmlns="${NS}" xmlns:r="${REL}"><workbookPr date1904="${model.dateSystem === 1904 ? 1 : 0}"/><bookViews><workbookView/></bookViews><sheets>${sheets.map((s, i) => `<sheet name="${x(s.meta.name)}" sheetId="${i + 1}"${s.meta.hidden ? ' state="hidden"' : ""} r:id="rId${i + 1}"/>`).join("")}</sheets>${defined ? `<definedNames>${defined}</definedNames>` : ""}<calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>`,
  ]);
  await add("xl/_rels/workbook.xml.rels", [
    HEADER +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="styles" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="metadata" Type="${REL}/sheetMetadata" Target="metadata.xml"/></Relationships>`,
  ]);
  await add("xl/styles.xml", [
    writeStyles(model.styles).replace(
      "</styleSheet>",
      writeDxfs(
        sheets.flatMap((s) => s.meta.conditionalFormats ?? []),
        164 + model.styles.length,
      ) + "</styleSheet>",
    ),
  ]);
  await add("xl/metadata.xml", [metadataXml]);
  for (let i = 0; i < sheets.length; i++) {
    await add(`xl/worksheets/sheet${i + 1}.xml`, sheetXml(model, sheets[i]));
    const objects = preserved[i];
    if (objects) {
      for (const [path, bytes] of [
        ...objects.parts,
        [
          `xl/worksheets/_rels/sheet${i + 1}.xml.rels`,
          objects.relationships,
        ] as const,
      ]) {
        const file = new ZipDeflate(path, { level: 3 });
        zip.add(file);
        file.push(bytes, true);
        await checkpoint();
      }
    }
    options.onProgress?.({
      stage: "worksheets",
      progress: (i + 1) / sheets.length,
    });
  }
  zip.end();
  if (failure) throw failure;
  const data = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
  let offset = 0;
  for (const chunk of chunks) {
    data.set(chunk, offset);
    offset += chunk.length;
  }
  return { data, diagnostics };
}
export async function writeXlsx(
  snapshot: WorkbookSnapshot,
  options: XlsxWriteOptions = {},
): Promise<{ data: Uint8Array; diagnostics: Diagnostic[] }> {
  return writeModelXlsx(new WorkbookModel({ snapshot }), options);
}

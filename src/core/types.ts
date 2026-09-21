export type ErrorCode =
  | "#DIV/0!"
  | "#VALUE!"
  | "#REF!"
  | "#NAME?"
  | "#NUM!"
  | "#N/A"
  | "#SPILL!"
  | "#CALC!"
  | "#CYCLE!";
export interface CellError {
  error: ErrorCode;
}
export type Scalar = string | number | boolean | null | CellError;
export type InputValue = string | number | boolean | null;
export interface Cell {
  value?: Scalar;
  formula?: string;
  style?: number;
}
export interface CellStyle {
  locked?: boolean;
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  background?: string;
  align?: "left" | "center" | "right";
  verticalAlign?: "top" | "center" | "bottom";
  wrap?: boolean;
  numberFormat?: string;
  border?: { top?: string; right?: string; bottom?: string; left?: string };
}
export interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}
export type FilterRule =
  | { column: number; operator: "in"; values: string[] }
  | {
      column: number;
      operator: "eq" | "contains" | "gt" | "lt" | "gte" | "lte" | "neq";
      value: InputValue;
    };
export interface ValidationRule {
  range: Rect;
  type: "list" | "decimal" | "whole" | "date" | "textLength";
  values?: string[];
  minimum?: number;
  maximum?: number;
  allowBlank?: boolean;
}
export interface ConditionalRule {
  range: Rect;
  operator: "eq" | "neq" | "gt" | "gte" | "lt" | "lte" | "between" | "contains";
  value: InputValue;
  second?: number;
  style: CellStyle;
}
export interface PreservedObjects {
  elements: string[];
  relationships: string;
  parts: Record<string, string>;
  contentTypes: string[];
}
export interface SheetMeta {
  id: string;
  name: string;
  hidden?: boolean;
  protected?: boolean;
  validations?: ValidationRule[];
  conditionalFormats?: ConditionalRule[];
  objects?: PreservedObjects;
  rowCount: number;
  columnCount: number;
  rowHeights: Record<number, number>;
  columnWidths: Record<number, number>;
  hiddenRows: number[];
  hiddenColumns: number[];
  frozenRows: number;
  frozenColumns: number;
  merges: Rect[];
  filter?: { range: Rect; rules: FilterRule[] };
  filteredRows?: number[];
}
export interface SheetSnapshot extends SheetMeta {
  cells: [number, Cell][];
}
export interface NamedRange {
  sheetId: string;
  range: Rect;
}
export interface WorkbookSnapshot {
  version: 1;
  sheets: SheetSnapshot[];
  styles: CellStyle[];
  names: Record<string, NamedRange>;
  dateSystem: 1900 | 1904;
  diagnostics?: Diagnostic[];
}
export interface Diagnostic {
  code: string;
  severity: "warning" | "error";
  message: string;
  sheetId?: string;
  range?: string;
  lossy?: boolean;
}
export interface RegionCell {
  row: number;
  column: number;
  value: Scalar;
  formula?: string;
  style: CellStyle;
  displayStyle?: CellStyle;
  spill?: { row: number; column: number };
}
export interface Region {
  range: Rect;
  cells: RegionCell[];
  revision: number;
}
export interface OperationContext {
  operationId?: string;
  origin?: string;
}
export interface ChangeEvent extends OperationContext {
  revision: number;
  changes: { sheetId: string; range?: Rect }[];
  source: "edit" | "undo" | "redo" | "import";
}
export interface ProgressEvent extends OperationContext {
  operation: string;
  stage: string;
  progress: number;
}
export interface WorkbookOptions {
  sheets?: { name: string; rows?: number; columns?: number }[];
  snapshot?: WorkbookSnapshot;
  workerUrl?: string | URL;
  workerFactory?: () => Worker;
  /** Trusted same-origin or CORS-enabled modules; evaluated inside each isolated Worker. */
  workerModules?: (string | URL)[];
  initializationTimeout?: number;
  /** Milliseconds; 0 disables the operation timeout. A timeout closes the Worker. */
  requestTimeout?: number;
  historyLimit?: number;
  dateSystem?: 1900 | 1904;
}
export interface OperationOptions extends OperationContext {
  signal?: AbortSignal;
  timeout?: number;
}
export interface SortKey {
  column: number;
  direction: "asc" | "desc";
}
export type Command =
  | {
      type: "setValues";
      sheetId: string;
      range: Rect;
      values: InputValue[][];
      parseFormulas?: boolean;
    }
  | {
      type: "setFormula";
      sheetId: string;
      row: number;
      column: number;
      formula: string;
    }
  | { type: "style"; sheetId: string; range: Rect; style: CellStyle }
  | { type: "clear"; sheetId: string; range: Rect; formats?: boolean }
  | {
      type: "addSheet";
      name: string;
      id?: string;
      rows?: number;
      columns?: number;
    }
  | { type: "deleteSheet"; sheetId: string }
  | { type: "renameSheet"; sheetId: string; name: string }
  | { type: "reorderSheet"; sheetId: string; index: number }
  | {
      type: "dimensions";
      sheetId: string;
      axis: "row" | "column";
      indexes: number[];
      size?: number;
      hidden?: boolean;
    }
  | {
      type: "structure";
      sheetId: string;
      axis: "row" | "column";
      index: number;
      count: number;
      delete?: boolean;
    }
  | { type: "merge"; sheetId: string; range: Rect; unmerge?: boolean }
  | { type: "freeze"; sheetId: string; rows: number; columns: number }
  | {
      type: "sort";
      sheetId: string;
      range: Rect;
      keys: SortKey[];
      header?: boolean;
    }
  | { type: "filter"; sheetId: string; range?: Rect; rules: FilterRule[] }
  | {
      type: "copy";
      sheetId: string;
      range: Rect;
      targetSheetId: string;
      targetRow: number;
      targetColumn: number;
      cut?: boolean;
      valuesOnly?: boolean;
    }
  | { type: "fill"; sheetId: string; source: Rect; target: Rect }
  | {
      type: "paste";
      sheetId: string;
      source: Rect;
      cells: RegionCell[];
      targetRow: number;
      targetColumn: number;
      mode?: "all" | "values" | "formulas" | "formats";
      transpose?: boolean;
      sourceRows?: number[];
      sourceColumns?: number[];
      targetRows?: number[];
      targetColumns?: number[];
      merges?: Rect[];
      validations?: ValidationRule[];
      conditionalFormats?: ConditionalRule[];
    }
  | { type: "duplicateSheet"; sheetId: string; id: string; name: string }
  | { type: "sheetVisibility"; sheetId: string; hidden: boolean }
  | { type: "protect"; sheetId: string; enabled: boolean }
  | {
      type: "validation";
      sheetId: string;
      range: Rect;
      rule?: Omit<ValidationRule, "range">;
    }
  | {
      type: "conditionalFormat";
      sheetId: string;
      range: Rect;
      rule?: Omit<ConditionalRule, "range">;
    }
  | { type: "renameName"; name: string; newName: string }
  | { type: "deleteName"; name: string }
  | { type: "defineName"; name: string; sheetId: string; range: Rect }
  | {
      type: "replace";
      sheetId: string;
      search: string;
      replacement: string;
      matchCase?: boolean;
      entireCell?: boolean;
    };

import { formatColor } from "../core/number-format";
import type { Rect, RegionCell, SheetMeta } from "../core/types";
import { columnName, contains, keyOf } from "../core/address";
import { formatValue } from "../core/format";
import { Axis, COLUMN_HEADER, ROW_HEADER } from "./geometry";
export interface RenderState {
  zoom?: number;
  sheet: SheetMeta;
  rows: Axis;
  columns: Axis;
  cells: Map<number, RegionCell>;
  selection: Rect;
  reference?: Rect;
  scrollLeft: number;
  scrollTop: number;
  width: number;
  height: number;
  dateSystem: 1900 | 1904;
  accent: string;
}
export function cellBox(
  row: number,
  col: number,
  state: RenderState,
): { x: number; y: number; width: number; height: number } {
  return {
    x:
      ROW_HEADER +
      state.columns.offsets[col] -
      (col < state.sheet.frozenColumns ? 0 : state.scrollLeft),
    y:
      COLUMN_HEADER +
      state.rows.offsets[row] -
      (row < state.sheet.frozenRows ? 0 : state.scrollTop),
    width: state.columns.size(col),
    height: state.rows.size(row),
  };
}
export function drawGrid(canvas: HTMLCanvasElement, state: RenderState): void {
  const { width, height, rows, columns, sheet, selection, cells } = state,
    dpr = (window.devicePixelRatio || 1) * (state.zoom ?? 1);
  if (
    canvas.width !== Math.round(width * dpr) ||
    canvas.height !== Math.round(height * dpr)
  ) {
    canvas.width = Math.round(width * dpr);
    canvas.height = Math.round(height * dpr);
    canvas.style.width = width + "px";
    canvas.style.height = height + "px";
  }
  const ctx = canvas.getContext("2d")!;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, width, height);
  const visibleRows = rows.visible(
      state.scrollTop,
      height - COLUMN_HEADER,
      sheet.frozenRows,
    ),
    visibleColumns = columns.visible(
      state.scrollLeft,
      width - ROW_HEADER,
      sheet.frozenColumns,
    );
  const frozenX = ROW_HEADER + columns.offsets[sheet.frozenColumns],
    frozenY = COLUMN_HEADER + rows.offsets[sheet.frozenRows];
  // Separate clips keep scrolled cells from painting over frozen panes.
  for (const frozenRow of [false, true])
    for (const frozenColumn of [false, true]) {
      const left = frozenColumn ? ROW_HEADER : frozenX,
        top = frozenRow ? COLUMN_HEADER : frozenY,
        right = frozenColumn ? frozenX : width,
        bottom = frozenRow ? frozenY : height;
      if (right <= left || bottom <= top) continue;
      ctx.save();
      ctx.beginPath();
      ctx.rect(left, top, right - left, bottom - top);
      ctx.clip();
      const paintedMerges = new Set<number>();
      for (const r of visibleRows.filter((r) => r.frozen === frozenRow))
        for (const c of visibleColumns.filter(
          (c) => c.frozen === frozenColumn,
        )) {
          const mergeIndex = sheet.merges.findIndex((m) =>
              contains(m, r.index, c.index),
            ),
            merge = sheet.merges[mergeIndex];
          if (merge && paintedMerges.has(mergeIndex)) continue;
          if (merge) paintedMerges.add(mergeIndex);
          const row = merge?.r1 ?? r.index,
            col = merge?.c1 ?? c.index,
            cell = cells.get(keyOf(row, col)),
            style = cell?.displayStyle ?? cell?.style ?? {};
          const box = cellBox(row, col, state);
          if (merge) {
            box.width =
              columns.offsets[merge.c2 + 1] - columns.offsets[merge.c1];
            box.height = rows.offsets[merge.r2 + 1] - rows.offsets[merge.r1];
          }
          ctx.fillStyle = style.background ?? "#fff";
          ctx.fillRect(box.x, box.y, box.width, box.height);
          if (contains(selection, r.index, c.index)) {
            ctx.fillStyle = "rgba(22,115,76,0.065)";
            ctx.fillRect(box.x, box.y, box.width, box.height);
          }
          ctx.strokeStyle = "#e8ece9";
          ctx.lineWidth = 1;
          ctx.strokeRect(box.x + 0.5, box.y + 0.5, box.width, box.height);
          if (style.border) {
            for (const side of ["top", "right", "bottom", "left"] as const)
              if (style.border[side]) {
                ctx.strokeStyle = style.border[side]!;
                ctx.beginPath();
                if (side === "top" || side === "bottom") {
                  const y = box.y + (side === "bottom" ? box.height : 0);
                  ctx.moveTo(box.x, y + 0.5);
                  ctx.lineTo(box.x + box.width, y + 0.5);
                } else {
                  const x = box.x + (side === "right" ? box.width : 0);
                  ctx.moveTo(x + 0.5, box.y);
                  ctx.lineTo(x + 0.5, box.y + box.height);
                }
                ctx.stroke();
              }
          }
          if (!cell || cell.value === null) continue;
          const text = formatValue(cell.value, style, state.dateSystem),
            align =
              style.align ??
              (typeof cell.value === "number"
                ? "right"
                : typeof cell.value === "boolean"
                  ? "center"
                  : "left");
          const fontSize = ((style.fontSize ?? 10) * 4) / 3;
          ctx.save();
          ctx.beginPath();
          ctx.rect(
            box.x + 3,
            box.y + 2,
            Math.max(0, box.width - 6),
            Math.max(0, box.height - 4),
          );
          ctx.clip();
          ctx.font = `${style.italic ? "italic " : ""}${style.bold ? "600 " : ""}${fontSize}px ${style.fontFamily ?? "Arial, sans-serif"}`;
          ctx.fillStyle =
            formatColor(cell?.value, style.numberFormat) ??
            style.color ??
            "#293b32";
          ctx.textAlign = align;
          ctx.textBaseline = "middle";
          const x =
            align === "right"
              ? box.x + box.width - 8
              : align === "center"
                ? box.x + box.width / 2
                : box.x + 8;
          const lines: string[] = [];
          for (const original of text.split("\n")) {
            if (!style.wrap) {
              lines.push(original);
              continue;
            }
            let line = "";
            for (const char of original) {
              if (line && ctx.measureText(line + char).width > box.width - 16) {
                lines.push(line);
                line = "";
              }
              line += char;
            }
            lines.push(line);
          }
          const lineHeight = fontSize * 1.35,
            blockHeight = lines.length * lineHeight;
          let y =
            style.verticalAlign === "top"
              ? box.y + lineHeight / 2 + 3
              : style.verticalAlign === "bottom"
                ? box.y + box.height - blockHeight + lineHeight / 2 - 3
                : box.y + (box.height - blockHeight) / 2 + lineHeight / 2;
          for (const line of lines) {
            ctx.fillText(line, x, y);
            if (style.underline) {
              const length = ctx.measureText(line).width,
                start =
                  align === "right"
                    ? x - length
                    : align === "center"
                      ? x - length / 2
                      : x;
              ctx.strokeStyle = ctx.fillStyle;
              ctx.beginPath();
              ctx.moveTo(start, y + fontSize / 2);
              ctx.lineTo(start + length, y + fontSize / 2);
              ctx.stroke();
            }
            y += lineHeight;
          }
          ctx.restore();
        }
      const start = cellBox(selection.r1, selection.c1, state),
        end = cellBox(selection.r2, selection.c2, state);
      ctx.strokeStyle = state.accent;
      ctx.lineWidth = 2;
      ctx.strokeRect(
        start.x + 1,
        start.y + 1,
        end.x + end.width - start.x - 2,
        end.y + end.height - start.y - 2,
      );
      ctx.fillStyle = state.accent;
      ctx.fillRect(end.x + end.width - 4, end.y + end.height - 4, 7, 7);
      if (state.reference) {
        const start = cellBox(state.reference.r1, state.reference.c1, state),
          end = cellBox(state.reference.r2, state.reference.c2, state);
        ctx.strokeStyle = "#3978d4";
        ctx.setLineDash([5, 3]);
        ctx.strokeRect(
          start.x + 1,
          start.y + 1,
          end.x + end.width - start.x - 2,
          end.y + end.height - start.y - 2,
        );
      }
      ctx.restore();
    }
  ctx.fillStyle = "#f6f8f7";
  ctx.fillRect(0, 0, width, COLUMN_HEADER);
  ctx.fillRect(0, 0, ROW_HEADER, height);
  ctx.strokeStyle = "#dfe6e1";
  ctx.lineWidth = 1;
  ctx.font = "11px Arial, sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  for (const frozen of [false, true]) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(
      frozen ? ROW_HEADER : frozenX,
      0,
      frozen ? frozenX - ROW_HEADER : width - frozenX,
      COLUMN_HEADER,
    );
    ctx.clip();
    for (const c of visibleColumns.filter((c) => c.frozen === frozen)) {
      const x = ROW_HEADER + c.start;
      if (c.index >= selection.c1 && c.index <= selection.c2) {
        ctx.fillStyle = "#e3f0e8";
        ctx.fillRect(x, 0, c.size, COLUMN_HEADER);
      }
      ctx.fillStyle = "#65786d";
      ctx.fillText(columnName(c.index), x + c.size / 2, COLUMN_HEADER / 2);
      if (
        sheet.filter &&
        c.size > 28 &&
        c.index >= sheet.filter.range.c1 &&
        c.index <= sheet.filter.range.c2
      ) {
        ctx.fillStyle = sheet.filter.rules.some(
          (rule) => rule.column === c.index,
        )
          ? state.accent
          : "#65786d";
        ctx.fillText("▾", x + c.size - 15, COLUMN_HEADER / 2);
      }
      ctx.strokeRect(x + 0.5, 0.5, c.size, COLUMN_HEADER);
    }
    ctx.restore();
    ctx.save();
    ctx.beginPath();
    ctx.rect(
      0,
      frozen ? COLUMN_HEADER : frozenY,
      ROW_HEADER,
      frozen ? frozenY - COLUMN_HEADER : height - frozenY,
    );
    ctx.clip();
    for (const r of visibleRows.filter((r) => r.frozen === frozen)) {
      const y = COLUMN_HEADER + r.start;
      if (r.index >= selection.r1 && r.index <= selection.r2) {
        ctx.fillStyle = "#e3f0e8";
        ctx.fillRect(0, y, ROW_HEADER, r.size);
      }
      ctx.fillStyle = "#65786d";
      ctx.fillText(String(r.index + 1), ROW_HEADER / 2, y + r.size / 2);
      ctx.strokeRect(0.5, y + 0.5, ROW_HEADER, r.size);
    }
    ctx.restore();
  }
  ctx.fillStyle = "#f0f4f1";
  ctx.fillRect(0, 0, ROW_HEADER, COLUMN_HEADER);
  ctx.fillStyle = "#b6c5bc";
  ctx.beginPath();
  ctx.moveTo(ROW_HEADER - 14, COLUMN_HEADER - 7);
  ctx.lineTo(ROW_HEADER - 6, COLUMN_HEADER - 15);
  ctx.lineTo(ROW_HEADER - 6, COLUMN_HEADER - 7);
  ctx.fill();
  ctx.strokeStyle = "#a0b7aa";
  if (sheet.frozenRows) {
    ctx.beginPath();
    ctx.moveTo(0, frozenY + 0.5);
    ctx.lineTo(width, frozenY + 0.5);
    ctx.stroke();
  }
  if (sheet.frozenColumns) {
    ctx.beginPath();
    ctx.moveTo(frozenX + 0.5, 0);
    ctx.lineTo(frozenX + 0.5, height);
    ctx.stroke();
  }
}

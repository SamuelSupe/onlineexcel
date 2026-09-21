import type { InputValue } from "./types";
import { dateSerial } from "../formula/values";
export type ColumnType =
  | "auto"
  | "text"
  | "number"
  | "date"
  | "boolean"
  | "percent";
export function parseInputValue(
  text: string,
  type: ColumnType = "auto",
  dateSystem: 1900 | 1904 = 1900,
): { value: InputValue; format?: string } {
  if (type === "text") return { value: text };
  if (text.startsWith("'")) return { value: text.slice(1) };
  if (!text) return { value: null };
  const value = text.trim();
  if ((type === "auto" || type === "boolean") && /^(TRUE|FALSE)$/i.test(value))
    return { value: /^true$/i.test(value) };
  if (type === "auto" || type === "date") {
    const date =
      /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(
        value,
      );
    if (date) {
      const [year, month, day, hour = 0, minute = 0, second = 0] = date
        .slice(1)
        .map((v) => (v === undefined ? undefined : Number(v)));
      const d = new Date(
        Date.UTC(year!, month! - 1, day!, hour, minute, second),
      );
      if (
        d.getUTCFullYear() === year &&
        d.getUTCMonth() === month! - 1 &&
        d.getUTCDate() === day &&
        hour < 24 &&
        minute < 60 &&
        second < 60
      )
        return {
          value: dateSerial(d, dateSystem),
          format: date[4] ? "yyyy-mm-dd hh:mm:ss" : "yyyy-mm-dd",
        };
    }
  }
  if (["auto", "number", "percent"].includes(type)) {
    const percent = /%$/.test(value);
    const currency = /^[$¥€£₩]/.exec(value)?.[0];
    const normalized = value.replace(/^[$¥€£₩]\s*/, "").replace(/%$/, "");
    const numeric =
      /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(normalized) ||
      /^[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?$/.test(normalized);
    const safeAuto =
      !/^[-+]?0\d/.test(normalized) &&
      normalized.replace(/[^0-9]/g, "").length <= 15;
    if (numeric && (type !== "auto" || safeAuto)) {
      const n =
        Number(normalized.replace(/,/g, "")) /
        (percent || type === "percent" ? 100 : 1);
      if (Number.isFinite(n))
        return {
          value: n,
          format:
            percent || type === "percent"
              ? "0.00%"
              : currency
                ? `"${currency}"#,##0.00`
                : normalized.includes(",")
                  ? "#,##0.00"
                  : undefined,
        };
    }
  }
  if (type !== "auto") throw new Error(`Invalid ${type} value: ${text}`);
  return { value: text };
}

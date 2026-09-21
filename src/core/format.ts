import { formatSections, numberSection, fractionText } from "./number-format";
import type { CellStyle, Scalar } from "./types";
import { isError, serialDate } from "../formula/values";
export function formatValue(
  value: Scalar,
  style: CellStyle = {},
  dateSystem: 1900 | 1904 = 1900,
): string {
  if (value === null) return "";
  if (isError(value)) return value.error;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "string") {
    const text = formatSections(style.numberFormat ?? "")[3];
    return text === undefined
      ? value
      : text.replace(
          /"([^"]*)"|\\(.)|@/g,
          (_, literal, escaped) => literal ?? escaped ?? value,
        );
  }
  const pattern = style.numberFormat ?? "General";
  if (pattern === "General" || pattern === "@")
    return String(Number(value.toPrecision(15)));
  const chosen = numberSection(pattern, value),
    section = chosen.section;
  if (!section) return "";
  if (chosen.absolute) value = Math.abs(value);
  const stripped = section.replace(/"[^"]*"/g, "").replace(/\\./g, "");
  if (/\[(h+|m+|s+)\]/i.test(section)) {
    const seconds = Math.round(Math.abs(value) * 86400);
    return section.replace(/\[(h+|m+|s+)\]|hh|mm|ss/gi, (token, elapsed) => {
      const unit = (elapsed ?? token)[0].toLowerCase();
      const n =
        unit === "h"
          ? Math.floor(seconds / 3600)
          : unit === "m"
            ? Math.floor(seconds / 60)
            : seconds;
      return String(elapsed ? n : n % (unit === "h" ? 24 : 60)).padStart(
        token.replace(/[\[\]]/g, "").length,
        "0",
      );
    });
  }
  const fraction = fractionText(Math.abs(value), section);
  if (fraction !== undefined) return (value < 0 ? "-" : "") + fraction;
  if (!/[0#?ydhms]/i.test(stripped))
    return section.replace(
      /"([^"]*)"|\\(.)/g,
      (_, literal, escaped) => literal ?? escaped,
    );
  if (/[ydhs]/i.test(stripped) || /m{2,}/i.test(stripped)) {
    const d = serialDate(value, dateSystem),
      pad = (n: number) => String(n).padStart(2, "0");
    const tokens =
      /"([^"]*)"|\\(.)|yyyy|yy|mmmm|mmm|mm|m|dd|d|hh|h|ss|s|AM\/PM/gi;
    let lastHour = false;
    return section.replace(tokens, (token, literal, escaped, offset) => {
      if (literal !== undefined) return literal;
      if (escaped !== undefined) return escaped;
      const t = token.toLowerCase();
      if (t === "yyyy") return String(d.getUTCFullYear());
      if (t === "yy") return pad(d.getUTCFullYear() % 100);
      if (t === "mmmm" || t === "mmm")
        return d.toLocaleString("en-US", {
          month: t === "mmmm" ? "long" : "short",
          timeZone: "UTC",
        });
      if (t === "mm" || t === "m") {
        const minute =
          lastHour || /^:s/i.test(section.slice(offset + token.length));
        const n = minute ? d.getUTCMinutes() : d.getUTCMonth() + 1;
        lastHour = false;
        return t === "mm" ? pad(n) : String(n);
      }
      if (t === "dd" || t === "d") {
        const day =
          dateSystem === 1900 && Math.floor(value) === 60 ? 29 : d.getUTCDate();
        return t === "dd" ? pad(day) : String(day);
      }
      if (t === "hh" || t === "h") {
        lastHour = true;
        let hour = d.getUTCHours();
        if (/AM\/PM/i.test(section)) hour = hour % 12 || 12;
        return t === "hh" ? pad(hour) : String(hour);
      }
      if (t === "ss" || t === "s")
        return t === "ss" ? pad(d.getUTCSeconds()) : String(d.getUTCSeconds());
      return d.getUTCHours() < 12 ? "AM" : "PM";
    });
  }
  if (/[eE][+-]0/.test(section)) {
    const decimals = Math.min(20, /\.([0#]+)/.exec(section)?.[1].length ?? 0);
    const digits = /[eE][+-](0+)/.exec(section)?.[1].length ?? 2;
    return value
      .toExponential(decimals)
      .replace(
        /e([+-])(\d+)/,
        (_, sign, exponent) => "E" + sign + exponent.padStart(digits, "0"),
      );
  }
  const percent = (stripped.match(/%/g) ?? []).length,
    decimals = /\.([0#]+)/.exec(stripped)?.[1] ?? "";
  const n = Math.abs(value) * 100 ** percent;
  let result = n.toLocaleString("en-US", {
    useGrouping: /[0#],[0#]/.test(stripped),
    minimumFractionDigits: Math.min(20, (decimals.match(/0/g) ?? []).length),
    maximumFractionDigits: Math.min(20, decimals.length),
  });
  const prefix =
    /^([^0#?]*)(?:[0#?])/
      .exec(section)?.[1]
      ?.replace(/"/g, "")
      .replace(/\\/g, "") ?? "";
  const suffix =
    /[0#?]([^0#?]*)$/
      .exec(section)?.[1]
      ?.replace(/"/g, "")
      .replace(/\\/g, "") ?? "";
  result = prefix + result + suffix;
  if (value < 0) result = "-" + result;
  return result;
}

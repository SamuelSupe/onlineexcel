import { define } from "./registry";
import { bool, error, flatten, isError, number, scalar, text } from "./values";
import { formatValue } from "../core/format";
for (const [name, fn] of Object.entries({
  UPPER: (s: string) => s.toUpperCase(),
  LOWER: (s: string) => s.toLowerCase(),
  TRIM: (s: string) => s.trim().replace(/ +/g, " "),
  CLEAN: (s: string) => s.replace(/[\x00-\x1f]/g, ""),
  PROPER: (s: string) =>
    s.toLowerCase().replace(/(^|[^\p{L}])\p{L}/gu, (s) => s.toUpperCase()),
}))
  define(name, 1, 1, "text", (a) => fn(text(a[0])), true);
define("LEN", 1, 1, "text", (a) => text(a[0]).length, true);
for (const name of ["LEFT", "RIGHT"])
  define(
    name,
    1,
    2,
    "text",
    (a) => {
      const s = text(a[0]),
        n = Math.trunc(number(a[1] ?? 1));
      return n < 0
        ? error("#VALUE!")
        : name === "LEFT"
          ? s.slice(0, n)
          : n
            ? s.slice(-n)
            : "";
    },
    true,
  );
define(
  "MID",
  3,
  3,
  "text",
  (a) => {
    const start = Math.trunc(number(a[1])),
      length = Math.trunc(number(a[2]));
    return start < 1 || length < 0
      ? error("#VALUE!")
      : text(a[0]).slice(start - 1, start - 1 + length);
  },
  true,
);
for (const name of ["FIND", "SEARCH"])
  define(
    name,
    2,
    3,
    "text",
    (a) => {
      const needle = text(a[0]),
        hay = text(a[1]),
        start = Math.trunc(number(a[2] ?? 1));
      if (start < 1 || start > hay.length + 1) return error("#VALUE!");
      let index: number;
      if (name === "FIND") index = hay.indexOf(needle, start - 1);
      else {
        let pattern = "";
        for (let i = 0; i < needle.length; i++) {
          const c = needle[i];
          if (c === "~" && i + 1 < needle.length)
            pattern += needle[++i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
          else
            pattern +=
              c === "*"
                ? ".*?"
                : c === "?"
                  ? "."
                  : c.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        }
        const match = new RegExp(pattern, "i").exec(hay.slice(start - 1));
        index = match ? start - 1 + match.index : -1;
      }
      return index < 0 ? error("#VALUE!") : index + 1;
    },
    true,
  );
define(
  "REPLACE",
  4,
  4,
  "text",
  (a) => {
    const start = Math.trunc(number(a[1])),
      length = Math.trunc(number(a[2]));
    return start < 1 || length < 0
      ? error("#VALUE!")
      : text(a[0]).slice(0, start - 1) +
          text(a[3]) +
          text(a[0]).slice(start - 1 + length);
  },
  true,
);
define(
  "SUBSTITUTE",
  3,
  4,
  "text",
  (a) => {
    const source = text(a[0]),
      old = text(a[1]),
      replacement = text(a[2]);
    if (!old) return source;
    if (a[3] === undefined) return source.split(old).join(replacement);
    const occurrence = Math.trunc(number(a[3]));
    if (occurrence < 1) return error("#VALUE!");
    let seen = 0;
    return source.replace(
      new RegExp(old.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g"),
      (match) => (++seen === occurrence ? replacement : match),
    );
  },
  true,
);
define(
  "REPT",
  2,
  2,
  "text",
  (a) => {
    const n = Math.trunc(number(a[1])),
      s = text(a[0]);
    return n < 0 || s.length * n > 32767 ? error("#VALUE!") : s.repeat(n);
  },
  true,
);
define("CONCAT", 1, 255, "text", (a) => [...flatten(a)].map(text).join(""));
define("TEXTJOIN", 3, 255, "text", (a) =>
  [...flatten(a.slice(2))]
    .filter((v) => !bool(a[1]) || (v !== null && v !== ""))
    .map(text)
    .join(text(a[0])),
);
define("EXACT", 2, 2, "text", (a) => text(a[0]) === text(a[1]), true);
define(
  "CHAR",
  1,
  1,
  "text",
  (a) => {
    const n = Math.trunc(number(a[0]));
    return n >= 1 && n <= 255 ? String.fromCharCode(n) : error("#VALUE!");
  },
  true,
  "Uses Unicode code points 1–255; platform-specific legacy code pages are not emulated.",
);
define(
  "CODE",
  1,
  1,
  "text",
  (a) => text(a[0]).charCodeAt(0) || error("#VALUE!"),
  true,
);
define(
  "UNICHAR",
  1,
  1,
  "text",
  (a) => {
    const n = Math.trunc(number(a[0]));
    return n < 1 || n > 0x10ffff || (n >= 0xd800 && n <= 0xdfff)
      ? error("#VALUE!")
      : String.fromCodePoint(n);
  },
  true,
);
define(
  "UNICODE",
  1,
  1,
  "text",
  (a) => text(a[0]).codePointAt(0) ?? error("#VALUE!"),
  true,
);
define(
  "VALUE",
  1,
  1,
  "text",
  (a) => {
    const s = text(a[0]).trim();
    const percent = s.endsWith("%");
    const n = Number(s.replace(/,/g, "").replace(/%$/, ""));
    return Number.isFinite(n) ? n / (percent ? 100 : 1) : error("#VALUE!");
  },
  true,
  "Invariant decimal syntax; locale-specific currency and date strings are not parsed.",
);
define(
  "NUMBERVALUE",
  1,
  3,
  "text",
  (a) => {
    const decimal = text(a[1] ?? "."),
      grouping = text(a[2] ?? ",");
    if (decimal === grouping) return error("#VALUE!");
    const s = text(a[0])
      .split(grouping)
      .join("")
      .replace(decimal, ".")
      .replace(/\s/g, "");
    const percent = /%+$/.exec(s)?.[0].length ?? 0;
    const n = Number(s.replace(/%+$/, ""));
    return Number.isFinite(n) ? n / 100 ** percent : error("#VALUE!");
  },
  true,
);
define(
  "TEXT",
  2,
  2,
  "text",
  (a, ctx) =>
    formatValue(scalar(a[0]), { numberFormat: text(a[1]) }, ctx.dateSystem),
  true,
);
define(
  "FIXED",
  1,
  3,
  "text",
  (a) => {
    const digits = Math.trunc(number(a[1] ?? 2)),
      n = number(a[0]);
    if (digits > 100 || digits < -100) return error("#VALUE!");
    const rounded =
      (Math.sign(n) * Math.round(Math.abs(n) * 10 ** digits)) / 10 ** digits;
    return rounded.toLocaleString("en-US", {
      useGrouping: !bool(a[2] ?? false),
      minimumFractionDigits: Math.max(0, digits),
      maximumFractionDigits: Math.max(0, digits),
    });
  },
  true,
);
define(
  "T",
  1,
  1,
  "information",
  (a) => {
    const v = scalar(a[0]);
    return typeof v === "string" || isError(v) ? v : "";
  },
  true,
);
define(
  "N",
  1,
  1,
  "information",
  (a) => {
    const v = scalar(a[0]);
    return typeof v === "number" || isError(v)
      ? v
      : typeof v === "boolean"
        ? +v
        : 0;
  },
  true,
);
for (const [name, test] of Object.entries({
  ISBLANK: (v: unknown) => v === null,
  ISNUMBER: (v: unknown) => typeof v === "number",
  ISTEXT: (v: unknown) => typeof v === "string",
  ISNONTEXT: (v: unknown) => typeof v !== "string",
  ISLOGICAL: (v: unknown) => typeof v === "boolean",
  ISERROR: isError,
  ISERR: (v: unknown) => isError(v) && v.error !== "#N/A",
  ISNA: (v: unknown) => isError(v) && v.error === "#N/A",
}))
  define(name, 1, 1, "information", (a) => test(scalar(a[0])), true);
define(
  "ISEVEN",
  1,
  1,
  "information",
  (a) => Math.trunc(number(a[0])) % 2 === 0,
  true,
);
define(
  "ISODD",
  1,
  1,
  "information",
  (a) => Math.abs(Math.trunc(number(a[0])) % 2) === 1,
  true,
);
define("TYPE", 1, 1, "information", (a) => {
  const v = a[0];
  return v && typeof v === "object"
    ? isError(v)
      ? 16
      : 64
    : typeof v === "string"
      ? 2
      : typeof v === "boolean"
        ? 4
        : 1;
});
define("NA", 0, 0, "information", () => error("#N/A"));
define("TRUE", 0, 0, "logical", () => true);
define("FALSE", 0, 0, "logical", () => false);
define("NOT", 1, 1, "logical", (a) => !bool(a[0]), true);
for (const name of ["AND", "OR", "XOR"])
  define(name, 1, 255, "logical", (a) => {
    const values = [...flatten(a)].filter(
      (v) => typeof v !== "string" && v !== null,
    );
    if (!values.length) return error("#VALUE!");
    const bs = values.map(bool);
    return name === "AND"
      ? bs.every(Boolean)
      : name === "OR"
        ? bs.some(Boolean)
        : bs.filter(Boolean).length % 2 === 1;
  });
// These functions are evaluated lazily by the interpreter so unused branches cannot fail.
for (const [name, min, max] of [
  ["IF", 2, 3],
  ["IFERROR", 2, 2],
  ["IFNA", 2, 2],
  ["IFS", 2, 254],
  ["SWITCH", 3, 255],
  ["CHOOSE", 2, 255],
] as const)
  define(name, min, max, "logical", () => error("#CALC!"));

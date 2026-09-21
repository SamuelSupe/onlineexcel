import { define } from "./registry";
import {
  dateSerial,
  DAY,
  error,
  flatten,
  number,
  serialDate,
  text,
} from "./values";
import type { FunctionContext } from "./registry";
function date(value: Parameters<typeof number>[0], ctx: FunctionContext): Date {
  return serialDate(number(value), ctx.dateSystem);
}
define(
  "DATE",
  3,
  3,
  "date",
  (a, ctx) => {
    let year = Math.trunc(number(a[0]));
    if (year >= 0 && year < 1900) year += 1900;
    if (year < 0 || year >= 10000) return error("#NUM!");
    const d = new Date(0);
    d.setUTCFullYear(
      year,
      Math.trunc(number(a[1])) - 1,
      Math.trunc(number(a[2])),
    );
    return dateSerial(d, ctx.dateSystem);
  },
  true,
);
define(
  "TIME",
  3,
  3,
  "date",
  (a) => {
    const seconds =
      Math.trunc(number(a[0])) * 3600 +
      Math.trunc(number(a[1])) * 60 +
      Math.trunc(number(a[2]));
    return seconds < 0 ? error("#NUM!") : (seconds % 86400) / 86400;
  },
  true,
);
for (const [name, get] of Object.entries({
  YEAR: (d: Date) => d.getUTCFullYear(),
  MONTH: (d: Date) => d.getUTCMonth() + 1,
  DAY: (d: Date) => d.getUTCDate(),
  HOUR: (d: Date) => d.getUTCHours(),
  MINUTE: (d: Date) => d.getUTCMinutes(),
  SECOND: (d: Date) => d.getUTCSeconds(),
}))
  define(
    name,
    1,
    1,
    "date",
    (a, ctx) =>
      number(a[0]) < 0
        ? error("#NUM!")
        : name === "DAY" &&
            ctx.dateSystem === 1900 &&
            Math.floor(number(a[0])) === 60
          ? 29
          : get(date(a[0], ctx)),
    true,
  );
define("TODAY", 0, 0, "date", (_, ctx) =>
  Math.floor(dateSerial(ctx.now, ctx.dateSystem)),
);
define("NOW", 0, 0, "date", (_, ctx) => dateSerial(ctx.now, ctx.dateSystem));
define(
  "DAYS",
  2,
  2,
  "date",
  (a) => Math.floor(number(a[0])) - Math.floor(number(a[1])),
  true,
);
define(
  "WEEKDAY",
  1,
  2,
  "date",
  (a, ctx) => {
    const day = date(a[0], ctx).getUTCDay(),
      type = number(a[1] ?? 1);
    if (type === 1) return day + 1;
    if (type === 2) return ((day + 6) % 7) + 1;
    if (type === 3) return (day + 6) % 7;
    if (type >= 11 && type <= 17)
      return ((day - ((type - 10) % 7) + 7) % 7) + 1;
    return error("#NUM!");
  },
  true,
);
function isoWeek(d: Date): number {
  const x = new Date(
    Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()),
  );
  x.setUTCDate(x.getUTCDate() + 4 - (x.getUTCDay() || 7));
  return Math.ceil(
    ((x.getTime() - Date.UTC(x.getUTCFullYear(), 0, 1)) / DAY + 1) / 7,
  );
}
define("ISOWEEKNUM", 1, 1, "date", (a, ctx) => isoWeek(date(a[0], ctx)), true);
define(
  "WEEKNUM",
  1,
  2,
  "date",
  (a, ctx) => {
    const d = date(a[0], ctx),
      type = number(a[1] ?? 1);
    if (type === 21) return isoWeek(d);
    const startDay =
      type === 1 || type === 17
        ? 0
        : type === 2
          ? 1
          : type >= 11 && type <= 16
            ? type - 10
            : -1;
    if (startDay < 0) return error("#NUM!");
    const first = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return (
      Math.floor(
        ((d.getTime() - first.getTime()) / DAY +
          ((first.getUTCDay() - startDay + 7) % 7)) /
          7,
      ) + 1
    );
  },
  true,
);
for (const name of ["EDATE", "EOMONTH"])
  define(
    name,
    2,
    2,
    "date",
    (a, ctx) => {
      const d = date(a[0], ctx),
        months = Math.trunc(number(a[1]));
      const end = new Date(
        Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + months + 1, 0),
      );
      if (name === "EDATE")
        end.setUTCDate(Math.min(d.getUTCDate(), end.getUTCDate()));
      return dateSerial(end, ctx.dateSystem);
    },
    true,
  );
define(
  "DATEVALUE",
  1,
  1,
  "date",
  (a, ctx) => {
    const s = text(a[0]);
    if (!/^\d{4}-\d{1,2}-\d{1,2}$/.test(s)) return error("#VALUE!");
    const d = new Date(s + "T00:00:00Z");
    return Number.isFinite(d.getTime())
      ? dateSerial(d, ctx.dateSystem)
      : error("#VALUE!");
  },
  true,
  "Accepts ISO YYYY-MM-DD; ambiguous locale-dependent strings are rejected.",
);
define(
  "TIMEVALUE",
  1,
  1,
  "date",
  (a) => {
    const match = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i.exec(
      text(a[0]),
    );
    if (!match) return error("#VALUE!");
    let h = +match[1];
    const m = +match[2],
      s = +(match[3] ?? 0);
    if (match[4]) h = (h % 12) + (match[4].toUpperCase() === "PM" ? 12 : 0);
    return h > 23 || m > 59 || s > 59
      ? error("#VALUE!")
      : (h * 3600 + m * 60 + s) / 86400;
  },
  true,
);
define("NETWORKDAYS", 2, 3, "date", (a, ctx) => {
  let start = Math.floor(number(a[0])),
    end = Math.floor(number(a[1])),
    sign = 1;
  if (start > end) {
    [start, end] = [end, start];
    sign = -1;
  }
  if (end - start > 1_000_000) return error("#NUM!");
  const holidays = new Set(
    [...flatten(a[2] === undefined ? [] : [a[2]])].map((v) =>
      Math.floor(number(v)),
    ),
  );
  let n = 0;
  for (let d = start; d <= end; d++) {
    const day = serialDate(d, ctx.dateSystem).getUTCDay();
    if (day !== 0 && day !== 6 && !holidays.has(d)) n++;
  }
  return sign * n;
});
define("WORKDAY", 2, 3, "date", (a, ctx) => {
  let d = Math.floor(number(a[0])),
    remaining = Math.abs(Math.trunc(number(a[1])));
  if (remaining > 1_000_000) return error("#NUM!");
  const sign = number(a[1]) < 0 ? -1 : 1,
    holidays = new Set(
      [...flatten(a[2] === undefined ? [] : [a[2]])].map((v) =>
        Math.floor(number(v)),
      ),
    );
  while (remaining) {
    d += sign;
    const day = serialDate(d, ctx.dateSystem).getUTCDay();
    if (day !== 0 && day !== 6 && !holidays.has(d)) remaining--;
  }
  return d;
});
define(
  "DAYS360",
  2,
  3,
  "date",
  (a, ctx) => {
    const start = date(a[0], ctx),
      end = date(a[1], ctx);
    let d1 = start.getUTCDate(),
      d2 = end.getUTCDate();
    const european = number(a[2] ?? 0) !== 0;
    if (european) {
      d1 = Math.min(d1, 30);
      d2 = Math.min(d2, 30);
    } else {
      const last = (d: Date) =>
        new Date(
          Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0),
        ).getUTCDate();
      if (d1 === last(start)) d1 = 30;
      if (d2 === last(end) && d1 >= 30) d2 = 30;
    }
    return (
      (end.getUTCFullYear() - start.getUTCFullYear()) * 360 +
      (end.getUTCMonth() - start.getUTCMonth()) * 30 +
      d2 -
      d1
    );
  },
  true,
);

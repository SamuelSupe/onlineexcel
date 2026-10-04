import { define } from "./registry";
import {
  dateSerial,
  DAY,
  error,
  flatten,
  number,
  serialDate,
  text,
  type Value,
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
function calendar(ctx: FunctionContext, holidayValues: Value | undefined) {
  const lastDate = dateSerial(new Date(Date.UTC(9999, 11, 31)), ctx.dateSystem);
  const day = (value: Value): number => {
    const serial = Math.floor(number(value));
    if (serial < 0 || serial > lastDate) throw error("#NUM!");
    return serial;
  };
  const weekdays = (start: number, end: number): number => {
    if (end < start) return 0;
    // Serial 60 repeats February 28 in the 1900 system; do not count across that discontinuity.
    if (ctx.dateSystem === 1900 && start < 60 && end >= 60)
      return weekdays(start, 59) + weekdays(60, end);
    const length = end - start + 1;
    let count = Math.floor(length / 7) * 5;
    const first = serialDate(start, ctx.dateSystem).getUTCDay();
    for (let i = 0; i < length % 7; i++) {
      const weekday = (first + i) % 7;
      if (weekday !== 0 && weekday !== 6) count++;
    }
    return count;
  };
  const holidays = new Set<number>();
  for (const value of flatten(
    holidayValues === undefined ? [] : [holidayValues],
  )) {
    const serial = day(value);
    if (weekdays(serial, serial)) holidays.add(serial);
  }
  const sorted = [...holidays].sort((a, b) => a - b);
  const before = (serial: number) => {
    let lo = 0,
      hi = sorted.length;
    while (lo < hi) {
      const mid = (lo + hi) >>> 1;
      if (sorted[mid] < serial) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  };
  return {
    day,
    lastDate,
    count: (start: number, end: number) =>
      end < start
        ? 0
        : weekdays(start, end) - (before(end + 1) - before(start)),
  };
}
define(
  "NETWORKDAYS",
  2,
  3,
  "date",
  (a, ctx) => {
    const dates = calendar(ctx, a[2]);
    let start = dates.day(a[0]),
      end = dates.day(a[1]),
      sign = 1;
    if (start > end) {
      [start, end] = [end, start];
      sign = -1;
    }
    return sign * dates.count(start, end);
  },
  false,
  "Dates and holidays must be within serial 0 through 9999-12-31 in the workbook date system; out-of-range dates return #NUM!.",
);
define(
  "WORKDAY",
  2,
  3,
  "date",
  (a, ctx) => {
    const dates = calendar(ctx, a[2]),
      start = dates.day(a[0]),
      days = Math.trunc(number(a[1])),
      remaining = Math.abs(days),
      sign = days < 0 ? -1 : 1;
    if (!remaining) return start;
    const count = (distance: number) =>
      sign > 0
        ? dates.count(start + 1, start + distance)
        : dates.count(start - distance, start - 1);
    let lo = 1,
      hi = sign > 0 ? dates.lastDate - start : start;
    if (count(hi) < remaining) return error("#NUM!");
    while (lo < hi) {
      const mid = Math.floor((lo + hi) / 2);
      if (count(mid) < remaining) lo = mid + 1;
      else hi = mid;
    }
    return start + sign * lo;
  },
  false,
  "Dates and holidays must be within serial 0 through 9999-12-31 in the workbook date system; out-of-range dates return #NUM!.",
);
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

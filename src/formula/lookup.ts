import { define } from "./registry";
import {
  array,
  bool,
  compare,
  error,
  flatten,
  isArray,
  isError,
  number,
  scalar,
  text,
  wildcardPattern,
} from "./values";
import { columnName } from "../core/address";
import type { Scalar } from "../core/types";
const sizeLimit = (rows: number, cols: number): boolean =>
  Number.isInteger(rows) &&
  Number.isInteger(cols) &&
  rows > 0 &&
  cols > 0 &&
  rows * cols <= 1_000_000;
define("ROW", 0, 1, "lookup", (a, ctx) =>
  a.length
    ? isArray(a[0]) && a[0].reference
      ? a[0].reference.row + 1
      : error("#VALUE!")
    : ctx.row + 1,
);
define("COLUMN", 0, 1, "lookup", (a, ctx) =>
  a.length
    ? isArray(a[0]) && a[0].reference
      ? a[0].reference.column + 1
      : error("#VALUE!")
    : ctx.column + 1,
);
define("ROWS", 1, 1, "lookup", (a) => array(a[0]).rows);
define("COLUMNS", 1, 1, "lookup", (a) => array(a[0]).columns);
define("ADDRESS", 2, 5, "lookup", (a) => {
  const r = Math.trunc(number(a[0])),
    c = Math.trunc(number(a[1])),
    type = number(a[2] ?? 1),
    a1 = bool(a[3] ?? true);
  if (r < 1 || c < 1 || r > 1048576 || c > 16384 || type < 1 || type > 4)
    return error("#VALUE!");
  const result = a1
    ? (type === 1 || type === 3 ? "$" : "") +
      columnName(c - 1) +
      (type <= 2 ? "$" : "") +
      r
    : `R${type <= 2 ? r : `[${r}]`}C${type === 1 || type === 3 ? c : `[${c}]`}`;
  return a[4] === undefined
    ? result
    : `'${text(a[4]).replace(/'/g, "''")}'!${result}`;
});
define("INDEX", 2, 3, "lookup", (a) => {
  const source = array(a[0]);
  let r = Math.trunc(number(a[1])),
    c = Math.trunc(number(a[2] ?? 1));
  if (a[2] === undefined && source.rows === 1) {
    c = r;
    r = 1;
  }
  if (r < 0 || c < 0 || r > source.rows || c > source.columns)
    return error("#REF!");
  if (!r || !c)
    return {
      rows: r ? 1 : source.rows,
      columns: c ? 1 : source.columns,
      get: (row: number, col: number) =>
        source.get(r ? r - 1 : row, c ? c - 1 : col),
    };
  return source.get(r - 1, c - 1);
});
function findIndex(
  needle: Scalar,
  values: Scalar[],
  mode: number,
  direction = 1,
): number {
  const indexes = values.map((_, i) => i);
  if (direction === -1) indexes.reverse();
  const pattern =
    mode === 2 && typeof needle === "string"
      ? wildcardPattern(needle)
      : undefined;
  const test = pattern
    ? (v: Scalar) =>
        (typeof v === "string" && pattern.test(v)) ||
        (v === null && needle === "")
    : (v: Scalar) => !isError(v) && compare(v, needle) === 0;
  for (const i of indexes) if (test(values[i])) return i;
  if (mode === 0 || mode === 2) return -1;
  let best = -1;
  for (const i of indexes) {
    if (isError(values[i])) continue;
    const order = compare(values[i], needle);
    if (
      (mode === -1 ? order < 0 : order > 0) &&
      (best === -1 ||
        (mode === -1
          ? compare(values[i], values[best]) > 0
          : compare(values[i], values[best]) < 0))
    )
      best = i;
  }
  return best;
}
define("MATCH", 2, 3, "lookup", (a) => {
  const source = array(a[1]),
    mode = number(a[2] ?? 1);
  if (source.rows > 1 && source.columns > 1) return error("#N/A");
  if (![-1, 0, 1].includes(mode)) return error("#N/A");
  const needle = scalar(a[0]),
    values = [...flatten([source])];
  const i = findIndex(
    needle,
    values,
    mode === 0 ? (typeof needle === "string" ? 2 : 0) : -mode,
  );
  return i < 0 ? error("#N/A") : i + 1;
});
define("XMATCH", 2, 4, "lookup", (a) => {
  const mode = number(a[2] ?? 0),
    direction = number(a[3] ?? 1),
    source = array(a[1]);
  if (
    ![-1, 0, 1, 2].includes(mode) ||
    ![-2, -1, 1, 2].includes(direction) ||
    (source.rows > 1 && source.columns > 1)
  )
    return error("#VALUE!");
  const i = findIndex(
    scalar(a[0]),
    [...flatten([source])],
    mode,
    direction < 0 ? -1 : 1,
  );
  return i < 0 ? error("#N/A") : i + 1;
});
define("XLOOKUP", 3, 6, "lookup", (a) => {
  const source = array(a[1]),
    target = array(a[2]),
    mode = number(a[4] ?? 0),
    direction = number(a[5] ?? 1);
  if (
    ![-1, 0, 1, 2].includes(mode) ||
    ![-2, -1, 1, 2].includes(direction) ||
    (source.rows > 1 && source.columns > 1)
  )
    return error("#VALUE!");
  const vertical = source.columns === 1;
  if (
    vertical ? source.rows !== target.rows : source.columns !== target.columns
  )
    return error("#VALUE!");
  const index = findIndex(
    scalar(a[0]),
    [...flatten([source])],
    mode,
    direction < 0 ? -1 : 1,
  );
  if (index < 0) return a[3] ?? error("#N/A");
  return {
    rows: vertical ? 1 : target.rows,
    columns: vertical ? target.columns : 1,
    get: (r: number, c: number) =>
      target.get(vertical ? index : r, vertical ? c : index),
  };
});
for (const name of ["VLOOKUP", "HLOOKUP"])
  define(name, 3, 4, "lookup", (a) => {
    const source = array(a[1]),
      index = Math.trunc(number(a[2])) - 1,
      vertical = name === "VLOOKUP";
    if (index < 0) return error("#VALUE!");
    if (index >= (vertical ? source.columns : source.rows))
      return error("#REF!");
    const values = Array.from(
      { length: vertical ? source.rows : source.columns },
      (_, i) => source.get(vertical ? i : 0, vertical ? 0 : i),
    );
    const needle = scalar(a[0]);
    const found = findIndex(
      needle,
      values,
      bool(a[3] ?? true) ? -1 : typeof needle === "string" ? 2 : 0,
    );
    return found < 0
      ? error("#N/A")
      : source.get(vertical ? found : index, vertical ? index : found);
  });
define("SEQUENCE", 1, 4, "array", (a) => {
  const rows = Math.trunc(number(a[0])),
    cols = Math.trunc(number(a[1] ?? 1)),
    start = number(a[2] ?? 1),
    step = number(a[3] ?? 1);
  return sizeLimit(rows, cols)
    ? {
        rows,
        columns: cols,
        get: (r: number, c: number) => start + (r * cols + c) * step,
      }
    : error("#NUM!");
});
define("TRANSPOSE", 1, 1, "array", (a) => {
  const source = array(a[0]);
  return {
    rows: source.columns,
    columns: source.rows,
    get: (r: number, c: number) => source.get(c, r),
  };
});
define("FILTER", 2, 3, "array", (a) => {
  const source = array(a[0]),
    include = array(a[1]),
    vertical = include.rows === source.rows && include.columns === 1;
  if (!vertical && !(include.rows === 1 && include.columns === source.columns))
    return error("#VALUE!");
  const indexes: number[] = [];
  for (let i = 0; i < (vertical ? source.rows : source.columns); i++)
    if (bool(include.get(vertical ? i : 0, vertical ? 0 : i))) indexes.push(i);
  if (!indexes.length) return a[2] ?? error("#CALC!");
  return {
    rows: vertical ? indexes.length : source.rows,
    columns: vertical ? source.columns : indexes.length,
    get: (r: number, c: number) =>
      source.get(vertical ? indexes[r] : r, vertical ? c : indexes[c]),
  };
});
define("SORT", 1, 4, "array", (a) => {
  const source = array(a[0]),
    index = Math.trunc(number(a[1] ?? 1)) - 1,
    order = number(a[2] ?? 1),
    byColumn = bool(a[3] ?? false);
  if (
    index < 0 ||
    index >= (byColumn ? source.rows : source.columns) ||
    ![-1, 1].includes(order)
  )
    return error("#VALUE!");
  const indexes = Array.from(
    { length: byColumn ? source.columns : source.rows },
    (_, i) => i,
  );
  indexes.sort(
    (x, y) =>
      order *
      compare(
        source.get(byColumn ? index : x, byColumn ? x : index),
        source.get(byColumn ? index : y, byColumn ? y : index),
      ),
  );
  return {
    rows: source.rows,
    columns: source.columns,
    get: (r: number, c: number) =>
      source.get(byColumn ? r : indexes[r], byColumn ? indexes[c] : c),
  };
});
define(
  "SORTBY",
  2,
  255,
  "array",
  (a) => {
    const source = array(a[0]),
      keys: { values: ReturnType<typeof array>; order: number }[] = [];
    for (let i = 1; i < a.length; i += 2) {
      const values = array(a[i]),
        order = number(a[i + 1] ?? 1);
      if (
        values.rows !== source.rows ||
        values.columns !== 1 ||
        ![-1, 1].includes(order)
      )
        return error("#VALUE!");
      keys.push({ values, order });
    }
    const indexes = Array.from({ length: source.rows }, (_, i) => i);
    indexes.sort((x, y) => {
      for (const { values, order } of keys) {
        const cmp = compare(values.get(x, 0), values.get(y, 0));
        if (cmp) return order * cmp;
      }
      return 0;
    });
    return {
      rows: source.rows,
      columns: source.columns,
      get: (r: number, c: number) => source.get(indexes[r], c),
    };
  },
  false,
  "Sort keys are vertical columns.",
);
define("UNIQUE", 1, 3, "array", (a) => {
  const source = array(a[0]),
    byColumn = bool(a[1] ?? false),
    exactlyOnce = bool(a[2] ?? false),
    groups = new Map<string, { index: number; count: number }>();
  for (let i = 0; i < (byColumn ? source.columns : source.rows); i++) {
    const values = Array.from(
      { length: byColumn ? source.rows : source.columns },
      (_, j) => {
        const v = source.get(byColumn ? j : i, byColumn ? i : j);
        return typeof v === "string" ? v.toUpperCase() : v;
      },
    );
    const key = JSON.stringify(values),
      group = groups.get(key);
    if (group) group.count++;
    else groups.set(key, { index: i, count: 1 });
  }
  const indexes = [...groups.values()]
    .filter((x) => !exactlyOnce || x.count === 1)
    .map((x) => x.index);
  return indexes.length
    ? {
        rows: byColumn ? source.rows : indexes.length,
        columns: byColumn ? indexes.length : source.columns,
        get: (r: number, c: number) =>
          source.get(byColumn ? r : indexes[r], byColumn ? indexes[c] : c),
      }
    : error("#CALC!");
});
for (const name of ["TAKE", "DROP"])
  define(name, 2, 3, "array", (a) => {
    const source = array(a[0]),
      rows = Math.trunc(number(a[1])),
      cols = a[2] === undefined ? undefined : Math.trunc(number(a[2]));
    const bounds = (n: number, count: number | undefined) =>
      count === undefined
        ? [0, n]
        : name === "TAKE"
          ? count >= 0
            ? [0, Math.min(n, count)]
            : [Math.max(0, n + count), n]
          : count >= 0
            ? [Math.min(n, count), n]
            : [0, Math.max(0, n + count)];
    const [r1, r2] = bounds(source.rows, rows),
      [c1, c2] = bounds(source.columns, cols);
    return r2 <= r1 || c2 <= c1
      ? error("#CALC!")
      : {
          rows: r2 - r1,
          columns: c2 - c1,
          get: (r: number, c: number) => source.get(r + r1, c + c1),
        };
  });
for (const name of ["CHOOSECOLS", "CHOOSEROWS"])
  define(name, 2, 255, "array", (a) => {
    const source = array(a[0]),
      columns = name === "CHOOSECOLS",
      length = columns ? source.columns : source.rows,
      indexes = [...flatten(a.slice(1))].map((v) => {
        const n = Math.trunc(number(v));
        return n < 0 ? length + n : n - 1;
      });
    if (indexes.some((i) => i < 0 || i >= length)) return error("#VALUE!");
    return {
      rows: columns ? source.rows : indexes.length,
      columns: columns ? indexes.length : source.columns,
      get: (r: number, c: number) =>
        source.get(columns ? r : indexes[r], columns ? indexes[c] : c),
    };
  });
for (const name of ["VSTACK", "HSTACK"])
  define(name, 1, 255, "array", (a) => {
    const sources = a.map(array),
      vertical = name === "VSTACK",
      rows = vertical
        ? sources.reduce((n, x) => n + x.rows, 0)
        : Math.max(...sources.map((x) => x.rows)),
      columns = vertical
        ? Math.max(...sources.map((x) => x.columns))
        : sources.reduce((n, x) => n + x.columns, 0);
    if (!sizeLimit(rows, columns)) return error("#NUM!");
    return {
      rows,
      columns,
      get: (r: number, c: number) => {
        for (const source of sources) {
          if (vertical ? r < source.rows : c < source.columns)
            return r < source.rows && c < source.columns
              ? source.get(r, c)
              : error("#N/A");
          if (vertical) r -= source.rows;
          else c -= source.columns;
        }
        return error("#N/A");
      },
    };
  });

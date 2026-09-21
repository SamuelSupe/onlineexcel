import { define } from "./registry";
import {
  array,
  criteria,
  error,
  finite,
  flatten,
  isArray,
  isError,
  number,
  numbers,
  type Value,
} from "./values";
const math = (name: string, fn: (n: number) => number) =>
  define(name, 1, 1, "math", (a) => finite(fn(number(a[0]))), true);
for (const [name, fn] of Object.entries({
  ABS: Math.abs,
  ACOS: Math.acos,
  ACOSH: Math.acosh,
  ASIN: Math.asin,
  ASINH: Math.asinh,
  ATAN: Math.atan,
  ATANH: Math.atanh,
  COS: Math.cos,
  COSH: Math.cosh,
  SIN: Math.sin,
  SINH: Math.sinh,
  TAN: Math.tan,
  TANH: Math.tanh,
  EXP: Math.exp,
  LN: Math.log,
  LOG10: Math.log10,
  SQRT: Math.sqrt,
  SIGN: Math.sign,
  INT: Math.floor,
}))
  math(name, fn);
math("DEGREES", (x) => (x * 180) / Math.PI);
math("RADIANS", (x) => (x * Math.PI) / 180);
math("EVEN", (x) => Math.sign(x) * Math.ceil(Math.abs(x) / 2) * 2);
math(
  "ODD",
  (x) => Math.sign(x || 1) * (Math.ceil((Math.abs(x) - 1) / 2) * 2 + 1),
);
math("FACT", (x) => {
  if (x < 0 || x > 170) return NaN;
  let v = 1;
  for (let n = 2; n <= Math.floor(x); n++) v *= n;
  return v;
});
define("PI", 0, 0, "math", () => Math.PI);
define(
  "POWER",
  2,
  2,
  "math",
  (a) => finite(number(a[0]) ** number(a[1])),
  true,
);
define(
  "ATAN2",
  2,
  2,
  "math",
  (a) =>
    number(a[0]) === 0 && number(a[1]) === 0
      ? error("#DIV/0!")
      : Math.atan2(number(a[1]), number(a[0])),
  true,
);
define(
  "LOG",
  1,
  2,
  "math",
  (a) =>
    finite(
      Math.log(number(a[0])) / Math.log(a[1] === undefined ? 10 : number(a[1])),
    ),
  true,
);
define(
  "MOD",
  2,
  2,
  "math",
  (a) => {
    const divisor = number(a[1]);
    return divisor
      ? number(a[0]) - divisor * Math.floor(number(a[0]) / divisor)
      : error("#DIV/0!");
  },
  true,
);
define(
  "QUOTIENT",
  2,
  2,
  "math",
  (a) =>
    number(a[1]) ? Math.trunc(number(a[0]) / number(a[1])) : error("#DIV/0!"),
  true,
);
for (const name of ["ROUND", "ROUNDUP", "ROUNDDOWN", "TRUNC"])
  define(
    name,
    name === "TRUNC" ? 1 : 2,
    2,
    "math",
    (a) => {
      const x = number(a[0]),
        digits = Math.trunc(number(a[1] ?? 0)),
        factor = 10 ** digits;
      const n = Math.abs(x) * factor;
      const rounded =
        name === "ROUND"
          ? Math.floor(n + 0.5 + Number.EPSILON * n)
          : name === "ROUNDUP"
            ? Math.ceil(n)
            : Math.floor(n);
      return finite((Math.sign(x) * rounded) / factor);
    },
    true,
  );
for (const name of ["CEILING.MATH", "FLOOR.MATH"])
  define(
    name,
    1,
    3,
    "math",
    (a) => {
      const n = number(a[0]),
        s = Math.abs(number(a[1] ?? 1)),
        mode = number(a[2] ?? 0);
      if (!s) return 0;
      const up = name === "CEILING.MATH";
      return (
        (n < 0 && mode !== 0
          ? up
            ? Math.floor(n / s)
            : Math.ceil(n / s)
          : up
            ? Math.ceil(n / s)
            : Math.floor(n / s)) * s
      );
    },
    true,
  );
define(
  "MROUND",
  2,
  2,
  "math",
  (a) => {
    const n = number(a[0]),
      m = number(a[1]);
    return n * m < 0
      ? error("#NUM!")
      : m
        ? Math.sign(n) * Math.floor(Math.abs(n / m) + 0.5) * Math.abs(m)
        : 0;
  },
  true,
);
define("RAND", 0, 0, "math", () => Math.random());
define("RANDBETWEEN", 2, 2, "math", (a) => {
  const lo = Math.ceil(number(a[0])),
    hi = Math.floor(number(a[1]));
  return lo > hi
    ? error("#NUM!")
    : lo + Math.floor(Math.random() * (hi - lo + 1));
});
const gcd = (a: number, b: number): number => {
  while (b) [a, b] = [b, a % b];
  return a;
};
for (const name of ["GCD", "LCM"])
  define(name, 1, 255, "math", (a) => {
    const ns = numbers(a).map(Math.floor);
    if (ns.some((n) => n < 0)) return error("#NUM!");
    return ns.reduce(
      (x, y) => (name === "GCD" ? gcd(x, y) : x && y ? (x / gcd(x, y)) * y : 0),
      name === "GCD" ? 0 : 1,
    );
  });
define(
  "COMBIN",
  2,
  2,
  "math",
  (a) => {
    const n = Math.floor(number(a[0]));
    let k = Math.floor(number(a[1]));
    if (n < 0 || k < 0 || k > n) return error("#NUM!");
    k = Math.min(k, n - k);
    let out = 1;
    for (let i = 1; i <= k; i++) out *= (n - i + 1) / i;
    return finite(out);
  },
  true,
);
const sum = (ns: number[]): number => ns.reduce((a, b) => a + b, 0);
const mean = (ns: number[]): number => sum(ns) / ns.length;
for (const name of [
  "SUM",
  "PRODUCT",
  "SUMSQ",
  "AVERAGE",
  "MIN",
  "MAX",
  "COUNT",
  "MEDIAN",
  "GEOMEAN",
  "HARMEAN",
  "AVEDEV",
  "DEVSQ",
])
  define(name, 1, 255, "statistics", (a) => {
    if (name === "COUNT") {
      let count = 0;
      for (const arg of a)
        for (const value of flatten([arg]))
          if (
            typeof value === "number" ||
            (!isArray(arg) &&
              (typeof value === "boolean" ||
                (typeof value === "string" &&
                  value.trim() !== "" &&
                  Number.isFinite(Number(value)))))
          )
            count++;
      return count;
    }
    const ns = numbers(a);
    if (name === "SUM") return sum(ns);
    if (name === "SUMSQ") return sum(ns.map((n) => n * n));
    if (name === "COUNT") return ns.length;
    if (name === "PRODUCT")
      return ns.length ? finite(ns.reduce((x, y) => x * y, 1)) : 0;
    if (name === "MIN" || name === "MAX")
      return ns.length
        ? ns.reduce((x, y) =>
            name === "MIN" ? Math.min(x, y) : Math.max(x, y),
          )
        : 0;
    if (!ns.length) return error("#DIV/0!");
    if (name === "AVERAGE") return mean(ns);
    if (name === "MEDIAN") {
      ns.sort((a, b) => a - b);
      const i = Math.floor(ns.length / 2);
      return ns.length % 2 ? ns[i] : (ns[i - 1] + ns[i]) / 2;
    }
    if (name === "GEOMEAN" || name === "HARMEAN")
      return ns.some((n) => n <= 0)
        ? error("#NUM!")
        : name === "GEOMEAN"
          ? finite(Math.exp(mean(ns.map(Math.log))))
          : ns.length / sum(ns.map((n) => 1 / n));
    const avg = mean(ns),
      deviations = ns.map((n) =>
        name === "AVEDEV" ? Math.abs(n - avg) : (n - avg) ** 2,
      );
    return name === "AVEDEV" ? mean(deviations) : sum(deviations);
  });
define(
  "COUNTA",
  1,
  255,
  "statistics",
  (a) => [...flatten(a)].filter((v) => v !== null).length,
);
define(
  "COUNTBLANK",
  1,
  1,
  "statistics",
  (a) => [...flatten(a)].filter((v) => v === null || v === "").length,
);
for (const name of ["VAR.S", "VAR.P", "STDEV.S", "STDEV.P"])
  define(name, 1, 255, "statistics", (a) => {
    const ns = numbers(a),
      d = ns.length - (name.endsWith(".S") ? 1 : 0);
    if (d <= 0) return error("#DIV/0!");
    const m = mean(ns),
      v = sum(ns.map((n) => (n - m) ** 2)) / d;
    return name.startsWith("STDEV") ? Math.sqrt(v) : v;
  });
for (const name of ["LARGE", "SMALL"])
  define(name, 2, 2, "statistics", (a) => {
    const ns = numbers([a[0]]).sort((x, y) =>
        name === "LARGE" ? y - x : x - y,
      ),
      k = Math.ceil(number(a[1]));
    return ns[k - 1] ?? error("#NUM!");
  });
for (const name of ["PERCENTILE.INC", "QUARTILE.INC"])
  define(name, 2, 2, "statistics", (a) => {
    const ns = numbers([a[0]]).sort((x, y) => x - y),
      p = number(a[1]) / (name === "QUARTILE.INC" ? 4 : 1);
    if (!ns.length || p < 0 || p > 1) return error("#NUM!");
    const index = p * (ns.length - 1),
      lo = Math.floor(index);
    return ns[lo] + (ns[Math.ceil(index)] - ns[lo]) * (index - lo);
  });
define("RANK.EQ", 2, 3, "statistics", (a) => {
  const n = number(a[0]),
    ns = numbers([a[1]]),
    asc = number(a[2] ?? 0) !== 0;
  return ns.includes(n)
    ? 1 + ns.filter((x) => (asc ? x < n : x > n)).length
    : error("#N/A");
});
define("MODE.SNGL", 1, 255, "statistics", (a) => {
  const counts = new Map<number, number>();
  for (const n of numbers(a)) counts.set(n, (counts.get(n) ?? 0) + 1);
  const entries = [...counts].sort((a, b) => b[1] - a[1] || a[0] - b[0]);
  return !entries.length || entries[0][1] === 1 ? error("#N/A") : entries[0][0];
});
function conditional(args: Value[], name: string): Value {
  const multiple = name.endsWith("IFS"),
    count = name.startsWith("COUNT"),
    average = name.startsWith("AVERAGE");
  const target = array(
    count ? args[0] : multiple ? args[0] : (args[2] ?? args[0]),
  );
  const pairs: [ReturnType<typeof array>, ReturnType<typeof criteria>][] = [];
  if (multiple) {
    const start = count ? 0 : 1;
    if ((args.length - start) % 2) return error("#VALUE!");
    for (let i = start; i < args.length; i += 2)
      pairs.push([array(args[i]), criteria(args[i + 1])]);
  } else pairs.push([array(args[0]), criteria(args[1])]);
  if (
    pairs.some(
      ([range]) =>
        range.rows !== target.rows || range.columns !== target.columns,
    )
  )
    return error("#VALUE!");
  let total = 0,
    matches = 0;
  for (let r = 0; r < target.rows; r++)
    for (let c = 0; c < target.columns; c++)
      if (pairs.every(([range, test]) => test(range.get(r, c)))) {
        const v = target.get(r, c);
        if (!count && isError(v)) return v;
        if (count) matches++;
        else if (typeof v === "number") {
          matches++;
          total += v;
        }
      }
  return count
    ? matches
    : average
      ? matches
        ? total / matches
        : error("#DIV/0!")
      : total;
}
for (const name of [
  "SUMIF",
  "SUMIFS",
  "COUNTIF",
  "COUNTIFS",
  "AVERAGEIF",
  "AVERAGEIFS",
])
  define(
    name,
    name.endsWith("IFS") && !name.startsWith("COUNT") ? 3 : 2,
    name.endsWith("IFS") ? 255 : name.startsWith("COUNT") ? 2 : 3,
    "statistics",
    (a) => conditional(a, name),
  );
define("SUMPRODUCT", 1, 255, "math", (a) => {
  const arrays = a.map(array),
    first = arrays[0];
  if (arrays.some((x) => x.rows !== first.rows || x.columns !== first.columns))
    return error("#VALUE!");
  let total = 0;
  for (let r = 0; r < first.rows; r++)
    for (let c = 0; c < first.columns; c++)
      total += arrays.reduce((p, v) => {
        const x = v.get(r, c);
        if (isError(x)) throw x;
        return p * (typeof x === "number" ? x : 0);
      }, 1);
  return total;
});
for (const name of [
  "CORREL",
  "COVARIANCE.P",
  "COVARIANCE.S",
  "SLOPE",
  "INTERCEPT",
  "RSQ",
])
  define(name, 2, 2, "statistics", (a) => {
    const aa = [...flatten([a[0]])],
      bb = [...flatten([a[1]])];
    if (aa.length !== bb.length) return error("#N/A");
    const x: number[] = [],
      y: number[] = [];
    for (let i = 0; i < aa.length; i++) {
      if (isError(aa[i])) return aa[i];
      if (isError(bb[i])) return bb[i];
      if (typeof aa[i] === "number" && typeof bb[i] === "number") {
        y.push(aa[i] as number);
        x.push(bb[i] as number);
      }
    }
    if (!x.length) return error("#DIV/0!");
    const mx = mean(x),
      my = mean(y);
    let xy = 0,
      xx = 0,
      yy = 0;
    for (let i = 0; i < x.length; i++) {
      xy += (x[i] - mx) * (y[i] - my);
      xx += (x[i] - mx) ** 2;
      yy += (y[i] - my) ** 2;
    }
    if (name.startsWith("COVARIANCE"))
      return x.length > (name.endsWith(".S") ? 1 : 0)
        ? xy / (x.length - (name.endsWith(".S") ? 1 : 0))
        : error("#DIV/0!");
    if (!xx || (!yy && ["CORREL", "RSQ"].includes(name)))
      return error("#DIV/0!");
    const correlation = xy / Math.sqrt(xx * yy);
    return name === "SLOPE"
      ? xy / xx
      : name === "INTERCEPT"
        ? my - (xy / xx) * mx
        : name === "RSQ"
          ? correlation ** 2
          : correlation;
  });

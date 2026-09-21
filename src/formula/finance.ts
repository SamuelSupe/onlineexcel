import { define } from "./registry";
import { error, finite, flatten, number, numbers } from "./values";
const payment = (
  rate: number,
  periods: number,
  present: number,
  future: number,
  type: number,
): number =>
  rate === 0
    ? -(present + future) / periods
    : (-(present * (1 + rate) ** periods + future) * rate) /
      ((1 + type * rate) * ((1 + rate) ** periods - 1));
define(
  "PMT",
  3,
  5,
  "financial",
  (a) =>
    finite(
      payment(
        number(a[0]),
        number(a[1]),
        number(a[2]),
        number(a[3] ?? 0),
        number(a[4] ?? 0),
      ),
    ),
  true,
);
define(
  "FV",
  3,
  5,
  "financial",
  (a) => {
    const r = number(a[0]),
      n = number(a[1]),
      p = number(a[2]),
      pv = number(a[3] ?? 0),
      type = number(a[4] ?? 0);
    return finite(
      r === 0
        ? -(pv + p * n)
        : -(pv * (1 + r) ** n + (p * (1 + type * r) * ((1 + r) ** n - 1)) / r),
    );
  },
  true,
);
define(
  "PV",
  3,
  5,
  "financial",
  (a) => {
    const r = number(a[0]),
      n = number(a[1]),
      p = number(a[2]),
      fv = number(a[3] ?? 0),
      type = number(a[4] ?? 0);
    return finite(
      r === 0
        ? -(fv + p * n)
        : -(fv + (p * (1 + type * r) * ((1 + r) ** n - 1)) / r) / (1 + r) ** n,
    );
  },
  true,
);
define(
  "NPER",
  3,
  5,
  "financial",
  (a) => {
    const r = number(a[0]),
      p = number(a[1]),
      pv = number(a[2]),
      fv = number(a[3] ?? 0),
      type = number(a[4] ?? 0);
    return finite(
      r === 0
        ? -(pv + fv) / p
        : Math.log(
            (p * (1 + r * type) - fv * r) / (pv * r + p * (1 + r * type)),
          ) / Math.log(1 + r),
    );
  },
  true,
);
define("NPV", 2, 255, "financial", (a) => {
  const rate = number(a[0]);
  return finite(
    numbers(a.slice(1)).reduce(
      (sum, v, i) => sum + v / (1 + rate) ** (i + 1),
      0,
    ),
  );
});
function root(fn: (rate: number) => number, guess: number): number {
  let x = guess;
  for (let i = 0; i < 100; i++) {
    const y = fn(x);
    if (Math.abs(y) < 1e-8) return x;
    const h = Math.max(1e-7, Math.abs(x) * 1e-5),
      derivative = (fn(x + h) - fn(x - h)) / (2 * h);
    if (!Number.isFinite(derivative) || Math.abs(derivative) < 1e-14) break;
    const next = x - y / derivative;
    if (!Number.isFinite(next) || next <= -1) {
      x = (x - 1) / 2;
      continue;
    }
    if (Math.abs(next - x) < 1e-12) return next;
    x = next;
  }
  return NaN;
}
define("IRR", 1, 2, "financial", (a) => {
  const ns = numbers([a[0]]);
  if (!ns.some((x) => x > 0) || !ns.some((x) => x < 0)) return error("#NUM!");
  return finite(
    root(
      (rate) => ns.reduce((sum, v, i) => sum + v / (1 + rate) ** i, 0),
      number(a[1] ?? 0.1),
    ),
  );
});
define("XNPV", 3, 3, "financial", (a) => {
  const rate = number(a[0]),
    values = [...flatten([a[1]])].map(number),
    dates = [...flatten([a[2]])].map(number);
  if (
    !values.length ||
    values.length !== dates.length ||
    dates.some((d) => d < dates[0])
  )
    return error("#NUM!");
  return finite(
    values.reduce(
      (sum, v, i) => sum + v / (1 + rate) ** ((dates[i] - dates[0]) / 365),
      0,
    ),
  );
});
define("XIRR", 2, 3, "financial", (a) => {
  const values = [...flatten([a[0]])].map(number),
    dates = [...flatten([a[1]])].map(number);
  if (
    !values.length ||
    values.length !== dates.length ||
    dates.some((d) => d < dates[0]) ||
    !values.some((x) => x > 0) ||
    !values.some((x) => x < 0)
  )
    return error("#NUM!");
  return finite(
    root(
      (rate) =>
        values.reduce(
          (sum, v, i) => sum + v / (1 + rate) ** ((dates[i] - dates[0]) / 365),
          0,
        ),
      number(a[2] ?? 0.1),
    ),
  );
});
define("RATE", 3, 6, "financial", (a) => {
  const n = number(a[0]),
    pmt = number(a[1]),
    pv = number(a[2]),
    fv = number(a[3] ?? 0),
    type = number(a[4] ?? 0);
  return finite(
    root(
      (r) =>
        r === 0
          ? pv + pmt * n + fv
          : pv * (1 + r) ** n +
            (pmt * (1 + r * type) * ((1 + r) ** n - 1)) / r +
            fv,
      number(a[5] ?? 0.1),
    ),
  );
});
define(
  "SLN",
  3,
  3,
  "financial",
  (a) =>
    number(a[2])
      ? (number(a[0]) - number(a[1])) / number(a[2])
      : error("#DIV/0!"),
  true,
);
define(
  "SYD",
  4,
  4,
  "financial",
  (a) => {
    const life = number(a[2]),
      period = number(a[3]);
    return life <= 0 || period <= 0 || period > life
      ? error("#NUM!")
      : ((number(a[0]) - number(a[1])) * (life - period + 1) * 2) /
          (life * (life + 1));
  },
  true,
);
define(
  "EFFECT",
  2,
  2,
  "financial",
  (a) => {
    const r = number(a[0]),
      n = Math.floor(number(a[1]));
    return r <= 0 || n < 1 ? error("#NUM!") : (1 + r / n) ** n - 1;
  },
  true,
);
define(
  "NOMINAL",
  2,
  2,
  "financial",
  (a) => {
    const r = number(a[0]),
      n = Math.floor(number(a[1]));
    return r <= 0 || n < 1 ? error("#NUM!") : ((1 + r) ** (1 / n) - 1) * n;
  },
  true,
);

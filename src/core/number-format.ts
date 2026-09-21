const colors: Record<string, string> = {
  black: "#000000",
  blue: "#0000ff",
  cyan: "#00ffff",
  green: "#008000",
  magenta: "#ff00ff",
  red: "#ff0000",
  white: "#ffffff",
  yellow: "#ffff00",
};
export function formatSections(pattern: string): string[] {
  const sections: string[] = [];
  let start = 0,
    quoted = false;
  for (let i = 0; i < pattern.length; i++) {
    if (pattern[i] === "\\") {
      i++;
      continue;
    }
    if (pattern[i] === '"') quoted = !quoted;
    if (pattern[i] === ";" && !quoted) {
      sections.push(pattern.slice(start, i));
      start = i + 1;
    }
  }
  sections.push(pattern.slice(start));
  return sections;
}
export function numberSection(
  pattern: string,
  value: number,
): { section: string; absolute: boolean; color?: string } {
  const sections = formatSections(pattern),
    hasConditions = sections.some((s) => /\[(?:[<>]=?|=|<>)-?[\d.]+\]/.test(s));
  let section: string;
  if (hasConditions) {
    section =
      sections.find((s) => {
        const m = /\[(<=|>=|<>|<|>|=)(-?[\d.]+)\]/.exec(s);
        if (!m) return true;
        const n = Number(m[2]);
        return m[1] === "<="
          ? value <= n
          : m[1] === ">="
            ? value >= n
            : m[1] === "<>"
              ? value !== n
              : m[1] === "<"
                ? value < n
                : m[1] === ">"
                  ? value > n
                  : value === n;
      }) ?? "";
  } else
    section =
      sections[
        value === 0 && sections.length >= 3
          ? 2
          : value < 0 && sections.length >= 2
            ? 1
            : 0
      ];
  const color = /\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow)\]/i.exec(
    section,
  )?.[1];
  section = section
    .replace(
      /\[(?:Black|Blue|Cyan|Green|Magenta|Red|White|Yellow|Color\d+|(?:[<>]=?|=|<>)-?[\d.]+)\]/gi,
      "",
    )
    .replace(/\[\$([^\]-]*)-?[\da-f]*\]/gi, (_, currency) =>
      currency ? `"${currency}"` : "",
    )
    .replace(/_./g, " ")
    .replace(/\*./g, "");
  return {
    section,
    absolute: !hasConditions && sections.length > 1 && value < 0,
    color: color ? colors[color.toLowerCase()] : undefined,
  };
}
export function formatColor(value: unknown, pattern?: string) {
  return typeof value === "number" && pattern
    ? numberSection(pattern, value).color
    : undefined;
}
export function fractionText(
  value: number,
  pattern: string,
): string | undefined {
  const match = /([#0?]+)?\s*([?0#]+)\/([?0#]+|\d+)/.exec(pattern);
  if (!match) return;
  const mixed = !!match[1],
    whole = mixed ? Math.floor(value) : 0,
    remainder = value - whole;
  let denominator =
      /^[0-9]+$/.test(match[3]) && /[1-9]/.test(match[3])
        ? Number(match[3])
        : 0,
    numerator = 0;
  if (denominator) numerator = Math.round(remainder * denominator);
  else {
    let error = Infinity;
    const max = Math.min(9999, 10 ** match[3].length - 1);
    for (let d = 1; d <= max; d++) {
      const n = Math.round(remainder * d),
        e = Math.abs(remainder - n / d);
      if (e < error) {
        error = e;
        denominator = d;
        numerator = n;
      }
      if (e < 1e-12) break;
    }
  }
  const nextWhole = whole + (mixed ? Math.floor(numerator / denominator) : 0);
  if (mixed) numerator %= denominator;
  const result = mixed
    ? `${nextWhole || ""}${nextWhole && numerator ? " " : ""}${numerator ? `${numerator}/${denominator}` : ""}` ||
      "0"
    : `${numerator}/${denominator}`;
  return (
    pattern.slice(0, match.index).replace(/"/g, "") +
    result +
    pattern.slice(match.index + match[0].length).replace(/"/g, "")
  );
}

function insideText(value: string, caret: number): boolean {
  let quoted = false;
  for (let i = 1; i < caret; i++) {
    if (value[i] !== '"') continue;
    if (quoted && value[i + 1] === '"') i++;
    else quoted = !quoted;
  }
  return quoted;
}

export function referenceSpan(
  value: string,
  start: number,
  end = start,
): { start: number; end: number } | undefined {
  if (!value.startsWith("=") || insideText(value, start)) return;
  const pattern =
    /(?<![\p{L}\p{N}_.])(?:(?:'(?:[^']|'')*'|[\p{L}_][\p{L}\p{N}_.]*)!)?\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?(?![\p{L}\p{N}_.(])/giu;
  for (const match of value.matchAll(pattern)) {
    const left = match.index,
      right = left + match[0].length;
    if (start >= left && end <= right && !insideText(value, left))
      return { start: left, end: right };
  }
}

export function referenceInsertion(
  value: string,
  start: number,
  end: number,
): { start: number; end: number } | undefined {
  if (!value.startsWith("=") || start < 1 || insideText(value, start)) return;
  const span = referenceSpan(value, start, end);
  if (span) return span;
  if (start !== end || /[=+\-*/^&(<>,;]\s*$/.test(value.slice(0, start)))
    return { start, end };
}

export function cycleReference(
  value: string,
  caret: number,
): { value: string; caret: number } | undefined {
  const span = referenceSpan(value, caret);
  if (!span) return;
  const ref = value.slice(span.start, span.end),
    separator = ref.lastIndexOf("!"),
    prefix = ref.slice(0, separator + 1),
    cells = ref.slice(separator + 1);
  const first = /^(\$?)[A-Z]+(\$?)/i.exec(cells)!;
  const mode = first[1] ? (first[2] ? 1 : 3) : first[2] ? 2 : 0;
  const next = cells.replace(
    /\$?([A-Z]+)\$?(\d+)/gi,
    (_, col, row) =>
      `${mode === 0 || mode === 2 ? "$" : ""}${col}${mode === 0 || mode === 1 ? "$" : ""}${row}`,
  );
  const replacement = prefix + next;
  return {
    value: value.slice(0, span.start) + replacement + value.slice(span.end),
    caret: span.start + replacement.length,
  };
}

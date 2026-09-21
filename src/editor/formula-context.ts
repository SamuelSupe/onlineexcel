export interface FormulaContext {
  call?: { name: string; argument: number };
  token?: { text: string; start: number; end: number };
}

/** Reads an unfinished formula at the caret without evaluating it. */
export function formulaContext(value: string, caret: number): FormulaContext {
  if (!value.startsWith("=") || caret < 1) return {};
  const prefix = value.slice(0, caret);
  const stack: { kind: string; name?: string; argument: number }[] = [];
  let quote = "";
  for (let i = 1; i < prefix.length; i++) {
    const char = prefix[i];
    if (quote) {
      if (char === quote) {
        if (prefix[i + 1] === quote) i++;
        else quote = "";
      }
      continue;
    }
    if (char === '"' || char === "'") quote = char;
    else if (char === "(" || char === "{") {
      stack.push({
        kind: char,
        name:
          char === "("
            ? /([A-Z_][A-Z0-9_.]*)\s*$/i
                .exec(prefix.slice(0, i))?.[1]
                .toUpperCase()
            : undefined,
        argument: 0,
      });
    } else if (char === ")" || char === "}") stack.pop();
    else if (char === "," && stack.at(-1)?.kind === "(")
      stack.at(-1)!.argument++;
  }
  const frame = [...stack].reverse().find((item) => item.name);
  const context: FormulaContext = frame?.name
    ? { call: { name: frame.name, argument: frame.argument } }
    : {};
  // Quoted text and sheet names must never be replaced by a function completion.
  if (quote) return context;
  const match = /[A-Z_][A-Z0-9_.]*$/i.exec(prefix);
  if (match && /[=+\-*/^&(<>,;{]\s*$/.test(prefix.slice(0, match.index))) {
    const end =
      caret + (/^[A-Z0-9_.]*/i.exec(value.slice(caret))?.[0].length ?? 0);
    if (!/^\s*!/.test(value.slice(end)))
      context.token = { text: match[0].toUpperCase(), start: match.index, end };
  } else if (/^=\s*$/.test(prefix)) {
    context.token = { text: "", start: caret, end: caret };
  }
  return context;
}

export function completeFunction(
  value: string,
  token: NonNullable<FormulaContext["token"]>,
  name: string,
  noArguments: boolean,
): { value: string; caret: number } {
  const suffix = value.slice(token.end);
  const open = /^\s*\(/.exec(suffix);
  const inserted = name + (open ? "" : noArguments ? "()" : "(");
  return {
    value: value.slice(0, token.start) + inserted + suffix,
    caret: token.start + inserted.length + (open?.[0].length ?? 0),
  };
}

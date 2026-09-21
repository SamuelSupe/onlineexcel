import { listFunctions } from "../formula/functions";
import { element } from "./dom";
import { localeMessages, type Locale } from "./locale";
import {
  completeFunction,
  formulaContext,
  type FormulaContext,
} from "./formula-context";

type FormulaInput = HTMLInputElement | HTMLTextAreaElement;

const popular = ["SUM", "AVERAGE", "IF", "COUNT", "XLOOKUP", "FILTER"];
export function createFormulaHelp(
  shell: HTMLElement,
  inputs: FormulaInput[],
  locale: Locale,
  signal: AbortSignal,
  onComplete: (source: FormulaInput) => void,
) {
  let catalog = listFunctions();
  const {
    labels: t,
    functions: descriptions,
    categories,
  } = localeMessages(locale);
  const panel = element("div", "formula-help");
  const title = element("div", "formula-help-title", "ƒx  " + t.formulaMode);
  const guidance = element(
    "div",
    "formula-help-guidance",
    t.formulaGuide + " " + t.formulaPickHint,
  );
  const list = element("div", "formula-suggestions");
  const detail = element("div", "formula-help-detail");
  const footer = element("div", "formula-help-footer");
  panel.hidden = true;
  guidance.id = "formula-guide-" + crypto.randomUUID();
  list.id = "formula-list-" + crypto.randomUUID();
  list.setAttribute("role", "listbox");
  list.setAttribute("aria-label", t.functionSuggestions);
  panel.append(title, guidance, list, detail, footer);
  shell.append(panel);
  let source: FormulaInput | undefined;
  let context: FormulaContext = {};
  let matches: typeof catalog = [];
  let selected = 0;
  const composing = new Set<FormulaInput>();

  function describe(item: (typeof catalog)[number]) {
    return (
      descriptions[item.name as keyof typeof descriptions] ??
      item.notes ??
      (categories[item.category as keyof typeof categories] ?? item.category) +
        t.functionSuffix
    );
  }
  function hide() {
    panel.hidden = true;
    if (source) {
      source.removeAttribute("aria-activedescendant");
      source.removeAttribute("aria-describedby");
      source.removeAttribute("aria-controls");
      source.removeAttribute("aria-autocomplete");
    }
    source = undefined;
    matches = [];
  }
  function position() {
    if (panel.hidden || !source) return;
    const bounds = shell.getBoundingClientRect();
    const anchor = source.getBoundingClientRect();
    panel.style.width =
      Math.min(410, Math.max(0, shell.clientWidth - 16)) + "px";
    panel.style.left =
      Math.max(
        8,
        Math.min(
          anchor.left - bounds.left,
          shell.clientWidth - panel.offsetWidth - 8,
        ),
      ) + "px";
    const below = shell.clientHeight - (anchor.bottom - bounds.top) - 8;
    const above = anchor.top - bounds.top - 8;
    const flip = below < panel.scrollHeight && above > below;
    panel.style.maxHeight = Math.max(0, flip ? above : below) + "px";
    panel.style.top =
      (flip
        ? anchor.top - bounds.top - panel.offsetHeight - 4
        : anchor.bottom - bounds.top + 4) + "px";
  }
  function showDetail() {
    detail.replaceChildren();
    const item =
      matches[selected] ??
      catalog.find((item) => item.name === context.call?.name);
    if (item) {
      const signature = element("div", "formula-signature");
      signature.append(item.name + "(");
      const args = item.signature ? item.signature.split(", ") : [];
      args.forEach((argument, index) => {
        if (index) signature.append(", ");
        const current =
          !matches.length &&
          context.call &&
          index === Math.min(context.call.argument, args.length - 1);
        signature.append(element(current ? "strong" : "span", "", argument));
      });
      signature.append(")");
      detail.append(
        signature,
        element("div", "formula-description", describe(item)),
      );
    } else if (context.call) {
      detail.append(
        element(
          "div",
          "formula-description",
          context.call.name + ": " + t.unknownFunction,
        ),
      );
    }
    detail.hidden = !detail.childNodes.length;
    footer.textContent = matches.length ? t.formulaKeys : t.formulaArguments;
    position();
  }
  function highlight() {
    for (const [index, option] of Array.from(list.children).entries()) {
      option.setAttribute("aria-selected", String(index === selected));
      if (index === selected)
        source?.setAttribute("aria-activedescendant", option.id);
    }
    showDetail();
  }
  function update(target: FormulaInput) {
    if (
      target.readOnly ||
      target.hidden ||
      !target.value.startsWith("=") ||
      composing.has(target)
    ) {
      hide();
      return;
    }
    if (source !== target) hide();
    source = target;
    context = formulaContext(
      target.value,
      target.selectionStart ?? target.value.length,
    );
    matches =
      target.selectionStart !== target.selectionEnd || !context.token
        ? []
        : context.token.text
          ? catalog
              .filter((item) => item.name.startsWith(context.token!.text))
              .sort((a, b) => {
                const rank = (name: string) =>
                  name === context.token!.text
                    ? -1
                    : popular.includes(name)
                      ? popular.indexOf(name)
                      : popular.length;
                return rank(a.name) - rank(b.name);
              })
              .slice(0, 6)
          : popular.map((name) => catalog.find((item) => item.name === name)!);
    selected = 0;
    list.replaceChildren();
    source.removeAttribute("aria-activedescendant");
    source.setAttribute("aria-describedby", guidance.id);
    source.setAttribute("aria-controls", list.id);
    source.setAttribute("aria-autocomplete", "list");
    matches.forEach((item, index) => {
      const option = element("button", "formula-option");
      option.type = "button";
      option.tabIndex = -1;
      option.id = list.id + "-" + index;
      option.setAttribute("role", "option");
      option.append(
        element("span", "formula-option-name", item.name),
        element("span", "", describe(item)),
      );
      option.addEventListener("pointerdown", (event) => event.preventDefault());
      option.addEventListener("click", () => {
        selected = index;
        accept();
      });
      list.append(option);
    });
    list.hidden = !matches.length;
    panel.hidden = false;
    highlight();
  }
  function accept() {
    const item = matches[selected];
    if (!source || !context.token || !item) return;
    const result = completeFunction(
      source.value,
      context.token,
      item.name,
      item.maxArgs === 0,
    );
    source.value = result.value;
    source.setSelectionRange(result.caret, result.caret);
    onComplete(source);
    update(source);
  }
  function keydown(event: KeyboardEvent) {
    if (
      panel.hidden ||
      event.isComposing ||
      event.keyCode === 229 ||
      !source ||
      composing.has(source)
    )
      return false;
    if (
      event.ctrlKey ||
      event.metaKey ||
      event.altKey ||
      event.shiftKey ||
      !matches.length
    )
      return false;
    if (event.key === "Tab") accept();
    else if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      selected =
        (selected + (event.key === "ArrowDown" ? 1 : -1) + matches.length) %
        matches.length;
      highlight();
    } else return false;
    event.preventDefault();
    return true;
  }
  for (const target of inputs) {
    for (const type of ["input", "focus", "click"])
      target.addEventListener(type, () => update(target), { signal });
    target.addEventListener(
      "keyup",
      (event) => {
        if (
          ["ArrowLeft", "ArrowRight", "Home", "End"].includes(
            (event as KeyboardEvent).key,
          )
        )
          update(target);
      },
      { signal },
    );
    target.addEventListener("blur", hide, { signal });
    target.addEventListener(
      "compositionstart",
      () => {
        composing.add(target);
        hide();
      },
      { signal },
    );
    target.addEventListener(
      "compositionend",
      () => {
        composing.delete(target);
        update(target);
      },
      { signal },
    );
  }
  signal.addEventListener(
    "abort",
    () => {
      hide();
      panel.remove();
    },
    { once: true },
  );
  return {
    update,
    hide,
    keydown,
    position,
    setCatalog: (value: ReturnType<typeof listFunctions>) => {
      catalog = value;
    },
  };
}

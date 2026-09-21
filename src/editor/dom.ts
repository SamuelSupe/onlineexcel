export function element<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className = "",
  text = "",
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  node.className = className;
  node.textContent = text;
  return node;
}
export interface Field {
  key: string;
  label: string;
  value?: string;
  type?: string;
  options?: { value: string; label: string }[];
}
export async function dialog(
  parent: HTMLElement,
  title: string,
  fields: Field[],
  labels: { cancel: string; confirm: string },
  options: { message?: string; confirm?: string; signal: AbortSignal },
): Promise<Record<string, string> | null> {
  return formDialog(parent, title, labels, options, (form) => {
    for (const field of fields) {
      const input = field.options ? element("select") : element("input");
      input.name = field.key;
      if (input instanceof HTMLSelectElement)
        for (const option of field.options ?? []) {
          const item = element("option", "", option.label);
          item.value = option.value;
          input.append(item);
        }
      else input.type = field.type ?? "text";
      input.value = field.value ?? field.options?.[0]?.value ?? "";
      if (input instanceof HTMLInputElement && input.type === "checkbox")
        input.checked = field.value === "true";
      labeledControl(form, field.label, input);
    }
    return () =>
      Object.fromEntries(
        fields.map((field) => {
          const input = form.elements.namedItem(field.key) as
            | HTMLInputElement
            | HTMLSelectElement;
          return [
            field.key,
            input instanceof HTMLInputElement && input.type === "checkbox"
              ? String(input.checked)
              : input.value,
          ];
        }),
      );
  });
}

export function labeledControl(
  parent: HTMLElement,
  text: string,
  control: HTMLElement,
) {
  const label = element("label", "", text);
  control.id ||= crypto.randomUUID();
  label.htmlFor = control.id;
  parent.append(label, control);
  return control;
}

export async function formDialog<T>(
  parent: HTMLElement,
  title: string,
  labels: { cancel: string; confirm: string },
  options: { message?: string; confirm?: string; signal: AbortSignal },
  build: (form: HTMLFormElement) => () => T | Promise<T>,
): Promise<T | null> {
  if (options.signal.aborted) return null;
  const overlay = element("div", "overlay"),
    form = element("form", "dialog");
  form.setAttribute("role", "dialog");
  form.setAttribute("aria-modal", "true");
  const heading = element("h2", "", title);
  heading.id = crypto.randomUUID();
  form.setAttribute("aria-labelledby", heading.id);
  form.append(heading);
  if (options.message) form.append(element("pre", "", options.message));
  const read = build(form);
  const failure = element("div", "dialog-error");
  failure.setAttribute("role", "alert");
  form.append(failure);
  const footer = element("footer"),
    cancel = element("button", "", labels.cancel),
    confirm = element("button", "primary", options.confirm ?? labels.confirm);
  cancel.type = "button";
  confirm.type = "submit";
  footer.append(cancel, confirm);
  form.append(footer);
  overlay.append(form);
  parent.append(overlay);
  return new Promise((resolve) => {
    let closed = false;
    const close = (value: T | null) => {
      if (closed) return;
      closed = true;
      options.signal.removeEventListener("abort", abort);
      overlay.remove();
      resolve(value);
    };
    const abort = () => close(null);
    options.signal.addEventListener("abort", abort, { once: true });
    cancel.onclick = () => close(null);
    form.onsubmit = async (event) => {
      event.preventDefault();
      confirm.disabled = true;
      try {
        close(await read());
      } catch (error) {
        failure.textContent = (error as Error).message;
      } finally {
        confirm.disabled = false;
      }
    };
    form.onkeydown = (event) => {
      if (event.key === "Escape") {
        event.stopPropagation();
        close(null);
      }
      if (event.key === "Tab") {
        const items = [
            ...form.querySelectorAll<HTMLElement>("input,select,button"),
          ],
          first = items[0],
          last = items[items.length - 1];
        if (
          event.shiftKey &&
          form.getRootNode() instanceof ShadowRoot &&
          (form.getRootNode() as ShadowRoot).activeElement === first
        ) {
          event.preventDefault();
          last.focus();
        } else if (
          !event.shiftKey &&
          (form.getRootNode() as ShadowRoot).activeElement === last
        ) {
          event.preventDefault();
          first.focus();
        }
      }
    };
    (
      (form.querySelector("input,select") as HTMLElement | null) ?? confirm
    ).focus();
  });
}
export function download(data: BlobPart, name: string, type: string): void {
  const url = URL.createObjectURL(new Blob([data], { type })),
    anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = name;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

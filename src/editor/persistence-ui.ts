import type { WorkbookPersistence } from "../runtime/persistence";
import { element } from "./dom";
import { icon } from "./icons";
import { errorMessage, labels, type Locale } from "./locale";

export function persistenceUI(
  parent: HTMLElement,
  persistence: WorkbookPersistence,
  locale: Locale,
  run: (task: () => unknown) => Promise<void>,
  draft: () => boolean,
) {
  const t = labels(locale),
    panel = element("div", "persistence-bar");
  const status = element("span"),
    detail = element("span", "persistence-detail");
  status.setAttribute("role", "status");
  const button = (text: string, action: () => unknown) => {
    const node = element("button", "tool-button");
    node.type = "button";
    node.append(icon("saveNow"), element("span", "", text));
    node.onclick = () => void run(action);
    return node;
  };
  const save = button(t.saveNow, () => persistence.save());
  const retry = button(t.retrySave, () => persistence.retry());
  const restore = button(t.restoreDraft, () => persistence.restore());
  const keep = button(t.keepCurrent, () => persistence.discardRecovery());
  panel.append(status, detail, save, retry, restore, keep);
  parent.append(panel);
  const update = () => {
    const state = persistence.getState();
    const recovery = state.recoverySavedAt !== undefined;
    const busy = ["loading", "saving", "retrying"].includes(state.status);
    status.textContent = recovery
      ? t.recoveryTitle
      : {
          loading: t.loading,
          recovery: t.recoveryTitle,
          saved: t.saved,
          dirty: t.unsaved,
          saving: t.saving,
          retrying: t.saveRetrying,
          error: t.saveFailed,
          disposed: t.cancelled,
        }[state.status];
    if (!recovery && state.status === "saved" && draft())
      status.textContent = t.unsaved;
    detail.textContent = recovery
      ? t.recoveryHint
      : state.error
        ? errorMessage(state.error, locale)
        : state.savedAt
          ? `${t.lastSaved}: ${new Date(state.savedAt).toLocaleString(locale)}`
          : "";
    if (recovery)
      detail.title = new Date(state.recoverySavedAt!).toLocaleString(locale);
    else detail.removeAttribute("title");
    save.hidden = recovery;
    save.disabled = busy || state.status === "disposed";
    retry.hidden = state.status !== "error" || recovery;
    restore.hidden = keep.hidden = !recovery;
    restore.disabled = keep.disabled = busy;
    panel.dataset.status = state.status;
  };
  const unsubscribe = persistence.subscribe(update);
  const controller = new AbortController();
  parent.addEventListener("input", update, { signal: controller.signal });
  parent.addEventListener("keydown", () => queueMicrotask(update), {
    signal: controller.signal,
  });
  return () => {
    unsubscribe();
    controller.abort();
    panel.remove();
  };
}

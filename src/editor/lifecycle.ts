import { validateZoom } from "./geometry";
import type { Editor, EditorOptions, EditorConfiguration } from "./index";
import type { Rect } from "../core/types";
import { WorkbookError } from "../runtime/errors";
export interface ViewState {
  zoom: number;
  sheetId: string;
  selection: Rect;
  top: number;
  left: number;
  sheets: [string, { selection: Rect; top: number; left: number }][];
}
export interface EditorView extends Omit<Editor, "setOptions"> {
  captureView(): ViewState;
  detach(): void;
}
export function manageEditor(
  container: HTMLElement,
  initial: EditorOptions,
  mount: (
    container: HTMLElement,
    options: EditorOptions,
    state?: ViewState,
  ) => EditorView,
): Editor {
  let options = { ...initial },
    view = mount(container, options),
    destroyed = false;
  let queue = Promise.resolve();
  function destroy(): void;
  function destroy(options: { commit: true }): Promise<void>;
  function destroy(config?: { commit: true }): void | Promise<void> {
    if (config?.commit)
      return queue.then(async () => {
        await view.destroy({ commit: true });
        destroyed = true;
      });
    destroyed = true;
    view.destroy();
  }
  return {
    get ready() {
      return view.ready;
    },
    select: (sheet, range) => queue.then(() => view.select(sheet, range)),
    getSelection: () => view.getSelection(),
    resize: () => view.resize(),
    getZoom: () => view.getZoom(),
    setZoom: (zoom) => queue.then(() => view.setZoom(zoom)),
    getEditState: () => view.getEditState(),
    commitEdit: () => queue.then(() => view.commitEdit()),
    cancelEdit: () => view.cancelEdit(),
    destroy,
    setOptions(next: EditorConfiguration) {
      const update = queue.then(async () => {
        if (destroyed)
          throw new WorkbookError("DISPOSED", "Editor is destroyed");
        await view.ready;
        if ("zoom" in next) validateZoom(next.zoom ?? 1);
        if (next.persistence && next.persistence.workbook !== initial.workbook)
          throw new Error("Persistence belongs to a different workbook");
        const changed = Object.entries(next).some(([key, value]) =>
          key === "zoom"
            ? view.getZoom() !== (value ?? 1)
            : options[key as keyof EditorConfiguration] !== value,
        );
        if (!changed) return;
        await view.commitEdit();
        if (destroyed)
          throw new WorkbookError("DISPOSED", "Editor is destroyed");
        const state = view.captureView();
        if ("zoom" in next) state.zoom = next.zoom ?? 1;
        // Rebuild only the view. The authoritative Worker, undo history and clipboard stay intact.
        view.detach();
        options = { ...options, ...next, workbook: initial.workbook };
        view = mount(container, options, state);
        await view.ready;
      });
      queue = update.catch(() => {});
      return update;
    },
  };
}

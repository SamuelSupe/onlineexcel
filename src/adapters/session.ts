import { createWorkbook } from "../index";
import {
  mountEditor,
  type Editor,
  type EditorConfiguration,
} from "../editor/index";
import type { Workbook } from "../runtime/client";
import type { ChangeEvent, WorkbookOptions } from "../core/types";
export interface EditorHandle {
  workbook: Workbook;
  editor: Editor;
}
export interface AdapterProps {
  /** External workbooks remain owned by the host and are never disposed by the adapter. */
  workbook?: Workbook;
  /** Applied on mount. Remount the component to create a different owned workbook. */
  workbookOptions?: WorkbookOptions;
  options?: EditorConfiguration;
  onReady?: (handle: EditorHandle) => void;
  onError?: (error: Error) => void;
  onChange?: (change: ChangeEvent) => void;
}
export function createEditorSession(
  container: HTMLElement,
  initial: AdapterProps,
  latest: () => AdapterProps,
) {
  let disposed = false,
    workbook: Workbook | undefined,
    editor: Editor | undefined;
  let unsubscribe: (() => void) | undefined;
  const report = (error: unknown) => {
    if (disposed) return;
    try {
      (latest().onError ?? latest().options?.onError)?.(error as Error);
    } catch (callbackError) {
      console.error("OnlineExcel error callback failed", callbackError);
    }
  };
  const configuration = (): EditorConfiguration => ({
    locale: undefined,
    readOnly: false,
    toolbar: true,
    styleNonce: undefined,
    toolbarItems: undefined,
    actions: undefined,
    onImport: undefined,
    onExport: undefined,
    ...latest().options,
    onError: report,
  });
  const ready = (async () => {
    workbook =
      initial.workbook ?? (await createWorkbook(initial.workbookOptions));
    if (disposed) {
      if (!initial.workbook) await workbook.dispose();
      return;
    }
    unsubscribe = workbook.on("change", (change) =>
      latest().onChange?.(change),
    );
    editor = mountEditor(container, {
      ...configuration(),
      workbook,
    });
    await editor.ready;
    if (!disposed) latest().onReady?.({ workbook, editor });
  })().catch((error) => {
    report(error);
    editor?.destroy();
    unsubscribe?.();
    if (!initial.workbook) void workbook?.dispose();
  });
  return {
    ready,
    update() {
      if (editor && !disposed)
        void editor.setOptions(configuration()).catch(report);
    },
    dispose() {
      disposed = true;
      unsubscribe?.();
      editor?.destroy();
      if (!initial.workbook) void workbook?.dispose();
    },
  };
}

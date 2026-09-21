import {
  defineComponent,
  h,
  onBeforeUnmount,
  onMounted,
  ref,
  toRaw,
  watch,
  type PropType,
} from "vue";
import {
  createEditorSession,
  type AdapterProps,
  type EditorHandle,
} from "./session";
import type { Workbook } from "../runtime/client";
import type { WorkbookOptions, ChangeEvent } from "../core/types";
import type { EditorConfiguration } from "../editor/index";
export type { EditorHandle, AdapterProps } from "./session";
export const OnlineExcel = defineComponent({
  name: "OnlineExcel",
  inheritAttrs: false,
  props: {
    workbook: Object as PropType<Workbook>,
    workbookOptions: Object as PropType<WorkbookOptions>,
    options: Object as PropType<EditorConfiguration>,
  },
  emits: {
    ready: (_handle: EditorHandle) => true,
    error: (_error: Error) => true,
    change: (_change: ChangeEvent) => true,
  },
  setup(props, { emit, attrs }) {
    const container = ref<HTMLElement>();
    let session: ReturnType<typeof createEditorSession> | undefined;
    const latest = (): AdapterProps => ({
      workbook: props.workbook ? toRaw(props.workbook) : undefined,
      workbookOptions: props.workbookOptions,
      options: props.options,
      onReady: (handle) => emit("ready", handle),
      onError: (error) => emit("error", error),
      onChange: (change) => emit("change", change),
    });
    const mount = () => {
      session?.dispose();
      session = createEditorSession(container.value!, latest(), latest);
    };
    onMounted(mount);
    watch(
      () => props.workbook,
      () => {
        if (container.value) mount();
      },
    );
    watch(
      () => props.options,
      () => session?.update(),
      { deep: true },
    );
    onBeforeUnmount(() => session?.dispose());
    return () =>
      h("div", {
        ...attrs,
        ref: container,
        style: [{ height: "100%" }, attrs.style],
      });
  },
});

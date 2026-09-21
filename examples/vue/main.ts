import { createApp, defineComponent, h, ref } from "vue";
import { OnlineExcel, type EditorHandle } from "../../src/adapters/vue";
import type { Locale } from "../../src/index";
const workbookOptions = {
  workerFactory: () =>
    new Worker(new URL("../../src/runtime/worker.ts", import.meta.url), {
      type: "module",
    }),
};
createApp(
  defineComponent({
    setup() {
      const locale = ref<Locale>("en-US"),
        error = ref("");
      const seed = async ({ workbook }: EditorHandle) => {
        const [sheet] = (await workbook.getMetadata()).sheets;
        await workbook.setValues(sheet.id, "A1:B3", [
          ["Vue adapter", "Value"],
          ["Quantity", 24],
          ["Total", "=B2*19.9"],
        ]);
      };
      return () =>
        h("div", [
          h(
            "button",
            {
              onClick: () =>
                (locale.value = locale.value === "en-US" ? "ko-KR" : "en-US"),
            },
            "Language / 언어",
          ),
          h("span", { role: "status" }, error.value),
          h(OnlineExcel, {
            style: "height:calc(100vh - 32px)",
            workbookOptions,
            options: { locale: locale.value },
            onReady: (handle) =>
              void seed(handle).catch((e) => (error.value = e.message)),
            onError: (e) => (error.value = e.message),
          }),
        ]);
    },
  }),
).mount("#root");

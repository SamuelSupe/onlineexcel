import React, { useState } from "react";
import { createRoot } from "react-dom/client";
import { OnlineExcel, type EditorHandle } from "../../src/adapters/react";
import type { Locale } from "../../src/index";
const workbookOptions = {
  workerFactory: () =>
    new Worker(new URL("../../src/runtime/worker.ts", import.meta.url), {
      type: "module",
    }),
};
async function seed({ workbook }: EditorHandle) {
  const [sheet] = (await workbook.getMetadata()).sheets;
  await workbook.setValues(sheet.id, "A1:B3", [
    ["React adapter", "Value"],
    ["Revenue", 1200],
    ["Tax", "=B2*0.1"],
  ]);
}
function App() {
  const [locale, setLocale] = useState<Locale>("en-US"),
    [error, setError] = useState("");
  return (
    <>
      <button onClick={() => setLocale(locale === "en-US" ? "ja-JP" : "en-US")}>
        Language / 言語
      </button>
      <span role="status">{error}</span>
      <OnlineExcel
        style={{ height: "calc(100vh - 32px)" }}
        workbookOptions={workbookOptions}
        options={{ locale }}
        onReady={(handle) =>
          void seed(handle).catch((e) => setError(e.message))
        }
        onError={(e) => setError(e.message)}
      />
    </>
  );
}
createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);

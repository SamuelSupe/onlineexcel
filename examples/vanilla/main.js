(async () => {
  const workbook = await OnlineExcel.createWorkbook({
    sheets: [{ name: "Standalone" }],
  });
  const [sheet] = (await workbook.getMetadata()).sheets;
  await workbook.setValues(sheet.id, "A1:B3", [
    ["Standalone script", "Strict CSP"],
    ["Quantity", 12],
    ["Total", "=B2*39.9"],
  ]);
  const editor = OnlineExcel.mountEditor(document.querySelector("#sheet"), {
    workbook,
    locale: new URLSearchParams(location.search).get("locale") || "en-US",
    styleNonce: "onlineexcel-example",
  });
  await editor.ready;
  window.addEventListener(
    "pagehide",
    () => {
      editor.destroy();
      void workbook.dispose();
    },
    { once: true },
  );
})().catch((error) => {
  document.body.textContent = error.message;
  console.error(error);
});

import {
  createWorkbook,
  createPersistence,
  createIndexedDBStorage,
  mountEditor,
  listFunctions,
  supportedLocales,
  type Locale,
} from "../src/index";
import type { Command, WorkbookSnapshot } from "../src/core/types";
import { keyOf } from "../src/core/address";
import "./style.css";
import { icon } from "../src/editor/icons";
const app = document.querySelector<HTMLDivElement>("#app")!;
app.innerHTML = `<header class="app-header"><div class="identity"><div class="logo">O<span>▦</span></div><div><h1>销售运营工作簿 <span>本地文件</span></h1><p>OnlineExcel <b>/</b> 2026 年度业务分析</p></div></div><div class="header-right"><a href="/examples/react/">React 示例</a><a href="/examples/vue/">Vue 示例</a><button id="benchmark"><span>大表基准</span></button><div class="avatar">OE</div></div></header><main id="editor"></main><footer class="app-footer"><span><i></i> 草稿自动保存到本机 · 可嵌入任何 Web 应用</span><span id="function-count"></span></footer>`;
for (const link of app.querySelectorAll(".header-right a"))
  link.prepend(icon("code"));
app.querySelector("#benchmark")!.prepend(icon("benchmark"));
const workbook = await createWorkbook({
  workerUrl: new URL("../src/runtime/worker.ts", import.meta.url),
  sheets: [{ name: "销售明细" }, { name: "分析看板" }, { name: "使用说明" }],
});
const metadata = await workbook.getMetadata(),
  [sales, summary, guide] = metadata.sheets.map((s) => s.id);
const headers = [
  "订单编号",
  "客户名称",
  "区域",
  "产品",
  "数量",
  "单价",
  "销售额",
  "订单状态",
  "签约日期",
  "客户负责人",
];
const clients = [
  "上海云栖科技",
  "北京北辰数据",
  "深圳澄明网络",
  "杭州星河系统",
  "成都远川信息",
  "苏州青禾智能",
  "南京方舟软件",
  "广州蓝图科技",
  "武汉知行科技",
  "西安长风数据",
  "厦门山海网络",
  "天津启明云科",
];
const regions = ["华东", "华北", "华南", "西南"],
  products = ["企业专业版", "数据分析套件", "云端协作版"],
  owners = ["陈一鸣", "林知夏", "周予安", "许清和"];
const data: (string | number)[][] = [headers];
for (let i = 0; i < 60; i++)
  data.push([
    `SO-2026-${String(1001 + i)}`,
    clients[i % clients.length],
    regions[i % 4],
    products[i % 3],
    10 + ((i * 7) % 80),
    [1280, 680, 980][i % 3],
    `=E${i + 2}*F${i + 2}`,
    i % 7 === 0 ? "待确认" : i % 5 === 0 ? "处理中" : "已完成",
    46265 + i,
    owners[i % 4],
  ]);
await workbook.setValues(sales, "A1:J61", data);
const commands: Command[] = [
  {
    type: "style",
    sheetId: sales,
    range: { r1: 0, c1: 0, r2: 0, c2: 9 },
    style: {
      background: "#eaf3ed",
      bold: true,
      color: "#24553c",
      fontSize: 11,
    },
  },
  { type: "dimensions", sheetId: sales, axis: "row", indexes: [0], size: 36 },
  {
    type: "dimensions",
    sheetId: sales,
    axis: "column",
    indexes: [0],
    size: 148,
  },
  {
    type: "dimensions",
    sheetId: sales,
    axis: "column",
    indexes: [1],
    size: 176,
  },
  {
    type: "dimensions",
    sheetId: sales,
    axis: "column",
    indexes: [3],
    size: 150,
  },
  {
    type: "dimensions",
    sheetId: sales,
    axis: "column",
    indexes: [2, 4],
    size: 84,
  },
  {
    type: "style",
    sheetId: sales,
    range: { r1: 1, c1: 5, r2: 60, c2: 6 },
    style: { numberFormat: "#,##0.00" },
  },
  {
    type: "style",
    sheetId: sales,
    range: { r1: 1, c1: 8, r2: 60, c2: 8 },
    style: { numberFormat: "yyyy-mm-dd" },
  },
  { type: "freeze", sheetId: sales, rows: 1, columns: 1 },
];
for (let r = 1; r <= 60; r++)
  if (r % 2 === 0)
    commands.push({
      type: "style",
      sheetId: sales,
      range: { r1: r, c1: 0, r2: r, c2: 9 },
      style: { background: "#f8faf9" },
    });
commands.push({
  type: "style",
  sheetId: sales,
  range: { r1: 1, c1: 7, r2: 60, c2: 7 },
  style: { color: "#16734c" },
});
await workbook.transaction(commands);
await workbook.setValues(summary, "A1:D8", [
  ["区域分析", null, null, null],
  ["区域", "销售总额", "订单数", "平均订单额"],
  ...regions.map((region, i) => [
    region,
    `=SUMIF('销售明细'!C2:C61,A${i + 3},'销售明细'!G2:G61)`,
    `=COUNTIF('销售明细'!C2:C61,A${i + 3})`,
    `=IFERROR(B${i + 3}/C${i + 3},0)`,
  ]),
  ["合计", "=SUM(B3:B6)", "=SUM(C3:C6)", "=B7/C7"],
  ["动态数组示例", null, null, null],
]);
await workbook.setFormula(summary, "A10", "=SORT(UNIQUE('销售明细'!D2:D61))");
await workbook.setStyle(summary, "A1:D2", {
  bold: true,
  background: "#eaf3ed",
});
await workbook.setStyle(summary, "B3:B7", { numberFormat: '"¥"#,##0.00' });
await workbook.setStyle(summary, "D3:D7", { numberFormat: '"¥"#,##0.00' });
await workbook.setDimensions(summary, "column", [0, 1, 2, 3], { size: 180 });
await workbook.setValues(
  guide,
  "A1:B10",
  [
    ["欢迎使用 OnlineExcel", "自主核心 · 浏览器内运行"],
    ["编辑", "双击单元格或直接输入；Enter 确认，Esc 取消"],
    ["公式", "=SUM(A1:A10)、跨表引用、动态数组"],
    ["填充", "拖动选区右下角的绿色方块"],
    ["保存", "自动保存到本机；重新打开可选择恢复草稿，也可导出 XLSX / JSON"],
    ["兼容性", "发现高级对象时显示差异，有损导出需要确认"],
    ["接入", "原生 JS / React / Vue 共用 Promise API"],
    ["文件", "文件不会上传到服务器"],
    ["性能", "顶部“大表基准”生成 100 万非空单元格"],
    ["说明", "函数支持参数与兼容差异见 docs/compatibility.md"],
  ],
  { parseFormulas: false },
);
await workbook.setDimensions(guide, "column", [0], { size: 170 });
await workbook.setDimensions(guide, "column", [1], { size: 650 });
await workbook.setStyle(guide, "A1:B1", { bold: true, background: "#eaf3ed" });
const requestedLocale = new URLSearchParams(location.search).get("locale");
const locale: Locale = supportedLocales.includes(requestedLocale as Locale)
  ? (requestedLocale as Locale)
  : "zh-CN";
const persistence = createPersistence(workbook, {
  key: "onlineexcel-demo",
  storage: createIndexedDBStorage(),
});
const editor = mountEditor(document.querySelector("#editor")!, {
  workbook,
  locale,
  persistence,
});
await editor.ready;
document.querySelector("#function-count")!.textContent =
  `${listFunctions().length} 个公式函数 · Worker 计算引擎`;
const benchmarkButton =
  document.querySelector<HTMLButtonElement>("#benchmark")!;
const benchmarkCaption = benchmarkButton.querySelector("span")!;
benchmarkButton.onclick = async () => {
  benchmarkButton.disabled = true;
  benchmarkCaption.textContent = "生成测试数据…";
  try {
    if (persistence.getState().status !== "disposed") {
      await persistence.save();
      persistence.dispose();
      await editor.setOptions({ persistence: undefined });
      app.querySelector(".app-footer span")!.textContent =
        "基准数据不自动保存 · 原草稿已保留，刷新可恢复";
    }
    const cells: WorkbookSnapshot["sheets"][number]["cells"] = [];
    for (let r = 0; r < 100000; r++)
      for (let c = 0; c < 10; c++)
        cells.push([
          keyOf(r, c),
          c === 9 && r < 50000
            ? { formula: `=SUM(A${r + 1}:I${r + 1})` }
            : {
                value:
                  c === 0 ? `ID-${r + 1}` : ((r * 31 + c * 17) % 10000) / 100,
              },
        ]);
    const current = await workbook.exportJSON();
    const base = current.sheets[0];
    const snapshot: WorkbookSnapshot = {
      version: 1,
      dateSystem: 1900,
      styles: [{}],
      names: {},
      sheets: [
        {
          ...base,
          id: "benchmark",
          name: "100万单元格",
          rowCount: 100000,
          columnCount: 100,
          rowHeights: {},
          columnWidths: {},
          hiddenRows: [],
          hiddenColumns: [],
          frozenRows: 1,
          frozenColumns: 1,
          merges: [],
          cells,
        },
      ],
    };
    const start = performance.now();
    await workbook.importJSON(snapshot);
    await editor.select("benchmark", "A1");
    benchmarkCaption.textContent = `载入 ${(performance.now() - start).toFixed(0)} ms`;
  } catch (error) {
    benchmarkCaption.textContent = (error as Error).message;
  } finally {
    benchmarkButton.disabled = false;
  }
};
// Exposed only by the demonstration page for integration experiments and repeatable browser QA.
Object.assign(window, {
  onlineExcelDemo: {
    workbook,
    editor,
    persistence,
    sheetIds: { sales, summary, guide },
  },
});

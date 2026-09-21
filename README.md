# OnlineExcel

浏览器内运行的可嵌入电子表格。TypeScript 编写，自主实现数据模型、公式计算、Canvas 网格和 XLSX 映射；不依赖 React、Vue 或第三方表格引擎。

**这是 0.1 版本的办公功能实现，不是桌面 Excel 的完整兼容替代品。** 功能和参数边界见 [兼容矩阵](docs/compatibility.md)、[公式清单](docs/functions.md)。

[下载版本](https://github.com/SamuelSupe/onlineexcel/releases) · [接入指南](docs/sdk.md) · [API](docs/api.md) · [验证记录](docs/validation.md)

## 安装发布包

在宿主应用中安装 [npm 包](https://www.npmjs.com/package/onlineexcel)：

```sh
npm install onlineexcel
npx onlineexcel-copy-assets public/onlineexcel
```

然后通过 `import { createWorkbook, mountEditor } from "onlineexcel"` 接入，并配置 `workerUrl: "/onlineexcel/worker.js"`。发布包包含 ESM、浏览器脚本、Worker、类型声明和 React/Vue 可选入口。也可以从 [GitHub Releases](https://github.com/SamuelSupe/onlineexcel/releases) 下载 `.tgz`，使用 `npm install ./onlineexcel-0.1.1.tgz` 安装；离线分发包的校验文件见对应 Release 的 `SHA256SUMS`。

## 本地运行

```sh
git clone https://github.com/SamuelSupe/onlineexcel.git
cd onlineexcel
npm ci
npm run dev
```

打开 `http://127.0.0.1:5173/`。演示包含销售明细、跨表汇总、动态数组，以及百万单元格测试数据生成入口。已提交的编辑自动保存到本机 IndexedDB，刷新后可选择恢复草稿或使用当前内容；导出按钮用于保存独立文件。大表基准会暂停自动保存，保留进入基准前的草稿。

```sh
npm run check       # 类型检查、行为测试、库构建
npm run functions   # 更新公式能力文档
npm run bench       # 固定数据集的核心性能基准
npm pack            # 生成可安装的 npm 包
```

工程默认使用 OrbStack 跑测试；复现环境和实测记录见 [验证说明](docs/validation.md)。不需要服务器、数据库或云服务。

SDK 现已提供 Worker 工厂与资产复制工具、结构化错误、带版本快照、分块和记录 API、动态编辑器配置、宿主业务校验、自定义函数，以及 `onlineexcel/react` / `onlineexcel/vue` 可选组件。完整接入契约和示例见 [SDK 指南](docs/sdk.md)。

提供可选自动保存和恢复提示、可筛选并定位单元格的问题面板、50%～200% 网格缩放、数据边界快捷键、格式刷和 F1 快捷键帮助。自动保存由宿主通过 `createPersistence(workbook, { key, storage })` 启用，再将 `persistence` 传给编辑器；默认库实例不访问存储。

## 嵌入到应用

```ts
import { createWorkbook, mountEditor } from "onlineexcel";

const workbook = await createWorkbook({
  // 先执行 npx onlineexcel-copy-assets public/onlineexcel。
  workerUrl: "/onlineexcel/worker.js",
  sheets: [{ name: "销售", rows: 100_000, columns: 100 }],
});
const editor = mountEditor(document.querySelector("#spreadsheet")!, {
  workbook,
  locale: "zh-CN",
});
await editor.ready;

const [sheet] = (await workbook.getMetadata()).sheets;
await workbook.setValues(sheet.id, "A1:B2", [
  ["数量", "单价"],
  [12, 39.9],
]);
await workbook.setFormula(sheet.id, "C2", "=A2*B2");
console.log(await workbook.getValues(sheet.id, "C2")); // [[478.8]]

const unsubscribe = workbook.on("change", (event) => {
  console.log(event.revision, event.changes);
});
await editor.commitEdit();
const { data, diagnostics } = await workbook.exportXlsx();
// data 为 Uint8Array。由宿主决定下载、上传或持久化位置。

unsubscribe();
await editor.destroy({ commit: true });
await workbook.dispose();
```

宿主容器需要明确高度，例如 `height: 600px`。CSS 已包含在编辑器中。`mountEditor` 不拥有 Workbook；移除视图后仍可使用 API，结束工作簿生命周期时再调用 `dispose()`。

在单元格或公式栏输入 `=` 会显示公式模式提示、示例和函数建议。继续输入函数名可筛选补全项，使用 `↑` / `↓` 选择，`Tab` 或鼠标点击补全；参数输入时会高亮当前参数。`Enter` 确认，`Esc` 取消。输入 `'=文字` 可以保留以等号开头的普通文本。提示随编辑器语言切换，并在输入法组合期间收起。

需要引用单元格时可点击或拖选网格，蓝色虚线标出本次引用；也可切换工作表插入跨表引用。`F4` 在相对、绝对和混合引用之间切换。查找支持大小写、整格匹配和 `F3` 连续查找并回绕；排序提供表头保护、范围扩展与多个关键字；筛选支持逐列条件和候选值多选。筛选后的复制与统计只包含可见单元格；可选择仅粘贴值、公式与值、格式或转置。工具栏随选区反映格式，公式草稿允许滚动寻找引用。

提供带编码与列类型选择的 CSV 预览、工作表复制与隐藏、名称管理器、下拉验证、基础条件格式和区域锁定。保护先解锁输入区再启用；它用于防止误编辑，不替代宿主权限控制。图片、传统注释和结构化表支持关联文件保留，编辑边界见兼容矩阵。

### 指定界面语言

语言包随库提供，无须额外下载。在 JS 创建编辑器时传入 `locale`：

```js
import { createWorkbook, mountEditor, supportedLocales } from "onlineexcel";

const workbook = await createWorkbook();
const editor = mountEditor(document.querySelector("#sheet"), {
  workbook,
  locale: "ja-JP",
});
await editor.ready;
```

| 语言 | `locale` |
| --- | --- |
| 简体中文（默认） | `zh-CN` |
| 繁體中文（台灣用語） | `zh-TW` |
| English | `en-US` |
| 日本語 | `ja-JP` |
| 한국어 | `ko-KR` |

`supportedLocales` 提供支持的语言列表，TypeScript 可导入 `Locale` 类型。每个编辑器独立设置语言，JS 传入未支持的语言值时回退到 `zh-CN`。界面语言包含工具栏、弹窗、状态栏、无障碍标签、公式提示及常见操作错误；工作表名称、单元格数据、数字格式、公式函数名和参数名不自动翻译。开发者 API 错误及文件兼容诊断原文保留不变。

本地演示可通过 `http://127.0.0.1:5173/?locale=ja-JP` 初始化为日语；将参数换成其他语言代码即可。React、Vue 和原生 JS 均使用相同的 `mountEditor` 配置。

### 无构建工具

将整个 `dist/` 目录部署到同源静态服务器：

```html
<div id="sheet" style="height:600px"></div>
<script src="/onlineexcel/onlineexcel.js"></script>
<script src="/my-spreadsheet-page.js"></script>
```

```js
// my-spreadsheet-page.js
const workbook = await OnlineExcel.createWorkbook();
const editor = OnlineExcel.mountEditor(document.querySelector("#sheet"), {
  workbook,
  locale: "zh-TW",
});
await editor.ready;
```

普通脚本中的上述代码需要放进 `async` 函数；也可以给 `my-spreadsheet-page.js` 添加 `type="module"`。默认 Worker 地址相对于库脚本；资产不在同目录时显式传入 `workerUrl`。

### React / Vue

- [React 示例](examples/react/main.tsx)：组件挂载后创建，在 effect 清理时销毁；包含 StrictMode 生命周期处理。
- [Vue 示例](examples/vue/main.ts)：`onMounted` 创建，`onBeforeUnmount` 释放。
- 原生 JS 接入见 [直接加载示例](examples/vanilla/main.js)。React/Vue 示例在仓库中使用源码入口；宿主分别导入 `onlineexcel/react` 和 `onlineexcel/vue`，并配置 Worker 资产地址。
- [安装包生产示例](examples/installed/README.md) 使用真实 npm 安装包及 Vite 构建，包含严格 CSP 和三个宿主的集成验证。

## 能力

- 多工作表；区域编辑、样式、合并、冻结、排序、筛选、查找替换、剪贴板和拖拽填充。
- 有界差异历史与原子事务；公式依赖、相对/绝对及跨表引用、命名区域、自动重算。
- 187 个函数，包括条件聚合、查找、日期、财务和动态数组；可通过 `listFunctions()` 查询签名与限制。
- XLSX、CSV、JSON 导入导出；保留未知公式原文，报告不支持的高级对象，有损导出必须显式允许。
- Worker 隔离计算与文件处理；主线程仅保留视图缓存；同页多实例和 Shadow DOM 样式隔离。
- 简体中文、繁体中文、英文、日语、韩语界面；按实例配置语言；只读视图、隐藏工具栏、CSS 变量主题和 CSP style nonce。

详细接口见 [API 文档](docs/api.md)。内部同步数据模型从 `onlineexcel/core` 导出，适合高级集成和测试；正常应用使用 Worker 驱动的顶层 API。

## 部署和 CSP

所有资产部署到宿主可访问的同源目录。Worker 会加载构建产物中的共享 chunk，不能只复制一个 `worker.js`。

公式解释器不使用 `eval` 或 `new Function`。严格 CSP 可使用 `script-src 'self'; worker-src 'self'; style-src 'nonce-随机值'` 并传入 `mountEditor(container, { workbook, styleNonce: '随机值' })`。本库不要求跨源隔离或 SharedArrayBuffer；仅发布包不需要 Vite 开发环境的额外 CSP 配置。

通过容器 CSS 变量调整外观：`--oe-accent`、`--oe-accent-soft`、`--oe-border`、`--oe-text`、`--oe-muted`、`--oe-font`。`readOnly` 是编辑界面限制，不是权限系统；宿主仍可调用写入 API。

## License

MIT。底层 ZIP/XML 依赖及许可证随 npm 包依赖链保留。

# 宿主集成与 SDK 契约

这些接口在 0.1.0 开发包中提供。原有 API 保持可用；React 和 Vue 是独立可选入口，核心包仍只依赖 fflate 和 saxes。

## 安装与 Worker 部署

```sh
npm install /path/to/onlineexcel-0.1.0.tgz
npx onlineexcel-copy-assets public/onlineexcel
```

命令复制浏览器脚本、Worker 和依赖的 chunk 文件，不删除目标目录已有文件。每次更新安装包后重新复制，整组资源一起部署；Worker 必须由同源地址或符合浏览器策略的部署路径加载，不能通过 `file://` 运行。

```ts
import { createWorkbook, mountEditor } from "onlineexcel";
const workbook = await createWorkbook({
  workerUrl: "/onlineexcel/worker.js",
  initializationTimeout: 15_000, // 默认；0 表示不设超时
  requestTimeout: 0, // 默认不限制长操作耗时
});
```

需要由宿主管理 Worker 构造时使用 `workerFactory: () => new Worker(url, { type: "module" })`，优先级高于 workerUrl。每次调用必须返回一个新 Worker，库拥有并最终终止它。Vite 等打包器也可以在工厂内分析 Worker 构造表达式；本仓库的源码 React/Vue 示例使用 `new Worker(new URL("../../src/runtime/worker.ts", import.meta.url), { type: "module" })`。安装包推荐上面的静态资产复制方案，无需依赖打包器对库内部 URL 的重写。

严格 CSP 可使用 `script-src 'self'; worker-src 'self'`，并为编辑器样式传入宿主生成的 `styleNonce`。库不使用 eval、Blob Worker 或动态函数构造。宿主模块是被信任的代码，其来源也必须满足 CSP。生产示例见 `examples/installed`，包含 npm 安装、Vite build、三个宿主和严格 CSP。

## 错误、取消和故障

数据 API 拒绝时可检查导出的 `WorkbookError`，不必匹配英文 message。属性包括 `code`、`operation`、`operationId`、`sheetId`、`range`、`outcome`、`details` 和 `diagnostics`；位置只在该操作提供了对应信息时存在。

| code                               | 含义                                                            |
| ---------------------------------- | --------------------------------------------------------------- |
| INVALID_ARGUMENT                   | 请求参数不合法，尚未提交                                        |
| VALIDATION_FAILED                  | 宿主业务校验未通过，整个事务回滚                                |
| OPERATION_FAILED                   | 模型/文件操作失败，具体原因见 message；单项模型规则仍使用此类别 |
| CANCELLED                          | 取消；保留 name=AbortError 以兼容已有处理                       |
| LOSSY_EXPORT                       | 导出需要先确认兼容报告，再显式传 allowLossy                     |
| MODULE_FAILED                      | 模块加载、注册或返回契约不合法                                  |
| REVISION_CONFLICT                  | 一致性分块读取期间工作簿发生变化                                |
| PROTOCOL_MISMATCH                  | Worker 协议、版本或消息格式不匹配                               |
| SERIALIZATION_FAILED               | 消息无法复制或反序列化                                          |
| WORKER_FAILED / TIMEOUT / DISPOSED | Worker 故障、超时或实例已释放                                   |

`outcome=not-executed` 表示尚未发送；`rolled-back` 表示 Worker 返回失败且没有提交该操作；`unknown` 表示宿主无法确认提交状态。超时和 Worker 故障会拒绝所有等待请求并终止实例，**不会自动重试写入，也不会自动从旧快照恢复**。应用可提示用户并从自己保存的快照创建新实例。默认只有初始化设超时，长任务可按需配置。

```ts
const scoped = workbook.withOptions({
  origin: "host-import",
  operationId: crypto.randomUUID(),
  signal: abortController.signal,
  timeout: 60_000,
});
await scoped.setFormula(sheetId, "C1", "=SUM(A1:B1)");
```

`withOptions` 是同一 Workbook 的轻量视图，共享状态、事件、历史和生命周期。调用该视图的 dispose 也会终止原实例。单次方法或 request 传入的选项覆盖视图默认值，互不影响并发请求。取消为协作式：已完成提交的操作仍返回成功；异步数据源自身的等待需要由数据源负责取消。

## 事件和可靠保存

change/progress/calculation 事件携带 `operationId` 和 `origin`。默认 origin 为 `api`，编辑器触发的操作为 `editor`，宿主可自定义。change.source 仍表示 edit / undo / redo / import，不与 origin 混用。Worker 的请求响应、变化事件及进度可按 operationId 关联；未显式指定时自动生成。底层同步 core 模型没有 Worker 请求，因此这些附加字段是可选的。

```ts
let latestRevision = 0;
const off = workbook.on("change", (event) => {
  latestRevision = event.revision;
});
await editor.commitEdit();
const savePoint = await workbook.createSavePoint();
await persist(savePoint); // snapshot 和 revision 作为同一个记录持久化
const stillDirty =
  latestRevision > savePoint.revision || editor.getEditState().dirty;
```

`createSavePoint()` 在同一个 Worker 队列步骤取得 `{snapshot, revision}`，不会发生先取版本再取快照的竞争。版本仅在一个 Workbook 实例内递增，新实例从 0 开始。宿主仍须串行化持久化请求或使用条件写入，避免较早的网络保存覆盖较新的保存；库不决定保存位置。exportJSON 仍只返回原有版本化快照格式。

## 可选自动保存与草稿恢复

```ts
import { createPersistence, createIndexedDBStorage } from "onlineexcel";

const persistence = createPersistence(workbook, {
  key: "sales-2026", // 宿主提供稳定的文档标识
  storage: createIndexedDBStorage("my-app-drafts"),
  debounceMs: 800,
  maxRetries: 2,
  retryDelayMs: 1000,
});
const editor = mountEditor(container, { workbook, persistence, locale: "zh-CN" });
const unsubscribe = persistence.subscribe((state) => {
  console.log(state.status, state.savedRevision, state.savedAt, state.error);
});
await persistence.ready;
// 有旧草稿时，编辑器显示“恢复草稿 / 使用当前内容”，不会自行覆盖。
// 无界面的宿主自行展示选择，再调用 restore() 或 discardRecovery()。

// 应用内离开页面前，在用户已处理恢复选择的前提下：
await persistence.save(); // 先提交绑定编辑器的当前输入，再等待实际存储完成
unsubscribe();
editor.destroy();
persistence.dispose();
await workbook.dispose();
```

不配置 persistence 时不访问任何存储，不显示保存栏。保存控制器与 Workbook 一一对应；可以绑定同一工作簿的多个编辑器。编辑器销毁只解绑视图，宿主负责 dispose 保存控制器及工作簿。

状态包括 loading、recovery、dirty、saving、retrying、saved、error、disposed。后台保存仅处理已提交的修改，不打断输入法或自动提交正在输入的草稿。显式 save()、保存按钮和 Ctrl/Command+S 会先提交；输入法组合未结束或提交失败时保留输入并报错。没有配置持久化时，Ctrl/Command+S 保持导出 XLSX 的行为。

首次加载失败时禁止写入，可调用 retry() 重试加载；发现已存草稿时暂停写入，直到宿主或用户作出选择。restore() 走原子 JSON 导入及业务校验，失败保留工作簿和编辑草稿；成功才清除编辑草稿。discardRecovery() 选择当前内容，下次保存替换旧记录。初始 ready Promise 只代表首次加载；其失败后的恢复请等待 retry()，不要重复等待已拒绝的 ready。

每个控制器串行写入，以保存点的版本判断是否已保存；保存期间的新修改继续排队。失败默认额外重试两次，间隔 1 秒、2 秒；耗尽后保留 error/未保存状态，等待显式 retry()。dispose() 清理计时器、订阅和离页监听，并发送 AbortSignal 取消存储请求；它不会保存剩余修改。

默认 beforeunload 在有未保存修改或编辑草稿时申请浏览器原生离页提示。它不能异步完成保存，也不覆盖 SPA 路由；宿主使用 hasUnsavedChanges() 和 save() 处理路由离开。浏览器崩溃时只能恢复最后一次成功写入的内容。

自定义存储实现 `PersistenceStorage`：`load(key, signal)` 返回 `SavedWorkbook | undefined`，`save(key, record, signal)` 持久化 `{version:1, savedAt, snapshot}`。save 必须在实际写入完成后 resolve，同键重复调用必须安全地替换同一记录，并响应取消。可使用宿主 API 或自己的本地存储；库不上传数据。网络条件写入所需的 ETag/令牌由宿主适配器维护，不能把实例内 revision 用作全局版本。

内置 IndexedDB 适用于同源本地草稿，每个文档键应只有一个写入控制器。多个标签页/实例同时写同键、跨设备同步及冲突合并不在契约内。浏览器存储配额和清理策略仍适用；大工作簿保存会复制完整快照，宿主应按规模配置间隔与存储方案。

## 诊断与编辑效率

状态栏和“视图”菜单均提供问题检查面板：按错误/警告和工作表筛选，显示影响、修复方向及可展开的原始诊断，点击定位跳到相关工作表与区域。隐藏工作表需先取消隐藏；没有位置的全局问题不显示跳转按钮。导入后有诊断会打开面板，有损导出的显式确认规则不变。

```ts
const editor = mountEditor(container, { workbook, zoom: 1.25 });
await editor.setZoom(1.5); // 0.5～2，真实缩放网格及输入框，先提交当前编辑
console.log(editor.getZoom());
await editor.setOptions({ locale: "ja-JP" }); // 保留当前缩放
const destination = await workbook.getNavigationTarget(sheetId, "A1", "down");
```

缩放是编辑器视图状态，不修改行高列宽，也不写入 XLSX/JSON。Ctrl/Command+方向键跳至连续数据末端或下一段数据，遇到空白区时寻找下一非空单元格，最后才到工作表边缘；自动跳过隐藏/筛选行列。加 Shift 扩展选区。

“开始 → 格式刷”捕获当前可见选区的基础格式和条件格式；点击单格按源尺寸展开，拖选区域按源图案重复，应用后退出，Esc 取消。保留目标内容、公式和数据校验，不复制合并结构或行列尺寸，遵守保护和合并限制，一次应用对应一次撤销。最多应用 200,000 格。“视图 → 快捷键帮助”或 F1 显示快捷键说明。新增按钮和文案支持全部五种语言。

## 编辑器配置和宿主操作

```ts
await editor.setOptions({ locale: "zh-TW", readOnly: true, toolbar: false });
```

更新前提交当前草稿；提交失败保留原视图。中文输入法组合未完成、正在导入/导出或编辑器已销毁时拒绝。成功后重建视图，保留工作簿、历史、当前选区、各工作表视图位置及剪贴板内容；不承诺保留工具栏活动页或打开的对话框。不能更换 workbook，切换工作表使用 select。setOptions 返回新视图 ready 后才完成。

```ts
const editor = mountEditor(container, {
  workbook,
  locale: "en-US",
  toolbarItems: ["bold", "font", "format", "open", "saveXlsx"],
  actions: [
    {
      id: "approve",
      label: "审批",
      icon: "saveJson",
      async run({ workbook, sheetId, range, signal }) {
        await workbook
          .withOptions({ signal })
          .setStyle(sheetId, range, { background: "#e9f4ee" });
      },
    },
  ],
  async onImport({ workbook, signal }) {
    const file = await hostChooseFile();
    if (file) await workbook.importXlsx(file, { signal });
  },
  async onExport(format, { workbook, signal }) {
    if (format === "xlsx")
      await hostSave(await workbook.exportXlsx({ signal }));
  },
});
```

toolbarItems 指定内置按钮或样式控件 ID，未指定时全部显示；空数组隐藏内置操作。自定义 action 使用内置 IconName 和文字，默认只在可编辑模式启用；`allowReadOnly: true` 表示宿主明确允许它在只读视图执行。隐藏按钮、只读视图都不代替宿主权限控制。回调接到带 `origin=editor` 的工作簿视图；一次 transaction 包含的多条命令对应一个撤销步骤，多次 API 调用仍是多个步骤。文件回调有取消 signal，须传给自己的 IO 或 Workbook 操作。

## 记录与分块数据

```ts
await workbook.setRecords(
  sheetId,
  [
    { sku: "001", amount: 12 },
    { sku: "002", amount: 20 },
  ],
  [
    { key: "sku", title: "编号" },
    { key: "amount", title: "金额" },
  ],
  { start: "A1", header: true },
);

for await (const records of workbook.readRecords(
  sheetId,
  "A2:B10001",
  [{ key: "sku" }, { key: "amount" }],
  { rowsPerChunk: 1000 },
)) {
  await consume(records);
}
for await (const region of workbook.readChunks(sheetId, "A1:Z10000")) {
  // 包含数据、公式、样式和 revision 的 Region
}
const result = await workbook.writeChunks(sheetId, "A1", asyncMatrixSource, {
  signal,
});
```

列 key 必须非空且唯一。写入接受 string / number / boolean / null，缺失字段转 null，默认把 `=...` 当文本，避免外部业务数据变成公式。日期请显式转换为日期序列值并设置格式。读取区域不自动跳过表头；错误单元格返回 `{error}`。

每块最多 200,000 个单元格。readChunks/readRecords 默认 `consistent=true`，后续块遇到其他操作改变版本时拒绝，已经交付的块由宿主丢弃或重试整个读取；设 false 可以接受实时变化。读取不复制整本工作簿。

writeChunks 支持 Iterable 或 AsyncIterable 矩阵源，逐块发送；每块一个原子事务和撤销步骤。失败或取消不撤回已成功的块，错误 `details.committed` 给出 rows、chunks、revision。需要整体原子性时，使用一次 setRecords/setValues/transaction 或导入快照。库不暗中缓存整个流以假装低内存的整体原子事务。

## 自定义公式与业务校验

创建时通过 `workerModules: ["/business.js"]` 加载可信 ES 模块。相对路径按 document.baseURI 解析；模块只在当前 Workbook 的 Worker 中注册，不影响其他实例。

```js
// business.js，无需导入库的内部注册器
export const functions = [
  {
    name: "ACME.DOUBLE",
    minArgs: 1,
    maxArgs: 1,
    signature: "value",
    description: "将输入数值乘以二",
    evaluate: ([value]) => Number(value) * 2,
  },
];
export function validate({ sheets, getRegion, operation, commands, origin }) {
  const sheetId = sheets[0].id;
  const value = getRegion(sheetId, "A1").cells[0]?.value;
  if (typeof value === "number" && value < 0)
    return [
      {
        code: "NEGATIVE_AMOUNT",
        severity: "error",
        sheetId,
        range: "A1",
        message: "金额不能为负数",
      },
    ];
}
```

函数名使用大写，建议采用业务命名空间，如 `ACME.DOUBLE`；禁止覆盖已存在的函数。函数是同步、确定性的计算，不能返回 Promise；不能靠闭包中的外部状态变化触发重算。参数为 Scalar 或二维 Scalar 数组，结果可返回标量或矩形二维数组；范围/数组最多 200,000 个单元格。上下文提供 row、column、dateSystem、now。异常和非法返回值变为 `#VALUE!`；需要传播具体错误时返回 `{error:"#N/A"}` 等标准错误。数组参与同样的溢出、遮挡和依赖计算。

`workbook.getFunctions()` 返回该实例的内置和自定义函数元数据，编辑器会用于补全和参数提示；顶层 listFunctions() 仍只描述主线程内置函数。自定义函数调用在兼容报告中标为 `CUSTOM_FUNCTION_REQUIRES_MODULE`。JSON/XLSX 保留公式文本，但不嵌入可执行代码；恢复时需要再次配置模块，Excel 本身不能执行这些 JS 函数。

validate 在普通事务计算完成、提交之前执行，也验证初始化及 JSON/XLSX 导入的候选工作簿。返回非空诊断数组即拒绝全部事务；必须同步返回，不能返回 Promise。上下文 commands 对普通事务是原始命令，对导入/初始化为空；getRegion 是候选状态的只读查询。undo/redo 恢复原有已接受历史，不重复执行当前业务校验。校验不提供安全隔离或访问控制，模块与宿主同样可信。

## React 和 Vue

```tsx
import { OnlineExcel } from "onlineexcel/react";
<OnlineExcel
  workbookOptions={{ workerUrl: "/onlineexcel/worker.js" }}
  options={{ locale: "zh-TW", readOnly: false }}
  style={{ height: 600 }}
  onReady={({ workbook, editor }) => {
    /* 可读写和订阅 */
  }}
  onChange={(event) => console.log(event.revision)}
  onError={(error) => console.error(error)}
/>;
```

```vue
<script setup>
import { OnlineExcel } from "onlineexcel/vue";
const workbookOptions = { workerUrl: "/onlineexcel/worker.js" };
</script>
<template>
  <OnlineExcel
    :workbook-options="workbookOptions"
    :options="{ locale: 'ja-JP' }"
    style="height: 600px"
    @ready="({ workbook, editor }) => onReady(workbook, editor)"
    @error="onError"
  />
</template>
```

两者都支持传入现有 `workbook`。传入时宿主拥有它，卸载组件仅销毁编辑器；未传入时适配器异步创建并拥有 Workbook，卸载同时终止 Worker，初始化中卸载也会清理迟到的实例。更换 workbook 会更换视图；workbookOptions 是创建参数，更换它需要通过组件 key 重新挂载。options 的变化调用 setOptions，不重建 Workbook。React StrictMode 的试挂载允许出现短暂的初始化，但不会留下孤儿实例。SSR 可以导入和渲染空容器，Worker 仅在客户端挂载后创建。

同步卸载默认丢弃草稿；需要保存时在路由守卫中先 await editor.commitEdit() 并保存，再允许卸载。onReady 提供的 editor 是稳定句柄，支持后续调用 setOptions。React 需要宿主安装 react，Vue 需要宿主安装 vue；其他入口不会加载框架。

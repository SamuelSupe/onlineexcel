# API

公开类型由 TypeScript 声明提供。单元格地址使用 A1，API 的行列索引和 `Rect` 坐标从 **0** 开始；范围端点包含在内。工作表使用稳定 ID，重命名不改变 ID。

## 创建与生命周期

`createWorkbook({ sheets?, snapshot?, workerUrl?, workerFactory?, workerModules?, initializationTimeout?, requestTimeout?, historyLimit?, dateSystem? }): Promise<Workbook>`

默认创建 Sheet1，100,000 行、100 列。默认 1900 日期系统，历史最多 100 个事务。每个 Workbook 对应独立 Worker；不支持 Node.js 顶层 API 或 SSR 阶段创建。

`mountEditor(container, { workbook, locale?, readOnly?, toolbar?, sheetId?, zoom?, persistence?, styleNonce?, onError?, ... }): Editor`

`locale` 使用 `Locale` 类型：`zh-CN`（默认简体中文）、`zh-TW`（繁体中文）、`en-US`、`ja-JP`、`ko-KR`。`supportedLocales` 和 `Locale` 同时从 `onlineexcel` 与 `onlineexcel/editor` 导出；浏览器脚本版使用 `OnlineExcel.supportedLocales`。语言包内置，不依赖远程加载。JS 传入未知语言值时回退到 `zh-CN`。

语言在挂载时按实例确定；同页不同实例互不影响。不修改工作簿数据、工作表名称、数字格式、公式函数名称或参数名称。切换已有视图的语言可调用 `await editor.setOptions({locale:"ja-JP"})`，工作簿数据、撤销历史、选区与缩放保留。开发者 API 错误与兼容性诊断仍返回原始消息，界面会翻译常见操作错误。

```js
const editor = mountEditor(container, { workbook, locale: "ko-KR" });
await editor.ready;
```

`Editor.ready` 等待首屏加载；`select(sheetId, range)` 定位选区；`getSelection()` 获取当前范围；`resize()` 主动更新尺寸；`destroy()` 只销毁视图。调用 `Workbook.dispose()` 结束 Worker 并拒绝仍在等待的请求。

`zoom` 默认 1，支持 0.5～2；`getZoom()` 读取当前比例，`await setZoom(value)` 提交编辑后缩放网格与输入框。缩放不写入文件。可选 `persistence` 显示保存状态、重试及草稿恢复选择；必须属于同一个 Workbook。`createPersistence`、`createIndexedDBStorage` 及其状态/存储契约见 [SDK 自动保存指南](sdk.md#可选自动保存与草稿恢复)。

`getEditState()` 返回 `{ editing, dirty, pending }`，分别表示正在编辑、草稿不同于原值（或最近提交失败）、写入尚未完成。这是编辑器草稿状态，不是工作簿相对上次持久化的修改状态；宿主仍应使用 revision 追踪持久化。

`await editor.commitEdit()` 提交调用时的草稿，等待 Worker 写入和重算以及视图刷新。输入法组合尚未结束、写入失败或编辑器已销毁时拒绝 Promise；不要在失败后继续保存或切页。它不自动把宿主保存位置标记为已保存，也不阻止用户随后继续编辑。`cancelEdit()` 丢弃尚未提交的草稿，不撤回已经发给 Worker 的写入。

`await editor.destroy({ commit: true })` 在提交期间阻止新的界面输入，成功后释放视图；提交失败时保留视图并拒绝 Promise。原有无参数 `destroy()` 仍是同步清理，会丢弃草稿，适用于明确取消或框架卸载清理。需要保留输入时，应在路由离开/关闭处理函数中先等待提交，不要指望 React/Vue 的同步清理回调等待异步保存。

```js
await editor.commitEdit();
const snapshot = await workbook.exportJSON();
await saveSnapshot(snapshot); // 宿主的持久化函数

await editor.destroy({ commit: true });
await workbook.dispose();
```

## 工作簿操作

| 方法                                                                       | 行为                                                                                                                                                    |
| -------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `getMetadata()`                                                            | 工作表配置、revision、撤销状态、日期系统、命名区域                                                                                                      |
| `getNavigationTarget(sheetId, cell, direction)` | direction 为 up/down/left/right，返回下一数据边界的零基 `{row,column}`；跳过隐藏行列，识别公式和溢出结果 |
| `getRegion(sheetId, range)`                                                | 稀疏单元格、计算结果、原公式、展开后的样式、溢出归属                                                                                                    |
| `getValues(sheetId, range)`                                                | 二维计算结果，空白为 null，错误为 `{ error: '#REF!' }` 等对象                                                                                           |
| `setValues(sheetId, range, values, options?)`                              | 二维数据必须匹配范围；默认将等号开头的字符串作为公式；`parseFormulas: false` 保留文本                                                                   |
| `setFormula(sheetId, cell, formula)`                                       | 写入等号开头的公式，未知或无效公式保留并诊断                                                                                                            |
| `setStyle(sheetId, range, style)`                                          | 合并样式属性；用 false 关闭 bold/italic/underline/wrap                                                                                                  |
| `clear(sheetId, range, formats?)`                                          | 默认清内容保留格式；formats=true 同时清格式                                                                                                             |
| `addSheet(name, { rows?, columns? }?)`                                     | 返回新工作表 ID                                                                                                                                         |
| `deleteSheet`, `renameSheet`, `reorderSheet`                               | 管理工作表；禁止删除最后一张                                                                                                                            |
| `setDimensions(sheetId, axis, indexes, { size?, hidden? })`                | 尺寸单位为 CSS px；axis 为 row 或 column                                                                                                                |
| `insertRows`, `deleteRows`, `insertColumns`, `deleteColumns`               | 参数为 sheetId、index、count；调整引用与相关元数据                                                                                                      |
| `merge`, `unmerge`                                                         | 合并禁止静默丢弃非锚点数据；需先显式清空                                                                                                                |
| `freeze(sheetId, rows, columns)`                                           | 从左上角冻结指定行列数，0 表示取消                                                                                                                      |
| `sort(sheetId, range, keys, header?)`                                      | keys 包含绝对列索引和 asc/desc；支持多关键字；header=true 排除第一行                                                                                    |
| `filter(sheetId, range, rules)`                                            | 第一行为标题；跨列及比较条件采用 AND，每列最多两个；另支持单列值列表多选，见下文；range=undefined 清除                                                  |
| `copyRange(sheetId, range, targetSheetId, target, { cut?, valuesOnly? }?)` | 复制时平移相对引用；剪切更新引用；valuesOnly 只粘贴结果                                                                                                 |
| `fill(sheetId, source, target)`                                            | 重复区域、平移公式；一维等差数字源向两侧扩展，其他模式重复；保留源值                                                                                                          |
| `defineName(name, sheetId, range)`                                         | 工作簿级名称，不区分大小写                                                                                                                              |
| `find(sheetId, search, matchCase?)`                                        | 检索值，最多返回 1,000 个位置                                                                                                                           |
| `findNext(sheetId, search, { after?, matchCase?, entireCell? }?)`          | 按行、列顺序返回下一个匹配或 null；after 为 A1 地址，从其后搜索，末尾回绕；包括动态数组结果，不受 find 的 1,000 条上限限制                              |
| `getDataRegion(sheetId, range)`                                            | 在 Worker 内沿非空相邻边界扩展数据区域；裁掉选区超出已用数据的空白边缘，空表返回原选区；适合向用户建议排序范围                                          |
| `getDistinctValues(sheetId, range, column)`                                | 跳过首行标题，返回 `{ values: string[], truncated }`；忽略大小写去重，空白为 `""`；超过 2,000 种值时返回空列表并置 truncated=true，调用方应改用条件筛选 |
| `replace(sheetId, search, replacement, options?)`                          | 替换文本值，不修改公式源码；支持 matchCase / entireCell                                                                                                 |
| `transaction(commands, { signal? }?)`                                      | 顺序执行 `Command[]`，成功成为一个历史项，失败整体回滚                                                                                                  |
| `undo()`, `redo()`                                                         | 返回 ChangeEvent 或 null                                                                                                                                |

合并区域仅左上角存放内容。`setValues`、`setFormula` 及保留合并的仅值/公式粘贴，若尝试在其他位置写入非空内容，会整批拒绝；清空和设置样式仍可作用于整个区域。大范围 `setValues` 以完整请求范围检查动态数组保护，内部分段不会拆开原子事务；取消仍整体回滚。

行列结构调整和剪切的引用更新基于已解析公式。未知语法保留原文并报告，不能承诺自动修正其引用。`copyRange`、排序和填充仍拒绝复杂合并区域；编辑器快照粘贴支持完整合并区域，不能只粘贴合并区域的一部分。

值列表规则为 `{ column, operator: "in", values: ["Amy", "Zoe", ""] }`，同一列表内为 OR，跨列为 AND；同列不能混用值列表与比较条件。列表比较使用计算值的文本（数值、布尔、错误值转为文本，空值为空字符串），忽略大小写，不使用单元格数字格式生成显示值。空列表隐藏所有数据行。比较条件仍使用 `{ column, operator: "eq" | "neq" | "contains" | "gt" | "lt" | "gte" | "lte", value }`。

边框四边的空字符串表示清除对应边框，例如 `setStyle(id, "A1:B2", { border: { top: "" } })`；其他颜色仍须为 `#RGB` 或 `#RRGGBB`。

界面排序默认勾选“首行为标题”，支持选择当前范围或相邻数据、多关键字排序；无标题的数据需取消该选项。筛选对话框可逐列编辑条件或候选值，已启用筛选的列头提供入口。批量行高/列宽作用于选中的全部行/列；自动适应使用实际字体与格式化文本测量，跳过涉及合并单元格的行/列，最终尺寸受 8–4096 px 约束。自定义数字格式保留格式原文，渲染支持范围仍以兼容矩阵为准。

`transaction` 的 `paste` 命令接受 `source: Rect`、`cells: RegionCell[]`、目标 `sheetId`、`targetRow` / `targetColumn`，可选 `mode: "all" | "values" | "formulas" | "formats"`、`transpose`、`merges`。`formulas` 模式复制公式和常量但保留目标样式；`values` 固化计算结果。公式逐格按目标与原始地址的位移调整；转置改变布局，相对引用仍按单元格位移规则处理。

`paste` 还接受源地址坐标下的 `validations` 和 `conditionalFormats`。`all` 模式复制两类规则，`formats` 模式只复制条件格式，`values`/`formulas` 模式保留目标规则。省略规则字段保留目标已有规则，传空数组则清除实际粘贴单元格上的该类规则。规则会按可见索引与转置映射，隐藏间隔的规则和合并区域不受影响；部分覆盖合并区域会整批拒绝。数据、规则和样式共享一个撤销事务。普通 `copyRange` 也复制规则，剪切会移动规则，`valuesOnly` 保持目标规则。

`sourceRows/sourceColumns` 指定源范围内可见索引，`targetRows/targetColumns` 指定实际写入的目标索引，必须严格递增；用于跳过筛选或手动隐藏的行列。编辑器的复制与选区统计只包含可见行列；粘贴也跳过隐藏目标。同页面实例间内部剪贴板保留值、公式和基础样式，不依赖源区域后续变化；浏览器外部剪贴板支持文本和基本 HTML 表格（含 rowspan、colspan 和 br 换行），普通粘贴创建实际合并区域，值/样式/合并共同撤销；合并目标不连续时拒绝整批写入。外部公式形文本默认不执行。内部快照通过 HTML 来源标记匹配，不用相同文本猜测来源。剪切仅限同工作簿、源工作簿版本未改变、连续可见的源与目标；不满足时明确拒绝，源数据不变。仅提供纯文本剪贴板的环境会按外部文本处理。

编辑器普通复制还保存当时的验证与条件格式规则。外部普通粘贴保留自动识别的日期、百分比等数字格式，并支持 HTML 单元格内联粗体、斜体和水平对齐；未提供的样式沿用目标样式。外部“仅值”粘贴和受保护工作表的粘贴不改变格式。

## 工作表、名称与表单规则

- `duplicateSheet(sheetId, name): Promise<string>`：复制数据、公式、样式、尺寸、筛选和规则；显式指向源表的引用改指复制表。不复制全局名称；含保留 OOXML 对象的表目前拒绝复制。
- `setSheetHidden(sheetId, hidden)`：隐藏或恢复工作表；至少保留一个可见表。编辑器切换标签会恢复各表上次的选区和滚动位置。
- `renameName(name, newName)` / `deleteName(name)`：更新已解析公式中的名称；删除后的引用变成 `#REF!`，撤销恢复。未知语法仍按诊断范围处理。
- `setValidation(sheetId, range, rule?)`：设置/移除选区规则，保留选区外的规则。规则含 `type: "list" | "decimal" | "whole" | "date" | "textLength"`、`values?`、`minimum?`、`maximum?`、`allowBlank?`。日期上下限使用工作簿日期序列值。列表字符串匹配忽略大小写。修改值/公式时在事务重算后验证，失败整体回滚；设置规则不清除已有无效值。
- `setConditionalFormat(sheetId, range, rule?)`：基础值比较或文本包含规则，包含 `operator`、`value`、可选 `second`（between）与 `style`。界面显示派生样式；`getRegion()` 的 `style` 保留基础样式，`displayStyle` 提供条件结果。尚不支持色阶、数据条、图标集、任意公式条件。
- `protectSheet(sheetId, enabled?)`：启用/取消工作表编辑保护。默认单元格锁定，使用 `setStyle(id, range, { locked: false })` 先开放输入区。保护同时约束 API、编辑器、粘贴及替换；结构和样式修改需先取消保护。保护可被宿主 API 取消，不是加密或权限安全边界。

上述命令支持事务、撤销/重做和 JSON；基础规则、锁定及隐藏属性支持 XLSX。XLSX 列表验证使用内联字符串，最多 255 字符且项目不能包含逗号；超过边界进入有损诊断，JSON 保留完整列表。

## 文件和快照

- `importXlsx(Blob | ArrayBuffer | Uint8Array, { signal? }?)`：在临时工作簿中解析和重算，成功后替换当前数据并清空历史；失败/取消保持原数据。
- `exportXlsx({ allowLossy?, signal? }?)`：返回 `{ data: Uint8Array, diagnostics }`。存在已知内容损失时默认拒绝并携带 diagnostics，调用方确认后设置 allowLossy=true。
- `exportJSON()` / `importJSON(snapshot, { signal? }?)`：版本为 1 的可序列化快照，保存原始值、公式、样式、工作表属性及兼容性诊断，不保存操作历史或公式旧缓存。
- `importCsv(text, { sheetId?, start?, delimiter?, columns?, header?, signal? }?)`：默认从 A1 写入；未指定列类型仍保留为文本。`columns` 逐列指定 `auto/text/number/date/boolean/percent`，`header=true` 保持首行为文本；类型转换失败整批回滚。自动类型保留前导零和超过 15 位的数字标识；日期仅接受明确 ISO 格式。来源文件中的公式形文本不执行。界面提供编码、分隔符、十行预览和列类型选择。
- `exportCsv(sheetId, { range?, delimiter?, signal? }?)`：返回 `{ text, diagnostics }`，只导出单表计算值，带 UTF-8 BOM。
- `getDiagnostics()`：返回导入兼容性诊断和当前未知公式/循环依赖诊断。

从 `onlineexcel/xlsx` 可独立调用 `readXlsx(data, options)` 和 `writeXlsx(snapshot, options)`，options 支持 signal / onProgress。顶层 Workbook API 在 Worker 内执行这些操作。

取消是协作式的：文件分块和大型粘贴在检查点处理中止；同步计算、排序等阶段结束后检查取消并回滚，不能抢占一条正在执行的 JS 指令。

## 事件

`on(event, callback)` 返回取消订阅函数。事件包含：

- `change`：revision、修改的 sheetId/range、edit/undo/redo/import 来源。修改范围表示直接编辑范围，公式重算可影响其他位置；宿主渲染应按 revision 失效缓存。
- `selection`：sheetId、range，由编辑器发出。
- `calculation`：calculating / idle。
- `progress`：operation、stage、progress（0–1）。
- `diagnostics`：当前诊断列表。
- `error`：Worker 故障；请求自身的错误仍以 Promise rejection 返回。

回调异常不会阻断其他订阅者。原始 API 参数与结果按 Worker 消息规则复制；库不会转移并清空调用方传入的 XLSX buffer。

## SDK 扩展

[SDK 指南](sdk.md) 包含 `withOptions`、结构化错误、operationId/origin、`createSavePoint`、`setRecords`、`readRecords`、`readChunks`、`writeChunks`、`editor.setOptions`、工具栏/文件回调、Worker 函数与业务校验模块、React/Vue 适配器。

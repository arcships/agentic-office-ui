# 0.5.2 发布说明

发布日期：2026-07-13

`0.5.2` 统一了 PPTX 与 PDF/DOCX 的普通浏览体验，并把 `PptxStage` 的 Surface 交互事件从类型声明补全为可运行、可验证的公开行为。

## PPTX Surface

- `PptxViewer mode="browse"` 和最小 Surface 使用纵向连续页面；`mode="present"` 使用单页舞台，左右键执行上/下一播放步骤，跨页由播放控制器完成。
- `PptxPreviewSession` 默认使用列表渲染；播放文档会话使用单页渲染。`renderMode: "list" | "slide"` 可显式选择，两种模式共用同一次解包结果。
- 列表渲染支持窗口化挂载；滚动、跳页和缩放保持 `activeIndex` 同步。
- `PptxStage` 实际发出 `selection-change`、`object-click` 和 `context-menu`：事件包含 `slideIndex`，对象事件包含稳定 `objectKey`，右键事件同时包含视口和容器坐标。
- 浏览模式使用上/下方向键与 PageUp/PageDown；左/右方向键保留给播放模式。

## 0.5.2 验证

- 12 页正式组合素材在浏览态形成纵向可滚动页面，在播放态只挂载一个单页舞台。
- 页面选中、对象点击和对象右键均通过组件测试与真实 Chromium 交互验证。

## 0.5.1 发布记录

发布日期：2026-07-13

`0.5.1` 统一提升八个公开包的版本，承接 `0.5.0` 的 XLSX/PDF 搜索 API 交付与 npm 发布制品。

# 0.5.0 发布说明

发布日期：2026-07-13

`0.5.0` 新增四格式统一 Surface 组件系列、统一事件模型、自适应缩放，以及多项 bug 修复。

## 0.5.0 包版本

- `@arcships/docx-core@0.5.0`
- `@arcships/vue-docx@0.5.0`
- `@arcships/xlsx-core@0.5.0`
- `@arcships/vue-xlsx@0.5.0`
- `@arcships/vue-pdf@0.5.0`
- `@arcships/vue-ui@0.5.0`
- `@arcships/pptx-core@0.5.0`
- `@arcships/vue-pptx@0.5.0`

## 新增 Surface 组件

四个格式各有一个最小嵌入渲染面，供第三方项目自建 toolbar/控制栏：

| 组件 | 格式 | 说明 |
|---|---|---|
| `DocxDocumentSurface` | DOCX | 分页文档面 + 批注 gutter，无 toolbar/缩略图 |
| `XlsxSheetSurface` | XLSX | 网格 + 图表/图片/绘图/选区叠加层 + 右键菜单 |
| `PptxStage` | PPTX | 纵向连续幻灯片 Surface；播放时切换为单页舞台 |
| `PdfSurface` | PDF | 全部页面垂直堆叠滚动，无翻页切换 |

## 统一 Surface API

### 样式定制
- 每个 surface 暴露 `--xxx-surface-bg` CSS 变量（背景色）
- DOCX/PDF 支持 `fitWidth` prop：自适应容器宽度缩放

### 统一事件
- `contextMenu`：右键菜单触发，携带容器相对坐标 + 格式特有的定位信息（页码/cell 地址/sheet 名/PPTX 对象）
- `selectionChange`：选中内容变化（PPTX 返回幻灯片页码，DOCX/PDF 基于浏览器 Selection API，XLSX 基于 cell range）
- `objectClick`：图片/图表/形状点击

### 统一 expose
- `scrollToPage(pageIndex)`：DOCX/PDF
- `scrollContainer`：DOM 引用（DOCX）
- `zoom` / `rotation`：读写（PDF）

## Bug 修复

- **DOCX**：修复缩放双重计算（`getBoundingClientRect` 已含 transform，上层又乘 zoomFactor）和批注 gutter 横向布局不足
- **XLSX**：修复编辑模式下单字符键无法连续输入（grid 按键冒泡重复触发）、只读模式下默认允许行列缩放、右键不再取消已有选区
- **PPTX**：普通浏览改为纵向连续页面；补齐页面选中、对象点击和上下文菜单运行时事件
- **PDF**：新增垂直滚动模式，修复单页切换体验

## PDF 文字选择层（已知限制）

`@embedpdf` 引擎的 `getPageTextRects` 对 CID 字体 + 缺失 ToUnicode CMap 的中文 PDF 无法正确输出文字。详见 [pdf-text-layer-issue.md](docs/pdf-text-layer-issue.md)，`0.5.0` 不含此功能。

## Demo 更新

- 新增四格式 surface demo 页：`/docx-surface`、`/xlsx-surface`、`/pptx-surface`、`/pdf-surface`
- 每个 demo 展示宿主自建 toolbar + surface 事件监听 + 状态栏实时反馈
- Dev 端口调整为 5173
- 新增统一 Surface API 设计文档：`docs/unified-surface-api-design.md`

## 0.4.0 发布记录

发布日期：待发布

`0.4.0` 将八个公开包统一升级。本次主要增加 Vue PPTX 的最小组合接口，同时保持完整 `PptxViewer` 兼容。

## 0.4.0 包版本

- `@arcships/docx-core@0.4.0`
- `@arcships/vue-docx@0.4.0`
- `@arcships/xlsx-core@0.4.0`
- `@arcships/vue-xlsx@0.4.0`
- `@arcships/vue-pdf@0.4.0`
- `@arcships/vue-ui@0.4.0`
- `@arcships/pptx-core@0.4.0`
- `@arcships/vue-pptx@0.4.0`

## PPTX 最小组合接口

- 新增 `usePptxDocument`，负责文件、页面、缩放和文档会话生命周期。
- 新增 `usePptxPlayback`，负责动画步骤、播放状态、媒体和播放控制器生命周期。
- 新增 `PptxStage`，只提供实际渲染元素，不绑定工具栏和业务界面。
- `PptxViewer` 已改用相同能力，原有属性、事件、搜索、全屏和播放方法保持兼容。
- 新增真实浏览器验证，覆盖逐步动画、快速换文件、重新挂载、暂停恢复、隐藏页、媒体和双播放器隔离。

## 0.4.0 验证和回退

同一提交生成并验证八个 npm 压缩包。PPTX 还要通过完整组件、最小组合方式和工作区外浏览器消费验证。发现回归时，八个包统一恢复到已发布的 `0.3.0`。

## 0.3.0 发布记录

发布日期：2026-07-12
状态：八个公开包已发布到 npm

`0.3.0` 将八个公开包统一升级，并首次公开发布 `@arcships/pptx-core` 和 `@arcships/vue-pptx`。

## 0.3.0 包版本

- `@arcships/docx-core@0.3.0`
- `@arcships/vue-docx@0.3.0`
- `@arcships/xlsx-core@0.3.0`
- `@arcships/vue-xlsx@0.3.0`
- `@arcships/vue-pdf@0.3.0`
- `@arcships/vue-ui@0.3.0`
- `@arcships/pptx-core@0.3.0`
- `@arcships/vue-pptx@0.3.0`

## PPTX

- 提供静态预览、搜索、缩略图、隐藏页和页面导航。
- 提供单击动画、时间安排、属性轨道、页面切换、Morph、媒体、重播、跳页和全屏。
- 补丁后的渲染器代码已经包含在 `@arcships/pptx-core/browser` 中，消费项目不安装第三方渲染器或 pnpm 补丁；普通运行依赖由 npm 自动安装。
- `@arcships/vue-pptx/style.css` 是公开样式入口。
- 无法精确执行的内容通过能力报告标记为近似、静态或未解析；本版本不宣称完整兼容 PowerPoint 全部动画。

## 0.3.0 验证和回退

同一提交必须生成八个真实压缩包。PPTX 两包除安装、类型检查和构建外，还要在工作区外实际打开 PPTX，并确认第一次点击执行动画而不直接换页。回退时八包统一恢复到已经发布的 `0.2.0`；如果应用已经使用 PPTX，则需要同时移除 PPTX 接入或改用保存的 `0.3.0` 候选包。

升级步骤见[从 0.2.0 升级到 0.3.0](docs/migration-0.3.md)。

## 0.2.0 发布记录

发布日期：2026-07-12
状态：六个公开包已发布到 npm，`latest` 指向 `0.2.0`

`0.2.0` 是六个原有公开包的稳定化版本。以下内容保留该版本的接口、行为和迁移说明。

## 包版本

以下公开包必须作为同一批 `0.2.0` 制品验证和发布：

- `@arcships/docx-core`
- `@arcships/vue-docx`
- `@arcships/xlsx-core`
- `@arcships/vue-xlsx`
- `@arcships/vue-pdf`
- `@arcships/vue-ui`

`@arcships/office-runtime` 仍是私有工作区实现，不发布，也不应出现在公开包的运行依赖或类型声明中。

## 主要变化

### DOCX

- 新增实例级 `createDocxRuntime`，由实例持有 Worker、WASM、来源规则、资源限制、任务和诊断。
- 公开 `./runtime`、`./wasm-url`、`./worker` 和 `./assets/docx_wasm_bg.wasm`。
- Viewer 和 Editor 使用统一文档渲染面；编辑、撤销重做、导入导出和缩略图继续通过高层组件使用。
- 输入、压缩包、XML、关系、图片、模型、布局页数、解析时间、历史记录和缓存具有明确的实例限制与错误。

### XLSX

- 新增实例级 `createXlsxRuntime`，统一 Worker/WASM 配置、取消、超时、限制和诊断。
- 公开 `./runtime`、`./wasm-url`、`./worker` 和 `./assets/duke_sheets_wasm_bg.wasm`。
- Worker 失败不会静默在主线程重跑并伪装成功。
- `@arcships/vue-xlsx/chart`、`./map` 和 `./webgl` 为可选入口；普通表格不需要先下载可选渲染实现。
- 大表使用真实滚动范围、累计偏移与帧内合并绘制；工作簿子对象按实例生命周期释放。
- 历史记录、图片、对象 URL 和缩略图同时受条目数和字节数约束。

### PDF

- `PdfViewer` 使用受控 PDF 引擎显示真实文档，不再用“已验证、仅下载”冒充查看成功。
- 正常 PDF 可以查看、翻页、缩放、旋转、使用缩略图、搜索和下载。
- PDF Runtime、Worker、页面图片地址和任务归属具体实例；来源切换或卸载时释放。
- 公开 `./assets/pdfium.wasm` 和 `./style.css`。

### 通用组件

- `SignaturePad`、`FileUpload`、`FileThumbnail`、`BoundingBoxCitations`、`LayoutBlocks`、`Spinner`、`Tooltip` 等通用组件保留公开事件与错误类型。
- 公开 `./style.css`。

## 资源限制

DOCX 和 XLSX 的输入、归档、XML、图片、模型、时间、缓存与历史限制归属各自 Runtime 或控制器实例。超限错误提供稳定 `code` 以及适用时的 `actual`、`allowed`、`limit` 和 `phase`，不会泄露本地绝对路径。

PDF 只有一个公开的文件拒绝条件：整份文件体积 `maxFileSize`，默认 `50 MiB`，宿主可以调整。超过当前值返回 `PDF_TOO_LARGE`，并且不启动 PDF 引擎。页数、单页像素、总内存、并发页或人为复杂度上限不是 PDF 的额外显示前置条件。

## 错误和诊断

- DOCX：`SOURCE_NOT_ALLOWED`、`FETCH_FAILED`、`LIMIT_EXCEEDED`、图片错误、`WORKER_UNAVAILABLE`、`WASM_LOAD_FAILED`、`PARSE_FAILED`、`TIMEOUT`、`ABORTED` 和 `RUNTIME_DISPOSED`。
- XLSX：来源、获取、格式、限制、图片、超时、Worker、取消和 Runtime 销毁错误；Worker 另有 `XlsxWorkerErrorCode`。
- PDF：`SOURCE_NOT_ALLOWED`、`FETCH_FAILED`、`INVALID_PDF`、`PDF_TOO_LARGE` 和 `ABORTED`。

完整字段、事件和诊断类型见[公开接口合同](docs/api/public-api-contract.md)。界面文字可以调整，程序判断必须使用稳定错误码。

## 公开资源和样式

候选压缩包应直接提供 JS、类型、样式、Worker 和 WASM。应用不得从 demo 公共目录或工作区 `dist` 手工复制资源。四个 Vue 包的新代码统一使用：

```ts
import "@arcships/vue-docx/style.css"
import "@arcships/vue-xlsx/style.css"
import "@arcships/vue-pdf/style.css"
import "@arcships/vue-ui/style.css"
```

## 弃用与兼容

以下旧入口从 `0.2.0` 开始标记弃用，但整个 `0.x` 继续保留，最早只能在 `1.0.0` 删除：

- DOCX/XLSX 根入口中的全局 `setWasmSource` 及相关默认运行配置。
- `vue-docx` 已经公开的低层页面、段落、表格、图片、工具栏和旧缩略图入口。
- `vue-xlsx` 已经公开的低层网格、工具栏、浮层和渲染组件。

新代码使用 Runtime 实例、高层 Viewer/Editor 和 `./style.css`。完整替代关系见[迁移说明](docs/migration-0.2.md)。

## 升级

不要只升级一个 Vue 包而保留旧核心包。相关包应统一升级到 `0.2.0`，重新接入公开样式和 Worker/WASM 资源，再执行类型、正式构建和浏览器验证。详细步骤见[从 0.1.x 迁移到 0.2.0](docs/migration-0.2.md)。

## 回退办法

1. 发现回归时先停止发布或灰度，不继续生成新制品。
2. `0.2.0` 仍保留旧公开入口；可以先把新调用切回兼容入口，确认问题是否来自迁移代码，不必立即降级。
3. 必须降级时，恢复上一批已经测试并保存 SHA-256 的六包制品；相关 Vue 包和核心包要作为一组恢复，不能混合版本。
4. 同时恢复该批制品对应的 Worker/WASM 地址和样式入口，不得把当前 `dist` 文件复制到旧项目冒充回退成功。
5. 用正式 preview 复跑受影响流程；只通过类型或构建不能结束回退。

`0.2.0` 的六个包已经由 GitHub Actions 发布，并从公共 npm 完成全新安装和导入验证。

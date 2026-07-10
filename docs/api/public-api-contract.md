# 五个公开包的接口合同

本文登记 `@extend-ai/docx-core`、`@extend-ai/xlsx-core`、`@extend-ai/vue-docx`、`@extend-ai/vue-xlsx` 和 `@extend-ai/vue-extend` 的公开入口、推荐用法、事件、错误与兼容期限。它用于约束 `0.2.0` 及后续版本的实现和发布检查。

当前源码清单中的版本仍是 `0.1.0`。版本号只能在正式发布流程中修改；本文和源码中的 `@deprecated Since 0.2.0` 表示从 `0.2.0` 开始执行的兼容安排，不代表已经发布了 `0.2.0`。

## 1. 适用规则

- 只有各包 `package.json` 的 `exports` 列出的路径是受支持的导入路径。`dist` 内其他文件和源码路径不属于公开入口。
- 根入口中已经导出的名称属于公开接口，即使名称看起来像内部组件，也不能在 `0.x` 中直接删除。
- 新代码优先使用高层组件、实例 Runtime 和 `./core` 纯数据入口。旧的全局 WASM 配置只用于兼容。
- 事件名、事件参数、错误码、属性默认值、Worker/WASM 地址和样式入口都属于合同。
- 已公开但准备收口的接口从 `0.2.0` 标记弃用，整个 `0.x` 保留，最早在 `1.0.0` 删除。删除前必须提供迁移示例，并用真实发布压缩包同时验证新旧入口。
- 从未导出的临时桥接不受公开接口版本周期保护，但必须登记负责人和删除日期。清理内部桥接不得导致公开兼容入口提前消失。

## 2. 精确导出路径

下表与五个包当前 `package.json` 的 `exports` 一一对应。未列出的深层路径必须由 Node 返回 `ERR_PACKAGE_PATH_NOT_EXPORTED`。

| 包 | 公开路径 | 类型文件 | 运行文件或资源 |
|---|---|---|---|
| `@extend-ai/docx-core` | `.` | `./dist/index.d.ts` | `./dist/index.js` |
|  | `./core` | `./dist/core.d.ts` | `./dist/core.js` |
|  | `./runtime` | `./dist/runtime.d.ts` | `./dist/runtime.js` |
|  | `./wasm-url` | `./dist/wasm-url.d.ts` | `./dist/wasm-url.js` |
|  | `./assets/docx_wasm_bg.wasm` | 不适用 | `./dist/assets/docx_wasm_bg.wasm` |
|  | `./worker` | 不适用 | `./dist/docx-import-worker.js` |
|  | `./package.json` | 不适用 | `./package.json` |
| `@extend-ai/xlsx-core` | `.` | `./dist/index.d.ts` | `./dist/index.js` |
|  | `./core` | `./dist/core.d.ts` | `./dist/core.js` |
|  | `./runtime` | `./dist/runtime.d.ts` | `./dist/runtime.js` |
|  | `./wasm-url` | `./dist/wasm-url.d.ts` | `./dist/wasm-url.js` |
|  | `./assets/duke_sheets_wasm_bg.wasm` | 不适用 | `./dist/assets/duke_sheets_wasm_bg.wasm` |
|  | `./worker` | 不适用 | `./dist/xlsx-worker.js` |
|  | `./package.json` | 不适用 | `./package.json` |
| `@extend-ai/vue-docx` | `.` | `./dist/index.d.ts` | `./dist/index.js` |
|  | `./style.css` | 不适用 | `./dist/style.css` |
| `@extend-ai/vue-xlsx` | `.` | `./dist/index.d.ts` | `./dist/index.js` |
|  | `./style.css` | 不适用 | `./dist/style.css` |
| `@extend-ai/vue-extend` | `.` | `./dist/index.d.ts` | `./dist/index.js` |
|  | `./assets/pdfium.wasm` | 不适用 | `./dist/pdfium.wasm` |
|  | `./style.css` | 不适用 | `./dist/style.css` |
|  | `./dist/index.css` | 不适用 | 兼容别名，实际指向 `./dist/style.css` |

Worker 和 WASM 必须来自安装后的真实包。消费项目不得从 demo 公共目录复制这些文件，也不得通过工作区源码别名绕过 `exports`。

## 3. 推荐的稳定入口

### 3.1 `@extend-ai/docx-core`

- `@extend-ai/docx-core/core`：平台无关的文档模型、标准化、克隆、布局及不可变段落、表格和文字样式命令。
- `@extend-ai/docx-core/runtime`：`createDocxRuntime`、`DocxRuntime`、`DocxRuntimeConfig`、来源与限制类型，以及 DOCX 导入的公开诊断和错误类型。
- `@extend-ai/docx-core/wasm-url`：包自带的 `bundledDocxWasmUrl`。
- 根入口：保留现有模型、布局、序列化、解析和编辑命令，供 `0.x` 兼容使用。新代码如只需要纯数据能力，应改用 `./core`；需要 Worker/WASM 时应改用 `./runtime`。

推荐实例写法：

```ts
import { createDocxRuntime } from "@extend-ai/docx-core/runtime"
import { bundledDocxWasmUrl } from "@extend-ai/docx-core/wasm-url"

const runtime = createDocxRuntime({
  wasmUrl: bundledDocxWasmUrl,
  createWorker: () => new Worker(workerUrl, { type: "module" }),
})

try {
  const result = await runtime.loadSource({ kind: "bytes", bytes })
  // 使用 result.model
} finally {
  runtime.dispose()
}
```

### 3.2 `@extend-ai/xlsx-core`

- `@extend-ai/xlsx-core/core`：平台无关的工作簿数据类型、颜色解析、A1 地址、图表公式和图片锚点计算。
- `@extend-ai/xlsx-core/runtime`：`createXlsxRuntime`、实例配置、Worker 客户端、来源检查、加载错误和诊断类型。
- `@extend-ai/xlsx-core/wasm-url`：包自带的 `bundledXlsxWasmUrl`。
- 根入口：保留现有数据、图片、图表、Worker 和控制器类型供 `0.x` 兼容使用。新代码应按纯数据或运行能力选择 `./core`、`./runtime`。

推荐实例写法：

```ts
import { createXlsxRuntime } from "@extend-ai/xlsx-core/runtime"
import { bundledXlsxWasmUrl } from "@extend-ai/xlsx-core/wasm-url"

const runtime = createXlsxRuntime({
  wasmSource: bundledXlsxWasmUrl,
  createWorker: () => new Worker(workerUrl, { type: "module" }),
})

try {
  const worker = runtime.createWorkerClient()
  // 通过实例拥有的 Worker 执行解析
} finally {
  runtime.dispose()
}
```

### 3.3 `@extend-ai/vue-docx`

稳定高层入口为：

- 组件：`DocxViewer`、`DocxEditor`。
- 状态与命令：`useDocxEditor`、`useDocxModel`。
- 公共能力组合函数：文档主题、段落样式、图片环绕、行距、边框、表单域、修订、批注、页面布局、分页和 `useDocxPageThumbnails`。
- 类型：`DocModel`、`DocxEditorController`、编辑选区、图片、表格、分页、修订、批注及上述组合函数的公开类型。
- 样式：`@extend-ai/vue-docx/style.css`。

页面、段落、表格、图片、工具栏和浮层的单独组件虽然在旧版本已经导出，但不再作为新功能的接入点。它们的兼容期限见第 6 节。

### 3.4 `@extend-ai/vue-xlsx`

稳定高层入口为：

- `XlsxViewer`。
- `useXlsxViewerController`、`useXlsxViewerThumbnails`。
- `XlsxFileSizeLimitExceededError`。
- `XlsxDiagnostic`、`XlsxLoadError`、`XlsxLoadErrorCode`、`XlsxSourceKind`、`XlsxSourceState`、`XlsxUrlPolicy`。
- 样式：`@extend-ai/vue-xlsx/style.css`。

网格、工具栏、功能区、公式栏、工作表标签、图表/图片/选区浮层、菜单及单独渲染函数是旧的低层出口，兼容期限见第 6 节。

### 3.5 `@extend-ai/vue-extend`

稳定入口为 `PdfViewer`、`SignaturePad`、`FileUpload`、`FileThumbnail`、`BoundingBoxCitations`、`LayoutBlocks`、`Spinner`、`Tooltip`，以及根入口导出的属性、PDF 来源、诊断、错误和 `FileUploadRejection` 类型。PDF 查看器默认创建自己专用的 `PdfRenderRuntime`，也可由宿主通过 `runtime` 属性注入实例；`pdfiumWasmUrl` 可覆盖包内默认资源地址。`createPdfRenderRuntime`、`bundledPdfiumWasmUrl`、`PdfRenderRuntimeConfig` 和页面、搜索、渲染类型是公开的实例配置入口。`DEFAULT_PDF_MAX_FILE_SIZE`、`PdfLoadOptions.maxFileSize` 和 `PdfViewer.maxFileSize` 是公开的整份文件体积配置入口。新代码使用 `@extend-ai/vue-extend/style.css`；`@extend-ai/vue-extend/dist/index.css` 只作为旧项目的兼容别名。

PDF Runtime 只接收已经取得并验证的 `ArrayBuffer`，不会再次请求原始 PDF 地址。默认使用包内 PDFium Worker，关闭外部字体回退；Worker 或 WASM 初始化失败会返回错误，不会把验证字节状态当成渲染成功。宿主自带严格内容安全策略时，需要允许 `worker-src blob:`，并允许读取 `@extend-ai/vue-extend/assets/pdfium.wasm` 对应的同源资源。

PDF 唯一的资源拒绝配置是整份文件体积 `maxFileSize`：默认 `50 MiB`，宿主可公开调整，不叠加第二层隐藏硬上限。本地 Blob 在读取前检查声明大小；URL 有可信 `Content-Length` 时先检查响应头，否则在取得正文后、创建 PDF 引擎前检查实际字节。超过当前值返回 `PDF_TOO_LARGE`，包含 `actual` 和 `allowed`，并且不显示旧页面或创建新页面。页数、单页像素、总内存和并发页不是公开拒绝条件。

## 4. 公开事件和诊断

### 4.1 DOCX

| 来源 | 事件或诊断 | 参数或关键字段 |
|---|---|---|
| `DocxViewer` | `load-start` | 无参数 |
| `DocxViewer` | `load-success` | `DocxImportResult` |
| `DocxViewer` | `load-error` | `Error`；若为导入错误，可读取公开 `code` |
| `DocxRuntimeConfig.onDiagnostic`、单次导入的 `onDiagnostic` | `load-start`、`worker-created`、`worker-success`、`worker-error`、`main-thread-start`、`main-thread-success`、`main-thread-fallback`、`aborted` | `requestId` 必有；可含 `runtimeId`、执行来源、脱敏后的资源地址、错误码、消息和耗时 |

诊断只提供执行信息，不得包含文档正文、压缩包内容或未脱敏的输入地址。诊断回调抛错不得改变加载结果。

### 4.2 XLSX

| 来源 | 事件或诊断 | 参数或关键字段 |
|---|---|---|
| `XlsxViewer` | `cellDoubleClick` | `XlsxCellAddress` |
| `UseXlsxViewerControllerOptions.onDiagnostic` | `load-start`、`load-deferred`、`load-resumed`、`load-success`、`load-error`、`load-cancelled`、`download` | `requestId` 必有；可含 `taskId`、来源、脱敏地址、字节数和 `XlsxLoadError` |
| `XlsxRuntimeConfig.onDiagnostic` | `worker-client-created`、`worker-client-disposed`、`runtime-disposed` | `runtimeId` 必有；可含 Worker/WASM 地址 |

### 4.3 `vue-extend`

| 组件 | 事件 | 参数 |
|---|---|---|
| `PdfViewer` | `document-load-success` | 页数 |
| `PdfViewer` | `document-load-error` | `PdfLoadError` |
| `PdfViewer` | `active-page-change` | 当前页码，从 1 开始 |
| `PdfViewer` | `diagnostic` | `PdfDiagnostic`，加载类型为 `load-start`、`load-success`、`load-error`、`load-cancelled`；页面类型为 `render-start`、`render-success`、`render-error`；搜索类型为 `search-start`、`search-success`、`search-error`；另有 `object-url-revoked` 和 `download`。可读取 `runtimeId`、`taskId`、页码、页数、缩放、旋转、匹配数和结构化错误 |
| `FileUpload` | `files-accepted` | 通过检查的 `File[]` |
| `FileUpload` | `files-rejected` | `FileUploadRejection[]`；每项包含 `code`、被拒绝的 `file` 和可显示的 `message` |
| `SignaturePad` | `update:signature` | 数据地址；空签名为 `null` |
| `BoundingBoxCitations` | `field-click` | `BoundingBoxField` |
| `LayoutBlocks` | `block-click` | `LayoutBlock` |

`FileThumbnail`、`Spinner` 和 `Tooltip` 当前没有业务事件。文件拒绝使用 `files-rejected`；宿主不能解析界面文字推断原因。

## 5. 公开错误

| 包 | 类型 | 错误码 | 含义 |
|---|---|---|---|
| `docx-core` | `DocxImportError` / `DocxImportErrorCode` | `ABORTED` | 调用方取消或旧任务失效 |
|  |  | `STALE_RESULT` | 低层导入结果已过期；Runtime 对外通常归一为 `ABORTED` |
|  |  | `SOURCE_NOT_ALLOWED` | 地址或来源不符合公开规则 |
|  |  | `FETCH_FAILED` | 受控获取失败 |
|  |  | `LIMIT_EXCEEDED` | 超过当前输入限制 |
|  |  | `WORKER_UNAVAILABLE` | Worker 无法创建或不可用 |
|  |  | `WASM_LOAD_FAILED` | WASM 地址、字节或初始化失败 |
|  |  | `PARSE_FAILED` | 文件无法解析为受支持的 DOCX |
|  |  | `RUNTIME_DISPOSED` | 已销毁的实例被再次使用 |
| `xlsx-core` / `vue-xlsx` | `XlsxLoadError` / `XlsxLoadErrorCode` | `SOURCE_NOT_ALLOWED`、`FETCH_FAILED`、`INVALID_WORKBOOK`、`ABORTED` | 分别表示来源被拒绝、获取失败、工作簿无效和取消 |
| `xlsx-core/runtime` | `XlsxRuntimeError` / `XlsxRuntimeErrorCode` | `RUNTIME_DISPOSED` | 已销毁的 XLSX Runtime 被再次使用；错误同时保留 `format: "xlsx"` 和 `runtimeId` |
| `vue-xlsx` | `XlsxFileSizeLimitExceededError` | 类的公开字段 | 本地输入超过控制器允许的文件大小 |
| `vue-extend` | `PdfSourceError` / `PdfLoadErrorCode` | `SOURCE_NOT_ALLOWED`、`FETCH_FAILED`、`INVALID_PDF`、`PDF_TOO_LARGE`、`ABORTED` | 分别表示来源被拒绝、获取失败、PDF 无效、整份文件超过当前 `maxFileSize` 和取消；体积错误同时提供 `actual`、`allowed` |
| `vue-extend` | `FileUploadRejection` / `FileUploadRejectionCode` | `FILE_TYPE_NOT_ACCEPTED`、`FILE_TOO_LARGE` | 文件类型不符合 `accept`，或文件超过 `maxSize` |

公开错误必须保留稳定的 `code`，界面文案可以改进但不能替代错误码。地址写入错误和诊断前必须脱敏。Worker、WASM 或解析失败不得静默转成成功；只有调用方明确开启且文件小于已配置上限时，DOCX 才能进行主线程回退，并必须发出 `main-thread-fallback` 诊断。

## 6. 兼容登记

登记字段中的 `kind` 只有 `public-api` 和 `internal-bridge` 两种。负责人是代码所有者，不表示某次会话的独立验收人。

| kind | 入口或实现 | 引入版本/日期 | 负责人 | 替代入口 | 删除期限 |
|---|---|---|---|---|---|
| `public-api` | `@extend-ai/docx-core` 根入口的 `setWasmSource` | `<=0.1.0` | `docx-core` Runtime | `createDocxRuntime({ wasmSource })` 或 `{ wasmUrl }` | `removeInVersion: 1.0.0`，整个 `0.x` 保留 |
| `public-api` | `@extend-ai/xlsx-core` 根入口的 `setWasmSource`、`canUseConfiguredWasmSourceInWorker`、`getConfiguredWorkerWasmSource`、`getSheetsWasmModule` | `<=0.1.0` | `xlsx-core` Runtime | `createXlsxRuntime({ wasmSource, createWorker })` | `removeInVersion: 1.0.0`，整个 `0.x` 保留 |
| `public-api` | `DocxEditorViewer` | `<=0.1.0` | `vue-docx` | `DocxEditor` | `removeInVersion: 1.0.0`，整个 `0.x` 保留 |
| `public-api` | `vue-docx` 已导出的页面、段落、表格、图片、字段、修订、菜单、工具栏、缩略图和拖放组件 | `<=0.1.0` | `vue-docx` | `DocxViewer`、`DocxEditor` | `removeInVersion: 1.0.0`，整个 `0.x` 保留 |
| `public-api` | `vue-docx` 的静态渲染函数和页面 surface registry | `<=0.1.0` | `vue-docx` | `DocxViewer`、`DocxEditor` | `removeInVersion: 1.0.0`，整个 `0.x` 保留 |
| `public-api` | `vue-docx` 的 `useDocxViewerThumbnails`、`UseDocxViewerThumbnailsOptions`、`DocxViewerThumbnails` | `<=0.1.0` | `vue-docx` | `useDocxPageThumbnails`、`UseDocxPageThumbnailsOptions`、`UseDocxPageThumbnailsResult` | `removeInVersion: 1.0.0`，整个 `0.x` 保留 |
| `public-api` | `vue-xlsx` 已导出的低层组件、`MemoChartSvg`、`MemoSurfaceChartComposite` 及四个对应渲染类型 | `<=0.1.0` | `vue-xlsx` | `XlsxViewer`、`useXlsxViewerController` | `removeInVersion: 1.0.0`，整个 `0.x` 保留 |
| `public-api` | `@extend-ai/vue-extend/dist/index.css` | `<=0.1.0` | `vue-extend` | `@extend-ai/vue-extend/style.css` | `removeInVersion: 1.0.0`，整个 `0.x` 保留 |
| `internal-bridge` | `docx-core/src/viewer/wasm-source.ts` 的 `createLegacyWorkerWasmBridge` / `legacyWorkerWasmBridge` | `2026-07-10` | `docx-core` Runtime | 由正式默认 `DocxRuntime` 适配器接管；公开旧入口继续保留 | `removeByDate: 2026-08-09` |
| `internal-bridge` | `xlsx-core/src/wasm.ts` 的 `createLegacyDefaultWasmRuntime` / `legacyDefaultWasmRuntime` | `2026-07-10` | `xlsx-core` Runtime | 由正式默认 `XlsxRuntime` 适配器接管；公开旧入口继续保留 | `removeByDate: 2026-08-09` |

如果内部桥接到期时仍有必要保留，必须在到期前把原因、负责人和新日期写回本表；不能用内部桥接的期限提前删除公开入口。

## 7. 发布门禁

每次改变公开接口或发布资源时至少检查：

1. 从当前源码构建，并对五个公开包执行真实 `npm pack`。
2. 在工作区外只安装生成的五个压缩包，不使用源码别名，不复制 demo 中的 Worker/WASM。
3. 检查本页列出的所有 `exports` 都能解析，未列出的代表性深层路径都被拒绝。
4. 检查根入口、`./core`、`./runtime`、`./wasm-url`、三个 `style.css`、两个独立 Worker 资源、三个 WASM，以及 PDF 包内生成的 Worker 代码。
5. 检查生成的 `.d.ts` 保留本页登记的 `@deprecated`，并且公开声明不引用私有 `@extend-ai/office-runtime` 或不受支持的深层路径。
6. 用正式 preview 验证 Worker/WASM 请求来自压缩包、无 404、WASM 类型和字节签名正确。
7. 黑盒消费页只通过本文公开入口、公开事件和公开错误观察结果，不读取模块私有变量。

本会话的开发后自测只能作为同会话复测；候选发布仍需另一个全新会话执行 `BB-RELEASE` 独立复核。

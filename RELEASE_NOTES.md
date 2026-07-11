# 0.2.0 候选发布说明

发布日期：尚未发布
候选状态：`@arcships` 六包已完成本会话发布自测，尚未发布

`0.2.0` 是六个公开包的稳定化候选版本。本文说明当前候选代码已经提供的接口和行为，不是发布公告。当前工作区已经用同一批 `@arcships` 真实压缩包完成可复现构建、工作区外消费、兼容矩阵和完整 `BB-RELEASE` 自测；但这仍是实现者会话结果，npm 组织权限、首次发布流程和全新会话独立复核完成前，不得把本版本描述为正式发布。

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

## 发布前仍需完成

- 重新登录 npm，并确认当前账号对 `@arcships` 组织具有六个新包的发布权限。
- 确定最终公开 GitHub 仓库，为六包填写与真实仓库一致的 `repository` 信息。
- 完成六个新包的首次受控发布，再逐包配置可信发布；同时把发布环境升级到 npm 当前支持的 Node/npm 版本。
- 处理候选标签晋级所需的合法凭证或改用 npm 支持的发布审批流程，不能假设可信发布身份可以直接执行 `dist-tag`。
- 全新会话对候选制品执行 `BB-RELEASE` 并独立复核。

上述条件完成前，本文件始终是候选说明，不是正式发布记录。

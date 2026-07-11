# Agentic Office UI

面向 Vue 3 的 DOCX、XLSX、PPTX 和 PDF 查看与编辑组件工作区。

## 当前状态

五个公开包的当前候选版本为 `0.2.0`，但**尚未发布**。仓库内的开发后自测不能替代候选发布复核；只有同一批真实压缩包完成 `BB-RELEASE`，并由全新会话独立复核后，才能给出发布结论。

当前状态、设计约束和验收方法分别以以下文档为准：

- [目标架构](docs/architecture-review-and-target-design.md)
- [稳定化路线图](docs/plan/stabilization-roadmap.md)
- [黑盒验收方案](docs/end-to-end-blackbox-test-plan.md)
- [Agent 执行手册](docs/testing/agent-execution-runbook.md)
- [公开接口合同](docs/api/public-api-contract.md)

完整文档入口见 [docs/INDEX.md](docs/INDEX.md)。旧迁移计划和上游对齐清单只保存历史背景，不能用来判断当前完成状态。

## 公开包

| 包 | 主要用途 |
|---|---|
| `@arcships/docx-core` | DOCX 模型、布局、命令、实例 Runtime、Worker 和 WASM 资源 |
| `@arcships/vue-docx` | `DocxViewer`、`DocxEditor` 和 Vue 组合函数 |
| `@arcships/xlsx-core` | XLSX 模型、公式、图片、图表、实例 Runtime、Worker 和 WASM 资源 |
| `@arcships/vue-xlsx` | `XlsxViewer`、控制器，以及按需图表、地图和 WebGL 入口 |
| `@arcships/vue-extend` | 可操作的 PDF 查看器，以及上传、签名、缩略图和版面组件 |

`@arcships/office-runtime` 是工作区私有实现，不是第六个公开包，消费端不应安装或导入它。

PPTX 网页预览与播放正在独立开发：`@arcships/pptx-core` 和 `@arcships/vue-pptx` 已提供文件加载、浏览、搜索、缩放、动画、页面切换、媒体、演示控制和全屏接口，但尚未进入 `0.2.0` 公开包合同。开发与验收入口见 [PPTX 播放开发准备与入口](docs/pptx-development-guide.md)。

## 使用候选版本

`0.2.0` 尚未发布到公共源。仓库开发使用冻结锁文件：

```bash
pnpm install --frozen-lockfile
pnpm typecheck
pnpm build
```

正式发布后，应用只安装自己需要的公开包，并保持相关 Vue 包与核心包为同一个 `0.2.x` 版本。例如：

```bash
pnpm add @arcships/docx-core@0.2.0 @arcships/vue-docx@0.2.0
pnpm add @arcships/xlsx-core@0.2.0 @arcships/vue-xlsx@0.2.0
pnpm add @arcships/vue-extend@0.2.0
```

不要混用 `0.1.x` 核心包和 `0.2.x` Vue 包，也不要从工作区源码或 demo 公共目录复制 Worker/WASM 来绕过包出口。

### 样式

三个 Vue 包只使用公开样式入口：

```ts
import "@arcships/vue-docx/style.css"
import "@arcships/vue-xlsx/style.css"
import "@arcships/vue-extend/style.css"
```

### Worker 和 WASM

需要显式资源地址时，从已安装包的公开入口交给构建工具处理：

```ts
import docxWorkerUrl from "@arcships/docx-core/worker?worker&url"
import docxWasmUrl from "@arcships/docx-core/assets/docx_wasm_bg.wasm?url"
import { createDocxRuntime } from "@arcships/docx-core/runtime"

const docxRuntime = createDocxRuntime({
  workerUrl: docxWorkerUrl,
  wasmUrl: docxWasmUrl,
})
```

XLSX 提供对应的 `@arcships/xlsx-core/worker` 和 `@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm`。完整写法见 [0.2.0 迁移说明](docs/migration-0.2.md)。

## PDF 行为

`PdfViewer` 使用受控 PDF 引擎显示真实页面。正常 PDF 可以查看、翻页、缩放、旋转、使用缩略图、搜索和下载；安全检查不会用“只能下载、不能查看”代替产品功能。

PDF 只有一个公开的文件拒绝上限：整份文件体积 `maxFileSize`，默认 `50 MiB`，宿主可以调整。超过当前值返回 `PDF_TOO_LARGE`，并且不启动引擎。页数、单页像素、总内存和并发页不是额外的显示前置条件。

DOCX 和 XLSX 另有实例级输入、压缩包、XML、图片、模型、历史记录和运行时间限制。错误码与配置项以[公开接口合同](docs/api/public-api-contract.md)为准。

## 验证命令

PPTX 日常开发可以先运行：

```bash
pnpm typecheck:pptx
pnpm build:pptx
pnpm test:pptx
```

完整仓库验证命令如下：

```bash
pnpm typecheck
pnpm build
pnpm test:unit
pnpm test:component
pnpm test:blackbox
pnpm test:stress
pnpm test:performance
pnpm test:consumer
pnpm test:docs
pnpm check
```

浏览器结论必须来自正式构建 preview。真实发布包结论必须来自工作区外安装的 `.tgz`，不能使用源码别名。兼容范围和待补测试见[兼容矩阵](docs/testing/compatibility-matrix.md)。

## 版本资料

- [0.2.0 候选发布说明](RELEASE_NOTES.md)
- [从 0.1.x 迁移到 0.2.0](docs/migration-0.2.md)
- [公开接口与弃用期限](docs/api/public-api-contract.md)

## 上游归属

本项目是面向 Vue 的实现，设计参考公开的 Extend UI / Extend AI React 包。原项目、包名、接口思路和设计方向属于 Extend 及其维护者。详细来源和同步规则见 [docs/upstream-extend-ui.md](docs/upstream-extend-ui.md)。

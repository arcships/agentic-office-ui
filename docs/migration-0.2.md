# 从 0.1.x 迁移到 0.2.0

## 版本状态

`0.2.0` 当前是候选版本，尚未发布。本文用于准备迁移和验证候选压缩包；公共源正式可用前，不要把示例中的版本号解释为已经发布。

## 升级原则

1. 相关 Vue 包和核心包统一使用 `0.2.0`，不要混用 `0.1.x` 与 `0.2.x`。
2. 只从 `package.json` 的公开 `exports` 导入，不导入源码或猜测 `dist` 文件名。
3. Worker、WASM 和样式必须来自安装后的包，不能从 `apps/demo/public` 复制。
4. 先保留旧入口完成迁移，再逐步切换到 Runtime 实例；旧公开入口在整个 `0.x` 保留，最早在 `1.0.0` 删除。
5. 类型和构建通过后，仍要使用正式 preview 验证正常、错误、取消和恢复路径。

## 1. 统一包版本

正式发布后，按应用需要安装对应包：

```bash
pnpm add @arcships/docx-core@0.2.0 @arcships/vue-docx@0.2.0
pnpm add @arcships/xlsx-core@0.2.0 @arcships/vue-xlsx@0.2.0
pnpm add @arcships/vue-extend@0.2.0
```

`@arcships/office-runtime` 是私有实现，不要加入应用依赖。

## 2. 使用公开样式入口

将内部 `dist` 样式路径改为稳定入口：

```ts
import "@arcships/vue-docx/style.css"
import "@arcships/vue-xlsx/style.css"
import "@arcships/vue-extend/style.css"
```

`@arcships/vue-extend/dist/index.css` 在 `0.x` 中仍能使用，但新代码应改为 `@arcships/vue-extend/style.css`。

## 3. DOCX 改用实例 Runtime

旧代码可能在模块加载前调用全局 `setWasmSource`。`0.2.0` 保留该入口，但新代码应把 Worker、WASM、限制和诊断交给当前 Runtime 实例：

```ts
import docxWorkerUrl from "@arcships/docx-core/worker?worker&url"
import docxWasmUrl from "@arcships/docx-core/assets/docx_wasm_bg.wasm?url"
import { createDocxRuntime } from "@arcships/docx-core/runtime"

const runtime = createDocxRuntime({
  workerUrl: docxWorkerUrl,
  wasmUrl: docxWasmUrl,
  onDiagnostic(event) {
    console.debug(event.type, event.requestId)
  },
})

try {
  const result = await runtime.loadSource({
    kind: "bytes",
    bytes: docxBytes,
    name: "document.docx",
  })
  // 使用 result.model
} finally {
  runtime.dispose()
}
```

Vite 配置保留 ES Worker：

```ts
import vue from "@vitejs/plugin-vue"
import { defineConfig } from "vite"

export default defineConfig({
  plugins: [vue()],
  resolve: { dedupe: ["vue"] },
  worker: { format: "es" },
})
```

如果应用只使用高层 `DocxViewer` 或 `DocxEditor`，组件会创建自己的默认 Runtime；需要自定义资源、限制或诊断时再显式注入公开配置。

## 4. XLSX 改用实例 Runtime

不要依赖模块级全局 WASM 状态，也不要在 Worker 失败后自行把大任务静默放回主线程：

```ts
import xlsxWorkerUrl from "@arcships/xlsx-core/worker?worker&url"
import xlsxWasmUrl from "@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm?url"
import { createXlsxRuntime } from "@arcships/xlsx-core/runtime"

const runtime = createXlsxRuntime({
  workerUrl: xlsxWorkerUrl,
  wasmSource: xlsxWasmUrl,
  onDiagnostic(event) {
    console.debug(event.type, event.runtimeId)
  },
})

try {
  const worker = runtime.createWorkerClient()
  // 使用实例拥有的 Worker 客户端
  worker.dispose()
} finally {
  runtime.dispose()
}
```

`useXlsxViewerController` 可以接收 Runtime、Worker 地址、是否只读、是否使用 Worker、来源规则和资源限制。Worker、WASM、来源限制与诊断要来自同一实例。

## 5. 按需使用 XLSX 图表能力

普通表格查看只使用 `@arcships/vue-xlsx` 根入口。需要直接使用可选渲染能力时，从对应入口导入：

```ts
import { MemoChartSvg } from "@arcships/vue-xlsx/chart"
import { resolveRegionMapFeature } from "@arcships/vue-xlsx/map"
import { MemoSurfaceChartComposite } from "@arcships/vue-xlsx/webgl"
```

不要为了“预热”在应用主入口同时导入三个可选入口，否则会失去按需加载效果。根入口中旧的渲染名称仍是异步兼容外观。

## 6. PDF 保持可看可用

`PdfViewer` 显示真实 PDF 页面，支持翻页、缩放、旋转、缩略图、搜索和下载。迁移时不要把它替换成“验证后只下载”的提示，也不要直接把原始 PDF 放入同源 iframe。

```vue
<script setup lang="ts">
import { PdfViewer } from "@arcships/vue-extend"
import "@arcships/vue-extend/style.css"

const fiftyMiB = 50 * 1024 * 1024
</script>

<template>
  <PdfViewer src="/documents/sample.pdf" :max-file-size="fiftyMiB" />
</template>
```

PDF 唯一的文件拒绝配置是整份文件体积 `maxFileSize`，默认 `50 MiB`，宿主可调整。超过当前值返回 `PDF_TOO_LARGE`，并且不启动引擎。不要再加入页数、单页像素、总内存或并发页的隐藏拒绝上限。

如需自定义 PDFium 地址，从 `@arcships/vue-extend/assets/pdfium.wasm` 取得包内资源，或通过公开 `pdfiumWasmUrl`/Runtime 配置覆盖。严格内容安全策略需要允许应用实际采用的 Worker 和同源 WASM 资源。

## 7. 处理资源限制和错误

DOCX/XLSX 的 Runtime 限制归属实例。应用可以基于默认值覆盖业务需要的输入、归档、XML、图片、模型、时间、缓存和历史上限，但不能通过吞掉错误或无上限回退来“提高兼容性”。

程序判断使用公开错误码：

- DOCX：来源、获取、资源、图片、Worker、WASM、解析、超时、取消和 Runtime 销毁错误。
- XLSX：来源、获取、格式、资源、图片、Worker、超时、取消和 Runtime 销毁错误。
- PDF：`SOURCE_NOT_ALLOWED`、`FETCH_FAILED`、`INVALID_PDF`、`PDF_TOO_LARGE`、`ABORTED`。

完整名称与字段见[公开接口合同](api/public-api-contract.md)。错误后应允许用户选择合法文件并恢复 `ready`，不能保留旧文档页面冒充恢复。

## 8. 替换已弃用入口

| `0.1.x` 入口 | `0.2.0` 推荐入口 | 兼容期限 |
|---|---|---|
| DOCX 根入口 `setWasmSource` | `createDocxRuntime({ wasmUrl, workerUrl })` | 整个 `0.x` 保留，最早 `1.0.0` 删除 |
| XLSX 根入口 `setWasmSource` 及默认 Worker/WASM 状态 | `createXlsxRuntime({ wasmSource, workerUrl })` | 整个 `0.x` 保留，最早 `1.0.0` 删除 |
| `DocxEditorViewer` | `DocxEditor` | 整个 `0.x` 保留，最早 `1.0.0` 删除 |
| `useDocxViewerThumbnails` | `useDocxPageThumbnails` | 整个 `0.x` 保留，最早 `1.0.0` 删除 |
| DOCX/XLSX 低层页面、网格、工具栏和浮层 | `DocxViewer`、`DocxEditor`、`XlsxViewer` | 整个 `0.x` 保留，最早 `1.0.0` 删除 |
| `@arcships/vue-extend/dist/index.css` | `@arcships/vue-extend/style.css` | 整个 `0.x` 保留，最早 `1.0.0` 删除 |

不要把 TypeScript 的弃用提示当成立即删除要求。先完成替换和真实压缩包验证，再在未来 `1.0.0` 决定是否移除。

## 9. 验证迁移

在应用自己的正式构建中至少完成：

```bash
pnpm exec vue-tsc --noEmit
pnpm exec vite build
pnpm exec vite preview --host 127.0.0.1 --port 4173
```

检查：

- 页面等待 `networkidle` 或明确 `ready`，没有未知控制台错误、页面异常或关键 404。
- DOCX/XLSX Worker 和 WASM 请求来自安装的候选包。
- 正常 DOCX、XLSX、PDF 可以打开和操作；损坏、超限、取消后可以恢复。
- 未使用图表、地图或 WebGL 时不请求对应功能包。
- 下载文件名正确；对象 URL、Worker 和临时资源在切换或卸载后释放。

仓库候选制品的完整验证命令和证据格式见[Agent 执行手册](testing/agent-execution-runbook.md)。

## 10. 回退

1. 停止发布或灰度，保存失败证据，不继续覆盖候选制品。
2. 优先把新调用切回 `0.2.0` 保留的兼容入口，判断问题是否来自新 Runtime 接线。
3. 必须降级时，相关五包恢复到上一批已测试制品；不要混用核心包和 Vue 包版本。
4. 同时恢复对应 Worker/WASM 和样式配置。
5. 用正式 preview 复测，确认功能和资源来源都恢复后再结束回退。

候选发布前还要按[兼容矩阵](testing/compatibility-matrix.md)验证三种浏览器，并由全新会话执行 `BB-RELEASE`。

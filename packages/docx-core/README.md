# @arcships/docx-core

DOCX、DOCM、DOTX、DOTM 文档模型与只读输入，布局引擎、DOCX 编辑命令，以及带 Worker/WASM 支持的实例级 Runtime。宏内容不会执行。

## 安装

```bash
pnpm add @arcships/docx-core
```

## 入口

| 路径 | 用途 |
|---|---|
| `@arcships/docx-core` | 根入口：模型、布局、序列化、解析、编辑命令，以及 DOCX 基础引用适配 |
| `@arcships/docx-core/core` | 平台无关的纯数据能力 |
| `@arcships/docx-core/runtime` | `createDocxRuntime` 实例入口 |
| `@arcships/docx-core/wasm-url` | 包自带的 `bundledDocxWasmUrl` |
| `@arcships/docx-core/worker` | 独立 Worker 资源 |
| `@arcships/docx-core/assets/docx_wasm_bg.wasm` | WASM 资源 |

## 快速示例

```ts
import { createDocxRuntime } from "@arcships/docx-core/runtime"
import { bundledDocxWasmUrl } from "@arcships/docx-core/wasm-url"

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

## 资源限制

实例级限制覆盖输入字节、压缩包条目、解压字节、XML、关系、图片、模型节点、布局页数和解析时间。超限返回结构化错误（含 `actual` / `allowed`），默认值由 `DEFAULT_DOCX_RUNTIME_LIMITS` 给出。

## 文字、页面与区域引用

根入口可把编辑器已有的 `DocxTextRange`、页面或规范化页面区域转换为 `@arcships/office-interaction` 的统一引用草稿：

```ts
import {
  createDocxTextReferenceDraft,
  resolveDocxReference,
} from "@arcships/docx-core"

const context = {
  revision: { format: "docx", documentId, revision },
  model,
  pageCount,
} as const

const draft = createDocxTextReferenceDraft(context, activeTextRange)
const result = resolveDocxReference(context, { ...draft, referenceId })
```

同修订会验证模型路径和文字指纹；修订变化后，唯一文字引用可迁移，重复文字返回 `ambiguous`，删除返回 `not-found`。DOCX 页面和人工区域受重排版影响，跨修订不会仅凭页码自动恢复。适配器不保存引用集合，也不决定确认后的产品动作。

## 文档

- [DOCX 使用指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/guide/docx.md)
- [API 导航](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/README.md)
- [公开接口合同](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/public-api-contract.md)
- [发布说明](https://github.com/arcships/agentic-office-ui/blob/master/RELEASE_NOTES.md)
- [迁移指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/migration-0.2.md)

## License

Apache-2.0

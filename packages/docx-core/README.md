# @arcships/docx-core

DOCX 文档模型、布局引擎、编辑命令，以及带 Worker/WASM 支持的实例级 Runtime。

## 安装

```bash
pnpm add @arcships/docx-core
```

## 入口

| 路径 | 用途 |
|---|---|
| `@arcships/docx-core` | 根入口：模型、布局、序列化、解析、编辑命令（0.x 兼容） |
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

## 文档

- [公开接口合同](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/public-api-contract.md)
- [0.2.0 候选发布说明](https://github.com/arcships/agentic-office-ui/blob/master/RELEASE_NOTES.md)
- [迁移指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/migration-0.2.md)

## License

MIT

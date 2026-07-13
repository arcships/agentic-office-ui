# @arcships/xlsx-core

Excel（XLSX、XLS、XLSB、XLSM、XLTX、XLTM）和 CSV 工作簿输入、模型、公式、图表、图片，以及带 Worker/WASM 支持的实例级 Runtime。XLSB 由 Vue 接入层按需转换后进入现有工作簿管线；宏内容不会执行。

## 安装

```bash
pnpm add @arcships/xlsx-core
```

## 入口

| 路径 | 用途 |
|---|---|
| `@arcships/xlsx-core` | 根入口：数据、图片、图表、Worker、控制器（0.x 兼容） |
| `@arcships/xlsx-core/core` | 平台无关的纯数据能力 |
| `@arcships/xlsx-core/runtime` | `createXlsxRuntime` 实例入口 |
| `@arcships/xlsx-core/wasm-url` | 包自带的 `bundledXlsxWasmUrl` |
| `@arcships/xlsx-core/worker` | 独立 Worker 资源 |
| `@arcships/xlsx-core/assets/duke_sheets_wasm_bg.wasm` | WASM 资源 |

## 快速示例

```ts
import { createXlsxRuntime } from "@arcships/xlsx-core/runtime"
import { bundledXlsxWasmUrl } from "@arcships/xlsx-core/wasm-url"

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

## 资源限制

实例级限制覆盖输入字节、压缩包、XML、关系、图片、解析时间、工作表 XML、共享字符串、行列、工作表数量和公式数量。Worker 失败不会静默转主线程伪装成功。

## 文档

- [XLSX 使用指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/guide/xlsx.md)
- [API 导航](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/README.md)
- [公开接口合同](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/public-api-contract.md)
- [迁移指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/migration-0.2.md)

## License

Apache-2.0

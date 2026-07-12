# Runtime 与资源所有权

Runtime 负责 Worker、WASM、来源规则、资源限制、任务和诊断。高层组件通常会创建自己的 Runtime；只有需要统一限制、定制 Worker 或集中诊断时才需要注入。

## 所有权规则

| 创建方式 | 谁销毁 |
|---|---|
| 组件内部创建 | 组件卸载时销毁 |
| 组合函数内部创建 | Vue 作用域销毁时销毁 |
| 调用方创建后传入组件或组合函数 | 调用方销毁 |

```ts
const runtime = createDocxRuntime({ limits })

onBeforeUnmount(() => runtime.dispose())
```

不要把已经销毁的 Runtime 传给新组件。多个查看器可以共享调用方 Runtime 的配置思路，但是否共享具体实例取决于该 Runtime 的任务和资源模型；默认做法是每个独立编辑或查看区域使用独立实例。

## 核心入口

- DOCX：`@arcships/docx-core/runtime`
- XLSX：`@arcships/xlsx-core/runtime`
- PDF：`createPdfRenderRuntime` 从 `@arcships/vue-pdf` 导入
- PPTX：浏览器会话从 `@arcships/pptx-core/browser` 创建，不另设 Runtime 名称

## 来源和错误

远程 URL 必须通过对应的来源规则。诊断和错误中的 URL 会脱敏；业务代码使用稳定错误码，不记录文件正文或完整敏感地址。

资源限制属于实例配置。不要通过修改导入对象的方式在运行中改变限制；创建 Runtime 时给出的配置会形成实例快照。

# @arcships/vue-docx

Vue 3 DOCX 查看与编辑组件。

## 安装

```bash
pnpm add @arcships/vue-docx @arcships/docx-core
```

## 使用

```vue
<script setup lang="ts">
import { DocxViewer } from "@arcships/vue-docx"
import "@arcships/vue-docx/style.css"
</script>

<template>
  <DocxViewer :src="docxUrl" />
</template>
```

## 组件

- `DocxViewer` — 文档查看器
- `DocxEditor` — 文档编辑器

## 组合函数

- `useDocxEditor` — 编辑器状态与命令
- `useDocxModel` — 文档模型访问
- `useDocxPageThumbnails` — 页面缩略图

## 事件

| 事件 | 参数 |
|---|---|
| `load-start` | 无 |
| `load-success` | `DocxImportResult` |
| `load-error` | `Error`（可读 `code`） |

## 文档

- [DOCX 使用指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/guide/docx.md)
- [组件手册](https://github.com/arcships/agentic-office-ui/blob/master/docs/components/README.md)
- [DOCX 组合函数](https://github.com/arcships/agentic-office-ui/blob/master/docs/composables/docx.md)
- [自定义 DOCX 界面](https://github.com/arcships/agentic-office-ui/blob/master/docs/custom-components/docx.md)
- [公开接口合同](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/public-api-contract.md)
- [迁移指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/migration-0.2.md)

## License

Apache-2.0

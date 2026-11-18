# @arcships/vue-docx

Vue 3 Word 文档组件。Viewer 支持 DOCX、DOCM、DOTX、DOTM 安全预览，以及 OpenDocument 文本（ODT、OTT、ODM、OTM、OTH、FODT）的只读预览；Editor 的原格式编辑仍以 DOCX 为主，宏内容不会执行。OpenDocument 文本没有 Word 包和写入器，打开后会显示只读原因，编辑命令和导出都会被拒绝。从 DOCM/DOTX/DOTM 导入后导出会保留原容器类型和宏部件。导入和导出都在专用 Worker 中执行，不会在 UI 主线程同步解析或序列化。

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

## 本地化

`DocxViewer`、`DocxEditor` 和 `DocxDocumentSurface` 都接受 `locale`（默认 `"en"`，内置 `"zh-CN"`）与 `messages` 覆盖。用户可见文案和 aria-label 来自 `@arcships/office-interaction` 的共享消息目录：

```vue
<DocxViewer
  :model="model"
  locale="zh-CN"
  :messages="{ docx: { toolbar: { comments: '备注' } } }"
/>
```

## 组件

- `DocxDocumentSurface` — 最小嵌入渲染区（支持受控文字/页面/区域选择，以及宿主持有的 50%–200% 手势缩放）
- `DocxViewer` — 文档查看器
- `DocxEditor` — 文档编辑器
- `DocxBlockHighlights` — 文档块高亮覆盖层（按块、块 id 或页码）
- `DocxWordHighlights` — 词/文字范围高亮覆盖层（按词或 `{ blockId, start, end }`）

`DocxViewer`、`DocxEditor` 和 `DocxDocumentSurface` 暴露 `revealBookmark(name)`：点击目录、交叉引用或其它 `#书签` 内部链接时，查看器会滚动到书签所在页与段落、短暂高亮目标，并在可编辑表面上放置光标。书签名可写成 `Name` 或 `#Name`，不区分大小写；找不到书签时只显示非阻塞提示，不会抛错。

## 组合函数

- `useDocxEditor` — 编辑器状态与命令（OpenDocument 模型的 `viewOnlyReason` 非空）
- 文件类型策略导出 `DOCX_FILE_ACCEPT`、`SUPPORTED_DOCUMENT_FILE_EXTENSIONS`、`SUPPORTED_DOCUMENT_FILE_MIME_TYPES`、`isWordFileName()`、`isOpenDocumentTextFileName()`、`isSupportedDocumentFileName()` 和 `OPEN_DOCUMENT_VIEW_ONLY_REASON`
- `useDocxModel` — 文档模型访问
- `useDocxPageThumbnails` — 页面缩略图

## 事件

| 事件 | 参数 |
|---|---|
| `load-start` | 无 |
| `load-success` | `DocxImportResult` |
| `load-error` | `Error`（可读 `code`） |

`DocxDocumentSurface` 统一支持 `documentId`、`selectionMode`、`emitReferenceCandidates`，并发出 `referenceConfirm`、`regionDraftChange`、`selectionCancel`、`referenceResolve` 和 `referenceError`。组件 expose 可描述、解析和滚动到引用；`getDocumentBlocks()`、`confirmBlockReference()` 和 `confirmWordReference()` 可以把块/词作为统一引用确认。组件不保存引用集合，也不调用 Agent。

书签导航需要挂载新的虚拟化页面时发出 `requestPageReveal(pageIndex)`；内部链接指向不存在的书签时发出 `bookmarkMissing(name)`。

## 文档

- [DOCX 使用指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/guide/docx.md)
- [组件手册](https://github.com/arcships/agentic-office-ui/blob/master/docs/components/README.md)
- [DOCX 组合函数](https://github.com/arcships/agentic-office-ui/blob/master/docs/composables/docx.md)
- [自定义 DOCX 界面](https://github.com/arcships/agentic-office-ui/blob/master/docs/custom-components/docx.md)
- [公开接口合同](https://github.com/arcships/agentic-office-ui/blob/master/docs/api/public-api-contract.md)
- [迁移指南](https://github.com/arcships/agentic-office-ui/blob/master/docs/migration-0.2.md)

## License

Apache-2.0

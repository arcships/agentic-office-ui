---
id: docx-005
scope: vue-docx
status: pending
depends-on: [docx-003]
---

# vue-docx DocxViewer 重写（简单只读渲染）

## objective

将上游 `index.tsx` 的 ReactDocxViewer（590行）重写为 `packages/vue-docx/src/DocxViewer.vue`，实现简单只读 DOCX 渲染。

## context

- `docs/upstream-docx-feature-alignment.md` 第一节 1.3（index.tsx）、第二节 2.7（Iris 报告：index.tsx）
- 上游：`/Users/eric8810/Code/extend-ui-upstream/react-docx/packages/react-viewer/src/index.tsx`

## path

- `packages/vue-docx/src/DocxViewer.vue`

## steps

1. 重写 useDocxModel hook → Vue composable（file → model 异步解析状态）
2. 重写 ReactDocxViewer → DocxViewer.vue：
   - `<template>`：接收 file/model prop，用 layoutDocument 分页，每页渲染为 `<section>` + `<p>`/`<table>` DOM
   - 渲染辅助：runTextStyle、renderParagraphRuns、renderTable、renderBlock
   - HIGHLIGHT_TO_CSS 颜色映射
3. 保留 importDocxBuffer 调用（wasm worker）
4. 保留 layoutDocument 分页调用

## verification

```bash
pnpm --filter @extend-ai/vue-docx typecheck
pnpm --filter @extend-ai/vue-docx build
```

浏览器验证：加载 .docx fixture，段落/表格/图片/标题/CJK 文本正确渲染。

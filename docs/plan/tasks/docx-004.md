---
id: docx-004
scope: vue-docx
status: pending
depends-on: [docx-003]
---

> ⚠️ **已失效（2026-07-07）**：本任务按旧方案（机械复制上游巨型文件）设计，已回退。
> 实际执行以 `docs/plan/README.md` 的「DOCX 模块化重做任务」为准。
> 架构设计见 `docs/docx-migration-architecture.md`。


# vue-docx DocxEditor 重写（~23584 行）

## objective

将上游 editor.tsx 32871-56454 行（DocxEditorViewer 组件）重写为 `packages/vue-docx/src/DocxEditor.vue`，实现 DOM 渲染 + 虚拟化 + contentEditable 编辑 + 交互。

## context

- `docs/upstream-docx-feature-alignment.md` 第二节 2.6（DocxEditorViewer）、第三节渲染层要点(35-43)、第四节 Phase 4
- 上游：editor.tsx 32871-56454行

## path

- `packages/vue-docx/src/DocxEditor.vue`

## steps

1. 提取纯函数到 docx-core（如尚未包含）：pageSegmentationPlan 编排逻辑、visiblePageIndexes 计算、pageStackVirtualSpacers 计算
2. Vue 组件重写：
   - `<template>`：viewer root + spacer divs + page wrapper/surface/header/body/footer/gutter + context menu
   - `<script setup>`：调用 useDocxEditor，分页 computed，虚拟化
   - `v-html` 替代 `dangerouslySetInnerHTML`（注意 contentEditable 冲突——复刻 draft 缓存 + suppress 机制）
3. 实现分页：测量驱动迭代（先估算→渲染→测量→重新分页），振荡检测
4. 实现虚拟化：@tanstack/vue-virtual 或自实现 visible range + spacer div
5. 实现 contentEditable + draft：
   - paragraphDraftsRef → ref(Map)
   - onInput → commitParagraphDraftFromElement
   - 导出前 flushPendingDraftsForExport
6. 实现 DOM 选区恢复：historyRestoreRequest nonce + watchEffect/nextTick 时序
7. 实现交互：表格 resize handle、图片 drag/wrap、右键菜单、缩略图
8. 实现 renderParagraphRuns 重写：JSX `<span>` → Vue render function（h()）

## verification

```bash
pnpm --filter @extend-ai/vue-docx typecheck
pnpm --filter @extend-ai/vue-docx build
```

浏览器验证：导入 .docx，编辑文本/格式/表格/图片，导出 .docx。

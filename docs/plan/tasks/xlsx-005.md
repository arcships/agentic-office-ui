---
id: xlsx-005
scope: vue-xlsx
status: pending
depends-on: [xlsx-003, xlsx-004]
---

# vue-xlsx XlsxViewer 重写

## objective

将上游 `XlsxViewer.tsx`（16615行）重写为 `packages/vue-xlsx/src/XlsxViewer.vue`，实现 canvas 渲染管线 + 虚拟化 + 交互 + toolbar。

## context

- `docs/upstream-xlsx-feature-alignment.md` 第二节 2.12（XlsxViewer 渲染层）、第三节渲染层要点(39-62)、第四节 Phase 6
- 上游：`XlsxViewer.tsx`

## path

设计文档（只读）：
- `docs/upstream-xlsx-feature-alignment.md` 第二节 2.12 + 第三节渲染层 + 第四节 Phase 6

- `packages/vue-xlsx/src/XlsxViewer.vue`
- `packages/vue-xlsx/src/XlsxViewerProvider.vue`（新建）

## steps

1. 提取纯函数到 xlsx-core（如果尚未在 xlsx-001 中包含）：
   - buildCellStyle、resolveInheritedCellStyle、mergeResolvedCellStyle
   - getCellDisplayValue、getCellNumericValue、getCellBooleanValue
   - 条件格式解析（resolveConditionalDataBar/ColorScale/IconForCell）
   - expandMergeAwareWindow、buildStickyOffsets、findIndexForOffsetPrefix
   - EMU 换算、resolveAnchoredRect
   - 缩略图 paint 逻辑
   - 键盘导航算法、选区算法
2. Vue 组件重写：
   - `<template>`：viewport div + canvas 元素（4 body + 5 header）+ overlay div + toolbar + sheet tabs
   - `<script setup>`：调用 useXlsxViewerController，canvas paint 用 requestAnimationFrame + 依赖监听
   - canvas 2D API 调用直接复用上游逻辑
   - 事件处理用 Vue @event 绑定
3. 实现 canvas 渲染管线：
   - 4 窗格分层（scroll/top/left/corner）
   - 前缀和二分查找 + 480px overscan + expandMergeAwareWindow
   - 签名比对跳过 + blit 滚动 + 脏矩形剔除
   - DPR 处理
4. 实现交互：选区拖拽、键盘导航、双击编辑、wheel zoom、表格 resize handle
5. 实现缩略图：独立 canvas，复用纯函数
6. 实现 toolbar：极简设计（文件名 + zoom + download + sheet tabs）

## verification

```bash
pnpm --filter @extend-ai/vue-xlsx typecheck
pnpm --filter @extend-ai/vue-xlsx build
```

浏览器验证：加载 welcome.xlsx / financial-model.xlsx / charts-images.xlsx，检查渲染/编辑/导出。

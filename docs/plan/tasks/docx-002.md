---
id: docx-002
scope: docx-core
status: done
depends-on: [docx-001]
---

> ⚠️ **已失效（2026-07-07）**：本任务按旧方案（机械复制上游巨型文件）设计，已回退。
> 实际执行以 `docs/plan/README.md` 的「DOCX 模块化重做任务」为准。
> 架构设计见 `docs/docx-migration-architecture.md`。


# docx-core editor helpers 复制+清理（~24953 行）

## objective

从上游 editor.tsx 提取前半部分（1-24953行，纯函数/类型区），sed 清理 73 处 React.CSSProperties 类型标注，抽出需重写的 JSX 函数。

## context

- `docs/upstream-docx-feature-alignment.md` 第一节 1.2（editor.tsx 拆分）、第二节 2.5（editor.tsx 前半部分对齐清单）、第四节 Phase 2
- 上游：`/Users/eric8810/Code/extend-ui-upstream/react-docx/packages/react-viewer/src/editor.tsx` 1-24953行

## path

设计文档（只读）：
- `docs/upstream-docx-feature-alignment.md` 第一节 1.2 + 第二节 2.5 + 第四节 Phase 2

- `packages/docx-core/src/editor-helpers.ts`（纯函数+类型）
- `packages/docx-core/src/editor-types.ts`（DocxEditorSelection/DocxTextRange 等类型，供 core/state.ts 引用）

## steps

1. 提取 editor.tsx 1-24953 行到 `editor-helpers.ts`
2. sed 清理：
   - `import * as React from "react"` → 删除
   - `React.CSSProperties` → `Record<string, string | number | undefined>`
   - `React.Dispatch<React.SetStateAction<string>>` → `(value: string | ((prev: string) => string)) => void`
   - `React.KeyboardEvent<HTMLElement>` → `{ key: string; shiftKey: boolean; ctrlKey: boolean; metaKey: boolean; preventDefault: () => void }`
   - `React.PointerEvent<HTMLElement>` → `Record<string, unknown>`
   - `React.MouseEvent<HTMLElement>` → `Record<string, unknown>`
   - `React.ReactNode` → `unknown`
3. 抽出 `renderParagraphRuns`（18126-19581，~1455行）和 `renderStaticHtml`（4902）——移到单独文件待 vue-docx 重写
4. 抽出 DocxEditorSelection/DocxTextRange/DocxTextRangeLocation 等类型到 `editor-types.ts`（供 core/state.ts 引用）
5. 保留的核心纯函数（零修改）：
   - 图片 wrap 几何：resolveDualWrappedFloatingImageGeometry, resolveParagraphDualWrappedTextLayout
   - pretext layout：buildParagraphPretextLayoutSource
   - 行高估算：estimateParagraphLineHeightPx, paragraphLineCountWithinWidth
   - 表格高度：estimateTableRowHeightsPx
   - 分页核心：buildDocumentPageNodeSegments
   - numbering：buildParagraphNumberingLabels
   - 页眉页脚：resolveHeaderPaginationReservePx
   - defaultStarterModel

## verification

```bash
pnpm --filter @arcships/docx-core typecheck
grep -r "React\." packages/docx-core/src/editor-helpers.ts  # 应为空
```

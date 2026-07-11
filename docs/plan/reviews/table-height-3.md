# table-height verify (round 3)

## 1. Typecheck

✅ `pnpm --filter @arcships/docx-core typecheck` — 零错误，通过。

## 2. 文件存在且非空

✅ `packages/docx-core/src/editor/helpers/line-height-table.ts` — 128 行，非空。
✅ 已在 `index.ts` barrel 文件中通过 `export * from "./line-height-table"` 导出。

## 3. React 残留

✅ 无 React/useState/useEffect/useMemo/useCallback/useRef/JSX/className 匹配。

## 4. 重复已有模块

✅ 无重复。`table-height.ts`（与 `line-height-table.ts` 逐字节重复的冗余文件）已在 fix #2 (`11b41a5`) 中删除。
✅ 全部 13 个导出函数仅在 `line-height-table.ts` 中定义，无跨文件重复。

### 函数清单

| 函数 | 类型 | 说明 |
|------|------|------|
| `resolveTableParagraphLineHeightPx` | wrapper | → `estimateParagraphLineHeightPx` |
| `resolveTableRowLineHeightPx` | 原创 | 遍历 cell 取最大行高 |
| `estimateTableLineHeightPx` | 原创 | 聚合所有 row 行高 |
| `resolveTableCellBlockStyle` | 原创 | twips→px padding 转换 |
| `isLineHeightRelevantParagraph` | wrapper | → `paragraphRendersTextFreeLine` / `paragraphContainsExplicitLineBreakText` |
| `resolveTableParagraphLineHeightContext` | 原创 | fontSize × scale 计算 |
| `resolveTableLineHeightTabStopsPx` | wrapper | → `resolveParagraphTabStopsPx` + origin 偏移 |
| `resolveTableParagraphDocGridLinePitchPx` | wrapper | → `resolveParagraphDocGridLinePitchPx` |
| `resolveTableParagraphTabStopsPx` | wrapper | → `resolveParagraphTabStopsPx` |
| `resolveTableParagraphFirstLineOriginPx` | wrapper | → `resolveParagraphFirstLineOriginPx` |
| `resolveTableParagraphFirstLineLeftTabStopsPx` | wrapper | → `resolveParagraphFirstLineLeftTabStopsPx` |
| `resolveTableParagraphNextTabStopPx` | wrapper | → `resolveNextTabStopPx` |
| `resolveTableParagraphDocGridSnapState` | wrapper | → `paragraphDocGridSnapState` |

### 变更历史

| Commit | 变更 |
|--------|------|
| `0bb5dd8` | 初始 commit，新增 `table-height.ts`（与 `line-height-table.ts` 重复） |
| `861bc78` | fix #1，新增 `table-height-1.md` review |
| `11b41a5` | fix #2，删除重复 `table-height.ts`，新增 `table-height-2.md` review |
| (本次) | round 3 verify，确认修复后全绿 |

### 备注

wrapper 函数是 thin pass-through，源自 `line-height.ts` 的同名函数。这是设计意图，非代码重复——提供 table-specific 命名空间，便于后续在 wrapper 层注入 table 差异化逻辑。

## 结论

✅ 通过。全部 4 项检查均通过。

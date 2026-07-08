# table-height verify

## 1. Typecheck

✅ `pnpm --filter @extend-ai/docx-core typecheck` — 零错误，通过。

## 2. 文件存在且非空

✅ `packages/docx-core/src/editor/helpers/table-height.ts` — 128 行，非空。

## 3. React 残留

✅ 无 React/useState/useEffect/useMemo/JSX 匹配。

## 4. 重复已有模块

🔴 BLOCKED: `table-height.ts` 是 `line-height-table.ts` 的逐字节副本。全部 13 个导出函数在 `line-height-table.ts` 中已存在且实现完全一致。没有任何文件 import `table-height.ts`。

重复函数清单:
- `resolveTableParagraphLineHeightPx`
- `resolveTableRowLineHeightPx`
- `estimateTableLineHeightPx`
- `resolveTableCellBlockStyle`
- `isLineHeightRelevantParagraph`
- `resolveTableParagraphLineHeightContext`
- `resolveTableLineHeightTabStopsPx`
- `resolveTableParagraphDocGridLinePitchPx`
- `resolveTableParagraphTabStopsPx`
- `resolveTableParagraphFirstLineOriginPx`
- `resolveTableParagraphFirstLineLeftTabStopsPx`
- `resolveTableParagraphNextTabStopPx`
- `resolveTableParagraphDocGridSnapState`

## 结论

🔥 阻塞。删除 `table-height.ts` 或修改为从 `line-height-table.ts` re-export 并补充差异逻辑。

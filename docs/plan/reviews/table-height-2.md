# table-height verify (round 2)

## 1. Typecheck

✅ `pnpm --filter @arcships/docx-core typecheck` — 零错误，通过。

## 2. 文件存在且非空

✅ `packages/docx-core/src/editor/helpers/table-height.ts` — 128 行，非空。

## 3. React 残留

✅ 无 React/useState/useEffect/useMemo/useCallback/useRef/JSX 匹配。

## 4. 重复已有模块

🔴 BLOCKED: `table-height.ts` 与 `line-height-table.ts` 逐字节相同（`diff` 输出为空）。全部 13 个导出函数在 `line-height-table.ts` 中已存在且实现完全一致。`table-height.ts` 未被任何文件 import，也未在 `index.ts` barrel 中导出。

自 review-1 以来无任何变更（git log 仅有初始 commit `0bb5dd8`）。

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

🔥 仍阻塞。删除 `table-height.ts`，或按 split plan 从 `line-height-table.ts` 中拆分出差异逻辑后删除重复内容。

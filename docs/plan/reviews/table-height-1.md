# table-height verify (round 3)

## 1. Typecheck

✅ `pnpm --filter @arcships/docx-core typecheck` — 零错误，通过。

## 2. 文件存在且非空

✅ `packages/docx-core/src/editor/helpers/table-height.ts` — 8 行 barrel，非空。
✅ `packages/docx-core/src/editor/helpers/line-height-table.ts` — 999 行（≤1000 行限制），非空。

`table-height.ts` 内容:

```ts
export * from "./line-height-table";
```

## 3. React 残留

✅ `table-height.ts` — 无 import，无 React 匹配。
✅ `line-height-table.ts` — 无 React/useState/useEffect/useMemo/useCallback/useRef/JSX 匹配。

## 4. 重复已有模块

✅ `table-height.ts` 已从逐字节副本（128 行，round 1/2）改为 barrel re-export（8 行），不再重复代码。

所有 25 个导出函数/接口仅定义在 `line-height-table.ts` 中，未在其他 helper 模块重复：

- `estimateTableHeightPx`
- `estimateDocNodeHeightPx`
- `estimateTableRowHeightsPx`
- `estimateTableCellContentHeightPx`
- `estimateTableCellSliceBoundaryLayoutPx`
- `resolveTableRowHeightCss`
- `resolveTableRowSliceHeightOnSafeBoundaryPx`
- `paragraphCanSplitAcrossPages`
- `paragraphHasExplicitSpacing`
- `wordLikeTableCellParagraph`
- `rowAllowsPageSplit`
- `rowHasDeepFlowContent`
- `rowHasNestedTableContent`
- `capSplitFriendlyTableRowEstimatePx`
- `tableStyleIdFromSourceXml`
- `tableHasVisibleBorders`
- `tableContainsParagraphsWithoutExplicitSpacing`
- `tableUsesWordLikeParagraphDefaults`
- `uniqueSortedPixelBoundaries`
- `tableCellSliceBoundaryIsSafe`
- `paragraphWidowControlEnabled`
- `paragraphIsOnlyExplicitPageBreak`
- `TableCellSliceBoundaryLayout` (interface)
- `injectEstimateParagraphHeightPx`
- `injectParagraphAvailableTextWidthPx`
- `injectParagraphLineCountWithinWidth`
- `injectParagraphBorderInsetPx`

⚠️ 注意: `index.ts` 同时导出 `line-height-table` 和 `table-height`（同一模块两次），TypeScript 会合并同名符号，无编译错误但 barrel 有冗余。如需清理，从 index.ts 移除 `export * from "./table-height";` 或 `export * from "./line-height-table";` 其一。

## 结论

✅ 通过。typecheck 零错误，无 React 残留，无代码重复。

# Review: line-height-table

## Checklist

| Check | Result |
|-------|--------|
| typecheck zero errors | ✅ `tsc --noEmit` passed with no output |
| File exists and non-empty | ✅ `packages/docx-core/src/editor/helpers/line-height-table.ts` (999 lines, 36KB) |
| React residues | ✅ No `import React`, no JSX, no hooks |
| Duplicated module functions | ✅ All 27 exports are unique to this file (grep confirmed no duplicates in `src/`) |

## Exports (27)

- `injectEstimateParagraphHeightPx` (injectable)
- `injectParagraphAvailableTextWidthPx` (injectable)
- `injectParagraphLineCountWithinWidth` (injectable)
- `injectParagraphBorderInsetPx` (injectable)
- `paragraphHasExplicitSpacing`
- `wordLikeTableCellParagraph`
- `estimateTableCellContentHeightPx`
- `rowAllowsPageSplit`
- `rowHasDeepFlowContent`
- `rowHasNestedTableContent`
- `capSplitFriendlyTableRowEstimatePx`
- `tableStyleIdFromSourceXml`
- `tableHasVisibleBorders`
- `tableContainsParagraphsWithoutExplicitSpacing`
- `tableUsesWordLikeParagraphDefaults`
- `estimateTableRowHeightsPx`
- `resolveTableRowHeightCss`
- `uniqueSortedPixelBoundaries`
- `estimateTableCellSliceBoundaryLayoutPx`
- `tableCellSliceBoundaryIsSafe`
- `resolveTableRowSliceHeightOnSafeBoundaryPx`
- `estimateTableHeightPx`
- `estimateDocNodeHeightPx`
- `paragraphWidowControlEnabled`
- `paragraphIsOnlyExplicitPageBreak`
- `paragraphCanSplitAcrossPages`
- `TableCellSliceBoundaryLayout` (type)

## Dependencies

Imports from sibling helpers: `cache-utils`, `constants`, `paragraph-inspect`, `paragraph-geometry`, `line-height`, `pretext-build`, `pretext-layout`, `cell-utils`. Imports from engine/types, `viewer/section-layout`, `layout/pagination`.

4 functions are injectable stubs awaiting not-yet-ported modules (`estimateParagraphHeightPx`, `paragraphAvailableTextWidthPx`, `paragraphLineCountWithinWidth`, `paragraphBorderInsetPx`).

## Verdict: PASS

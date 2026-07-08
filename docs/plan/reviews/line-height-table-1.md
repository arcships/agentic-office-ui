# Review: line-height-table

**Date:** 2026-07-08
**Status:** ✅ PASS

---

## 1. Typecheck

`pnpm --filter @extend-ai/docx-core typecheck` — **zero errors**.

## 2. File existence

| File | Lines | Status |
|---|---|---|
| `packages/docx-core/src/editor/helpers/line-height-table.ts` | 141 | ✅ Present, non-empty |

## 3. React residuals

Grep for `react`, `useState`, `useEffect`, `useMemo`, `useCallback`, `useRef`, `JSX`, `className`, `onClick`, `onChange` on `line-height-table.ts` — **zero matches**. No React import, no JSX, no hook usage.

## 4. Duplicate functions

All 13 exported functions from `line-height-table.ts` are **unique** to this file. No other file in `packages/docx-core/src` defines functions with the same names.

Wrapper functions that delegate to `line-height.ts` helpers are intentional table-specific entry points, not redundant duplicates:

| Function | Delegates to |
|---|---|
| `resolveTableParagraphLineHeightPx` | `estimateParagraphLineHeightPx` |
| `resolveTableParagraphDocGridLinePitchPx` | `resolveParagraphDocGridLinePitchPx` |
| `resolveTableParagraphTabStopsPx` | `resolveParagraphTabStopsPx` |
| `resolveTableParagraphFirstLineOriginPx` | `resolveParagraphFirstLineOriginPx` |
| `resolveTableParagraphFirstLineLeftTabStopsPx` | `resolveParagraphFirstLineLeftTabStopsPx` |
| `resolveTableParagraphNextTabStopPx` | `resolveNextTabStopPx` |
| `resolveTableParagraphDocGridSnapState` | `paragraphDocGridSnapState` |

Novel functions (no delegation):

| Function | Purpose |
|---|---|
| `resolveTableRowLineHeightPx` | Max line-height across all cell paragraphs in a row |
| `estimateTableLineHeightPx` | Max line-height across all rows in a table |
| `resolveTableCellBlockStyle` | Converts cell padding twips → px RecordStyle |
| `resolveTableParagraphLineHeightContext` | fontSizePx × singleLineScale for a paragraph |
| `resolveTableLineHeightTabStopsPx` | First-line left tab stops shifted by origin |
| `isLineHeightRelevantParagraph` | Paragraph renders text line OR contains explicit line breaks |

## Verdict

All checks pass. Typecheck clean, file present and non-empty, no React residuals, no duplicate function definitions.

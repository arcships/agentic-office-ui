# pagination-plan-core — verify review #1

**Status: pass**

## 1. Typecheck

`pnpm --filter @arcships/docx-core typecheck` → zero errors.

## 2. Files

| File | Exists | Non-empty | Lines | Notes |
|---|---|---|---|---|
| `src/editor/helpers/pagination-plan-core.ts` | ✅ | ✅ | 990 | Core module (new) |
| `src/editor/helpers/pagination-plan-iterate.ts` | ✅ | ✅ | 707 | Imported by pagination-plan-core |
| `src/layout/pagination.ts` | ✅ | ✅ | 690 | Imported by pagination-plan-core |
| `src/viewer/pagination-breaks.ts` | ✅ | ✅ | 224 | Pre-existing; not imported by new files |

## 3. React residues

Grep for `react\|jsx\|useState\|useEffect\|useMemo\|useCallback\|useRef` across the new file — one match: the migration comment `// Migrated from @extend-ai/react-docx editor.tsx lines:`. Zero React imports, hooks, or JSX.

## 4. Duplication checks

### pagination-plan-core.ts exports (9):

| Export | Present in `layout/pagination.ts`? | Present in `viewer/pagination-breaks.ts`? | Verdict |
|---|---|---|---|
| `resolveLineRangeWithinVerticalSlice` | No | No | Unique |
| `resolveTableCellParagraphVisualBottomPx` | No | No | Unique |
| `tableCellParagraphFitsFullyWithinSlice` | No | No | Unique |
| `resolveParagraphSegmentClipBleedPx` | No | No | Unique |
| `resolveFallbackParagraphSegmentClipBleedPx` | No | No | Unique |
| `resolveParagraphSegmentNonFlowReservePx` | No | No | Private fn in `page-segmentation-core.ts`; exported implementation is unique to this file |
| `buildRenderColumnSegmentsForPageSection` | No | No | Unique |
| `buildDocumentPageNodeSegmentsFromLastRenderedPageBreakHints` | No | No | Unique |
| `mergeTrailingPagesToTargetCount` | No | No | Unique |

Primitives (`paragraphHasLastRenderedPageBreak`, `paragraphHasPageBreakBefore`, `effectiveParagraphAfterSpacingPx`, `effectiveParagraphBeforeSpacingPx`) are imported from `layout/pagination.ts` and `pagination-plan-iterate.ts` — no reimplementation.

### Pre-existing (not introduced by this commit)

`viewer/pagination-breaks.ts` duplicates `TableExplicitPageBreakInfo`, `collectTableExplicitPageBreakInfo`, and `collectTopLevelExplicitPageBreakStartNodeIndexes` from `layout/pagination.ts`. `viewer/index.ts` explicitly excludes `pagination-breaks.ts` from the barrel due to symbol collision. This pre-dates the current commit.

## 5. Observations

- `pagination-plan-core.ts` imports from `pagination-plan-iterate.ts` (2 functions), `layout/pagination.ts` (2 functions), plus `page-segmentation-core`, `page-segmentation-table`, `line-height`, and `constants` — all within `docx-core`.
- The module is the primary column-render orchestrator, composing segment clip/boundary helpers and page segmentation plan assembly.
- `tableCellParagraphFitsFullyWithinSlice` is dedicated to table-cell inline paragraph segmentation during column rendering.

## Verdict

All four checks pass. No blockers.

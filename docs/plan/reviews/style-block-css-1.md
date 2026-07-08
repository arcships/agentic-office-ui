# style-block-css Review

**Date**: 2026-07-08
**Module**: `packages/docx-core/src/editor/helpers/style-block-css.ts`
**Status**: ✅ pass

## Checklist

| Check | Result |
|---|---|
| Typecheck zero errors | ✅ `tsc --noEmit` passed with no output |
| File exists and non-empty | ✅ 51 lines, 4 exports |
| No React residuals | ✅ Zero React imports |
| No duplicate functions | ✅ All 4 exports defined only in this module |

## Module summary

Extracts paragraph border CSS helpers from upstream `editor.tsx:15036-15079`. Four exports:

- `paragraphBorderToCss` — delegates to `tableBorderToCss` (shared shape between `ParagraphBorderStyle` and `TableBorderStyle`)
- `paragraphBorderPaddingPx` — border space in pixels via `pointsToPixels`
- `paragraphBorderStrokeWidthPx` — stroke width in pixels from `sizeEighthPt` (eighth-points), min 0.5px
- `paragraphBorderInsetPx` — total inset = stroke + padding

## Dependencies

- Imports from `table-utils.ts`: `normalizeBorderType`, `tableBorderToCss` — intentional shared utilities
- Imports from `ooxml-helpers.ts`: `pointsToPixels` — standard OOXML conversion
- Consumer: `line-height-table.ts` uses `paragraphBorderInsetPx` (4 call sites)

## Type compatibility

`paragraphBorderToCss` casts `ParagraphBorderStyle` to `TableBorderStyle` — both share `type`, `color`, `sizeEighthPt`, `spacePt`, `shadow` fields. No runtime risk.

## Verdict

Module is clean, correctly extracted, typecheck-clean, and has no duplication with existing helpers.

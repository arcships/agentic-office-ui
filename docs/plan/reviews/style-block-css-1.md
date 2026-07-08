# style-block-css review

**File:** `packages/docx-core/src/editor/helpers/style-block-css.ts`
**Date:** 2026-07-08
**Status:** ✅ pass

## 1. Typecheck

```sh
pnpm --filter @extend-ai/docx-core typecheck
# tsc --noEmit
```

Zero errors. All type imports (`ParagraphBorderSet`, `ParagraphIndent`, `ParagraphStyleDefinition`, `TableBorderSet`, `TableCellStyle`, `TableRowStyle`) resolve correctly from `../../engine/types`. Import of `twipsToPixels` from `../../viewer/section-layout` resolves correctly.

## 2. File existence and content

File exists at `packages/docx-core/src/editor/helpers/style-block-css.ts` (115 lines, 5 functions). Non-empty.

Barrel export registered in `index.ts` line 46: `export * from "./style-block-css"`.

## 3. React residuals

No React residuals. The only mentions of `react`, `jsx`, `tsx` are in comments documenting the upstream extraction source (`editor.tsx`). No `className`, no hooks (`useState`, `useEffect`, `useMemo`, `useCallback`), no `createElement`, no JSX syntax. Module is framework-agnostic.

## 4. Duplicate function check

All 5 functions exist only in this file — no duplicates anywhere in `packages/docx-core/src`:

| Function | Defined in | Duplicated elsewhere? |
|---|---|---|
| `toFiniteNumber` | style-block-css.ts (private) | No |
| `twipsToCssPx` | style-block-css.ts (private) | No |
| `borderStyleToCss` | style-block-css.ts (private) | No |
| `paragraphStyleToCss` | style-block-css.ts (exported) | No |
| `tableCellStyleToCss` | style-block-css.ts (exported) | No |

`twipsToPixels` is imported from `viewer/section-layout.ts` (the canonical exported definition) rather than redefined — correct re-use.

No overlap with the sister module `style-to-css.ts`, which handles run-level/inline CSS conversion (font, color, decoration, highlight, link) while `style-block-css.ts` handles block-level CSS conversion (paragraph margins/indents/borders, table cell padding/height/borders).

## 5. Module register

Listed as completed in `index.ts` header comment (line 23: `style-block-css.ts — paragraphStyleToCss / tableCellStyleToCss`) and exported (line 46).

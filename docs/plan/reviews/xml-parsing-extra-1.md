# Review: xml-parsing-extra-1

**Date:** 2026-07-08

## Checks

### 1. Typecheck

`pnpm --filter @arcships/docx-core typecheck` — zero errors.

### 2. File existence and non-empty

`packages/docx-core/src/editor/helpers/xml-parsing-extra.ts` — 280 lines, non-empty.

### 3. React residuals

No React imports, hooks, JSX, or `.tsx` usage. The single grep match is a source-line reference comment (`Upstream editor.tsx: lines 15081-15324`).

### 4. Duplicate functions

`paragraphExplicitIndentTwips` and `resolveListParagraphIndent` appear only in `xml-parsing-extra.ts` across the entire `packages/docx-core/src` tree. No duplication with existing modules.

Dependencies used:
- `paragraphExplicitIndentBySourceXml`, `setCacheEntry` from `cache-utils.ts`
- `effectiveNumberingNumIdForParagraph`, `findNumberingLevelDefinition` from `numbering.ts`
- `LIST_LEVEL_STEP_TWIPS` from `constants.ts`
- Types from `engine/types`

Barrel export in `helpers/index.ts` — correct.

## Verdict

**PASS**

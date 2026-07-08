# Review: xml-parsing-extra

**Status: BLOCKED**

**Module:** `packages/docx-core/src/editor/helpers/xml-parsing-extra.ts` (427 lines)

## Checklist

### 1. Typecheck — PASS
`pnpm --filter @extend-ai/docx-core typecheck` exits zero. No type errors.

### 2. File exists and is non-empty — PASS
`xml-parsing-extra.ts` exists at 427 lines with substantive content.

### 3. React residuals — PASS
No `import React`, no hooks (`useState`, `useEffect`, etc.), no JSX elements found.

### 4. Duplicate functions — FAIL
All 15 exported functions from `xml-parsing-extra.ts` are duplicates of exports already present in `xml-parsing.ts`:

| Function | xml-parsing.ts | xml-parsing-extra.ts |
|---|---|---|
| `extractBalancedTagRanges` | exported (returns `XmlBalancedTagRange[]`) | exported (returns inline type) |
| `trackedChangeKindFromTagName` | exported | exported (identical) |
| `normalizeTrackedChangeSnippet` | exported | exported (identical) |
| `formatTrackedChangeDate` | exported | exported (identical) |
| `stripTextBoxContentFromRunXml` | exported | exported (identical) |
| `parseTrackedRunTokens` | exported | exported (identical) |
| `xmlBooleanFlag` | exported | exported (identical) |
| `xmlColorValue` | exported | exported (identical) |
| `parseRunStyleFromRunXml` | exported | exported (identical) |
| `balancedTagXmlBlocks` | exported | exported (identical) |
| `mergeTextRunStyles` | exported | exported (identical) |
| `parseParagraphAlignmentFromXml` | exported | exported (identical) |
| `parseDrawingImageTransformFromSourceXml` | exported | exported (identical) |
| `joinCssTransforms` | exported | exported (identical logic) |
| `resolveImageRenderTransformStyle` | exported | exported (identical logic) |

The barrel file `index.ts` re-exports from both modules (`export * from "./xml-parsing"` at line 59 + 15 named re-exports from `"./xml-parsing-extra"` at lines 60-76), creating 15 duplicate barrel exports.

The only novel content in `xml-parsing-extra.ts`:
- `import { twipsToPixels } from "../../viewer/section-layout"` + `void twipsToPixels(0)` side-effect call (line 13, 378)
- Private `decodeXmlText` (unexported, replicates `xml-parsing.ts` exported version)
- Private `XML_NAME_ESCAPE_PATTERN` constant

## Recommendation
Remove `xml-parsing-extra.ts` and its barrel re-exports from `index.ts`. If the `twipsToPixels(0)` side-effect is needed, place it in a module that actually uses `twipsToPixels`, or add the import to `xml-parsing.ts` directly.

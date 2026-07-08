// render/ barrel — Vue renderParagraphRuns + static-html.
//
// Provides the canonical Vue h()-based paragraph run renderer
// and static HTML serialization for copy/export scenarios.

export { renderParagraphRuns, type ParagraphRunRenderOptions } from "./paragraph-runs"
export { renderStaticHtml } from "./static-html"

// Sub-module exports for advanced consumers
export { renderTextRun, type TextRunRenderContext, type TextRunRenderResult } from "./paragraph-runs-text"
export { renderImageRun, type ImageRunRenderContext, type ImageRunRenderResult, type ImageRunRenderExtra, type ImageInteraction } from "./paragraph-runs-image"
export { renderFormFieldRun, type FieldRunRenderContext, type FieldRunRenderResult } from "./paragraph-runs-field"

// Field resolution helpers
export {
  resolveFieldText,
  resolvePageFieldText,
  resolveStyleRefFieldText,
  normalizeFieldComparableText,
  numberToRoman,
  type FieldResolutionContext,
  type FieldResolutionState,
} from "./paragraph-runs-field-resolve"

/**
 * @extend-ai/docx-core
 *
 * Framework-agnostic DOCX document engine.
 * Inspired by the public @extend-ai/react-docx API; see docs/upstream-extend-ui.md for attribution.
 *
 * Provides:
 * - Document model types (DocModel, ParagraphNode, TableNode, etc.)
 * - Layout engine (pagination, paragraph layout, table layout)
 * - Immutable document editing (insert, remove, format, clone)
 * - OOXML parsing/serialization (via WASM bridge)
 * - No React/Vue dependency
 */

// Types
export type * from "./types"

// Layout engine
export {
  DEFAULT_LAYOUT_OPTIONS,
  DEFAULT_PAGE_OVERFLOW_TOLERANCE_PX,
  DEFAULT_MIN_PARAGRAPH_LINE_HEIGHT_PX,
  headingScale,
  paragraphToLayout,
  tableToLayout,
  layoutDocument,
  estimateBlockHeight,
  paragraphBeforeSpacingPx,
  paragraphAfterSpacingPx,
  resolveParagraphBeforeSpacingPx,
  paragraphHasExplicitPageBreak,
  paragraphHasPageBreakBefore,
  sectionBreakAfterParagraphStartsNewPage,
  sectionBreakPropertiesStartNewPage,
  sectionTitlePageEnabled,
  collectDocxHardPageBreakStartNodeIndexes,
  collectTopLevelExplicitPageBreakStartNodeIndexes,
  resolveSectionIndexForNodeIndex,
  resolveSectionPropertiesXmlForNodeIndex,
  parseSectionLayout,
  resolveDocumentLayout,
  resolveDocxPageThumbnailResolution,
} from "./layout"

export type { DocumentLayoutMetrics } from "./layout"

// Document model editing
export {
  cloneDocModel,
  normalizeDocModel,
  insertParagraph,
  removeParagraph,
  duplicateParagraph,
  copyParagraphs,
  pasteParagraphs,
  updateParagraphText,
  updateTableCellText,
  updateTableCellParagraphText,
  replaceText,
  setParagraphHeading,
  setParagraphAlignment,
  applyRunStyle,
  toggleRunStyleFlag,
  setRunHighlight,
  setRunColor,
  serializeParagraphsForClipboard,
  parseParagraphsFromClipboard,
  splitParagraphChildrenAtTextOffsets,
  createBlankDocumentModel,
  defaultStarterModel,
} from "./model"

// WASM & OOXML
export {
  setWasmSource,
  getWasmSource,
  initWasm,
  parseDocx,
  packageToArrayBuffer,
  createMinimalDocxPackage,
  getPart,
  withPart,
  buildDocModel,
  buildDocModelFromBytes,
  modelToDocumentXml,
  serializeDocModel,
  serializeDocx,
} from "./wasm"


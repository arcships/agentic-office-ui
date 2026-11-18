// engine barrel — DOCX engine layer (wasm + ooxml-core + serializer + doc-model).
// All zero-React modules ported from upstream @extend-ai/react-docx commit 6f70b92.

// Document model types + clone + normalize (the contract layer)
export * from "./types";
// The granular run/paragraph clone helpers stay internal: editor code imports
// them from "./clone". The document-copy contract is `cloneDocModel` (full
// clone), `cloneDocModelShallow` (copy-on-write shell), `cloneDocNode` (single
// node), `cloneDocModelNodes` (several top-level nodes) and
// `cloneDocModelSectionLists` (header/footer sections).
export {
  cloneDocModel,
  cloneDocModelNodes,
  cloneDocModelSectionLists,
  cloneDocModelShallow,
  cloneDocNode,
  cloneNumberingDefinitions,
} from "./clone";
export { normalizeDocModel } from "./normalize";

// Comment ranges as character offsets (parse + edit-time maintenance)
export {
  parseParagraphCommentRanges,
  mergeParagraphCommentRanges,
  remapParagraphCommentRanges,
  resolveParagraphCommentRanges,
  shiftParagraphCommentRanges,
  splitParagraphCommentRanges,
} from "./comment-ranges";

// DocModel construction (buildDocModel / buildDocModelFromBytes)
export { buildDocModel, buildDocModelFromBytes } from "./doc-model";

// Note numbering/placement properties (w:footnotePr / w:endnotePr)
export {
  parseNoteNumberingProperties,
  parseNoteProperties,
  mergeNoteNumberingProperties,
} from "./note-properties";

// Page background colour and pattern parsing (w:background)
export { parseDocumentBackground } from "./page-background";
export type { ParsedDocumentBackground } from "./page-background";

// Document layout inputs from word/settings.xml (compatibility mode, default
// tab stop) and the paragraph rules that depend on them (HTML auto spacing).
export {
  attachDocumentDefaults,
  parseCompatibilityMode,
  parseDefaultTabStopTwips,
  parseParagraphAutoSpacing,
} from "./document-defaults";
export type { ParsedParagraphAutoSpacing } from "./document-defaults";
export {
  WORD_DEFAULT_TAB_STOP_TWIPS,
  allowsWrappedLeftTabRefit,
  documentHonoursAutoSpacing,
  documentLayoutCacheKey,
  resolveDefaultTabStopPx,
  resolveDocumentLayoutDefaults,
  tabGridLayoutDefaults,
  twipsToPixelsExact,
  twipsToWholePixels,
} from "./layout-defaults";
export type {
  DocumentLayoutDefaults,
  DocumentLayoutDefaultsSource,
} from "./layout-defaults";
export {
  HTML_AUTO_SPACING_PX,
  HTML_AUTO_SPACING_TWIPS,
  paragraphsShareListNumbering,
  paragraphsSuppressAutoSpacing,
  resolveParagraphAutoSpacingInStory,
  resolveParagraphAutoSpacingPx,
  resolveParagraphSpacingPx,
} from "./paragraph-auto-spacing";
export type {
  ParagraphAutoSpacingContext,
  ParagraphAutoSpacingPx,
} from "./paragraph-auto-spacing";

// WASM bridge: initWasm / setWasmSource / wasmParseDocx / wasmSerializeDocx / ...
export type { WasmSource, WasmOoxmlPart, WasmOoxmlPackage, LegacyWasmOoxmlPackage } from "./wasm";
export {
  setWasmSource,
  initWasm,
  docModelToWasmJson,
  wasmPackageToMaps,
  mapsToWasmPackage,
  wasmParseDocx,
  wasmBuildDocModelFromPackage,
  wasmBuildDocModelFromBytes,
  wasmSerializeDocx,
  wasmModelToDocumentXml,
  wasmPackageToArrayBuffer,
} from "./wasm";

// OOXML package operations
export type { OoxmlPart, OoxmlPackage } from "./ooxml-core";
export {
  parseDocx,
  packageToArrayBuffer,
  createMinimalDocxPackage,
  getPart,
  withPart,
} from "./ooxml-core";

// Serialization: model → document.xml / OoxmlPackage / ArrayBuffer
export { modelToDocumentXml, serializeDocModel, serializeDocx } from "./serializer";
export {
  DocxSerializationError,
  prepareModelForSerialization
} from "./prepare-model";
export type { DocxSerializationErrorCode } from "./prepare-model";
export {
  hasCollapsedTableGrid,
  projectCollapsedTableGrid
} from "./table-grid-projection";
export type {
  GridCell,
  GridProjectionOptions,
  GridProjectionResult,
  GridRow,
  GridTable,
  LogicalBoundaryRange
} from "./table-grid-projection";

// Section page setup (w:pgSz / w:pgMar) read + rewrite helpers
export {
  DEFAULT_DOCX_PAGE_MARGINS,
  DEFAULT_DOCX_PAGE_SIZE,
  DEFAULT_SECTION_PROPERTIES_XML,
  DOCX_PAGE_SIZE_PRESETS,
  TWIPS_PER_INCH,
  matchDocxPageSizePreset,
  normalizePageSizeOrientation,
  readSectionPageMargins,
  readSectionPageSize,
  resolveDocxPageSizePreset,
  writeSectionPageMargins,
  writeSectionPageSize,
} from "./section-properties";
export type {
  DocxPageMargins,
  DocxPageMarginsUpdate,
  DocxPageOrientation,
  DocxPageSize,
  DocxPageSizePreset,
  DocxPageSizeUpdate,
} from "./section-properties";

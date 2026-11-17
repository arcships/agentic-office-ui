// engine barrel — DOCX engine layer (wasm + ooxml-core + serializer + doc-model).
// All zero-React modules ported from upstream @extend-ai/react-docx commit 6f70b92.

// Document model types + clone + normalize (the contract layer)
export * from "./types";
// The granular run/paragraph clone helpers stay internal: editor code imports
// them from "./clone". `cloneDocModel`, `cloneDocModelShallow` and
// `cloneDocNode` are the document-copy contract (full clone vs copy-on-write).
export { cloneDocModel, cloneDocModelShallow, cloneDocNode } from "./clone";
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

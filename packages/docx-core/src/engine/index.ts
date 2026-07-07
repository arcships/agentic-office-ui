// engine barrel — DOCX engine layer (wasm + ooxml-core + serializer + doc-model).
// All zero-React modules ported from upstream @extend-ai/react-docx commit 6f70b92.

// Document model types + clone + normalize (the contract layer)
export * from "./types";
export { cloneDocModel } from "./clone";
export { normalizeDocModel } from "./normalize";

// DocModel construction (buildDocModel / buildDocModelFromBytes)
export { buildDocModel, buildDocModelFromBytes } from "./doc-model";

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

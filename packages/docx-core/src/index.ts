// docx-core — engine layer + layout + editor helpers (zero React dependency)

// --- Engine layer ---
export * from "./ooxml-core";
export * from "./types";
export { cloneDocModel } from "./clone";
export { normalizeDocModel } from "./normalize";
export { buildDocModel, buildDocModelFromBytes } from "./doc-model";
// wasm low-level API (note: initWasm/setWasmSource/WasmSource are re-exported from wasm-source below)
export {
  wasmParseDocx,
  wasmBuildDocModelFromBytes,
  wasmBuildDocModelFromPackage,
  wasmSerializeDocx,
  wasmModelToDocumentXml,
  wasmPackageToArrayBuffer,
  wasmPackageToMaps,
  mapsToWasmPackage,
  docModelToWasmJson,
  type WasmOoxmlPart,
  type WasmOoxmlPackage,
} from "./wasm";
export * from "./serializer";

// --- Editor operations ---
export * from "./editor-ops";

// --- Editor types (shared with vue-docx) ---
export * from "./editor-types";

// --- Layout layer ---
// layout-core re-exports pagination + page-segmentation
export * from "./layout-core";
export * from "./layout-engine";
// layout-snapshot has some name overlaps with layout-core/layout-engine;
// re-export only the non-conflicting members
export {
  type LayoutInvalidationScope,
  type LayoutParagraphLineRange,
  type LayoutTableRowRange,
  type LayoutTableRowSlice,
  type LayoutNodeSegment,
  type FloatingObjectDraftState,
  type SelectionState,
  type LayoutLineFragment,
  type LayoutRegion,
  type BuildLayoutSnapshotArgs,
  type LayoutEditOperation,
  type LayoutPoint,
  type LayoutOffsetTarget,
  type SnapshotSelectionRect,
  mapPointToDocOffset,
  mapOffsetToCaretRect,
  resolveSelectionRectsForNode,
  withLineFragments,
  applyEditOperation,
} from "./layout-snapshot";

// --- Section layout ---
export * from "./section-layout";

// --- Pretext layout ---
export * from "./pretext-layout";

// --- Page count reconciliation ---
export * from "./page-count-reconciliation";

// --- Content signature ---
export * from "./content-signature";

// --- Image render ---
export * from "./image-render";

// --- Thumbnail raster ---
export * from "./thumbnail-raster";

// --- DOCX import ---
export * from "./docx-import";

// --- WASM source (worker-safe initWasm/setWasmSource wrapping wasm.ts) ---
export * from "./wasm-source";

// --- Editor state ---
export * from "./core/state";

// --- Canvas ---
export * from "./canvas/types";
export * from "./canvas/layout-diagnostics";

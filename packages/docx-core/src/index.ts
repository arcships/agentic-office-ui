// docx-core — upstream DOCX engine port (modular remigration in progress).
//
// Migration source: @extend-ai/react-docx @ commit 6f70b92
// Upstream path: /Users/eric8810/Code/extend-ui-upstream/react-docx/packages
//
// This port rebuilds the engine module-by-module with sane file boundaries,
// instead of copying upstream's 56k-line editor.tsx verbatim.
//
// Layered structure:
//   engine/   — wasm + ooxml-core + serializer + doc-model (types/clone/normalize)
//   editor/   — editor-ops (pure paragraph/text editing operations)
//   layout/   — pagination + page-segmentation + layout-engine (pending)
//   viewer/   — pretext-layout + thumbnail-raster + docx-import (pending)
//   canvas/   — layout-diagnostics + canvas types (pending)

// Engine layer (complete)
export * from "./engine";

// Editor operations (complete)
export * from "./editor/editor-ops";

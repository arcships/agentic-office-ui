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
//   layout/   — pagination + page-segmentation + layout-engine (complete)
//   viewer/   — pretext-layout + thumbnail-raster + docx-import (complete)
//   canvas/   — layout-diagnostics + canvas types (complete)

// Engine layer (complete)
export * from "./engine";

// Layout layer (complete)
export {
  collectTableExplicitPageBreakInfo,
  collectTopLevelExplicitPageBreakStartNodeIndexes,
} from "./layout";
export * from "./layout";

// Editor operations — paragraph-ops + table-ops + run-style-ops (split from editor-ops.ts)
export * from "./editor";

// Editor helpers (modular split in progress — see editor/helpers/index.ts)
export * from "./editor/helpers";

// Viewer layer (complete)
export * from "./viewer";

// Canvas layer (complete)
export * from "./canvas/types";
export * from "./canvas/layout-diagnostics";

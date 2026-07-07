// viewer barrel — text layout, thumbnails, import, and viewer helpers.
// All zero-React modules ported from upstream @extend-ai/react-docx commit 6f70b92.
//
// NOTE: wasm-source, layout-snapshot, and pagination-breaks are NOT re-exported
// here because they define symbols that collide with engine/ and layout/ barrels.
// Import them directly: import { ... } from "@extend-ai/docx-core/viewer/wasm-source"
// or through the editor helpers that compose them.

// Pretext layout (text layout around exclusions)
export * from "./pretext-layout";

// Pretext multi-item layout (inline mixed-font text layout)
export * from "./pretext-items-layout";

// Pretext selection/hit-testing (caret rect, selection rects, offset at point)
export * from "./pretext-selection";

// Thumbnail rasterization (direct draw + DOM → SVG → canvas pipeline)
export * from "./thumbnail-raster";

// Thumbnail cache (LRU surface cache + SerialIdleTaskQueue)
export * from "./thumbnail-cache";

// DOCX import (main thread + worker dual path)
export * from "./docx-import";

// Content signing (FNV-1a structural hash for change detection)
export * from "./content-signature";

// Section layout (page margins, borders, twips conversion)
export * from "./section-layout";

// Image rendering (TIFF → PNG, EMF/WMF placeholder)
export * from "./image-render";

// Page count reconciliation (height-scaling page count alignment)
export * from "./page-count-reconciliation";
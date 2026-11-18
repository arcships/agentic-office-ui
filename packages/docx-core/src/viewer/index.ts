// viewer barrel — text layout, thumbnails, import, and viewer helpers.
// All zero-React modules ported from upstream @extend-ai/react-docx commit 6f70b92.
//
// NOTE: wasm-source, layout-snapshot, and pagination-breaks are internal source
// modules because their symbols collide with the public engine/layout entries.
// Consumers must use the package root (for example `setWasmSource`) instead of
// unsupported source-file deep imports.

// Inline line metrics (ascent/descent/baseline, suppressible bottom space)
export * from "./line-metrics";

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

// DOCX export (dedicated Worker; headless inline fallback)
export * from "./docx-export";

// Content signing (FNV-1a structural hash for change detection)
export * from "./content-signature";

// Section layout (page margins, borders, twips conversion)
export * from "./section-layout";

// Page background pattern/shading paint (CSS + canvas)
export * from "./page-pattern-paint";

// Image rendering (TIFF → PNG, EMF/WMF placeholder)
export * from "./image-render";

// SVG picture layer (sanitised vector rendering + unsupported-feature fallback)
export * from "./svg-picture";

// Page count reconciliation (deprecated height-scaling helper; kept for 0.x)
export * from "./page-count-reconciliation";

// Measure→repaginate scroll anchoring helpers (#99)
export * from "./pagination-scroll-anchor";

// Text-measurement cache invalidation (web-font loading, font-set changes)
export * from "./font-metrics";

// DOCX font subsystem: advances, kerning, OpenType features, script slots and
// document font loading.
export * from "./font-advances";
export * from "./font-kerning";
export * from "./font-features";
export * from "./font-table";
export * from "./script-fonts";
export * from "./document-fonts";

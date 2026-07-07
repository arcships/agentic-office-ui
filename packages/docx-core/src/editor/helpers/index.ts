// editor/helpers barrel — DOCX editor helper modules (modular split in progress).
//
// Migration source: @extend-ai/react-docx commit 6f70b92, react-viewer/src/editor.tsx
// lines 1-24953 (pure functions / types / constants region, zero React hooks).
//
// Each module is framework-agnostic (React types replaced with plain
// equivalents) and kept ≤1000 lines per the migration hard constraints.
//
// Completed modules:
//   constants.ts          — global constants + theme/handle/wrap-mode tables
//   performance.ts        — import performance tracing
//   cache-utils.ts        — shared cache Maps/WeakMaps + cache helpers
//   editor-types.ts       — public API types (part 1: selection/range/history/...)
//   editor-types-extra.ts — public API types (part 2: controller/menu/viewer/...)
//   ooxml-helpers.ts      — OOXML parsing, embedded fonts, twips/points, xmlAttribute
//   zoom-utils.ts         — zoom/scroll/clamp utilities
//   dom-helpers.ts        — scheduled DOM writes + selection placement
//   default-model.ts      — defaultStarterModel / createBlankDocumentModel
//
// Pending modules (see docs/docx-editor-helpers-split-plan.md):
//   header-footer, page-measurement, paragraph-inspect, paragraph-geometry,
//   paragraph-tracked, pretext-build, pretext-measure, drop-cap, letterhead,
//   line-height, line-height-table, table-height, pagination-plan-core,
//   pagination-plan-iterate, pagination-plan-stabilize, style-to-css,
//   style-block-css, xml-parsing, xml-parsing-extra, synthetic-textbox,
//   tracked-changes, tracked-changes-gutter, field-helpers, numbering,
//   table-utils, table-utils-extra, text-mutation, selection-helpers,
//   selection-restore, section-manipulation, state

export * from "./constants";
export * from "./performance";
export * from "./cache-utils";
export * from "./editor-types";
export * from "./editor-types-extra";
export * from "./ooxml-helpers";
export * from "./zoom-utils";
export * from "./dom-helpers";
export * from "./default-model";
export * from "./state";
export * from "./page-measurement";
export * from "./xml-parsing";
export * from "./style-to-css";
export * from "./paragraph-inspect";
export * from "./text-mutation";
export * from "./numbering";
export * from "./table-utils";
export * from "./table-utils-extra";
export * from "./paragraph-toc";
export * from "./field-helpers";
export * from "./synthetic-textbox";
export * from "./paragraph-geometry";
export * from "./drop-cap";
export * from "./letterhead";
export * from "./line-height";
export * from "./header-footer";
export * from "./paragraph-tracked";

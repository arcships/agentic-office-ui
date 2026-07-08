// editor/helpers barrel — DOCX editor helper modules (modular split in progress).
//
// Migration source: @extend-ai/react-docx commit 6f70b92, react-viewer/src/editor.tsx
// lines 1-24953 (pure functions / types / constants region, zero React hooks).
//
// Each module is framework-agnostic (React types replaced with plain
// equivalents) and kept ≤1000 lines per the migration hard constraints.
//
// Completed modules (31):
//   constants.ts           — global constants + theme/handle/wrap-mode tables
//   performance.ts         — import performance tracing
//   cache-utils.ts         — shared cache Maps/WeakMaps + cache helpers
//   editor-types.ts        — public API types (part 1: selection/range/history/...)
//   editor-types-extra.ts  — public API types (part 2: controller/menu/viewer/...)
//   ooxml-helpers.ts       — OOXML parsing, embedded fonts, twips/points, xmlAttribute
//   zoom-utils.ts          — zoom/scroll/clamp utilities
//   dom-helpers.ts         — scheduled DOM writes + selection placement
//   default-model.ts       — defaultStarterModel / createBlankDocumentModel
//   state.ts               — editor state machine
//   page-measurement.ts    — page height measurement
//   xml-parsing.ts         — XML parsing (part 1)
//   style-to-css.ts        — runStyleToCss (part 1)
//   style-block-css.ts     — paragraph border CSS helpers
//   paragraph-inspect.ts   — paragraph property extraction (part 1)
//   text-mutation.ts       — text mutation helpers
//   numbering.ts           — list numbering
//   table-utils.ts         — table utilities (part 1)
//   table-utils-extra.ts   — table utilities (part 2)
//   paragraph-toc.ts       — paragraph table-of-contents leaf (cycle-breaker)
//   field-helpers.ts       — form-field helpers
//   synthetic-textbox.ts   — synthetic textbox
//   paragraph-geometry.ts  — paragraph geometry / wrap calculation
//   drop-cap.ts            — drop-cap layout
//   letterhead.ts          — letterhead layout
//   line-height.ts         — estimateParagraphLineHeightPx (part 1)
//   line-height-table.ts   — table row/cell height + paragraph spacing helpers
//   table-height.ts        — re-export barrel for table-height-estimate (split plan)
//   header-footer.ts       — header/footer reserve
//   paragraph-tracked.ts   — paragraph tracked-change extraction
//   selection-helpers.ts   — selection/cursor helpers (clone/normalize/compare)
//   selection-restore.ts   — DOM selection restore heuristics
//   section-manipulation.ts — section paragraph/image mutation at location
//
// Pending modules (see docs/docx-editor-helpers-split-plan.md) — 8 remaining:
//   pagination-plan-core, pagination-plan-iterate, pagination-plan-stabilize,
//   xml-parsing-extra

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
export {
  extractBalancedTagRanges,
  trackedChangeKindFromTagName,
  normalizeTrackedChangeSnippet,
  formatTrackedChangeDate,
  stripTextBoxContentFromRunXml,
  parseTrackedRunTokens,
  xmlBooleanFlag,
  xmlColorValue,
  parseRunStyleFromRunXml,
  balancedTagXmlBlocks,
  mergeTextRunStyles,
  parseParagraphAlignmentFromXml,
  parseDrawingImageTransformFromSourceXml,
  joinCssTransforms,
  resolveImageRenderTransformStyle,
} from "./xml-parsing";
export * from "./style-to-css";
export * from "./style-block-css";
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
export * from "./line-height-table";
export * from "./table-height";
export * from "./header-footer";
export * from "./paragraph-tracked";
export * from "./selection-helpers";
export * from "./selection-restore";
export * from "./section-manipulation";
export * from "./pretext-build";
export * from "./pretext-measure";
export * from "./tracked-changes";
export * from "./tracked-changes-gutter";

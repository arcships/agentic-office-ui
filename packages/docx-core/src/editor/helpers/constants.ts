// Global constants and compile-time configuration for the DOCX editor.
// Upstream editor.tsx: lines 98-351, 892-980, 1284-1348, 1383-1431,
// 1494-1508, 20078-20084.
//
// All values are framework-agnostic. Style records use the plain
// `Record<string, string | number | undefined>` shape (the project-wide
// replacement for `React.CSSProperties`).

import type { HeadingLevel, ParagraphStyleDefinition } from "../../engine/types";
import type { DocxDocumentTheme } from "./editor-types";
import type {
  DocxContextMenuAction,
  DocxContextMenuActionId,
  DocxImageWrapMode,
  DocxTableContextMenuAction
} from "./editor-types-extra";

export const HIGHLIGHT_TO_CSS: Record<string, string> = {
  yellow: "#fff59d",
  green: "#bbf7d0",
  cyan: "#a5f3fc",
  magenta: "#f5d0fe",
  red: "#fecaca",
  blue: "#bfdbfe",
  black: "#111827",
  white: "#ffffff",
};

export const DEFAULT_WORD_HEADING_STYLES: Record<
  HeadingLevel,
  Record<string, string | number | undefined>
> = {
  1: {
    fontFamily: '"Calibri Light", Calibri, sans-serif',
    fontSize: "16pt",
    fontWeight: 600,
    color: "#2f5496",
  },
  2: {
    fontFamily: '"Calibri Light", Calibri, sans-serif',
    fontSize: "13pt",
    fontWeight: 600,
    color: "#2f5496",
  },
  3: {
    fontFamily: '"Calibri", sans-serif',
    fontSize: "12pt",
    fontWeight: 600,
    color: "#1f3763",
  },
  4: {
    fontFamily: '"Calibri", sans-serif',
    fontSize: "11pt",
    fontWeight: 600,
    color: "#1f3763",
  },
  5: {
    fontFamily: '"Calibri", sans-serif',
    fontSize: "11pt",
    fontWeight: 600,
    color: "#1f3763",
  },
  6: {
    fontFamily: '"Calibri", sans-serif',
    fontSize: "11pt",
    fontWeight: 600,
    color: "#1f3763",
  },
};

export const DEFAULT_WORD_HEADING_RUN_STYLES: Record<
  HeadingLevel,
  NonNullable<ParagraphStyleDefinition["runStyle"]>
> = {
  1: {
    fontFamily: "Calibri Light",
    fontSizePt: 16,
    bold: true,
    color: "#2f5496",
  },
  2: {
    fontFamily: "Calibri Light",
    fontSizePt: 13,
    bold: true,
    color: "#2f5496",
  },
  3: {
    fontFamily: "Calibri",
    fontSizePt: 12,
    bold: true,
    color: "#1f3763",
  },
  4: {
    fontFamily: "Calibri",
    fontSizePt: 11,
    bold: true,
    color: "#1f3763",
  },
  5: {
    fontFamily: "Calibri",
    fontSizePt: 11,
    bold: true,
    color: "#1f3763",
  },
  6: {
    fontFamily: "Calibri",
    fontSizePt: 11,
    bold: true,
    color: "#1f3763",
  },
};

export const DOC_PAGE_BREAK_GAP = 28;
export const PAGE_OVERFLOW_TOLERANCE_PX = 2;
export const DEFAULT_PARAGRAPH_FONT_SIZE_PT = 11;
export const SCRIPT_FONT_SCALE = 0.65;
// Word defaults to single line spacing unless the document/style overrides it.
export const DEFAULT_PARAGRAPH_LINE_MULTIPLE = 1;
// Browser line box metrics run taller at single-spacing but converge by ~1.08.
export const WORD_SINGLE_LINE_AUTO_SCALE = 0.88;
export const WORD_SINGLE_LINE_AUTO_SCALE_SANS = 0.9;
export const WORD_SINGLE_LINE_AUTO_SCALE_SERIF = 1.08;
// A text-free paragraph occupies exactly one line at the paragraph mark's
// natural (line-height: normal) font metrics. The wrap-compensation scales
// above deliberately undersize lines to balance wrapped-line overcounting,
// but an empty paragraph has no wrapping to compensate for.
export const WORD_EMPTY_PARAGRAPH_LINE_SCALE = 1.21;
export const WORD_EMPTY_PARAGRAPH_LINE_SCALE_SERIF = 1.15;
export const WORD_EMPTY_PARAGRAPH_LINE_SCALE_SANS = 1.15;
export const WORD_AUTO_LINE_SCALE_BLEND_END_MULTIPLE = 1.08;
export const MIN_AUTO_LINE_MULTIPLE = 0.1;
export const MIN_PARAGRAPH_LINE_HEIGHT_PX = 14;
export const FLOATING_HEADER_FOOTER_CLEARANCE_BUFFER_PX = 18;
export const MEASURED_PAGE_HEADER_CLEARANCE_BUFFER_PX = 8;
export const MEASURED_PAGE_FOOTER_CLEARANCE_BUFFER_PX = 24;
export const UNOVERLAPPED_FOOTER_MIN_CLEARANCE_PX = 8;
export const FLOATING_FOOTER_BASELINE_CLEARANCE_RESERVE_PX = 8;
export const MIN_VISIBLE_FLOW_FOOTER_PAGINATION_RESERVE_PX = 8;
export const PAGINATION_MEASUREMENT_INTERACTION_DEBOUNCE_MS = 180;
export const MEASURED_BODY_FOOTER_OVERLAP_STABILITY_THRESHOLD = 1;
export const WORD_TABLE_CELL_PARAGRAPH_AUTO_LINE_TWIPS = 240;
export const WORD_TABLE_CELL_PARAGRAPH_BEFORE_TWIPS = 0;
export const WORD_TABLE_CELL_PARAGRAPH_AFTER_TWIPS = 0;
export const DOCX_IMPORT_PERFORMANCE_PREFIX = "@arcships/docx-core.import";

export const TABLE_ROW_SLICE_VISUAL_BLEED_PX = 1;
export const TABLE_CELL_SLICE_FULLY_VISIBLE_BOTTOM_BUFFER_PX = 4;
export const DEFAULT_SPLIT_PARAGRAPH_LINE_TWIPS = 259;
export const DEFAULT_SPLIT_PARAGRAPH_AFTER_TWIPS = 160;
export const PAGE_CONTENT_MEASUREMENT_IGNORE_ATTRIBUTE = "data-docx-pagination-ignore";
export const WORD_TABLE_CELL_FALLBACK_PADDING_PX = {
  top: 4,
  right: 4,
  bottom: 4,
  left: 4,
} as const;
export const DEFAULT_UNSPECIFIED_DOCX_FONT_FAMILY = "Times New Roman";
export const MAX_TRAILING_SECTION_TAIL_PARAGRAPHS = 18;
export const MAX_TRAILING_SECTION_TAIL_OVERFLOW_PX = 700;
export const SPLITTABLE_TABLE_ROW_ESTIMATE_EXTRA_LINE_COUNT = 2;
export const SPLITTABLE_TABLE_ROW_DEEP_CONTENT_NODE_THRESHOLD = 8;
export const PAGE_BREAK_XML_PATTERN = /<w:br\b[^>]*w:type="page"[^>]*\/?>/i;
export const COLUMN_BREAK_XML_PATTERN = /<w:br\b[^>]*w:type="column"[^>]*\/?>/i;
export const LAST_RENDERED_PAGE_BREAK_XML_PATTERN =
  /<w:lastRenderedPageBreak\b[^>]*\/?>/i;
export const PAGE_BREAK_BEFORE_XML_PATTERN = /<w:pageBreakBefore\b[^>]*\/?>/i;
export const SECTION_PROPERTIES_XML_PATTERN = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/i;
export const SECTION_TYPE_XML_PATTERN = /<w:type\b[^>]*w:val="([^"]+)"/i;
export const BOOKMARK_START_XML_PATTERN = /<w:bookmarkStart\b[^>]*w:name="([^"]+)"/gi;
export const FOOTNOTE_REFERENCE_XML_PATTERN =
  /<w:footnoteReference\b[^>]*w:id="(-?\d+)"/gi;
export const ENDNOTE_REFERENCE_XML_PATTERN =
  /<w:endnoteReference\b[^>]*w:id="(-?\d+)"/gi;
export const XML_CACHE_MAX_ENTRIES = 4000;
export const TEXT_MEASURE_CACHE_MAX_ENTRIES = 12000;
export const DEFAULT_TAB_STOP_PX = 48;
export const TAB_LEADER_ZONE_GAP_PX = 20;
export const EMPTY_PARAGRAPH_EXTRA_HEIGHT_PX = 0;
export const PARAGRAPH_SEGMENT_TOP_BLEED_PX = 22;
export const PARAGRAPH_SEGMENT_DESCENDER_BLEED_PX = 6;
export const PARAGRAPH_SEGMENT_VISUAL_SAFETY_PX = 24;
export const LAST_RENDERED_PAGE_BREAK_HINT_MAX_REMAINING_SPACE_RATIO = 0.18;
export const LAST_RENDERED_PAGE_BREAK_HINT_MIN_REMAINING_SPACE_PX =
  MIN_PARAGRAPH_LINE_HEIGHT_PX * 3;
export const LAST_RENDERED_PAGE_BREAK_HINT_MAX_REMAINING_SPACE_PX = 120;
export const PARAGRAPH_SEGMENT_FALLBACK_TOP_BLEED_MAX_PX = 4;
export const PARAGRAPH_SEGMENT_FALLBACK_BOTTOM_BLEED_MAX_PX = 0;
export const PARAGRAPH_SEGMENT_FALLBACK_VISUAL_SAFETY_PX = 4;
export const INITIAL_PAGINATION_PREMEASURE_PAGE_LIMIT = 8;
export const INITIAL_PAGINATION_PAGE_COUNT_OSCILLATION_DISTINCT_THRESHOLD = 2;
export const INITIAL_PAGINATION_PAGE_COUNT_OSCILLATION_CHANGE_THRESHOLD = 4;
export const INITIAL_PAGINATION_BACKGROUND_REFINEMENT_DELAY_MS = 96;
export const DEFAULT_PAGE_VIRTUALIZATION_OVERSCAN = 2;
export const LARGE_TABLE_PAGE_VIRTUALIZATION_OVERSCAN = 0;
export const LARGE_TABLE_PAGE_ADJACENT_RENDER_COUNT = 1;
export const DEFAULT_PAGE_VIRTUALIZATION_SETTLE_DELAY_MS = 350;
export const PAGE_SCROLL_MEASUREMENT_SUSPEND_MS = 220;
export const ENABLE_TABLE_ROW_SLICING = true;
export const TABLE_ROW_HEIGHT_PAGINATION_ESTIMATE_PADDING_RATIO = 1.08;
export const TABLE_ROW_HEIGHT_PAGINATION_ESTIMATE_PADDING_MIN_ROWS = 15;
export const MIN_TABLE_ROW_SLICE_REMAINING_HEIGHT_PX =
  MIN_PARAGRAPH_LINE_HEIGHT_PX * 2;
export const TABLE_ROW_SLICE_BOUNDARY_TOLERANCE_PX = 2;
export const TOP_AND_BOTTOM_VERTICAL_DRAG_SNAP_PX = 10;
export const HEADER_FOOTER_INACTIVE_OPACITY = 0.5;
export const LETTERHEAD_INDENT_MIN_TWIPS = 900;
export const LETTERHEAD_INDENT_MAX_TWIPS = 4200;
export const LETTERHEAD_MAX_TEXT_LENGTH = 96;
export const LETTERHEAD_FRAME_NEARBY_NODE_DISTANCE = 3;
export const ACTIVE_RANGE_SYNC_KEYS = new Set([
  "ArrowLeft",
  "ArrowRight",
  "ArrowUp",
  "ArrowDown",
  "Home",
  "End",
  "PageUp",
  "PageDown",
  "Backspace",
  "Delete",
  "Enter",
  "Tab",
]);
export const MODIFIER_ONLY_KEYS = new Set([
  "Shift",
  "Control",
  "Meta",
  "Alt",
  "CapsLock",
]);

// Surface theme styling (upstream lines 892-924).
export const DOC_SURFACE_STYLE_BY_THEME: Record<
  DocxDocumentTheme,
  Record<string, string | number | undefined>
> = {
  light: {
    backgroundColor: "#ffffff",
    color: "#111827",
    colorScheme: "light",
    border: "none",
    boxShadow:
      "0 2px 10px rgba(15, 23, 42, 0.08), 0 1px 2px rgba(15, 23, 42, 0.05)",
  },
  dark: {
    backgroundColor: "#0a0a0a",
    color: "#f3f4f6",
    colorScheme: "dark",
    border: "none",
    boxShadow:
      "0 2px 10px rgba(2, 6, 23, 0.55), 0 1px 2px rgba(2, 6, 23, 0.45)",
  },
};

export const NIGHT_READER_INVERSION_FILTER =
  "invert(0.95) hue-rotate(180deg) saturate(0.9) brightness(0.9) contrast(0.94)";

export function appendCssFilters(
  ...filters: Array<string | undefined | null>
): string | undefined {
  const resolvedFilters = filters
    .map((filter) => filter?.trim())
    .filter((filter): filter is string => Boolean(filter));
  return resolvedFilters.length > 0 ? resolvedFilters.join(" ") : undefined;
}

// Table resize/move handle geometry + base doc style (upstream lines 926-969).
export const TABLE_RESIZE_HANDLE_SIZE = 8;
export const TABLE_RESIZE_HANDLE_HIT_SIZE = 16;
export const TABLE_RESIZE_BORDER_SIZE = 1;
export const TABLE_MOVE_HANDLE_SIZE = 14;
export const TABLE_MOVE_HANDLE_HIT_SIZE = 18;
export const TABLE_HANDLE_HOVER_OUTSET_PX = 24;
export const TABLE_HANDLE_SAFEZONE_TOP_PX = TABLE_MOVE_HANDLE_HIT_SIZE + 4;
export const TABLE_HANDLE_SAFEZONE_LEFT_PX = TABLE_MOVE_HANDLE_HIT_SIZE + 4;
export const TABLE_HANDLE_SAFEZONE_RIGHT_PX =
  Math.ceil(TABLE_RESIZE_HANDLE_HIT_SIZE / 2) + 4;
export const TABLE_HANDLE_SAFEZONE_BOTTOM_PX = TABLE_RESIZE_HANDLE_HIT_SIZE + 8;
export const TABLE_RESIZE_HANDLE_ALIGNMENT_OFFSET_PX = 0.5;
export const TABLE_MOVE_DRAG_THRESHOLD_PX = 3;
export const TABLE_RESIZE_HANDLE_STYLE: Record<string, string | number | undefined> = {
  width: TABLE_RESIZE_HANDLE_SIZE,
  height: TABLE_RESIZE_HANDLE_SIZE,
  borderRadius: 2,
  border: "1px solid #d4d4d8",
  backgroundColor: "#ffffff",
  boxShadow: "0 1px 2px rgba(0, 0, 0, 0.14)",
  pointerEvents: "none",
};
export const TABLE_MOVE_HANDLE_STYLE: Record<string, string | number | undefined> = {
  width: TABLE_MOVE_HANDLE_SIZE,
  height: TABLE_MOVE_HANDLE_SIZE,
  borderRadius: 4,
  border: "1px solid #d4d4d8",
  backgroundColor: "#ffffff",
  boxShadow: "0 1px 2px rgba(0, 0, 0, 0.14)",
  color: "#4b5563",
  display: "grid",
  placeItems: "center",
  pointerEvents: "none",
};

// DEFAULT_DOC_PAGE_WIDTH / DEFAULT_DOC_PAGE_HEIGHT are re-exported from
// section-layout; import them from there directly when needed. The base doc
// style below mirrors upstream which referenced those constants in-scope.
import {
  DEFAULT_DOC_PAGE_WIDTH,
  DEFAULT_DOC_PAGE_HEIGHT
} from "../../viewer/section-layout";

export const BASE_DOC_STYLE: Record<string, string | number | undefined> = {
  position: "relative",
  margin: "0 auto",
  width: DEFAULT_DOC_PAGE_WIDTH,
  minHeight: DEFAULT_DOC_PAGE_HEIGHT,
  display: "block",
  gap: 0,
  transition: "box-shadow 0.2s ease",
};

export const TRACKED_CHANGE_GUTTER_WIDTH_PX = 300;
export const TRACKED_CHANGE_GUTTER_CARD_LEFT_PX = 48;
export const TRACKED_CHANGE_GUTTER_CARD_RIGHT_PX = 12;
export const TRACKED_CHANGE_GUTTER_CARD_GAP_PX = 4;
export const TRACKED_CHANGE_GUTTER_CARD_MIN_HEIGHT_PX = 30;
export const TRACKED_CHANGE_GUTTER_BEND_OFFSET_PX = 8;
export const TRACKED_CHANGE_GUTTER_CONNECTOR_LANE_GAP_PX = 7;
export const TRACKED_CHANGE_GUTTER_CONNECTOR_LANE_COUNT = 5;
export const INITIAL_PAGINATION_STABILITY_IDLE_MS = 240;

// Default context-menu configurations (upstream lines 1284-1348).
export const DEFAULT_TABLE_CONTEXT_MENU_ACTIONS: DocxTableContextMenuAction[] = [
  { id: "insert-row-above", label: "Insert row above" },
  { id: "insert-row-below", label: "Insert row below" },
  { id: "insert-column-left", label: "Insert column left" },
  { id: "insert-column-right", label: "Insert column right" },
  { id: "delete-row", label: "Delete row", destructive: true },
  { id: "delete-column", label: "Delete column", destructive: true },
  { id: "delete-table", label: "Delete table", destructive: true },
];

export function usesCommandKeyShortcutLabel(): boolean {
  if (typeof navigator === "undefined") {
    return false;
  }

  const navigatorWithUAData = navigator as Navigator & {
    userAgentData?: {
      platform?: string;
    };
  };
  const userAgentDataPlatform =
    typeof navigatorWithUAData.userAgentData?.platform === "string"
      ? navigatorWithUAData.userAgentData.platform
      : "";
  const platform =
    typeof navigator.platform === "string" ? navigator.platform : "";
  const userAgent =
    typeof navigator.userAgent === "string" ? navigator.userAgent : "";
  const combinedPlatform = `${userAgentDataPlatform} ${platform}`.toLowerCase();
  const normalizedUserAgent = userAgent.toLowerCase();

  return (
    /mac|iphone|ipad|ipod/.test(combinedPlatform) ||
    /macintosh|mac os x|iphone|ipad|ipod/.test(normalizedUserAgent)
  );
}

export function clipboardShortcutLabel(key: string): string {
  return usesCommandKeyShortcutLabel() ? `⌘${key}` : `Ctrl+${key}`;
}

export const DEFAULT_CONTEXT_MENU_CLIPBOARD_ACTIONS: DocxContextMenuAction[] = [
  { id: "cut", label: "Cut", shortcut: clipboardShortcutLabel("X") },
  { id: "copy", label: "Copy", shortcut: clipboardShortcutLabel("C") },
  { id: "paste", label: "Paste", shortcut: clipboardShortcutLabel("V") },
];

export const DEFAULT_CONTEXT_MENU_IMAGE_LAYER_ACTIONS: DocxContextMenuAction[] = [
  {
    id: "image-bring-to-front",
    label: "Bring to Front",
    children: [
      { id: "image-bring-to-front", label: "Bring to Front" },
      { id: "image-bring-forward", label: "Bring Forward" },
    ],
  },
  {
    id: "image-send-to-back",
    label: "Send to Back",
    children: [
      { id: "image-send-to-back", label: "Send to Back" },
      { id: "image-send-backward", label: "Send Backward" },
    ],
  },
];

// Image wrap-mode action table + Word z-index constants (upstream 1383-1431).
export const DOCX_IMAGE_WRAP_MODE_ACTIONS: Array<{
  actionId: DocxContextMenuActionId;
  label: string;
  mode: DocxImageWrapMode;
  separatorBefore?: boolean;
}> = [
  {
    actionId: "image-wrap-inline",
    label: "In Line with Text",
    mode: "inline",
    separatorBefore: true,
  },
  {
    actionId: "image-wrap-square",
    label: "Square",
    mode: "square",
  },
  {
    actionId: "image-wrap-tight",
    label: "Tight",
    mode: "tight",
  },
  {
    actionId: "image-wrap-through",
    label: "Through",
    mode: "through",
  },
  {
    actionId: "image-wrap-top-and-bottom",
    label: "Top and Bottom",
    mode: "topAndBottom",
  },
  {
    actionId: "image-wrap-behind-text",
    label: "Behind Text",
    mode: "behindText",
    separatorBefore: true,
  },
  {
    actionId: "image-wrap-in-front-of-text",
    label: "In Front of Text",
    mode: "inFrontOfText",
  },
];

export const WORD_IMAGE_Z_INDEX_MIN = 0;
export const WORD_IMAGE_Z_INDEX_STEP = 65536;
export const WORD_IMAGE_Z_INDEX_MAX = 251658240;
export const WORD_IMAGE_Z_INDEX_DEFAULT = WORD_IMAGE_Z_INDEX_MAX;

// Page/toolbar defaults (upstream 1494-1508).
export const DEFAULT_PAGE_NUMBER_START = 1;
export const DEFAULT_TOOLBAR_BORDER_SIZE_EIGHTH_PT = 8;
export const DEFAULT_TOOLBAR_BORDER_COLOR = "#000000";

// List prefix regexes (upstream 20078-20084).
export const UNORDERED_LIST_PREFIX_PATTERN = /^\s*•\s+/;
export const ORDERED_LIST_PREFIX_PATTERN = /^\s*\d+\.\s+/;
export const LIST_PREFIX_PATTERN = /^\s*(?:•\s+|\d+\.\s+)/;
export const ORDERED_LIST_PREFIX_CAPTURE_PATTERN = /^(\s*)(\d+)\.\s+/;
export const LIST_LEVEL_STEP_TWIPS = 720;
export const DEFAULT_LIST_HANGING_TWIPS = 360;
export const MAX_FALLBACK_LIST_LEVEL = 8;

import {
  cloneDocModel,
  type DocModel,
  type DocumentNoteDefinition,
  type FooterSection,
  type FormFieldRunNode,
  type HeaderSection,
  type HeadingLevel,
  type NumberingDefinitionSet,
  type NumberingLevelDefinition,
  type ParagraphStyleDefinition,
  type ParagraphAlignment,
  type ParagraphBorderSet,
  type ParagraphBorderStyle,
  type ParagraphIndent,
  type ImageRunNode,
  type ParagraphNode,
  type TableBorderSet,
  type TableBorderStyle,
  type TableCellStyle,
  type TableCellContentNode,
  type TableNode,
  type TableRowStyle,
  type TextRunNode,
} from "./doc-model";
import {
  splitParagraphChildrenAtTextOffsets,
  updateParagraphText,
  updateTableCellParagraphTextRecursive,
  updateTableCellParagraphText,
  updateTableCellText,
} from "./editor-ops";
import { type OoxmlPackage } from "./ooxml-core";
import {
  collectTableExplicitPageBreakInfo,
  collectTopLevelExplicitPageBreakStartNodeIndexes,
} from "./pagination-breaks";
import {
  reconcilePageCountCandidateToTargetCountByScalingHeight,
  resolveMeasuredBodyFooterOverlapLatchState,
  shouldAllowStoredPageCountReduction,
} from "./page-count-reconciliation";
import {
  DEFAULT_DOCUMENT_LAYOUT,
  DEFAULT_DOC_PAGE_HEIGHT,
  DEFAULT_DOC_PAGE_WIDTH,
  pageMarginPaddingStyle,
  resolveDocumentLayout,
  parseSectionLayout,
  parseSectionPageBorders,
  TWIPS_PER_PIXEL,
  type DocumentLayoutMetrics,
  twipsToPixels,
} from "./section-layout";
import {
  imageUsesPlaceholderFallback,
  resolveRenderableImageSource,
  subscribeRenderableImageSourceUpdates,
  unsupportedImageFallbackLabel,
} from "./image-render";
import {
  layoutItemsWithPretextAroundExclusions,
  layoutTextWithPretextAroundExclusions,
  measurePretextPlainTextLineCount,
  type PretextLayoutItem,
  type PretextExclusionRect,
  type PretextLineFragment,
  type PretextSelectionRect,
  type PretextVariableWidthLayout,
  resolveCaretRectAtOffset,
  resolveOffsetAtPoint,
  resolveSelectionRects,
  sliceLayoutToLineRange,
} from "./pretext-layout";
import {
  docModelThumbnailMetadataSignature,
  docNodeContentSignature,
} from "./content-signature";
import {
  blitDocxThumbnailSurface,
  DOCX_THUMBNAIL_EXCLUDE_ATTRIBUTE,
  type DocxPageThumbnailRenderSnapshot,
  type DocxPageThumbnailSnapshotElement,
  type DocxPageThumbnailTableCellSnapshot,
  type DocxPageThumbnailTextRunSnapshot,
  DocxThumbnailSurfaceCache,
  rasterizeDocxThumbnailSurface,
  renderDocxThumbnailSnapshotSurface,
  SerialIdleTaskQueue,
} from "./thumbnail-raster";
import type {
  DocxEditorSelection,
  DocxTextRange,
  DocxTextRangeBoundary,
  DocxTextRangeLocation
} from "./editor-types";

// Minimal ambient declaration for the Node.js `Buffer` global. `toBase64` below
// uses `typeof Buffer !== "undefined"` to pick the Node fast path at runtime;
// this keeps typechecking portable without pulling in @types/node.
declare const Buffer: {
  from(data: Uint8Array): { toString(encoding: string): string };
} | undefined;

const HIGHLIGHT_TO_CSS: Record<string, string> = {
  yellow: "#fff59d",
  green: "#bbf7d0",
  cyan: "#a5f3fc",
  magenta: "#f5d0fe",
  red: "#fecaca",
  blue: "#bfdbfe",
  black: "#111827",
  white: "#ffffff",
};

export type DocxDocumentTheme = "light" | "dark";
export type DocxHeadingStyleMap = Partial<
  Record<HeadingLevel, Record<string, string | number | undefined>>
>;

const DEFAULT_WORD_HEADING_STYLES: Record<HeadingLevel, Record<string, string | number | undefined>> = {
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

const DEFAULT_WORD_HEADING_RUN_STYLES: Record<
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

const DOC_PAGE_BREAK_GAP = 28;
const PAGE_OVERFLOW_TOLERANCE_PX = 2;
const DEFAULT_PARAGRAPH_FONT_SIZE_PT = 11;
const SCRIPT_FONT_SCALE = 0.65;
// Word defaults to single line spacing unless the document/style overrides it.
const DEFAULT_PARAGRAPH_LINE_MULTIPLE = 1;
// Browser line box metrics run taller at single-spacing but converge by ~1.08.
const WORD_SINGLE_LINE_AUTO_SCALE = 0.88;
const WORD_SINGLE_LINE_AUTO_SCALE_SANS = 0.9;
const WORD_SINGLE_LINE_AUTO_SCALE_SERIF = 1.08;
// A text-free paragraph occupies exactly one line at the paragraph mark's
// natural (line-height: normal) font metrics. The wrap-compensation scales
// above deliberately undersize lines to balance wrapped-line overcounting,
// but an empty paragraph has no wrapping to compensate for.
const WORD_EMPTY_PARAGRAPH_LINE_SCALE = 1.21;
const WORD_EMPTY_PARAGRAPH_LINE_SCALE_SERIF = 1.15;
const WORD_EMPTY_PARAGRAPH_LINE_SCALE_SANS = 1.15;
const WORD_AUTO_LINE_SCALE_BLEND_END_MULTIPLE = 1.08;
const MIN_AUTO_LINE_MULTIPLE = 0.1;
const MIN_PARAGRAPH_LINE_HEIGHT_PX = 14;
const FLOATING_HEADER_FOOTER_CLEARANCE_BUFFER_PX = 18;
const MEASURED_PAGE_HEADER_CLEARANCE_BUFFER_PX = 8;
const MEASURED_PAGE_FOOTER_CLEARANCE_BUFFER_PX = 24;
const UNOVERLAPPED_FOOTER_MIN_CLEARANCE_PX = 8;
const FLOATING_FOOTER_BASELINE_CLEARANCE_RESERVE_PX = 8;
const MIN_VISIBLE_FLOW_FOOTER_PAGINATION_RESERVE_PX = 8;
const PAGINATION_MEASUREMENT_INTERACTION_DEBOUNCE_MS = 180;
const MEASURED_BODY_FOOTER_OVERLAP_STABILITY_THRESHOLD = 1;
const WORD_TABLE_CELL_PARAGRAPH_AUTO_LINE_TWIPS = 240;
const WORD_TABLE_CELL_PARAGRAPH_BEFORE_TWIPS = 0;
const WORD_TABLE_CELL_PARAGRAPH_AFTER_TWIPS = 0;
const DOCX_IMPORT_PERFORMANCE_PREFIX = "react-docx.import";

function markDocxImportPerformance(name: string): void {
  if (
    typeof performance === "undefined" ||
    typeof performance.mark !== "function"
  ) {
    return;
  }

  try {
    performance.mark(name);
  } catch {
    // Performance marks are diagnostic-only.
  }
}

function measureDocxImportPerformance(
  name: string,
  startMark: string,
  endMark: string
): void {
  if (
    typeof performance === "undefined" ||
    typeof performance.measure !== "function"
  ) {
    return;
  }

  try {
    performance.measure(name, startMark, endMark);
  } catch {
    // A missing mark should not affect import.
  }
}

function createDocxImportPerformanceTraceName(fileName: string): string {
  const normalizedName = fileName.replace(/[^a-z0-9._-]+/gi, "_").slice(0, 80);
  return `${DOCX_IMPORT_PERFORMANCE_PREFIX}.${Date.now()}.${normalizedName}`;
}
const TABLE_ROW_SLICE_VISUAL_BLEED_PX = 1;
const TABLE_CELL_SLICE_FULLY_VISIBLE_BOTTOM_BUFFER_PX = 4;
const DEFAULT_SPLIT_PARAGRAPH_LINE_TWIPS = 259;
const DEFAULT_SPLIT_PARAGRAPH_AFTER_TWIPS = 160;
const PAGE_CONTENT_MEASUREMENT_IGNORE_ATTRIBUTE = "data-docx-pagination-ignore";
const WORD_TABLE_CELL_FALLBACK_PADDING_PX = {
  top: 4,
  right: 4,
  bottom: 4,
  left: 4,
} as const;
const DEFAULT_UNSPECIFIED_DOCX_FONT_FAMILY = "Times New Roman";
const MAX_TRAILING_SECTION_TAIL_PARAGRAPHS = 18;
const MAX_TRAILING_SECTION_TAIL_OVERFLOW_PX = 700;
const SPLITTABLE_TABLE_ROW_ESTIMATE_EXTRA_LINE_COUNT = 2;
const SPLITTABLE_TABLE_ROW_DEEP_CONTENT_NODE_THRESHOLD = 8;
const PAGE_BREAK_XML_PATTERN = /<w:br\b[^>]*w:type="page"[^>]*\/?>/i;
const COLUMN_BREAK_XML_PATTERN = /<w:br\b[^>]*w:type="column"[^>]*\/?>/i;
const LAST_RENDERED_PAGE_BREAK_XML_PATTERN =
  /<w:lastRenderedPageBreak\b[^>]*\/?>/i;
const PAGE_BREAK_BEFORE_XML_PATTERN = /<w:pageBreakBefore\b[^>]*\/?>/i;
const SECTION_PROPERTIES_XML_PATTERN = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/i;
const SECTION_TYPE_XML_PATTERN = /<w:type\b[^>]*w:val="([^"]+)"/i;
const BOOKMARK_START_XML_PATTERN = /<w:bookmarkStart\b[^>]*w:name="([^"]+)"/gi;
const FOOTNOTE_REFERENCE_XML_PATTERN =
  /<w:footnoteReference\b[^>]*w:id="(-?\d+)"/gi;
const ENDNOTE_REFERENCE_XML_PATTERN =
  /<w:endnoteReference\b[^>]*w:id="(-?\d+)"/gi;
const XML_CACHE_MAX_ENTRIES = 4000;
const TEXT_MEASURE_CACHE_MAX_ENTRIES = 12000;
const DEFAULT_TAB_STOP_PX = 48;
const TAB_LEADER_ZONE_GAP_PX = 20;
const EMPTY_PARAGRAPH_EXTRA_HEIGHT_PX = 0;
const PARAGRAPH_SEGMENT_TOP_BLEED_PX = 22;
const PARAGRAPH_SEGMENT_DESCENDER_BLEED_PX = 6;
const PARAGRAPH_SEGMENT_VISUAL_SAFETY_PX = 24;
const LAST_RENDERED_PAGE_BREAK_HINT_MAX_REMAINING_SPACE_RATIO = 0.18;
const LAST_RENDERED_PAGE_BREAK_HINT_MIN_REMAINING_SPACE_PX =
  MIN_PARAGRAPH_LINE_HEIGHT_PX * 3;
const LAST_RENDERED_PAGE_BREAK_HINT_MAX_REMAINING_SPACE_PX = 120;
const PARAGRAPH_SEGMENT_FALLBACK_TOP_BLEED_MAX_PX = 4;
const PARAGRAPH_SEGMENT_FALLBACK_BOTTOM_BLEED_MAX_PX = 0;
const PARAGRAPH_SEGMENT_FALLBACK_VISUAL_SAFETY_PX = 4;
const INITIAL_PAGINATION_PREMEASURE_PAGE_LIMIT = 8;
const INITIAL_PAGINATION_PAGE_COUNT_OSCILLATION_DISTINCT_THRESHOLD = 2;
const INITIAL_PAGINATION_PAGE_COUNT_OSCILLATION_CHANGE_THRESHOLD = 4;
const INITIAL_PAGINATION_BACKGROUND_REFINEMENT_DELAY_MS = 96;
const DEFAULT_PAGE_VIRTUALIZATION_OVERSCAN = 2;
const LARGE_TABLE_PAGE_VIRTUALIZATION_OVERSCAN = 0;
const LARGE_TABLE_PAGE_ADJACENT_RENDER_COUNT = 1;
const DEFAULT_PAGE_VIRTUALIZATION_SETTLE_DELAY_MS = 350;
const PAGE_SCROLL_MEASUREMENT_SUSPEND_MS = 220;
const ENABLE_TABLE_ROW_SLICING = true;
const TABLE_ROW_HEIGHT_PAGINATION_ESTIMATE_PADDING_RATIO = 1.08;
const TABLE_ROW_HEIGHT_PAGINATION_ESTIMATE_PADDING_MIN_ROWS = 15;
const MIN_TABLE_ROW_SLICE_REMAINING_HEIGHT_PX =
  MIN_PARAGRAPH_LINE_HEIGHT_PX * 2;
const TABLE_ROW_SLICE_BOUNDARY_TOLERANCE_PX = 2;
const TOP_AND_BOTTOM_VERTICAL_DRAG_SNAP_PX = 10;
const HEADER_FOOTER_INACTIVE_OPACITY = 0.5;
const LETTERHEAD_INDENT_MIN_TWIPS = 900;
const LETTERHEAD_INDENT_MAX_TWIPS = 4200;
const LETTERHEAD_MAX_TEXT_LENGTH = 96;
const LETTERHEAD_FRAME_NEARBY_NODE_DISTANCE = 3;
const ACTIVE_RANGE_SYNC_KEYS = new Set([
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
const MODIFIER_ONLY_KEYS = new Set([
  "Shift",
  "Control",
  "Meta",
  "Alt",
  "CapsLock",
]);

const paragraphBreakFlagsBySourceXml = new Map<
  string,
  {
    explicitPageBreak: boolean;
    explicitColumnBreak: boolean;
    lastRenderedPageBreak: boolean;
    pageBreakBefore: boolean;
    sectionBreakStartsNewPage: boolean;
  }
>();
const paragraphEstimatedHeightBySourceXml = new Map<
  string,
  Map<number | string, number>
>();
const tableEstimatedHeightBySourceXml = new Map<string, Map<number, number>>();
const tableEstimatedRowHeightsByNode = new WeakMap<
  TableNode,
  Map<number, number[]>
>();
const paragraphExplicitIndentBySourceXml = new Map<
  string,
  ParagraphIndent | null
>();
const paragraphDropCapBySourceXml = new Map<
  string,
  {
    type: "drop" | "margin";
    lines?: number;
  } | null
>();
interface ParagraphTrackedInlineChange {
  id: string;
  kind: DocxTrackedChangeKind;
  author?: string;
  date?: string;
  text?: string;
}

interface ParagraphTrackedDeletionSegment {
  text: string;
  change: ParagraphTrackedInlineChange;
  style?: TextRunNode["style"] | FormFieldRunNode["style"];
}

interface ParagraphTrackedMarkup {
  inlineChangeByVisibleChildIndex: Array<
    ParagraphTrackedInlineChange | undefined
  >;
  deletedSegmentsByVisibleChildIndex: Map<
    number,
    ParagraphTrackedDeletionSegment[]
  >;
  changes: ParagraphTrackedInlineChange[];
}

interface ParagraphCommentMarkup {
  commentIdsByVisibleChildIndex: Array<number[] | undefined>;
}

const paragraphTrackedMarkupBySourceXml = new Map<
  string,
  ParagraphTrackedMarkup | null
>();
const paragraphCommentMarkupBySourceXml = new Map<
  string,
  ParagraphCommentMarkup | null
>();
let paragraphMeasureCanvasContext: CanvasRenderingContext2D | undefined;
const textWidthByFontAndValue = new Map<string, number>();
const estimatedTextAdvanceWidthByFontAndValue = new Map<string, number>();
const pretextWordBreakModeByText = new Map<string, "normal" | "keep-all">();
const paragraphBaseFontSizePxByParagraph = new WeakMap<ParagraphNode, number>();
const paragraphDominantFontFamilyByParagraph = new WeakMap<
  ParagraphNode,
  string | null
>();

interface TableSpacingTwips {
  topTwips?: number;
  rightTwips?: number;
  bottomTwips?: number;
  leftTwips?: number;
}

function setCacheEntry<K, V>(cache: Map<K, V>, key: K, value: V): void {
  if (!cache.has(key) && cache.size >= XML_CACHE_MAX_ENTRIES) {
    const oldestKey = cache.keys().next().value as K | undefined;
    if (oldestKey !== undefined) {
      cache.delete(oldestKey);
    }
  }
  cache.set(key, value);
}

function widthCacheKeyPx(widthPx?: number): number {
  if (!Number.isFinite(widthPx) || (widthPx as number) <= 0) {
    return -1;
  }

  return Math.max(1, Math.round(widthPx as number));
}

function heightEstimateCacheKeyPx(
  widthPx?: number,
  docGridLinePitchPx?: number,
  disableDocGridSnap = false
): number {
  const widthKey = widthCacheKeyPx(widthPx);
  const docGridKey =
    Number.isFinite(docGridLinePitchPx) && (docGridLinePitchPx as number) > 0
      ? Math.max(0, Math.round(docGridLinePitchPx as number))
      : 0;

  return (
    (widthKey + 2) * 10_000 + docGridKey * 2 + (disableDocGridSnap ? 1 : 0)
  );
}

function xmlAttribute(tagXml: string, attribute: string): string | undefined {
  const escaped = attribute.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = tagXml.match(
    new RegExp(`${escaped}=(?:"([^"]+)"|'([^']+)')`, "i")
  );
  return match?.[1] ?? match?.[2];
}

function headingLevelFromStyleLabel(value?: string): HeadingLevel | undefined {
  if (!value) {
    return undefined;
  }

  const match = value.match(/(?:^|[\s_-])(?:heading|h)\s*([1-6])(?:$|[\s_-])/i);
  if (!match?.[1]) {
    return undefined;
  }

  const parsed = Number(match[1]);
  return parsed >= 1 && parsed <= 6 ? (parsed as HeadingLevel) : undefined;
}

function resolveParagraphStyleHeadingLevel(
  styleDefinition: Pick<
    ParagraphStyleDefinition,
    "headingLevel" | "id" | "name"
  >
): HeadingLevel | undefined {
  if (
    Number.isFinite(styleDefinition.headingLevel) &&
    (styleDefinition.headingLevel as number) >= 1 &&
    (styleDefinition.headingLevel as number) <= 6
  ) {
    return styleDefinition.headingLevel as HeadingLevel;
  }

  return (
    headingLevelFromStyleLabel(styleDefinition.id) ??
    headingLevelFromStyleLabel(styleDefinition.name)
  );
}

function normalizeParagraphStyleDefinitionsForUi(
  styles: ParagraphStyleDefinition[]
): ParagraphStyleDefinition[] {
  return styles.map((styleDefinition) => {
    const headingLevel = resolveParagraphStyleHeadingLevel(styleDefinition);
    const headingRunStyle = headingLevel
      ? DEFAULT_WORD_HEADING_RUN_STYLES[headingLevel]
      : undefined;
    const mergedRunStyle = headingRunStyle
      ? {
          ...headingRunStyle,
          ...(styleDefinition.runStyle ?? {}),
        }
      : styleDefinition.runStyle;
    const runStyleChanged =
      mergedRunStyle !== styleDefinition.runStyle ||
      (headingRunStyle !== undefined && styleDefinition.runStyle === undefined);
    const headingChanged = headingLevel !== styleDefinition.headingLevel;

    if (!runStyleChanged && !headingChanged) {
      return styleDefinition;
    }

    return {
      ...styleDefinition,
      headingLevel,
      runStyle: mergedRunStyle,
    };
  });
}

function paragraphDropCap(
  paragraph: ParagraphNode
): NonNullable<NonNullable<ParagraphNode["style"]>["dropCap"]> | undefined {
  if (paragraph.style?.dropCap) {
    return paragraph.style.dropCap;
  }

  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return undefined;
  }

  const cached = paragraphDropCapBySourceXml.get(xml);
  if (cached !== undefined) {
    return cached ?? undefined;
  }

  const framePrTag = xml.match(/<w:framePr\b[^>]*\/?>/i)?.[0];
  if (!framePrTag) {
    setCacheEntry(paragraphDropCapBySourceXml, xml, null);
    return undefined;
  }

  const dropCapRaw = xmlAttribute(framePrTag, "w:dropCap")
    ?.trim()
    .toLowerCase();
  let type: "drop" | "margin" | undefined;
  if (dropCapRaw === "drop" || dropCapRaw === "margin") {
    type = dropCapRaw;
  }
  if (!type) {
    setCacheEntry(paragraphDropCapBySourceXml, xml, null);
    return undefined;
  }

  const linesRaw = xmlAttribute(framePrTag, "w:lines");
  const parsedLines = linesRaw ? Number(linesRaw) : Number.NaN;
  const wrap = xmlAttribute(framePrTag, "w:wrap")?.trim();
  const horizontalAnchor = xmlAttribute(framePrTag, "w:hAnchor")?.trim();
  const verticalAnchor = xmlAttribute(framePrTag, "w:vAnchor")?.trim();
  const xTwipsRaw = xmlAttribute(framePrTag, "w:x");
  const yTwipsRaw = xmlAttribute(framePrTag, "w:y");
  const horizontalSpaceTwipsRaw = xmlAttribute(framePrTag, "w:hSpace");
  const verticalSpaceTwipsRaw = xmlAttribute(framePrTag, "w:vSpace");
  const xTwips = xTwipsRaw ? Number(xTwipsRaw) : Number.NaN;
  const yTwips = yTwipsRaw ? Number(yTwipsRaw) : Number.NaN;
  const horizontalSpaceTwips = horizontalSpaceTwipsRaw
    ? Number(horizontalSpaceTwipsRaw)
    : Number.NaN;
  const verticalSpaceTwips = verticalSpaceTwipsRaw
    ? Number(verticalSpaceTwipsRaw)
    : Number.NaN;
  const resolved = {
    type,
    lines:
      Number.isFinite(parsedLines) && parsedLines > 0
        ? Math.round(parsedLines)
        : undefined,
    ...(wrap ? { wrap } : undefined),
    ...(horizontalAnchor ? { horizontalAnchor } : undefined),
    ...(verticalAnchor ? { verticalAnchor } : undefined),
    ...(Number.isFinite(xTwips) ? { xTwips: Math.round(xTwips) } : undefined),
    ...(Number.isFinite(yTwips) ? { yTwips: Math.round(yTwips) } : undefined),
    ...(Number.isFinite(horizontalSpaceTwips)
      ? { horizontalSpaceTwips: Math.round(horizontalSpaceTwips) }
      : undefined),
    ...(Number.isFinite(verticalSpaceTwips)
      ? { verticalSpaceTwips: Math.round(verticalSpaceTwips) }
      : undefined),
  };
  setCacheEntry(paragraphDropCapBySourceXml, xml, resolved);
  return resolved;
}

type LetterheadFloatSide = "left" | "right";

interface LetterheadColumnGroup {
  startOffset: number;
  endOffset: number;
  leftSegments: DocumentPageNodeSegment[];
  rightSegments: DocumentPageNodeSegment[];
  entries: Array<{
    segment: DocumentPageNodeSegment;
    side: LetterheadFloatSide;
  }>;
}

function paragraphHasLegacyFrame(paragraph: ParagraphNode): boolean {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return false;
  }

  return /<w:framePr\b[^>]*\/?>/i.test(xml);
}

function paragraphLetterheadSideFromIndent(
  paragraph: ParagraphNode
): LetterheadFloatSide | undefined {
  const leftIndentTwipsRaw = paragraph.style?.indent?.leftTwips;
  const rightIndentTwipsRaw = paragraph.style?.indent?.rightTwips;
  const leftIndentTwips =
    Number.isFinite(leftIndentTwipsRaw) && (leftIndentTwipsRaw as number) > 0
      ? Math.round(leftIndentTwipsRaw as number)
      : 0;
  const rightIndentTwips =
    Number.isFinite(rightIndentTwipsRaw) && (rightIndentTwipsRaw as number) > 0
      ? Math.round(rightIndentTwipsRaw as number)
      : 0;
  const leftCandidate =
    rightIndentTwips >= LETTERHEAD_INDENT_MIN_TWIPS &&
    rightIndentTwips <= LETTERHEAD_INDENT_MAX_TWIPS;
  const rightCandidate =
    leftIndentTwips >= LETTERHEAD_INDENT_MIN_TWIPS &&
    leftIndentTwips <= LETTERHEAD_INDENT_MAX_TWIPS;

  if (leftCandidate && !rightCandidate) {
    return "left";
  }

  if (rightCandidate && !leftCandidate) {
    return "right";
  }

  if (leftCandidate && rightCandidate) {
    return rightIndentTwips >= leftIndentTwips ? "left" : "right";
  }

  return undefined;
}

function paragraphHasNearbyLegacyFrame(
  nodes: DocModel["nodes"],
  nodeIndex: number
): boolean {
  const startIndex = Math.max(
    0,
    nodeIndex - LETTERHEAD_FRAME_NEARBY_NODE_DISTANCE
  );
  const endIndex = Math.min(
    nodes.length - 1,
    nodeIndex + LETTERHEAD_FRAME_NEARBY_NODE_DISTANCE
  );

  for (let index = startIndex; index <= endIndex; index += 1) {
    const node = nodes[index];
    if (!node || node.type !== "paragraph") {
      continue;
    }
    if (paragraphHasLegacyFrame(node)) {
      return true;
    }
  }

  return false;
}

export function paragraphLetterheadFloatSideAtNodeIndex(
  nodes: DocModel["nodes"],
  nodeIndex: number
): LetterheadFloatSide | undefined {
  const node = nodes[nodeIndex];
  if (!node || node.type !== "paragraph") {
    return undefined;
  }

  const side = paragraphLetterheadSideFromIndent(node);
  if (!side) {
    return undefined;
  }

  const normalizedText = paragraphText(node).replace(/\s+/g, " ").trim();
  if (normalizedText.length > LETTERHEAD_MAX_TEXT_LENGTH) {
    return undefined;
  }

  if (
    paragraphHasLegacyFrame(node) ||
    paragraphHasNearbyLegacyFrame(nodes, nodeIndex)
  ) {
    return side;
  }

  return undefined;
}

export function paragraphLetterheadColumnGroupAtSegmentOffset(
  nodes: DocModel["nodes"],
  nodeSegments: DocumentPageNodeSegment[],
  startOffset: number
): LetterheadColumnGroup | undefined {
  if (startOffset < 0 || startOffset >= nodeSegments.length) {
    return undefined;
  }

  const leftSegments: DocumentPageNodeSegment[] = [];
  const rightSegments: DocumentPageNodeSegment[] = [];
  const entries: Array<{
    segment: DocumentPageNodeSegment;
    side: LetterheadFloatSide;
  }> = [];
  let cursor = startOffset;

  while (cursor < nodeSegments.length) {
    const segment = nodeSegments[cursor];
    if (
      !segment ||
      segment.tableRowRange ||
      paragraphSegmentHasPartialLineRange(segment.paragraphLineRange)
    ) {
      break;
    }

    const side = paragraphLetterheadFloatSideAtNodeIndex(
      nodes,
      segment.nodeIndex
    );
    if (!side) {
      break;
    }

    if (side === "left") {
      leftSegments.push(segment);
    } else {
      rightSegments.push(segment);
    }
    entries.push({
      segment,
      side,
    });
    cursor += 1;
  }

  if (
    cursor <= startOffset ||
    leftSegments.length === 0 ||
    rightSegments.length === 0
  ) {
    return undefined;
  }

  return {
    startOffset,
    endOffset: cursor,
    leftSegments,
    rightSegments,
    entries,
  };
}

export function letterheadColumnGroupContainerStyle(): Record<string, string | number | undefined> {
  return {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    columnGap: 28,
    alignItems: "start",
    width: "100%",
    maxWidth: "100%",
    boxSizing: "border-box",
  };
}

export function letterheadColumnStackStyle(): Record<string, string | number | undefined> {
  return {
    minWidth: 0,
    maxWidth: "100%",
  };
}

export function letterheadParagraphStyleAdjustments(
  side: LetterheadFloatSide,
  oppositeSide?: LetterheadFloatSide,
  options?: {
    suppressLetterheadColumnLayout?: boolean;
  }
): Record<string, string | number | undefined> {
  void side;
  void oppositeSide;
  if (options?.suppressLetterheadColumnLayout) {
    return {
      width: "100%",
      boxSizing: "border-box",
      marginLeft: 0,
      marginRight: 0,
    };
  }

  return {};
}

function clampNumber(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value));
}

function hasVerticalScrollOverflow(value: string): boolean {
  return (
    value === "auto" ||
    value === "scroll" ||
    value === "overlay" ||
    value === "hidden"
  );
}

function nearestScrollableAncestor(
  element: HTMLElement | null
): HTMLElement | null {
  let current: HTMLElement | null = element;
  while (current) {
    const style = window.getComputedStyle(current);
    if (
      hasVerticalScrollOverflow(style.overflowY) ||
      hasVerticalScrollOverflow(style.overflow)
    ) {
      return current;
    }
    current = current.parentElement;
  }
  const scrollingElement =
    typeof document !== "undefined" ? document.scrollingElement : null;
  if (scrollingElement instanceof HTMLElement) {
    return scrollingElement;
  }
  return typeof document !== "undefined" ? document.documentElement : null;
}

function resolveEffectiveZoomScale(element: HTMLElement): number {
  let current: HTMLElement | null = element;
  let scale = 1;
  while (current) {
    const zoomRaw = window.getComputedStyle(current).zoom;
    if (zoomRaw && zoomRaw !== "normal") {
      const zoom = Number.parseFloat(zoomRaw);
      if (Number.isFinite(zoom) && zoom > 0) {
        scale *= zoom;
      }
    }
    current = current.parentElement;
  }

  return Number.isFinite(scale) && scale > 0 ? scale : 1;
}

function normalizePageVirtualizationZoomScale(
  value: unknown
): number | undefined {
  const scale = Number(value);
  return Number.isFinite(scale) && scale > 0 ? scale : undefined;
}

const DOC_SURFACE_STYLE_BY_THEME: Record<
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

const NIGHT_READER_INVERSION_FILTER =
  "invert(0.95) hue-rotate(180deg) saturate(0.9) brightness(0.9) contrast(0.94)";

function appendCssFilters(
  ...filters: Array<string | undefined | null>
): string | undefined {
  const resolvedFilters = filters
    .map((filter) => filter?.trim())
    .filter((filter): filter is string => Boolean(filter));
  return resolvedFilters.length > 0 ? resolvedFilters.join(" ") : undefined;
}

const TABLE_RESIZE_HANDLE_SIZE = 8;
const TABLE_RESIZE_HANDLE_HIT_SIZE = 16;
const TABLE_RESIZE_BORDER_SIZE = 1;
const TABLE_MOVE_HANDLE_SIZE = 14;
const TABLE_MOVE_HANDLE_HIT_SIZE = 18;
const TABLE_HANDLE_HOVER_OUTSET_PX = 24;
const TABLE_HANDLE_SAFEZONE_TOP_PX = TABLE_MOVE_HANDLE_HIT_SIZE + 4;
const TABLE_HANDLE_SAFEZONE_LEFT_PX = TABLE_MOVE_HANDLE_HIT_SIZE + 4;
const TABLE_HANDLE_SAFEZONE_RIGHT_PX =
  Math.ceil(TABLE_RESIZE_HANDLE_HIT_SIZE / 2) + 4;
const TABLE_HANDLE_SAFEZONE_BOTTOM_PX = TABLE_RESIZE_HANDLE_HIT_SIZE + 8;
const TABLE_RESIZE_HANDLE_ALIGNMENT_OFFSET_PX = 0.5;
const TABLE_MOVE_DRAG_THRESHOLD_PX = 3;
const TABLE_RESIZE_HANDLE_STYLE: Record<string, string | number | undefined> = {
  width: TABLE_RESIZE_HANDLE_SIZE,
  height: TABLE_RESIZE_HANDLE_SIZE,
  borderRadius: 2,
  border: "1px solid #d4d4d8",
  backgroundColor: "#ffffff",
  boxShadow: "0 1px 2px rgba(0, 0, 0, 0.14)",
  pointerEvents: "none",
};
const TABLE_MOVE_HANDLE_STYLE: Record<string, string | number | undefined> = {
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

const BASE_DOC_STYLE: Record<string, string | number | undefined> = {
  position: "relative",
  margin: "0 auto",
  width: DEFAULT_DOC_PAGE_WIDTH,
  minHeight: DEFAULT_DOC_PAGE_HEIGHT,
  display: "block",
  gap: 0,
  transition: "box-shadow 0.2s ease",
};

const TRACKED_CHANGE_GUTTER_WIDTH_PX = 300;
const TRACKED_CHANGE_GUTTER_CARD_LEFT_PX = 48;
const TRACKED_CHANGE_GUTTER_CARD_RIGHT_PX = 12;
const TRACKED_CHANGE_GUTTER_CARD_GAP_PX = 4;
const TRACKED_CHANGE_GUTTER_CARD_MIN_HEIGHT_PX = 30;
const TRACKED_CHANGE_GUTTER_BEND_OFFSET_PX = 8;
const TRACKED_CHANGE_GUTTER_CONNECTOR_LANE_GAP_PX = 7;
const TRACKED_CHANGE_GUTTER_CONNECTOR_LANE_COUNT = 5;
const INITIAL_PAGINATION_STABILITY_IDLE_MS = 240;

function scheduleDomWrite(callback: () => void): void {
  if (
    typeof window !== "undefined" &&
    typeof window.requestAnimationFrame === "function"
  ) {
    window.requestAnimationFrame(() => {
      callback();
    });
    return;
  }

  setTimeout(callback, 0);
}

function normalizeMeasuredTableRowHeightPx(heightPx: number): number {
  if (!Number.isFinite(heightPx) || heightPx <= 0) {
    return MIN_PARAGRAPH_LINE_HEIGHT_PX;
  }

  return Math.max(MIN_PARAGRAPH_LINE_HEIGHT_PX, Math.round(heightPx));
}

function reconcileMeasuredTableRowHeightsForImportPagination(
  table: TableNode,
  measuredRowHeights: number[],
  maxAvailableWidthPx?: number,
  pageContentHeightPx?: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number
): number[] | undefined {
  if (measuredRowHeights.length !== table.rows.length) {
    return undefined;
  }

  const estimatedRowHeights = estimateTableRowHeightsPx(
    table,
    maxAvailableWidthPx,
    numberingDefinitions,
    docGridLinePitchPx,
    pageContentHeightPx
  );
  if (estimatedRowHeights.length !== table.rows.length) {
    return undefined;
  }

  return measuredRowHeights.map((heightPx, rowIndex) => {
    const row = table.rows[rowIndex];
    const normalizedMeasuredHeightPx =
      normalizeMeasuredTableRowHeightPx(heightPx);
    const normalizedPageContentHeightPx =
      Number.isFinite(pageContentHeightPx) &&
      (pageContentHeightPx as number) > 0
        ? Math.max(120, Math.round(pageContentHeightPx as number))
        : undefined;
    const estimatedHeightPx = Math.max(
      MIN_PARAGRAPH_LINE_HEIGHT_PX,
      Math.round(estimatedRowHeights[rowIndex] ?? MIN_PARAGRAPH_LINE_HEIGHT_PX)
    );
    const explicitHeightPx = twipsToPixels(row.style?.heightTwips);
    const normalizedExplicitHeightPx =
      Number.isFinite(explicitHeightPx) && (explicitHeightPx as number) > 0
        ? Math.max(
            MIN_PARAGRAPH_LINE_HEIGHT_PX,
            Math.round(explicitHeightPx as number)
          )
        : undefined;
    const allowedGrowthPx = rowHasDeepFlowContent(row)
      ? Math.max(
          MIN_PARAGRAPH_LINE_HEIGHT_PX,
          Math.round(estimatedHeightPx * 0.2)
        )
      : Math.max(4, Math.round(MIN_PARAGRAPH_LINE_HEIGHT_PX * 0.5));
    const explicitMinimumHeightPx =
      normalizedExplicitHeightPx ?? MIN_PARAGRAPH_LINE_HEIGHT_PX;
    if (rowAllowsPageSplit(row) && row.style?.heightRule !== "exact") {
      return Math.max(explicitMinimumHeightPx, normalizedMeasuredHeightPx);
    }

    if (
      Number.isFinite(normalizedPageContentHeightPx) &&
      normalizedMeasuredHeightPx >
        (normalizedPageContentHeightPx as number) + PAGE_OVERFLOW_TOLERANCE_PX
    ) {
      return Math.max(explicitMinimumHeightPx, normalizedMeasuredHeightPx);
    }

    const minimumHeightPx = Math.max(
      estimatedHeightPx,
      explicitMinimumHeightPx
    );
    const maximumHeightPx = Math.max(
      minimumHeightPx,
      (normalizedExplicitHeightPx ?? estimatedHeightPx) + allowedGrowthPx
    );
    return clampNumber(
      normalizedMeasuredHeightPx,
      minimumHeightPx,
      maximumHeightPx
    );
  });
}

export function resolveTableMeasuredRowHeightsForPagination(
  nodes: DocModel["nodes"],
  tableMeasuredRowHeights: Record<number, number[]>,
  options?: {
    allowMeasuredImportPagination?: boolean;
    activeDraftKeys?: string[];
    numberingDefinitions?: NumberingDefinitionSet;
    pageContentWidthPxByNodeIndex?: Map<number, number | undefined>;
    pageContentHeightPxByNodeIndex?: Map<number, number | undefined>;
    docGridLinePitchPxByNodeIndex?: Map<number, number | undefined>;
  }
): Record<number, number[]> | undefined {
  const activeTableIndexes = new Set<number>();

  (options?.activeDraftKeys ?? []).forEach((draftKey) => {
    const [tableIndexText] = draftKey.split(":");
    const tableIndex = Number.parseInt(tableIndexText ?? "", 10);
    if (Number.isFinite(tableIndex) && tableIndex >= 0) {
      activeTableIndexes.add(tableIndex);
    }
  });

  const includeOnlyActiveTables = activeTableIndexes.size > 0;
  if (
    includeOnlyActiveTables === false &&
    options?.allowMeasuredImportPagination !== true
  ) {
    return undefined;
  }

  const next: Record<number, number[]> = {};
  Object.entries(tableMeasuredRowHeights).forEach(
    ([tableIndexText, measuredRowHeights]) => {
      const tableIndex = Number.parseInt(tableIndexText, 10);
      const tableNode = nodes[tableIndex];
      if (
        !Number.isFinite(tableIndex) ||
        !tableNode ||
        tableNode.type !== "table" ||
        !Array.isArray(measuredRowHeights) ||
        measuredRowHeights.length !== tableNode.rows.length
      ) {
        return;
      }

      if (includeOnlyActiveTables && !activeTableIndexes.has(tableIndex)) {
        return;
      }

      const nextHeights = includeOnlyActiveTables
        ? measuredRowHeights.map((heightPx) =>
            normalizeMeasuredTableRowHeightPx(heightPx)
          )
        : reconcileMeasuredTableRowHeightsForImportPagination(
            tableNode,
            measuredRowHeights,
            options?.pageContentWidthPxByNodeIndex?.get(tableIndex),
            options?.pageContentHeightPxByNodeIndex?.get(tableIndex),
            options?.numberingDefinitions,
            options?.docGridLinePitchPxByNodeIndex?.get(tableIndex)
          ) ??
          measuredRowHeights.map((heightPx) =>
            normalizeMeasuredTableRowHeightPx(heightPx)
          );

      next[tableIndex] = nextHeights;
    }
  );

  return Object.keys(next).length > 0 ? next : undefined;
}

function placeCaretInsideElementDom(element: HTMLElement): void {
  const selection = window.getSelection();
  if (!selection) {
    return;
  }

  const range = document.createRange();
  range.selectNodeContents(element);
  range.collapse(false);
  selection.removeAllRanges();
  selection.addRange(range);
}

function selectionOffsetsWithinElementDom(
  element: HTMLElement
): { start: number; end: number } | undefined {
  const selection = window.getSelection();
  if (!selection || selection.rangeCount === 0) {
    return undefined;
  }

  const selectedRange = selection.getRangeAt(0);
  if (
    !element.contains(selectedRange.startContainer) ||
    !element.contains(selectedRange.endContainer)
  ) {
    return undefined;
  }

  try {
    const textLengthWithoutNumberingLabels = (range: Range): number => {
      const fragment = range.cloneContents();
      fragment
        .querySelectorAll("[data-docx-numbering-label='true']")
        .forEach((label) => {
          label.remove();
        });
      return fragment.textContent?.length ?? 0;
    };

    const startRange = document.createRange();
    startRange.setStart(element, 0);
    startRange.setEnd(selectedRange.startContainer, selectedRange.startOffset);

    const endRange = document.createRange();
    endRange.setStart(element, 0);
    endRange.setEnd(selectedRange.endContainer, selectedRange.endOffset);

    const startOffset = textLengthWithoutNumberingLabels(startRange);
    const endOffset = textLengthWithoutNumberingLabels(endRange);

    return {
      start: Math.min(startOffset, endOffset),
      end: Math.max(startOffset, endOffset),
    };
  } catch {
    return undefined;
  }
}

function setSelectionWithinElementByTextOffsetsDom(
  element: HTMLElement,
  startOffset: number,
  endOffset: number
): void {
  const selection = window.getSelection();
  if (!selection) {
    return;
  }

  const resolveDomPosition = (
    container: HTMLElement,
    targetOffset: number
  ): {
    node: Node;
    offset: number;
  } => {
    const safeOffset = Math.max(0, Math.round(targetOffset));
    const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
    let traversed = 0;
    let currentTextNode: Node | null = walker.nextNode();

    while (currentTextNode) {
      if (
        currentTextNode instanceof Text &&
        currentTextNode.parentElement?.closest(
          "[data-docx-numbering-label='true']"
        )
      ) {
        currentTextNode = walker.nextNode();
        continue;
      }

      const textLength = currentTextNode.textContent?.length ?? 0;
      if (traversed + textLength >= safeOffset) {
        return {
          node: currentTextNode,
          offset: Math.max(0, safeOffset - traversed),
        };
      }

      traversed += textLength;
      currentTextNode = walker.nextNode();
    }

    return {
      node: container,
      offset: container.childNodes.length,
    };
  };

  try {
    const range = document.createRange();
    const normalizedStart = Math.max(0, Math.min(startOffset, endOffset));
    const normalizedEnd = Math.max(
      normalizedStart,
      Math.max(startOffset, endOffset)
    );
    const start = resolveDomPosition(element, normalizedStart);
    const end = resolveDomPosition(element, normalizedEnd);
    range.setStart(start.node, start.offset);
    range.setEnd(end.node, end.offset);
    selection.removeAllRanges();
    selection.addRange(range);
  } catch {
    placeCaretInsideElementDom(element);
  }
}

const DEFAULT_TABLE_CONTEXT_MENU_ACTIONS: DocxTableContextMenuAction[] = [
  { id: "insert-row-above", label: "Insert row above" },
  { id: "insert-row-below", label: "Insert row below" },
  { id: "insert-column-left", label: "Insert column left" },
  { id: "insert-column-right", label: "Insert column right" },
  { id: "delete-row", label: "Delete row", destructive: true },
  { id: "delete-column", label: "Delete column", destructive: true },
  { id: "delete-table", label: "Delete table", destructive: true },
];

function usesCommandKeyShortcutLabel(): boolean {
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

function clipboardShortcutLabel(key: string): string {
  return usesCommandKeyShortcutLabel() ? `⌘${key}` : `Ctrl+${key}`;
}

const DEFAULT_CONTEXT_MENU_CLIPBOARD_ACTIONS: DocxContextMenuAction[] = [
  { id: "cut", label: "Cut", shortcut: clipboardShortcutLabel("X") },
  { id: "copy", label: "Copy", shortcut: clipboardShortcutLabel("C") },
  { id: "paste", label: "Paste", shortcut: clipboardShortcutLabel("V") },
];

const DEFAULT_CONTEXT_MENU_IMAGE_LAYER_ACTIONS: DocxContextMenuAction[] = [
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

export type DocxImageWrapMode =
  | "inline"
  | "square"
  | "tight"
  | "through"
  | "topAndBottom"
  | "behindText"
  | "inFrontOfText";

export interface DocxImageWrapState {
  mode: DocxImageWrapMode;
  moveWithText: boolean;
  fixedPositionOnPage: boolean;
}

export interface DocxImageWrapMenuOption {
  actionId: DocxContextMenuActionId | (string & {});
  label: string;
  checked: boolean;
  disabled?: boolean;
  onSelect: () => void;
}

export interface UseDocxImageWrapMenuResult {
  state: DocxImageWrapState;
  wrapOptions: DocxImageWrapMenuOption[];
  positioningOptions: DocxImageWrapMenuOption[];
  editWrapBoundaryOption: DocxImageWrapMenuOption;
  moreLayoutOptionsOption: DocxImageWrapMenuOption;
  setMode: (mode: DocxImageWrapMode) => void;
  setMoveWithText: (moveWithText: boolean) => void;
}

const DOCX_IMAGE_WRAP_MODE_ACTIONS: Array<{
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

const WORD_IMAGE_Z_INDEX_MIN = 0;
const WORD_IMAGE_Z_INDEX_STEP = 65536;
const WORD_IMAGE_Z_INDEX_MAX = 251658240;
const WORD_IMAGE_Z_INDEX_DEFAULT = WORD_IMAGE_Z_INDEX_MAX;

export function resolveFloatingZIndex(floating: {
  behindDocument?: boolean;
  zIndex?: number;
}): number {
  const rawZIndex = Number.isFinite(floating.zIndex)
    ? (floating.zIndex as number)
    : floating.behindDocument
    ? WORD_IMAGE_Z_INDEX_MIN
    : 0;
  const normalized = rawZIndex / WORD_IMAGE_Z_INDEX_STEP;
  if (floating.behindDocument) {
    return normalized;
  }
  return WORD_IMAGE_Z_INDEX_STEP + normalized;
}

export function resolveUiActiveTextRange(
  selection: DocxEditorSelection,
  activeTextRange: DocxTextRange | undefined,
  lastInViewerActiveTextRange: DocxTextRange | undefined
): DocxTextRange | undefined {
  if (activeTextRange) {
    return activeTextRange;
  }
  if (
    lastInViewerActiveTextRange &&
    selection.kind === "paragraph" &&
    lastInViewerActiveTextRange.start.location.kind === "paragraph" &&
    selection.nodeIndex === lastInViewerActiveTextRange.start.location.nodeIndex
  ) {
    return lastInViewerActiveTextRange;
  }
  return undefined;
}

export function shouldPreservePendingRunStyleBetweenRanges(
  previousRange: DocxTextRange,
  nextRange: DocxTextRange
): boolean {
  const previous = normalizeTextRange(previousRange);
  const next = normalizeTextRange(nextRange);

  if (compareTextRangeBoundaries(previous.start, previous.end) < 0) {
    return false;
  }
  if (compareTextRangeBoundaries(next.start, next.end) < 0) {
    return false;
  }
  if (!sameParagraphLocation(previous.start.location, next.start.location)) {
    return false;
  }

  return true;
}

interface SectionColumnLayout {
  count: number;
  gapPx: number;
  widthsPx?: number[];
}

const DEFAULT_PAGE_NUMBER_START = 1;

function createDefaultEditorTableBorders(): TableBorderSet {
  return {
    top: { type: "single", sizeEighthPt: 4, color: "#d1d5db" },
    right: { type: "single", sizeEighthPt: 4, color: "#d1d5db" },
    bottom: { type: "single", sizeEighthPt: 4, color: "#d1d5db" },
    left: { type: "single", sizeEighthPt: 4, color: "#d1d5db" },
    insideH: { type: "single", sizeEighthPt: 4, color: "#d1d5db" },
    insideV: { type: "single", sizeEighthPt: 4, color: "#d1d5db" },
  };
}

const DEFAULT_TOOLBAR_BORDER_SIZE_EIGHTH_PT = 8;
const DEFAULT_TOOLBAR_BORDER_COLOR = "#000000";

function twipsToSignedPixels(twips?: number): number | undefined {
  if (!Number.isFinite(twips)) {
    return undefined;
  }

  return Math.round((twips as number) / TWIPS_PER_PIXEL);
}

function pointsToPixels(points?: number): number | undefined {
  if (!Number.isFinite(points)) {
    return undefined;
  }

  return Math.max(0, Number((((points as number) * 96) / 72).toFixed(2)));
}

type EmbeddedFontFaceDescriptor = {
  family: string;
  style: "normal" | "italic";
  weight: string;
  source: ArrayBuffer;
};

function relationshipPartNameForOoxmlPart(partName: string): string {
  const segments = partName.split("/");
  const fileName = segments.pop() ?? partName;
  const directory = segments.join("/");
  return directory.length > 0
    ? `${directory}/_rels/${fileName}.rels`
    : `_rels/${fileName}.rels`;
}

function resolveRelativeOoxmlPartName(
  basePartName: string,
  target: string
): string {
  if (/^[a-z]+:/i.test(target)) {
    return target;
  }

  const normalizedBasePartName = basePartName.replace(/^\/+/, "");
  if (target.startsWith("/")) {
    return target.replace(/^\/+/, "");
  }

  const baseSegments = normalizedBasePartName.split("/");
  baseSegments.pop();
  target.split("/").forEach((segment) => {
    if (!segment || segment === ".") {
      return;
    }

    if (segment === "..") {
      if (baseSegments.length > 0) {
        baseSegments.pop();
      }
      return;
    }

    baseSegments.push(segment);
  });

  return baseSegments.join("/");
}

function parseOoxmlRelationships(
  pkg: OoxmlPackage,
  partName: string
): Map<string, string> {
  const relationships = new Map<string, string>();
  const relationshipsXml = pkg.parts.get(
    relationshipPartNameForOoxmlPart(partName)
  )?.content;
  if (!relationshipsXml) {
    return relationships;
  }

  const relationshipPattern =
    /<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*\/?>/gi;
  let match: RegExpExecArray | null;
  while ((match = relationshipPattern.exec(relationshipsXml))) {
    const relationshipId = match[1]?.trim();
    const target = match[2]?.trim();
    if (!relationshipId || !target) {
      continue;
    }

    relationships.set(
      relationshipId,
      resolveRelativeOoxmlPartName(partName, target)
    );
  }

  return relationships;
}

function deobfuscateEmbeddedFontData(
  fontData: Uint8Array,
  fontKey?: string
): ArrayBuffer {
  const output = Uint8Array.from(fontData);
  const normalizedFontKey = (fontKey ?? "").replace(/[{}-]/g, "");
  if (!/^[0-9a-f]{32}$/i.test(normalizedFontKey)) {
    return output.buffer.slice(
      output.byteOffset,
      output.byteOffset + output.byteLength
    );
  }

  const keyBytes = Uint8Array.from(
    normalizedFontKey.match(/../g)?.map((pair) => Number.parseInt(pair, 16)) ??
      []
  ).reverse();
  const xorLength = Math.min(32, output.length);
  for (let index = 0; index < xorLength; index += 1) {
    output[index] ^= keyBytes[index % keyBytes.length] ?? 0;
  }

  return output.buffer.slice(
    output.byteOffset,
    output.byteOffset + output.byteLength
  );
}

function collectEmbeddedFontFaceDescriptors(
  pkg: OoxmlPackage
): EmbeddedFontFaceDescriptor[] {
  const fontTableXml = pkg.parts.get("word/fontTable.xml")?.content;
  if (!fontTableXml) {
    return [];
  }

  const fontRelationships = parseOoxmlRelationships(pkg, "word/fontTable.xml");
  const descriptors: EmbeddedFontFaceDescriptor[] = [];
  const fontPattern =
    /<w:font\b[^>]*w:name="([^"]+)"[^>]*>([\s\S]*?)<\/w:font>/gi;
  const fontVariants = [
    { tagName: "embedRegular", style: "normal" as const, weight: "400" },
    { tagName: "embedBold", style: "normal" as const, weight: "700" },
    { tagName: "embedItalic", style: "italic" as const, weight: "400" },
    { tagName: "embedBoldItalic", style: "italic" as const, weight: "700" },
  ];

  let fontMatch: RegExpExecArray | null;
  while ((fontMatch = fontPattern.exec(fontTableXml))) {
    const family = fontMatch[1]?.trim();
    const fontXml = fontMatch[2] ?? "";
    if (!family) {
      continue;
    }

    fontVariants.forEach((variant) => {
      const tagXml =
        fontXml.match(
          new RegExp(`<w:${variant.tagName}\\b[^>]*\\/?>`, "i")
        )?.[0] ?? "";
      const relationshipId = xmlAttribute(tagXml, "r:id");
      if (!relationshipId) {
        return;
      }

      const partName = fontRelationships.get(relationshipId);
      const fontData = partName ? pkg.binaryAssets.get(partName) : undefined;
      if (!fontData) {
        return;
      }

      descriptors.push({
        family,
        style: variant.style,
        weight: variant.weight,
        source: deobfuscateEmbeddedFontData(
          fontData,
          xmlAttribute(tagXml, "w:fontKey")
        ),
      });
    });
  }

  return descriptors;
}

function parseSectionColumns(
  sectionPropertiesXml?: string
): SectionColumnLayout | undefined {
  if (!sectionPropertiesXml) {
    return undefined;
  }

  const columnsTag = sectionPropertiesXml.match(/<w:cols\b[^>]*\/?>/i)?.[0];
  if (!columnsTag) {
    return undefined;
  }

  const numberOfColumnsRaw = columnsTag.match(/w:num="([\d.]+)"/i)?.[1];
  const numberOfColumns = numberOfColumnsRaw
    ? Math.round(Number(numberOfColumnsRaw))
    : 1;
  if (!Number.isFinite(numberOfColumns) || numberOfColumns <= 1) {
    return undefined;
  }
  const columnCount = Math.max(2, numberOfColumns);

  const columnTags = [...sectionPropertiesXml.matchAll(/<w:col\b[^>]*\/>/gi)];
  const widthsTwips = columnTags
    .map((match) => Number(match[0].match(/w:w="([\d.]+)"/i)?.[1]))
    .filter((value) => Number.isFinite(value) && value > 0);
  const widthsPx = widthsTwips.map((value) =>
    Math.max(1, Math.round(twipsToPixels(value) ?? 0))
  );

  // Gap resolution: explicit space on the cols tag, then per-col space, then
  // — when every column width is declared — derived from the section body
  // width (Word keeps widths + gaps equal to the body), then Word's default.
  const colsTagSpaceRaw = columnsTag.match(/w:space="([\d.]+)"/i)?.[1];
  const firstColSpaceRaw = columnTags
    .map((match) => match[0].match(/w:space="([\d.]+)"/i)?.[1])
    .find((value) => value !== undefined);
  let columnGapTwips =
    colsTagSpaceRaw !== undefined
      ? Number(colsTagSpaceRaw)
      : firstColSpaceRaw !== undefined
      ? Number(firstColSpaceRaw)
      : undefined;
  if (columnGapTwips === undefined && widthsTwips.length === columnCount) {
    const pageWidthTwips = Number(
      sectionPropertiesXml.match(/<w:pgSz\b[^>]*w:w="([\d.]+)"/i)?.[1]
    );
    const marginTag = sectionPropertiesXml.match(/<w:pgMar\b[^>]*\/?>/i)?.[0];
    const marginLeftTwips = Number(marginTag?.match(/w:left="([\d.-]+)"/i)?.[1]);
    const marginRightTwips = Number(
      marginTag?.match(/w:right="([\d.-]+)"/i)?.[1]
    );
    if (
      Number.isFinite(pageWidthTwips) &&
      Number.isFinite(marginLeftTwips) &&
      Number.isFinite(marginRightTwips)
    ) {
      const bodyTwips = pageWidthTwips - marginLeftTwips - marginRightTwips;
      const totalWidthTwips = widthsTwips.reduce((sum, value) => sum + value, 0);
      columnGapTwips = Math.max(
        0,
        (bodyTwips - totalWidthTwips) / Math.max(1, columnCount - 1)
      );
    }
  }
  if (columnGapTwips === undefined || !Number.isFinite(columnGapTwips)) {
    columnGapTwips = 720;
  }
  const columnGapPx = twipsToPixels(columnGapTwips) ?? 24;

  return {
    count: columnCount,
    gapPx: Math.max(0, columnGapPx),
    ...(widthsPx.length === columnCount ? { widthsPx } : undefined),
  };
}

export function resolveSectionPaginationContentWidthPx(
  layout: Pick<DocumentLayoutMetrics, "pageWidthPx" | "marginsPx">,
  sectionPropertiesXml?: string
): number {
  const bodyWidthPx = Math.max(
    120,
    layout.pageWidthPx - layout.marginsPx.left - layout.marginsPx.right
  );
  const columns = parseSectionColumns(sectionPropertiesXml);
  if (!columns) {
    return bodyWidthPx;
  }

  const totalGapPx =
    Math.max(0, columns.gapPx) * Math.max(0, columns.count - 1);
  return Math.max(120, Math.round((bodyWidthPx - totalGapPx) / columns.count));
}

function parseSectionPageNumberStart(sectionPropertiesXml?: string): number {
  if (!sectionPropertiesXml) {
    return DEFAULT_PAGE_NUMBER_START;
  }

  const pageNumberTag = sectionPropertiesXml.match(
    /<w:pgNumType\b[^>]*\/?>/i
  )?.[0];
  if (!pageNumberTag) {
    return DEFAULT_PAGE_NUMBER_START;
  }

  const startRaw = pageNumberTag.match(/\bw:start="(\d+)"/i)?.[1];
  if (!startRaw) {
    return DEFAULT_PAGE_NUMBER_START;
  }

  const start = Number(startRaw);
  if (!Number.isFinite(start) || start <= 0) {
    return DEFAULT_PAGE_NUMBER_START;
  }

  return Math.max(1, Math.round(start));
}

function parseSectionPageNumberStartOverride(
  sectionPropertiesXml?: string
): number | undefined {
  if (!sectionPropertiesXml) {
    return undefined;
  }

  const pageNumberTag = sectionPropertiesXml.match(
    /<w:pgNumType\b[^>]*\/?>/i
  )?.[0];
  if (!pageNumberTag) {
    return undefined;
  }

  const startRaw = pageNumberTag.match(/\bw:start="(\d+)"/i)?.[1];
  if (!startRaw) {
    return undefined;
  }

  const start = Number(startRaw);
  if (!Number.isFinite(start) || start <= 0) {
    return undefined;
  }

  return Math.max(1, Math.round(start));
}

function parseSectionPageNumberFormat(
  sectionPropertiesXml?: string
): string | undefined {
  if (!sectionPropertiesXml) {
    return undefined;
  }

  const pageNumberTag = sectionPropertiesXml.match(
    /<w:pgNumType\b[^>]*\/?>/i
  )?.[0];
  const format = pageNumberTag?.match(/\bw:fmt="([^"]+)"/i)?.[1]?.trim();
  return format && format.length > 0 ? format : undefined;
}

function parseSectionStartType(
  sectionPropertiesXml?: string
): string | undefined {
  if (!sectionPropertiesXml) {
    return undefined;
  }

  const sectionType = sectionPropertiesXml
    .match(/<w:type\b[^>]*w:val="([^"]+)"/i)?.[1]
    ?.trim();
  return sectionType && sectionType.length > 0
    ? sectionType.toLowerCase()
    : undefined;
}

interface ResolvedDocumentSection {
  startNodeIndex: number;
  sectionPropertiesXml?: string;
  headerSections: HeaderSection[];
  footerSections: FooterSection[];
}

function normalizeSectionReferenceType(referenceType?: string): string {
  const normalized = referenceType?.trim().toLowerCase();
  return normalized && normalized.length > 0 ? normalized : "default";
}

function inheritSectionReferences<T extends HeaderSection | FooterSection>(
  sections: ResolvedDocumentSection[],
  sectionKey: "headerSections" | "footerSections"
): ResolvedDocumentSection[] {
  const inheritedByType = new Map<string, T>();

  return sections.map((section) => {
    const explicitSections = section[sectionKey] as T[];
    if (explicitSections.length > 0) {
      explicitSections.forEach((entry) => {
        inheritedByType.set(
          normalizeSectionReferenceType(entry.referenceType),
          entry
        );
      });
    }

    return {
      ...section,
      [sectionKey]: [...inheritedByType.values()],
    };
  });
}

function resolveInheritedSectionHeaderFooterReferences(
  sections: ResolvedDocumentSection[]
): ResolvedDocumentSection[] {
  if (sections.length === 0) {
    return sections;
  }

  return inheritSectionReferences(
    inheritSectionReferences(sections, "headerSections"),
    "footerSections"
  );
}

function resolveDocumentSectionsFromMetadata(
  metadata: DocModel["metadata"]
): ResolvedDocumentSection[] {
  const normalizedSections = (metadata.sections ?? [])
    .map(
      (section): ResolvedDocumentSection => ({
        startNodeIndex:
          Number.isFinite(section.startNodeIndex) &&
          (section.startNodeIndex as number) >= 0
            ? Math.round(section.startNodeIndex as number)
            : 0,
        sectionPropertiesXml: section.sectionPropertiesXml,
        headerSections: section.headerSections ?? [],
        footerSections: section.footerSections ?? [],
      })
    )
    .sort((left, right) => left.startNodeIndex - right.startNodeIndex);

  if (normalizedSections.length > 0) {
    if (normalizedSections[0].startNodeIndex > 0) {
      normalizedSections.unshift({
        startNodeIndex: 0,
        sectionPropertiesXml: normalizedSections[0].sectionPropertiesXml,
        headerSections: normalizedSections[0].headerSections,
        footerSections: normalizedSections[0].footerSections,
      });
    }
    return resolveInheritedSectionHeaderFooterReferences(normalizedSections);
  }

  return [
    {
      startNodeIndex: 0,
      sectionPropertiesXml: metadata.sectionPropertiesXml,
      headerSections: metadata.headerSections ?? [],
      footerSections: metadata.footerSections ?? [],
    },
  ];
}

interface PaginationSectionMetrics {
  startNodeIndex: number;
  pageContentWidthPx: number;
  pageContentHeightPx: number;
  pageContentHeightMultiplier?: number;
  docGridLinePitchPx?: number;
}

function paragraphHasHeaderFooterReserveRelevantContent(
  paragraph: ParagraphNode
): boolean {
  if (paragraphHasVisibleText(paragraph) || paragraphHasFormField(paragraph)) {
    return true;
  }

  if (!paragraphHasImage(paragraph)) {
    return false;
  }

  return paragraph.children.some((child) => {
    if (child.type !== "image") {
      return false;
    }

    // Inline/wrapped images are part of the header/footer flow and should keep
    // reserve. Absolute behind-text anchors (watermarks/background art) should not.
    if (!child.floating || shouldRenderWrappedFloatingImage(child)) {
      return true;
    }

    if (!shouldRenderAbsoluteFloatingImage(child)) {
      return true;
    }

    return shouldReserveHeaderFooterFloatingImageSpace(child);
  });
}

function sectionHasVisibleHeaderContent(
  section: ResolvedDocumentSection
): boolean {
  const headerSections = section.headerSections ?? [];
  return headerSections.some((headerSection) =>
    (headerSection.nodes ?? []).some(
      (node) =>
        node.type === "table" ||
        paragraphHasHeaderFooterReserveRelevantContent(node)
    )
  );
}

function sectionHasVisibleFooterContent(
  section: ResolvedDocumentSection
): boolean {
  const footerSections = section.footerSections ?? [];
  return footerSections.some((footerSection) =>
    (footerSection.nodes ?? []).some(
      (node) =>
        node.type === "table" ||
        paragraphHasHeaderFooterReserveRelevantContent(node)
    )
  );
}

function resolveHeaderFooterAbsoluteFloatingTopPx(
  image: ImageRunNode,
  layout: Pick<DocumentLayoutMetrics, "marginsPx">
): number | undefined {
  if (!shouldRenderAbsoluteFloatingImage(image) || !image.floating) {
    return undefined;
  }

  const floating = image.floating;
  const verticalRelativeTo = floating.verticalRelativeTo?.trim().toLowerCase();
  if (!Number.isFinite(floating.yPx)) {
    return undefined;
  }

  if (verticalRelativeTo === "page") {
    return Math.round(floating.yPx as number);
  }

  if (verticalRelativeTo === "margin") {
    return Math.round((floating.yPx as number) + layout.marginsPx.top);
  }

  return undefined;
}

function shouldReserveHeaderFooterFloatingImageSpace(
  image: ImageRunNode
): boolean {
  return (
    shouldRenderAbsoluteFloatingImage(image) &&
    Boolean(image.floating) &&
    image.floating?.behindDocument !== true
  );
}

function resolveFooterParagraphFloatingBoundaryTopPx(
  paragraph: ParagraphNode,
  layout: Pick<DocumentLayoutMetrics, "marginsPx">
): number | undefined {
  if (!paragraphHasImage(paragraph)) {
    return undefined;
  }

  let boundaryTopPx: number | undefined;

  paragraph.children.forEach((child) => {
    if (
      child.type !== "image" ||
      !child.floating ||
      !shouldReserveHeaderFooterFloatingImageSpace(child)
    ) {
      return;
    }

    const imageTopPx = resolveHeaderFooterAbsoluteFloatingTopPx(child, layout);
    if (!Number.isFinite(imageTopPx)) {
      return;
    }

    const distTPx = Math.max(0, Math.round(child.floating.distTPx ?? 0));
    const candidateTopPx = Math.max(
      0,
      Math.round((imageTopPx as number) - distTPx)
    );
    boundaryTopPx =
      boundaryTopPx === undefined
        ? candidateTopPx
        : Math.min(boundaryTopPx, candidateTopPx);
  });

  return boundaryTopPx;
}

function resolveFooterNodesFloatingBoundaryTopPx(
  footerNodes: DocModel["nodes"],
  layout: Pick<DocumentLayoutMetrics, "marginsPx">
): number | undefined {
  let boundaryTopPx: number | undefined;

  footerNodes.forEach((node) => {
    if (node.type !== "paragraph") {
      return;
    }

    const paragraphBoundaryTopPx = resolveFooterParagraphFloatingBoundaryTopPx(
      node,
      layout
    );
    if (!Number.isFinite(paragraphBoundaryTopPx)) {
      return;
    }

    boundaryTopPx =
      boundaryTopPx === undefined
        ? Math.round(paragraphBoundaryTopPx as number)
        : Math.min(boundaryTopPx, Math.round(paragraphBoundaryTopPx as number));
  });

  return boundaryTopPx;
}

function estimateHeaderFooterParagraphFloatingReservePx(
  paragraph: ParagraphNode,
  layout: Pick<
    DocumentLayoutMetrics,
    "pageWidthPx" | "pageHeightPx" | "marginsPx"
  > &
    Partial<Pick<DocumentLayoutMetrics, "headerDistancePx">>,
  region: "header" | "footer"
): number {
  if (!paragraphHasImage(paragraph)) {
    return 0;
  }

  const nominalBodyTopPx = layout.marginsPx.top;
  const nominalBodyBottomPx = layout.pageHeightPx - layout.marginsPx.bottom;

  if (region === "footer") {
    const floatingBoundaryTopPx = resolveFooterParagraphFloatingBoundaryTopPx(
      paragraph,
      layout
    );
    return Number.isFinite(floatingBoundaryTopPx)
      ? Math.max(
          0,
          nominalBodyBottomPx -
            Math.round(floatingBoundaryTopPx as number) +
            FLOATING_HEADER_FOOTER_CLEARANCE_BUFFER_PX
        )
      : 0;
  }

  return paragraph.children.reduce((largest, child) => {
    if (child.type !== "image" || !child.floating) {
      return largest;
    }

    const floating = child.floating;
    const imageHeightPx =
      Number.isFinite(child.heightPx) && (child.heightPx as number) > 0
        ? Math.round(child.heightPx as number)
        : Number.isFinite(child.widthPx) && (child.widthPx as number) > 0
        ? Math.round(child.widthPx as number)
        : MIN_PARAGRAPH_LINE_HEIGHT_PX;
    const distTPx = Math.max(0, Math.round(floating.distTPx ?? 0));
    const distBPx = Math.max(0, Math.round(floating.distBPx ?? 0));

    let topPx = resolveHeaderFooterAbsoluteFloatingTopPx(child, layout);
    if (!Number.isFinite(topPx)) {
      const verticalRelativeTo = floating.verticalRelativeTo
        ?.trim()
        .toLowerCase();
      const isParagraphRelativeAnchor =
        verticalRelativeTo === undefined ||
        verticalRelativeTo === "" ||
        verticalRelativeTo === "paragraph" ||
        verticalRelativeTo === "line";
      if (region === "header" && isParagraphRelativeAnchor) {
        const headerAnchorOriginPx = Math.max(
          0,
          Math.round(layout.headerDistancePx ?? layout.marginsPx.top)
        );
        topPx = headerAnchorOriginPx + Math.round(floating.yPx ?? 0);
      }
    }
    if (!Number.isFinite(topPx)) {
      return largest;
    }

    const resolvedTopPx = Math.round(topPx as number);
    const resolvedBottomPx = resolvedTopPx + imageHeightPx + distTPx + distBPx;
    const reserveBehindDocHeaderBand =
      region === "header" &&
      floating.behindDocument === true &&
      Math.max(1, Math.round(child.widthPx ?? 0)) >=
        Math.round(layout.pageWidthPx * 0.5) &&
      resolvedTopPx <= nominalBodyTopPx + 192 &&
      resolvedBottomPx > nominalBodyTopPx;
    if (
      !shouldReserveHeaderFooterFloatingImageSpace(child) &&
      !reserveBehindDocHeaderBand
    ) {
      return largest;
    }

    const bodyOverlapPx = Math.max(0, resolvedBottomPx - nominalBodyTopPx);
    if (bodyOverlapPx <= 0) {
      return largest;
    }

    return Math.max(
      largest,
      bodyOverlapPx + FLOATING_HEADER_FOOTER_CLEARANCE_BUFFER_PX
    );
  }, 0);
}

export function resolveHeaderPaginationReservePx(
  headerSections: HeaderSection[],
  layout: Pick<
    DocumentLayoutMetrics,
    | "pageWidthPx"
    | "pageHeightPx"
    | "marginsPx"
    | "headerDistancePx"
    | "docGridLinePitchPx"
  >
): number {
  const visibleHeaderSections = (headerSections ?? []).filter((headerSection) =>
    (headerSection.nodes ?? []).some(
      (node) =>
        node.type === "table" ||
        paragraphHasHeaderFooterReserveRelevantContent(node)
    )
  );
  if (visibleHeaderSections.length === 0) {
    return 0;
  }

  const availableWidthPx = Math.max(
    24,
    layout.pageWidthPx - layout.marginsPx.left - layout.marginsPx.right
  );
  const estimatedHeaderHeightPx = visibleHeaderSections.reduce(
    (largestHeightPx, headerSection) => {
      const visibleNodes = (headerSection.nodes ?? []).filter(
        (node) =>
          node.type === "table" ||
          paragraphHasHeaderFooterReserveRelevantContent(node)
      );
      if (visibleNodes.length === 0) {
        return largestHeightPx;
      }

      const nodeHeightsPx = visibleNodes.reduce((sum, node) => {
        const floatingReservePx =
          node.type === "paragraph"
            ? estimateHeaderFooterParagraphFloatingReservePx(
                node,
                layout,
                "header"
              )
            : 0;
        return (
          sum +
          Math.max(
            estimateDocNodeHeightPx(
              node,
              availableWidthPx,
              undefined,
              layout.docGridLinePitchPx
            ),
            floatingReservePx
          )
        );
      }, 0);
      const interParagraphGapPx = Math.max(0, visibleNodes.length - 1) * 8;
      return Math.max(
        largestHeightPx,
        Math.round(nodeHeightsPx + interParagraphGapPx)
      );
    },
    0
  );

  if (estimatedHeaderHeightPx <= 0) {
    return 0;
  }

  const headerBodyOverlapPx =
    estimatedHeaderHeightPx + layout.headerDistancePx - layout.marginsPx.top;
  return Math.max(0, Math.round(headerBodyOverlapPx));
}

export function resolveFooterPaginationReservePx(
  footerSections: FooterSection[],
  layout: Pick<
    DocumentLayoutMetrics,
    | "pageWidthPx"
    | "pageHeightPx"
    | "marginsPx"
    | "footerDistancePx"
    | "docGridLinePitchPx"
  >
): number {
  const visibleFooterSections = (footerSections ?? []).filter((footerSection) =>
    (footerSection.nodes ?? []).some(
      (node) =>
        node.type === "table" ||
        paragraphHasHeaderFooterReserveRelevantContent(node)
    )
  );
  if (visibleFooterSections.length === 0) {
    return 0;
  }

  const availableWidthPx = Math.max(
    24,
    layout.pageWidthPx - layout.marginsPx.left - layout.marginsPx.right
  );
  const estimatedFooterHeightPx = visibleFooterSections.reduce(
    (largestHeightPx, footerSection) => {
      const visibleNodes = (footerSection.nodes ?? []).filter(
        (node) =>
          node.type === "table" ||
          paragraphHasHeaderFooterReserveRelevantContent(node)
      );
      if (visibleNodes.length === 0) {
        return largestHeightPx;
      }

      const nodeHeightsPx = visibleNodes.reduce((sum, node) => {
        const floatingReservePx =
          node.type === "paragraph"
            ? estimateHeaderFooterParagraphFloatingReservePx(
                node,
                layout,
                "footer"
              )
            : 0;
        return (
          sum +
          Math.max(
            estimateDocNodeHeightPx(
              node,
              availableWidthPx,
              undefined,
              layout.docGridLinePitchPx
            ),
            floatingReservePx
          )
        );
      }, 0);
      const interParagraphGapPx = Math.max(0, visibleNodes.length - 1) * 8;
      return Math.max(
        largestHeightPx,
        Math.round(nodeHeightsPx + interParagraphGapPx)
      );
    },
    0
  );

  if (estimatedFooterHeightPx <= 0) {
    return 0;
  }

  const nominalBodyBottomPx = layout.pageHeightPx - layout.marginsPx.bottom;
  const floatingFooterBoundaryTopPx = visibleFooterSections.reduce<
    number | undefined
  >((smallestTopPx, footerSection) => {
    const sectionFloatingTopPx = resolveFooterNodesFloatingBoundaryTopPx(
      footerSection.nodes ?? [],
      layout
    );
    if (!Number.isFinite(sectionFloatingTopPx)) {
      return smallestTopPx;
    }

    return smallestTopPx === undefined
      ? Math.round(sectionFloatingTopPx as number)
      : Math.min(smallestTopPx, Math.round(sectionFloatingTopPx as number));
  }, undefined);
  const hasFloatingFooterBodyIntrusionRisk = visibleFooterSections.some(
    (footerSection) =>
      (footerSection.nodes ?? []).some(
        (node) =>
          node.type === "paragraph" &&
          estimateHeaderFooterParagraphFloatingReservePx(
            node,
            layout,
            "footer"
          ) > 0
      )
  );

  const footerBodyOverlapPx =
    estimatedFooterHeightPx + layout.footerDistancePx - layout.marginsPx.bottom;
  const explicitFooterBoundaryReservePx = Math.max(
    0,
    Math.round(footerBodyOverlapPx + UNOVERLAPPED_FOOTER_MIN_CLEARANCE_PX)
  );
  const floatingFooterBoundaryReservePx = Number.isFinite(
    floatingFooterBoundaryTopPx
  )
    ? Math.max(
        0,
        Math.round(
          nominalBodyBottomPx -
            Math.round(floatingFooterBoundaryTopPx as number) +
            UNOVERLAPPED_FOOTER_MIN_CLEARANCE_PX
        )
      )
    : 0;

  const resolvedReservePx = hasFloatingFooterBodyIntrusionRisk
    ? Math.max(
        explicitFooterBoundaryReservePx,
        floatingFooterBoundaryReservePx,
        FLOATING_FOOTER_BASELINE_CLEARANCE_RESERVE_PX
      )
    : Math.max(
        explicitFooterBoundaryReservePx,
        floatingFooterBoundaryReservePx
      );

  return Math.max(
    estimatedFooterHeightPx > 0
      ? MIN_VISIBLE_FLOW_FOOTER_PAGINATION_RESERVE_PX
      : 0,
    resolvedReservePx
  );
}

function resolveMeasuredPageContentHeightDiagnostics(params: {
  pageLayout: Pick<
    DocumentLayoutMetrics,
    "pageHeightPx" | "marginsPx" | "footerDistancePx"
  >;
  fallbackHeightPx: number;
  headerHeightPx: number;
  currentMeasuredHeightPx?: number;
  bodyTopPx?: number;
  bodyRenderedBottomPx?: number;
  footerTopPx?: number;
  skipBodyBottomAdjustment?: boolean;
}): {
  heightPx: number;
  bodyOverrunsFooter: boolean;
} {
  const {
    pageLayout,
    fallbackHeightPx,
    headerHeightPx,
    currentMeasuredHeightPx,
    bodyTopPx,
    bodyRenderedBottomPx,
    footerTopPx,
    skipBodyBottomAdjustment = false,
  } = params;
  const effectiveBodyRenderedBottomPx = skipBodyBottomAdjustment
    ? undefined
    : bodyRenderedBottomPx;
  const effectiveCurrentMeasuredHeightPx = skipBodyBottomAdjustment
    ? undefined
    : currentMeasuredHeightPx;

  const nominalBodyBottomPx =
    pageLayout.pageHeightPx - pageLayout.marginsPx.bottom;
  const footerOverlapPx = Number.isFinite(footerTopPx)
    ? Math.max(0, Math.round(nominalBodyBottomPx - (footerTopPx as number)))
    : 0;
  const allowedBodyBottomPx = nominalBodyBottomPx - footerOverlapPx;
  const hardFooterBottomLimitPx =
    Number.isFinite(footerTopPx) && footerOverlapPx === 0
      ? Math.round(
          (footerTopPx as number) - UNOVERLAPPED_FOOTER_MIN_CLEARANCE_PX
        )
      : undefined;
  const guardedAllowedBodyBottomPx = Number.isFinite(hardFooterBottomLimitPx)
    ? Math.min(allowedBodyBottomPx, hardFooterBottomLimitPx as number)
    : allowedBodyBottomPx;
  const renderedBodyOverrunPx = Number.isFinite(effectiveBodyRenderedBottomPx)
    ? Math.max(
        0,
        Math.round(
          (effectiveBodyRenderedBottomPx as number) - guardedAllowedBodyBottomPx
        )
      )
    : 0;
  const measuredBodyToFooterGapPx = Number.isFinite(
    effectiveBodyRenderedBottomPx
  )
    ? Math.max(
        0,
        Math.round(
          guardedAllowedBodyBottomPx - (effectiveBodyRenderedBottomPx as number)
        )
      )
    : undefined;
  const measuredFooterClearanceBufferPx =
    Number.isFinite(footerTopPx) && Number.isFinite(measuredBodyToFooterGapPx)
      ? Math.max(
          0,
          MEASURED_PAGE_FOOTER_CLEARANCE_BUFFER_PX -
            (measuredBodyToFooterGapPx as number)
        )
      : 0;
  const correctedAllowedBodyBottomPx = Math.max(
    0,
    guardedAllowedBodyBottomPx - measuredFooterClearanceBufferPx
  );
  const iterativeMeasuredHeightPx =
    renderedBodyOverrunPx > 0 &&
    !Number.isFinite(bodyTopPx) &&
    Number.isFinite(effectiveCurrentMeasuredHeightPx) &&
    (effectiveCurrentMeasuredHeightPx as number) > 0
      ? Math.max(
          120,
          Math.round(
            (effectiveCurrentMeasuredHeightPx as number) -
              renderedBodyOverrunPx -
              measuredFooterClearanceBufferPx
          )
        )
      : undefined;
  const bodyOverrunsFooter = renderedBodyOverrunPx > 0;

  if (
    Number.isFinite(bodyTopPx) &&
    correctedAllowedBodyBottomPx > (bodyTopPx as number)
  ) {
    const measuredHeaderClearanceBufferPx =
      headerHeightPx > 0 ? MEASURED_PAGE_HEADER_CLEARANCE_BUFFER_PX : 0;
    const correctedHeightPx = Math.max(
      120,
      Math.round(correctedAllowedBodyBottomPx - (bodyTopPx as number))
    );
    return {
      heightPx: Number.isFinite(iterativeMeasuredHeightPx)
        ? Math.min(
            Math.max(120, correctedHeightPx - measuredHeaderClearanceBufferPx),
            Math.max(
              120,
              (iterativeMeasuredHeightPx as number) -
                measuredHeaderClearanceBufferPx
            )
          )
        : Math.max(120, correctedHeightPx - measuredHeaderClearanceBufferPx),
      bodyOverrunsFooter,
    };
  }

  const measuredHeaderClearanceBufferPx =
    headerHeightPx > 0 ? MEASURED_PAGE_HEADER_CLEARANCE_BUFFER_PX : 0;
  const correctedFallbackHeightPx = Math.max(
    120,
    fallbackHeightPx -
      headerHeightPx -
      measuredHeaderClearanceBufferPx -
      footerOverlapPx -
      renderedBodyOverrunPx -
      measuredFooterClearanceBufferPx
  );
  return {
    heightPx: Number.isFinite(iterativeMeasuredHeightPx)
      ? Math.min(correctedFallbackHeightPx, iterativeMeasuredHeightPx as number)
      : correctedFallbackHeightPx,
    bodyOverrunsFooter,
  };
}

export function resolveMeasuredPageContentHeightPx(params: {
  pageLayout: Pick<
    DocumentLayoutMetrics,
    "pageHeightPx" | "marginsPx" | "footerDistancePx"
  >;
  fallbackHeightPx: number;
  headerHeightPx: number;
  currentMeasuredHeightPx?: number;
  bodyTopPx?: number;
  bodyRenderedBottomPx?: number;
  footerTopPx?: number;
  skipBodyBottomAdjustment?: boolean;
}): number {
  return resolveMeasuredPageContentHeightDiagnostics(params).heightPx;
}

export function resolveMeasuredBodyRenderedBottomPx(
  descendants: Array<{
    bottomPx: number;
    widthPx: number;
    heightPx: number;
    ignore?: boolean;
  }>
): number | undefined {
  let visualBottomPx: number | undefined;

  descendants.forEach((descendant) => {
    if (descendant.ignore) {
      return;
    }
    if (!Number.isFinite(descendant.bottomPx)) {
      return;
    }
    if (descendant.widthPx <= 0 && descendant.heightPx <= 0) {
      return;
    }

    visualBottomPx =
      visualBottomPx === undefined
        ? descendant.bottomPx
        : Math.max(visualBottomPx, descendant.bottomPx);
  });

  return visualBottomPx;
}

export function stabilizeMeasuredPageContentHeights(
  current: number[],
  next: number[],
  options?: {
    currentPageIdentityKeys?: string[];
    nextPageIdentityKeys?: string[];
  }
): number[] {
  return next.map((heightPx, pageIndex) => {
    const roundedNextHeightPx = Math.round(heightPx);
    const currentHeightPx = current[pageIndex];
    const currentPageIdentityKey =
      options?.currentPageIdentityKeys?.[pageIndex];
    const nextPageIdentityKey = options?.nextPageIdentityKeys?.[pageIndex];
    const canPreserveConservativeHeight =
      currentPageIdentityKey === undefined ||
      nextPageIdentityKey === undefined ||
      currentPageIdentityKey === nextPageIdentityKey;
    return Number.isFinite(currentHeightPx)
      ? canPreserveConservativeHeight
        ? Math.min(Math.round(currentHeightPx as number), roundedNextHeightPx)
        : roundedNextHeightPx
      : roundedNextHeightPx;
  });
}

function documentPageNodeSegmentIdentityKey(
  segment: DocumentPageNodeSegment
): string {
  const tableRowRangeKey = segment.tableRowRange
    ? `${segment.tableRowRange.startRowIndex}-${segment.tableRowRange.endRowIndex}`
    : "none";
  const tableRowSliceKey = segment.tableRowSlice
    ? `${segment.tableRowSlice.rowIndex}-${Math.round(
        segment.tableRowSlice.startOffsetPx
      )}-${Math.round(segment.tableRowSlice.sliceHeightPx)}`
    : "none";
  const paragraphLineRangeKey = segment.paragraphLineRange
    ? `${segment.paragraphLineRange.startLineIndex}-${segment.paragraphLineRange.endLineIndex}-${segment.paragraphLineRange.totalLineCount}`
    : "none";

  return `${segment.nodeIndex}|${tableRowRangeKey}|${tableRowSliceKey}|${paragraphLineRangeKey}`;
}

function documentPageNodeSegmentsIdentityKey(
  pageSegments: DocumentPageNodeSegment[]
): string {
  return pageSegments.map(documentPageNodeSegmentIdentityKey).join("::");
}

function buildPaginationSectionMetrics(
  sections: ResolvedDocumentSection[],
  fallbackLayout: DocumentLayoutMetrics
): PaginationSectionMetrics[] {
  const fallbackWidthPx =
    resolveSectionPaginationContentWidthPx(fallbackLayout);
  const fallbackHeightPx = Math.max(
    120,
    fallbackLayout.pageHeightPx -
      fallbackLayout.marginsPx.top -
      fallbackLayout.marginsPx.bottom
  );

  if (sections.length === 0) {
    return [
      {
        startNodeIndex: 0,
        pageContentWidthPx: fallbackWidthPx,
        pageContentHeightPx: fallbackHeightPx,
        pageContentHeightMultiplier: 1,
      },
    ];
  }

  return sections
    .map((section) => {
      const layout = parseSectionLayout(section.sectionPropertiesXml);
      const sectionColumns = parseSectionColumns(section.sectionPropertiesXml);
      const pageContentHeightMultiplier = Math.max(
        1,
        sectionColumns?.count ?? 1
      );
      const hasHeaderContent = sectionHasVisibleHeaderContent(section);
      const hasFooterContent = sectionHasVisibleFooterContent(section);
      const headerPaginationReservePx = hasHeaderContent
        ? resolveHeaderPaginationReservePx(section.headerSections ?? [], layout)
        : 0;
      const footerPaginationReservePx = hasFooterContent
        ? resolveFooterPaginationReservePx(section.footerSections ?? [], layout)
        : 0;
      return {
        startNodeIndex: Math.max(0, Math.round(section.startNodeIndex)),
        pageContentWidthPx: resolveSectionPaginationContentWidthPx(
          layout,
          section.sectionPropertiesXml
        ),
        pageContentHeightPx: Math.max(
          120,
          (layout.pageHeightPx -
            layout.marginsPx.top -
            layout.marginsPx.bottom -
            headerPaginationReservePx -
            footerPaginationReservePx) *
            pageContentHeightMultiplier
        ),
        pageContentHeightMultiplier,
        docGridLinePitchPx: layout.docGridLinePitchPx,
      };
    })
    .sort((left, right) => left.startNodeIndex - right.startNodeIndex);
}

function resolvePaginationSectionMetricsIndexForNodeIndex(
  metricsBySection: PaginationSectionMetrics[],
  nodeIndex: number,
  previousSectionIndex: number
): number {
  if (metricsBySection.length === 0) {
    return 0;
  }

  const safePrevious = Math.max(
    0,
    Math.min(previousSectionIndex, metricsBySection.length - 1)
  );
  let sectionIndex = safePrevious;
  if (nodeIndex < metricsBySection[sectionIndex].startNodeIndex) {
    sectionIndex = 0;
  }

  while (
    sectionIndex + 1 < metricsBySection.length &&
    metricsBySection[sectionIndex + 1].startNodeIndex <= nodeIndex
  ) {
    sectionIndex += 1;
  }

  return sectionIndex;
}

function scalePaginationSectionMetricsHeights(
  metricsBySection: PaginationSectionMetrics[],
  heightScale: number
): PaginationSectionMetrics[] {
  if (!Number.isFinite(heightScale) || Math.abs(heightScale - 1) < 0.001) {
    return metricsBySection;
  }

  return metricsBySection.map((metrics) => ({
    ...metrics,
    pageContentHeightPx: Math.max(
      120,
      Math.round(metrics.pageContentHeightPx * heightScale)
    ),
  }));
}

function scaleMeasuredPageContentHeights(
  measuredPageContentHeightsPxByPageIndex: number[] | undefined,
  heightScale: number
): number[] | undefined {
  if (
    !measuredPageContentHeightsPxByPageIndex ||
    measuredPageContentHeightsPxByPageIndex.length === 0
  ) {
    return measuredPageContentHeightsPxByPageIndex;
  }

  if (!Number.isFinite(heightScale) || Math.abs(heightScale - 1) < 0.001) {
    return measuredPageContentHeightsPxByPageIndex;
  }

  return measuredPageContentHeightsPxByPageIndex.map((heightPx) =>
    Math.max(120, Math.round(heightPx * heightScale))
  );
}

export function resolvePageContentHeightPxForPageSegments(
  pageSegments: DocumentPageNodeSegment[],
  pageIndex: number,
  defaultPageContentHeightPx: number,
  metricsBySection: PaginationSectionMetrics[],
  measuredPageContentHeightsPxByPageIndex?: number[],
  measuredPageContentIdentityKeysByPageIndex?: string[],
  pageIdentityKey?: string
): number {
  const pageContainsOnlySplitParagraphSegments =
    documentPageContainsOnlySplitParagraphSegments(pageSegments);
  const measuredHeightPx = measuredPageContentHeightsPxByPageIndex?.[pageIndex];
  const measuredHeightMatchesCurrentPage =
    pageIdentityKey === undefined ||
    measuredPageContentIdentityKeysByPageIndex?.[pageIndex] === undefined ||
    measuredPageContentIdentityKeysByPageIndex?.[pageIndex] === pageIdentityKey;
  if (
    Number.isFinite(measuredHeightPx) &&
    measuredHeightMatchesCurrentPage &&
    !pageContainsOnlySplitParagraphSegments
  ) {
    return Math.max(120, Math.round(measuredHeightPx as number));
  }

  const firstNodeIndex = pageSegments[0]?.nodeIndex ?? 0;
  const metricsIndex = resolvePaginationSectionMetricsIndexForNodeIndex(
    metricsBySection,
    firstNodeIndex,
    0
  );

  return Math.max(
    120,
    Math.round(
      metricsBySection[metricsIndex]?.pageContentHeightPx ??
        defaultPageContentHeightPx
    )
  );
}

export function resolveRenderPageContentHeightPxForPageSegments(params: {
  pageSegments: DocumentPageNodeSegment[];
  pageIndex: number;
  defaultPageContentHeightPx: number;
  metricsBySection: PaginationSectionMetrics[];
  measuredPageContentHeightsPxByPageIndex?: number[];
  measuredPageContentIdentityKeysByPageIndex?: string[];
  pageIdentityKey?: string;
  useMeasuredPageContentHeights?: boolean;
  pageContentHeightScale?: number;
}): number {
  const firstNodeIndex = params.pageSegments[0]?.nodeIndex ?? 0;
  const metricsIndex = resolvePaginationSectionMetricsIndexForNodeIndex(
    params.metricsBySection,
    firstNodeIndex,
    0
  );
  const sectionHeightMultiplier = Math.max(
    1,
    Math.round(
      params.metricsBySection[metricsIndex]?.pageContentHeightMultiplier ?? 1
    )
  );
  const pageContainsOnlySplitParagraphSegments =
    documentPageContainsOnlySplitParagraphSegments(params.pageSegments);
  const measuredHeightPx =
    params.useMeasuredPageContentHeights === false
      ? undefined
      : params.measuredPageContentHeightsPxByPageIndex?.[params.pageIndex];
  const measuredHeightMatchesCurrentPage =
    params.pageIdentityKey === undefined ||
    params.measuredPageContentIdentityKeysByPageIndex?.[params.pageIndex] ===
      undefined ||
    params.measuredPageContentIdentityKeysByPageIndex?.[params.pageIndex] ===
      params.pageIdentityKey;
  const usesMeasuredVisualHeight =
    Number.isFinite(measuredHeightPx) &&
    measuredHeightMatchesCurrentPage &&
    !pageContainsOnlySplitParagraphSegments;
  let resolvedHeightPx = resolvePageContentHeightPxForPageSegments(
    params.pageSegments,
    params.pageIndex,
    params.defaultPageContentHeightPx,
    params.metricsBySection,
    params.useMeasuredPageContentHeights === false
      ? undefined
      : params.measuredPageContentHeightsPxByPageIndex,
    params.useMeasuredPageContentHeights === false
      ? undefined
      : params.measuredPageContentIdentityKeysByPageIndex,
    params.pageIdentityKey
  );
  if (usesMeasuredVisualHeight || sectionHeightMultiplier <= 1) {
    return resolvedHeightPx;
  }

  return Math.max(120, Math.round(resolvedHeightPx / sectionHeightMultiplier));
}

export function documentPageContainsOnlySplitParagraphSegments(
  pageSegments: DocumentPageNodeSegment[]
): boolean {
  return (
    pageSegments.length > 0 &&
    pageSegments.every(
      (segment) =>
        !segment.tableRowRange &&
        paragraphSegmentHasPartialLineRange(segment.paragraphLineRange)
    )
  );
}

function estimateDocumentNoteSectionHeightPx(
  notes: DocumentNoteDefinition[],
  pageContentWidthPx: number,
  numberingDefinitions?: NumberingDefinitionSet
): number {
  if (notes.length === 0) {
    return 0;
  }

  let totalHeightPx = 19;
  notes.forEach((note, noteIndex) => {
    if (noteIndex > 0) {
      totalHeightPx += 6;
    }

    const noteParagraphs =
      note.nodes?.filter(
        (node): node is ParagraphNode => node.type === "paragraph"
      ) ?? [];
    if (noteParagraphs.length === 0) {
      totalHeightPx += Math.round(12 * 1.35);
      return;
    }

    noteParagraphs.forEach((paragraph, paragraphIndex) => {
      if (paragraphIndex > 0) {
        totalHeightPx += 4;
      }
      totalHeightPx += estimateParagraphHeightPx(
        paragraph,
        pageContentWidthPx,
        numberingDefinitions
      );
    });
  });

  return Math.max(0, totalHeightPx);
}

function collectReferencedFootnotesForPageSegments(
  model: DocModel,
  pageSegments: DocumentPageNodeSegment[],
  footnotesById: Map<number, DocumentNoteDefinition>
): DocumentNoteDefinition[] {
  const referencedIds: number[] = [];
  const seen = new Set<number>();

  pageSegments.forEach((segment) => {
    const node = model.nodes[segment.nodeIndex];
    if (!node) {
      return;
    }

    nodeReferencedNoteIds(
      node,
      "footnote",
      segment.tableRowRange,
      segment.paragraphLineRange
    ).forEach((referenceId) => {
      if (seen.has(referenceId)) {
        return;
      }
      seen.add(referenceId);
      referencedIds.push(referenceId);
    });
  });

  return referencedIds
    .map((referenceId) => footnotesById.get(referenceId))
    .filter((note): note is DocumentNoteDefinition => Boolean(note));
}

function applyEstimatedFootnoteReserveToPages(
  model: DocModel,
  pages: DocumentPageNodeSegment[][],
  defaultPageContentHeightPx: number,
  pageContentWidthPx: number,
  metricsBySection: PaginationSectionMetrics[],
  numberingDefinitions: NumberingDefinitionSet | undefined,
  footnotes: DocumentNoteDefinition[],
  buildPages: (
    pageContentHeightsOverride?: number[]
  ) => DocumentPageNodeSegment[][]
): DocumentPageNodeSegment[][] {
  if (pages.length === 0 || footnotes.length === 0) {
    return pages;
  }

  const footnotesById = new Map(footnotes.map((note) => [note.id, note]));
  let currentPages = pages;

  for (let iteration = 0; iteration < 3; iteration += 1) {
    const reserveOverrides = currentPages.map((pageSegments, pageIndex) => {
      const pageFootnotes = collectReferencedFootnotesForPageSegments(
        model,
        pageSegments,
        footnotesById
      );
      const baseHeightPx = resolvePageContentHeightPxForPageSegments(
        pageSegments,
        pageIndex,
        defaultPageContentHeightPx,
        metricsBySection
      );
      if (pageFootnotes.length === 0) {
        return baseHeightPx;
      }

      const reservePx = estimateDocumentNoteSectionHeightPx(
        pageFootnotes,
        pageContentWidthPx,
        numberingDefinitions
      );
      return Math.max(120, baseHeightPx - reservePx);
    });

    const nextPages = buildPages(reserveOverrides);
    const stable =
      nextPages.length === currentPages.length &&
      nextPages.every((pageSegments, pageIndex) => {
        const previousSegments = currentPages[pageIndex] ?? [];
        if (pageSegments.length !== previousSegments.length) {
          return false;
        }
        return pageSegments.every((segment, segmentIndex) => {
          const previous = previousSegments[segmentIndex];
          return (
            previous?.nodeIndex === segment.nodeIndex &&
            previous?.tableRowRange?.startRowIndex ===
              segment.tableRowRange?.startRowIndex &&
            previous?.tableRowRange?.endRowIndex ===
              segment.tableRowRange?.endRowIndex &&
            previous?.paragraphLineRange?.startLineIndex ===
              segment.paragraphLineRange?.startLineIndex &&
            previous?.paragraphLineRange?.endLineIndex ===
              segment.paragraphLineRange?.endLineIndex
          );
        });
      });

    currentPages = nextPages;
    if (stable) {
      break;
    }
  }

  return currentPages;
}

export type DocxListType = "unordered" | "ordered";

interface DocxTableCellLocation {
  tableIndex: number;
  rowIndex: number;
  cellIndex: number;
}

interface DocxTableCellSelectionRange {
  tableIndex: number;
  anchorRowIndex: number;
  anchorCellIndex: number;
  focusRowIndex: number;
  focusCellIndex: number;
}

function parseDocumentIndexFromAttribute(
  raw: string | null
): number | undefined {
  const parsed = raw ? Number.parseInt(raw, 10) : Number.NaN;
  if (!Number.isFinite(parsed)) {
    return undefined;
  }

  return parsed;
}

function parseParagraphLocationFromElement(
  element: Element | null
): ParagraphLocation | undefined {
  if (!element) {
    return undefined;
  }

  const kind = element.getAttribute("data-docx-paragraph-kind");
  if (kind === "paragraph") {
    const nodeIndex = parseDocumentIndexFromAttribute(
      element.getAttribute("data-docx-paragraph-node-index")
    );
    if (nodeIndex === undefined) {
      return undefined;
    }
    return { kind: "paragraph", nodeIndex };
  }

  if (kind === "table-cell") {
    const tableIndex = parseDocumentIndexFromAttribute(
      element.getAttribute("data-docx-table-index")
    );
    const rowIndex = parseDocumentIndexFromAttribute(
      element.getAttribute("data-docx-row-index")
    );
    const cellIndex = parseDocumentIndexFromAttribute(
      element.getAttribute("data-docx-cell-index")
    );
    const paragraphIndex = parseDocumentIndexFromAttribute(
      element.getAttribute("data-docx-paragraph-index")
    );

    if (
      tableIndex === undefined ||
      rowIndex === undefined ||
      cellIndex === undefined ||
      paragraphIndex === undefined
    ) {
      return undefined;
    }

    return {
      kind: "table-cell",
      tableIndex,
      rowIndex,
      cellIndex,
      paragraphIndex,
    };
  }

  return undefined;
}

function parseTableCellLocationFromElement(
  element: Element | null
): DocxTableCellLocation | undefined {
  if (!element) {
    return undefined;
  }

  const cellElement = element.closest("[data-docx-table-cell='true']");
  if (!cellElement) {
    return undefined;
  }

  const tableIndex = parseDocumentIndexFromAttribute(
    cellElement.getAttribute("data-docx-table-index")
  );
  const rowIndex = parseDocumentIndexFromAttribute(
    cellElement.getAttribute("data-docx-row-index")
  );
  const cellIndex = parseDocumentIndexFromAttribute(
    cellElement.getAttribute("data-docx-cell-index")
  );
  if (
    tableIndex === undefined ||
    rowIndex === undefined ||
    cellIndex === undefined
  ) {
    return undefined;
  }

  return {
    tableIndex,
    rowIndex,
    cellIndex,
  };
}

function tableCellSelectionRangeBounds(range: DocxTableCellSelectionRange): {
  startRowIndex: number;
  endRowIndex: number;
  startCellIndex: number;
  endCellIndex: number;
} {
  return {
    startRowIndex: Math.min(range.anchorRowIndex, range.focusRowIndex),
    endRowIndex: Math.max(range.anchorRowIndex, range.focusRowIndex),
    startCellIndex: Math.min(range.anchorCellIndex, range.focusCellIndex),
    endCellIndex: Math.max(range.anchorCellIndex, range.focusCellIndex),
  };
}

function isSingleTableCellSelectionRange(
  range: DocxTableCellSelectionRange
): boolean {
  return (
    range.anchorRowIndex === range.focusRowIndex &&
    range.anchorCellIndex === range.focusCellIndex
  );
}

function sameTableCellSelectionRange(
  a: DocxTableCellSelectionRange | undefined,
  b: DocxTableCellSelectionRange | undefined
): boolean {
  if (!a || !b) {
    return a === b;
  }

  return (
    a.tableIndex === b.tableIndex &&
    a.anchorRowIndex === b.anchorRowIndex &&
    a.anchorCellIndex === b.anchorCellIndex &&
    a.focusRowIndex === b.focusRowIndex &&
    a.focusCellIndex === b.focusCellIndex
  );
}

function isCellWithinTableSelectionRange(
  range: DocxTableCellSelectionRange | undefined,
  tableIndex: number,
  rowIndex: number,
  cellIndex: number
): boolean {
  if (!range || range.tableIndex !== tableIndex) {
    return false;
  }

  const bounds = tableCellSelectionRangeBounds(range);
  return (
    rowIndex >= bounds.startRowIndex &&
    rowIndex <= bounds.endRowIndex &&
    cellIndex >= bounds.startCellIndex &&
    cellIndex <= bounds.endCellIndex
  );
}

function textRangeTouchesTable(
  range: DocxTextRange | undefined,
  tableIndex: number
): boolean {
  if (!range) {
    return false;
  }

  return [range.start.location, range.end.location].some(
    (location) =>
      location.kind === "table-cell" && location.tableIndex === tableIndex
  );
}

function selectedTableCellLocations(
  table: TableNode,
  range: DocxTableCellSelectionRange
): Array<{ rowIndex: number; cellIndex: number }> {
  const bounds = tableCellSelectionRangeBounds(range);
  const selected: Array<{ rowIndex: number; cellIndex: number }> = [];

  for (let rowIndex = 0; rowIndex < table.rows.length; rowIndex += 1) {
    if (rowIndex < bounds.startRowIndex || rowIndex > bounds.endRowIndex) {
      continue;
    }

    const row = table.rows[rowIndex];
    if (!row) {
      continue;
    }

    for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex += 1) {
      if (
        cellIndex < bounds.startCellIndex ||
        cellIndex > bounds.endCellIndex
      ) {
        continue;
      }

      const cell = row.cells[cellIndex];
      if (!cell || cell.style?.vMergeContinuation) {
        continue;
      }

      selected.push({ rowIndex, cellIndex });
    }
  }

  return selected;
}

function tableSelectableCellExtents(table: TableNode):
  | {
      first: { rowIndex: number; cellIndex: number };
      last: { rowIndex: number; cellIndex: number };
    }
  | undefined {
  let first: { rowIndex: number; cellIndex: number } | undefined;
  let maxSelectableRowIndex = -1;
  let maxSelectableCellIndex = -1;

  for (let rowIndex = 0; rowIndex < table.rows.length; rowIndex += 1) {
    const row = table.rows[rowIndex];
    if (!row) {
      continue;
    }

    for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex += 1) {
      const cell = row.cells[cellIndex];
      if (!cell || cell.style?.vMergeContinuation) {
        continue;
      }

      if (!first) {
        first = { rowIndex, cellIndex };
      }
      if (rowIndex > maxSelectableRowIndex) {
        maxSelectableRowIndex = rowIndex;
      }
      if (cellIndex > maxSelectableCellIndex) {
        maxSelectableCellIndex = cellIndex;
      }
    }
  }

  if (!first || maxSelectableRowIndex < 0 || maxSelectableCellIndex < 0) {
    return undefined;
  }

  return {
    first,
    last: {
      rowIndex: maxSelectableRowIndex,
      cellIndex: maxSelectableCellIndex,
    },
  };
}

function tableSelectionCoversWholeTable(
  table: TableNode,
  range: DocxTableCellSelectionRange
): boolean {
  const selectedKeys = new Set(
    selectedTableCellLocations(table, range).map(
      (location) => `${location.rowIndex}:${location.cellIndex}`
    )
  );
  if (selectedKeys.size === 0) {
    return false;
  }

  let totalSelectableCells = 0;
  for (let rowIndex = 0; rowIndex < table.rows.length; rowIndex += 1) {
    const row = table.rows[rowIndex];
    if (!row) {
      continue;
    }

    for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex += 1) {
      const cell = row.cells[cellIndex];
      if (!cell || cell.style?.vMergeContinuation) {
        continue;
      }

      totalSelectableCells += 1;
      if (!selectedKeys.has(`${rowIndex}:${cellIndex}`)) {
        return false;
      }
    }
  }

  return totalSelectableCells > 0 && selectedKeys.size === totalSelectableCells;
}

function textLengthFromRange(range: Range): number {
  const fragment = range.cloneContents();
  fragment
    .querySelectorAll("[data-docx-numbering-label='true']")
    .forEach((label) => {
      label.remove();
    });
  return fragment.textContent?.length ?? 0;
}

function compareParagraphLocations(
  a: ParagraphLocation,
  b: ParagraphLocation
): number {
  if (a.kind === "paragraph" && b.kind === "paragraph") {
    return Math.sign(a.nodeIndex - b.nodeIndex);
  }

  if (a.kind === "paragraph" && b.kind === "table-cell") {
    if (a.nodeIndex !== b.tableIndex) {
      return Math.sign(a.nodeIndex - b.tableIndex);
    }

    return -1;
  }

  if (a.kind === "table-cell" && b.kind === "paragraph") {
    if (a.tableIndex !== b.nodeIndex) {
      return Math.sign(a.tableIndex - b.nodeIndex);
    }

    return 1;
  }

  if (a.kind === "table-cell" && b.kind === "table-cell") {
    if (a.tableIndex !== b.tableIndex) {
      return Math.sign(a.tableIndex - b.tableIndex);
    }

    if (a.rowIndex !== b.rowIndex) {
      return Math.sign(a.rowIndex - b.rowIndex);
    }

    if (a.cellIndex !== b.cellIndex) {
      return Math.sign(a.cellIndex - b.cellIndex);
    }

    return Math.sign(a.paragraphIndex - b.paragraphIndex);
  }

  return 0;
}

function compareTextRangeBoundaries(
  a: DocxTextRangeBoundary,
  b: DocxTextRangeBoundary
): number {
  const locationCompare = compareParagraphLocations(a.location, b.location);
  if (locationCompare !== 0) {
    return locationCompare;
  }

  return Math.sign(a.offset - b.offset);
}

function normalizeTextRange(range: DocxTextRange): DocxTextRange {
  if (compareTextRangeBoundaries(range.start, range.end) <= 0) {
    return {
      start: range.start,
      end: range.end,
    };
  }

  return {
    start: range.end,
    end: range.start,
  };
}

interface DocxHistorySnapshot {
  model: DocModel;
  selection: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
}

interface DocxHistoryRestoreRequest {
  nonce: number;
  selection: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
}

interface DocxEditorTransactionContext {
  model: DocModel;
  selection: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
  pendingRunStyle?: TextRunNode["style"];
}

interface DocxEditorTransactionPatch {
  model?: DocModel;
  selection?: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
  pendingRunStyle?: TextRunNode["style"];
  status?: string;
  clearSelectedFormField?: boolean;
  pushHistory?: boolean;
}

type ParagraphLocation = DocxTextRangeLocation;

export type DocxImageLocation = ParagraphLocation & {
  childIndex: number;
};

export type DocxSectionRegion = "header" | "footer";

export interface DocxSectionParagraphLocation {
  region: DocxSectionRegion;
  partName: string;
  nodeIndex: number;
  rowIndex?: number;
  cellIndex?: number;
  paragraphIndex?: number;
}

export type DocxSectionImageLocation = DocxSectionParagraphLocation & {
  childIndex: number;
};

export type DocxFormFieldLocation = ParagraphLocation & {
  childIndex: number;
};

export interface DocxSelectedFormField {
  location: DocxFormFieldLocation;
  field: FormFieldRunNode;
}

export type DocxImageDropTarget = ParagraphLocation & {
  childIndex: number;
};

export type DocxTrackedChangeKind =
  | "insertion"
  | "deletion"
  | "move-from"
  | "move-to"
  | "format-change"
  | "paragraph-format-change";

export interface DocxTrackedChange {
  id: string;
  inlineAnchorId?: string;
  kind: DocxTrackedChangeKind;
  author?: string;
  date?: string;
  text?: string;
  nodeIndex: number;
  location: DocxTextRangeLocation;
}

/**
 * A document comment anchored to a text range.
 *
 * Comment definitions come from `word/comments.xml` (threading and resolution
 * state from `word/commentsExtended.xml`); the anchor location is resolved
 * from the paragraph that carries the `commentReference` run.
 */
export interface DocxComment {
  /** Stable identifier (unique per rendered anchor). */
  id: string;
  /** Comment id from `word/comments.xml` (`w:id`). */
  commentId: number;
  author?: string;
  initials?: string;
  date?: string;
  /** Plain-text comment body. */
  text: string;
  /** Comment id this comment replies to, when part of a thread. */
  parentId?: number;
  /** True when the comment thread is marked done. */
  resolved?: boolean;
  /** Plain-text excerpt of the commented document range. */
  anchorText?: string;
  nodeIndex: number;
  location: DocxTextRangeLocation;
}

export type DocxLineSpacingRule = "auto" | "exact" | "atLeast";
export type DocxSelectionSessionKind =
  | "idle"
  | "pointer"
  | "keyboard"
  | "composition"
  | "command"
  | "history-restore";

export interface DocxLineSpacingInfo {
  lineRule: DocxLineSpacingRule;
  lineTwips?: number;
  multiple: number;
}

export type DocxBorderContext = "paragraph" | "table";

export type DocxBorderPreset =
  | "bottom"
  | "top"
  | "left"
  | "right"
  | "none"
  | "all"
  | "outside"
  | "inside"
  | "inside-horizontal"
  | "inside-vertical"
  | "diagonal-down"
  | "diagonal-up"
  | "horizontal-line";

export type DocxBorderPresetState = Record<DocxBorderPreset, boolean>;

/**
 * Initial state for `useDocxEditor`.
 *
 * These values are read when the editor controller is created. Later document
 * imports update the controller through `editor.importDocxFile`.
 */
export interface UseDocxEditorOptions {
  /**
   * Model used for `editor.newDocument()` and for the initial empty editor.
   *
   * @defaultValue `defaultStarterModel`
   */
  starterModel?: DocModel;
  /**
   * Initial display name before a user imports a file.
   *
   * @defaultValue `"(new document)"`
   */
  initialFileName?: string;
  /**
   * Initial status text exposed on `editor.status`.
   *
   * @defaultValue `"Ready"`
   */
  initialStatus?: string;
  /**
   * Initial document surface theme.
   *
   * @defaultValue `"light"`
   */
  initialDocumentTheme?: DocxDocumentTheme;
  /**
   * Whether tracked changes are visible when the editor first mounts.
   *
   * @defaultValue `false`
   */
  initialShowTrackedChanges?: boolean;
  /**
   * Whether document comments are visible when the editor first mounts.
   *
   * @defaultValue `false`
   */
  initialShowComments?: boolean;
}

/**
 * Controller returned by `useDocxEditor`.
 *
 * Pass this object to `DocxEditorViewer`, then wire toolbar buttons and custom
 * controls to the methods exposed here.
 *
 * @example
 * ```tsx
 * const editor = useDocxEditor();
 *
 * return (
 *   <>
 *     <button onClick={() => editor.toggleBold()}>Bold</button>
 *     <DocxEditorViewer editor={editor} />
 *   </>
 * );
 * ```
 */
export interface DocxEditorController {
  model: DocModel;
  documentLoadNonce: number;
  fileName: string;
  status: string;
  importError?: Error;
  isImporting: boolean;
  documentTheme: DocxDocumentTheme;
  selection: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
  historyRestoreRequest?: DocxHistoryRestoreRequest;
  selectedFormField?: DocxSelectedFormField;
  selectedParagraph?: ParagraphNode;
  selectedRunStyle?: TextRunNode["style"];
  selectedLink?: string;
  pendingRunStyle?: TextRunNode["style"];
  selectionSessionKind: DocxSelectionSessionKind;
  suppressNextDomSelectionRestore: () => void;
  beginSelectionSession: (
    kind: Exclude<DocxSelectionSessionKind, "idle">,
    options?: {
      settleAfterMs?: number;
    }
  ) => void;
  clearSelectionSession: (expectedKind?: DocxSelectionSessionKind) => void;
  selectedParagraphStyleId?: string;
  selectedLineSpacing: DocxLineSpacingInfo;
  selectedBorderContext: DocxBorderContext;
  activeBorderPresets: DocxBorderPresetState;
  availableParagraphStyles: ParagraphStyleDefinition[];
  trackedChanges: DocxTrackedChange[];
  showTrackedChanges: boolean;
  comments: DocxComment[];
  showComments: boolean;
  currentPage: number;
  totalPages: number;
  hasUnorderedList: boolean;
  hasOrderedList: boolean;
  canUndo: boolean;
  canRedo: boolean;
  registerPendingExportModelTransformer: (
    transformer?: (model: DocModel) => DocModel
  ) => void;
  setStatus: (value: string | ((prev: string) => string)) => void;
  setDocumentTheme: (theme: DocxDocumentTheme) => void;
  setShowTrackedChanges: (showTrackedChanges: boolean) => void;
  setShowComments: (showComments: boolean) => void;
  syncPaginationInfo: (pagination: DocxPaginationInfo) => void;
  toggleShowTrackedChanges: () => void;
  toggleShowComments: () => void;
  importDocxFile: (file: File) => Promise<void>;
  newDocument: () => void;
  exportDocx: () => void;
  undo: () => void;
  redo: () => void;
  setHeading: (heading?: HeadingLevel) => void;
  setParagraphStyle: (styleId?: string) => void;
  setLineSpacing: (lineMultiple: number) => void;
  setFontFamily: (fontFamily: string) => void;
  setFontSize: (fontSizePt: number) => void;
  toggleBold: () => void;
  toggleItalic: () => void;
  toggleUnderline: () => void;
  toggleStrike: () => void;
  toggleSuperscript: () => void;
  toggleSubscript: () => void;
  setHighlight: (highlight?: string) => void;
  setTextColor: (color?: string) => void;
  setLink: (link?: string) => void;
  selectFormField: (location?: DocxFormFieldLocation) => void;
  toggleFormCheckbox: (location: DocxFormFieldLocation) => void;
  setFormFieldValue: (location: DocxFormFieldLocation, value: string) => void;
  updateFormFieldWidget: (
    location: DocxFormFieldLocation,
    patch: Partial<NonNullable<FormFieldRunNode["widget"]>>
  ) => void;
  applyBorderPreset: (preset: DocxBorderPreset) => void;
  setAlignment: (align?: ParagraphAlignment) => void;
  toggleList: (listType: DocxListType) => void;
  adjustSelectedListDepth: (levelDelta: number, draftText?: string) => boolean;
  insertListItemAfterSelection: (
    draftText: string,
    startOffset: number,
    endOffset?: number,
    targetLocation?: ParagraphLocation
  ) =>
    | {
        paragraphIndex: number;
        caretOffset: number;
      }
    | undefined;
  splitParagraphAtSelection: (
    draftText: string,
    startOffset: number,
    endOffset?: number,
    targetLocation?: ParagraphLocation
  ) =>
    | {
        paragraphIndex: number;
        caretOffset: number;
      }
    | undefined;
  insertTable: () => void;
  insertImageFile: (file: File) => Promise<void>;
  appendParagraph: (text?: string) => number;
  resizeImage: (
    location: DocxImageLocation,
    widthPx: number,
    heightPx: number
  ) => void;
  setSyntheticTextBoxText: (location: DocxImageLocation, text: string) => void;
  setImageWrapMode: (
    location: DocxImageLocation,
    mode: DocxImageWrapMode,
    seedFloating?: Partial<NonNullable<ImageRunNode["floating"]>>
  ) => void;
  moveFloatingImage: (
    location: DocxImageLocation,
    patch: Partial<NonNullable<ImageRunNode["floating"]>>
  ) => void;
  moveSectionFloatingImage: (
    location: DocxSectionImageLocation,
    patch: Partial<NonNullable<ImageRunNode["floating"]>>
  ) => void;
  moveParagraphDropCap: (
    nodeIndex: number,
    patch: Partial<NonNullable<NonNullable<ParagraphNode["style"]>["dropCap"]>>
  ) => void;
  setParagraphDropCapFontSizePt: (
    nodeIndex: number,
    fontSizePt: number
  ) => void;
  setParagraphDropCapText: (nodeIndex: number, text: string) => void;
  moveImage: (source: DocxImageLocation, target: DocxImageDropTarget) => void;
  setActiveTextRange: (range?: DocxTextRange) => void;
  selectParagraph: (nodeIndex: number) => void;
  selectTableCell: (
    tableIndex: number,
    rowIndex: number,
    cellIndex: number
  ) => void;
  clearTableCellContents: (
    tableIndex: number,
    cells: Array<{ rowIndex: number; cellIndex: number }>
  ) => void;
  insertTableRow: (
    tableIndex: number,
    rowIndex: number,
    direction: "above" | "below"
  ) => void;
  insertTableColumn: (
    tableIndex: number,
    cellIndex: number,
    direction: "left" | "right",
    rowIndex?: number
  ) => void;
  deleteTableRow: (tableIndex: number, rowIndex: number) => void;
  deleteTableColumn: (
    tableIndex: number,
    cellIndex: number,
    rowIndex?: number
  ) => void;
  deleteTable: (tableIndex: number) => void;
  moveTable: (tableIndex: number, targetNodeIndex: number) => void;
  moveEmbeddedTableToBody: (
    tableRuntimeKey: string,
    targetNodeIndex: number
  ) => void;
  replaceExpandedSelection: (
    text: string,
    range?: DocxTextRange
  ) => DocxTextRange | undefined;
  deleteExpandedSelection: (range?: DocxTextRange) => DocxTextRange | undefined;
  commitParagraphText: (nodeIndex: number, text: string) => void;
  commitTableCellText: (
    tableIndex: number,
    rowIndex: number,
    cellIndex: number,
    text: string
  ) => void;
  commitTableCellParagraphTextRecursive: (
    tableIndex: number,
    rowIndex: number,
    cellIndex: number,
    paragraphIndex: number,
    text: string
  ) => void;
  commitSectionParagraphText: (
    location: DocxSectionParagraphLocation,
    text: string
  ) => void;
}

export type DocxTableContextMenuActionId =
  | "insert-row-above"
  | "insert-row-below"
  | "insert-column-left"
  | "insert-column-right"
  | "delete-row"
  | "delete-column"
  | "delete-table";

export type DocxContextMenuActionId =
  | DocxTableContextMenuActionId
  | "cut"
  | "copy"
  | "paste"
  | "image-wrap-inline"
  | "image-wrap-square"
  | "image-wrap-tight"
  | "image-wrap-through"
  | "image-wrap-top-and-bottom"
  | "image-wrap-behind-text"
  | "image-wrap-in-front-of-text"
  | "image-edit-wrap-boundary"
  | "image-move-with-text"
  | "image-fix-position-on-page"
  | "image-more-layout-options"
  | "image-bring-to-front"
  | "image-bring-forward"
  | "image-in-front-of-text"
  | "image-send-to-back"
  | "image-send-backward"
  | "image-behind-text";

export interface DocxTableContextMenuContext {
  /** Zero-based table index in the document body. */
  tableIndex: number;
  /** Zero-based row index under the pointer. */
  rowIndex: number;
  /** Zero-based cell index under the pointer. */
  cellIndex: number;
}

export interface DocxTableContextMenuAction {
  /** Built-in table command id. */
  id: DocxTableContextMenuActionId;
  /** Human-readable menu label. */
  label: string;
  /** True when the action removes content or structure. */
  destructive?: boolean;
}

export interface DocxContextMenuAction {
  /** Built-in command id, or a custom id from your menu renderer. */
  id: DocxContextMenuActionId | (string & {});
  /** Human-readable menu label. */
  label: string;
  /** Optional keyboard shortcut text to display next to the label. */
  shortcut?: string;
  /** True when the action removes content or structure. */
  destructive?: boolean;
  /** Prevents the action from being selected. */
  disabled?: boolean;
  /** Marks toggle-style menu items as active. */
  checked?: boolean;
  /** Adds a visual separator before this item in default-style menus. */
  separatorBefore?: boolean;
  /** Nested submenu actions. */
  children?: DocxContextMenuAction[];
}

/**
 * Arguments passed to `renderTableContextMenu`.
 */
export interface DocxTableContextMenuRenderProps {
  /** Table cell that opened the menu. */
  context: DocxTableContextMenuContext;
  /** Built-in table actions available for this context. */
  actions: DocxTableContextMenuAction[];
  /** Runs a built-in table action. */
  runAction: (actionId: DocxTableContextMenuActionId) => void;
  /** Closes the menu without running an action. */
  closeMenu: () => void;
  /** Viewport coordinates where the menu should be placed. */
  position: {
    x: number;
    y: number;
  };
  /** Current document theme, useful for custom menu styling. */
  documentTheme: DocxDocumentTheme;
}

export interface DocxContextMenuContext {
  /** Type of target that opened the menu. */
  kind: "text" | "table" | "image";
  /** Active text range when the menu was opened from text. */
  activeTextRange?: DocxTextRange;
  /** Text location under the pointer when available. */
  location?: DocxTextRangeLocation;
  /** Table context when `kind` is `"table"`. */
  tableContext?: DocxTableContextMenuContext;
  /** Image context when `kind` is `"image"`. */
  image?:
    | {
        location: DocxImageLocation;
        floating?: NonNullable<ImageRunNode["floating"]>;
        wrap?: DocxImageWrapState;
      }
    | undefined;
}

/**
 * Arguments passed to `renderContextMenu`.
 */
export interface DocxContextMenuRenderProps {
  /** Target context that opened the menu. */
  context: DocxContextMenuContext;
  /** Built-in actions available for this context. */
  actions: DocxContextMenuAction[];
  /** Runs a built-in or custom action id. */
  runAction: (actionId: DocxContextMenuActionId | (string & {})) => void;
  /** Closes the menu without running an action. */
  closeMenu: () => void;
  /** Viewport coordinates where the menu should be placed. */
  position: {
    x: number;
    y: number;
  };
  /** Current document theme, useful for custom menu styling. */
  documentTheme: DocxDocumentTheme;
}

/**
 * Controls how `DocxEditorViewer` windows page DOM nodes.
 *
 * Virtualization is enabled by default for multi-page documents. It keeps
 * offscreen pages as lightweight placeholders, which is important for large
 * DOCX files with many tables or complex paragraphs.
 */
export interface DocxPageVirtualizationOptions {
  /**
   * Enables or disables internal page virtualization.
   *
   * Disable only when you need every page mounted at once, for example when
   * generating thumbnails for all pages from the live DOM.
   *
   * @defaultValue `true`
   */
  enabled?: boolean;
  /**
   * Number of pages to keep mounted before and after the visible viewport.
   *
   * Higher values reduce mount/unmount churn while scrolling but increase
   * initial DOM work.
   *
   * @defaultValue `2`
   */
  overscan?: number;
  /**
   * Compatibility option from earlier releases.
   *
   * Page virtualization now starts immediately, so this option is accepted but
   * does not delay the initial page window.
   *
   * @defaultValue `0`
   */
  settleDelayMs?: number;
  /**
   * Explicit scroll container for internal page virtualization.
   *
   * Omit this to let the viewer discover the nearest scrollable ancestor.
   * Provide it when the host app owns scrolling, such as when the viewer is
   * mounted inside a custom scroll-area viewport.
   */
  scrollElement?: HTMLElement | null;
  /**
   * Effective visual scale applied around the viewer.
   *
   * Omit this to let the viewer infer CSS `zoom` from its ancestor chain.
   * Provide it when toolbar zoom is controlled outside the viewer so virtual
   * page offsets update synchronously with the selected zoom.
   */
  zoomScale?: number;
}

/**
 * Externally controlled visible page window.
 *
 * Provide this when your app owns scrolling or virtualization and wants
 * `DocxEditorViewer` to mount only a specific page range.
 */
export interface DocxVisiblePageRange {
  /**
   * Zero-based first page index to render.
   *
   * @example
   * ```tsx
   * <DocxEditorViewer
   *   editor={editor}
   *   visiblePageRange={{ startPageIndex: 0, endPageIndex: 2 }}
   * />
   * ```
   */
  startPageIndex: number;
  /** Zero-based last page index to render. */
  endPageIndex: number;
}

/**
 * Props for the full DOCX editor/viewer surface.
 *
 * Use this component with a controller from `useDocxEditor`.
 *
 * @example
 * ```tsx
 * const editor = useDocxEditor();
 *
 * return (
 *   <DocxEditorViewer
 *     editor={editor}
 *     mode="read-only"
 *     pageVirtualization={{ overscan: 3 }}
 *   />
 * );
 * ```
 */
export interface DocxEditorViewerProps {
  /**
   * Editor controller returned by `useDocxEditor`.
   */
  editor: DocxEditorController;
  /**
   * CSS class applied to the outer viewer root.
   */
  className?: string;
  /**
   * Inline styles applied to the outer viewer root.
   */
  style?: Record<string, string | number | undefined>;
  /**
   * Background color of each rendered page surface.
   *
   * @defaultValue Viewer theme page color.
   */
  pageBackgroundColor?: string;
  /**
   * Background color shown between pages.
   *
   * @defaultValue `"transparent"`
   */
  pageGapBackgroundColor?: string;
  /**
   * Hides the document behind `loadingState` until initial pagination settles.
   *
   * Enable only when you prefer stable first paint over immediate approximate
   * layout.
   *
   * @defaultValue `false`
   */
  deferInitialPaginationPaint?: boolean;
  /**
   * Custom content shown while initial pagination is settling.
   *
   * Used only when `deferInitialPaginationPaint` is true.
   *
   * @defaultValue A compact `"Loading..."` pill.
   */
  loadingState?: unknown;
  /**
   * Configures internal page virtualization.
   *
   * @defaultValue `{ enabled: true, overscan: 2 }`
   */
  pageVirtualization?: DocxPageVirtualizationOptions;
  /**
   * Controlled page range to mount.
   *
   * When provided, internal page virtualization is bypassed and only this
   * range is rendered.
   */
  visiblePageRange?: DocxVisiblePageRange;
  /**
   * Called whenever the viewer's resolved page count changes.
   */
  onPageCountChange?: (pageCount: number) => void;
  /**
   * Called when the viewer needs a page to become visible, such as after a
   * bookmark or cross-reference navigation.
   *
   * If omitted, the internal virtualizer scrolls the page into view.
   */
  onRequestPageReveal?: (pageIndex: number) => void;
  /**
   * Overrides visual styles for heading levels.
   */
  headingStyles?: DocxHeadingStyleMap;
  /**
   * Interaction mode for the viewer.
   *
   * @defaultValue `"edit"`
   */
  mode?: DocxEditorViewerMode;
  /**
   * Overrides whether tracked changes are shown.
   *
   * If omitted, the value from `useDocxTrackChanges(editor)` or
   * `editor.showTrackedChanges` is used.
   */
  showTrackedChanges?: boolean;
  /**
   * Custom renderer for tracked-change cards in the page gutter.
   */
  renderTrackedChangeCard?: (
    props: DocxTrackedChangeCardRenderProps
  ) => unknown;
  /**
   * Overrides whether document comments are shown.
   *
   * If omitted, the value from `useDocxComments(editor)` or
   * `editor.showComments` is used.
   */
  showComments?: boolean;
  /**
   * Custom renderer for comment cards in the page gutter.
   */
  renderCommentCard?: (props: DocxCommentCardRenderProps) => unknown;
  /**
   * Custom renderer for table context menus.
   *
   * Call `props.runAction(action.id)` to execute built-in actions.
   */
  renderTableContextMenu?: (
    props: DocxTableContextMenuRenderProps
  ) => unknown;
  /**
   * Custom renderer for text, image, and table context menus.
   *
   * Call `props.closeMenu()` after handling a custom action.
   */
  renderContextMenu?: (props: DocxContextMenuRenderProps) => unknown;
  /**
   * Called when a form field is double-clicked.
   *
   * Use this to open a custom field settings panel.
   */
  onFormFieldDoubleClick?: (location: DocxFormFieldLocation) => void;
}

export interface DocxTrackedChangeCardRenderProps {
  /** Tracked-change data represented by the card. */
  change: DocxTrackedChange;
  /** Short display label for the change kind. */
  kindLabel: string;
  /** Plain-text excerpt for the changed content. */
  snippet: string;
  /** Formatted change date, if the source document provided one. */
  formattedDate?: string;
  /** Accent color chosen for this change kind. */
  accentColor: string;
  /** Current document theme. */
  documentTheme: DocxDocumentTheme;
  /** Zero-based page index that owns the card. */
  pageIndex: number;
  /** Positioning style computed by the viewer. Apply this to the card root. */
  style: Record<string, string | number | undefined>;
}

export interface DocxCommentCardRenderProps {
  /** Comment data represented by the card. */
  comment: DocxComment;
  /** Plain-text comment body (already normalized for display). */
  snippet: string;
  /** Formatted comment date, if the source document provided one. */
  formattedDate?: string;
  /** Accent color chosen for comments. */
  accentColor: string;
  /** Current document theme. */
  documentTheme: DocxDocumentTheme;
  /** Zero-based page index that owns the card. */
  pageIndex: number;
  /** Positioning style computed by the viewer. Apply this to the card root. */
  style: Record<string, string | number | undefined>;
}

/**
 * Viewer interaction mode.
 *
 * - `"edit"` allows editing when the document state permits it.
 * - `"read-only"` disables editing affordances while preserving navigation.
 */
export type DocxEditorViewerMode = "edit" | "read-only";

export interface UseDocxDocumentThemeResult {
  documentTheme: DocxDocumentTheme;
  isDarkDocument: boolean;
  setDocumentTheme: (theme: DocxDocumentTheme) => void;
  toggleDocumentTheme: () => void;
}

export interface UseDocxParagraphStylesResult {
  paragraphStyles: ParagraphStyleDefinition[];
  selectedParagraphStyleId?: string;
  setParagraphStyle: (styleId?: string) => void;
}

export interface UseDocxLineSpacingResult {
  lineSpacing: DocxLineSpacingInfo;
  setLineSpacing: (lineMultiple: number) => void;
}

export interface UseDocxBordersResult {
  borderContext: DocxBorderContext;
  activeBorderPresets: DocxBorderPresetState;
  applyBorderPreset: (preset: DocxBorderPreset) => void;
}

export interface UseDocxFormFieldsResult {
  formFields: DocxSelectedFormField[];
  selectedFormField?: DocxSelectedFormField;
  selectFormField: (location?: DocxFormFieldLocation) => void;
  setFormFieldValue: (location: DocxFormFieldLocation, value: string) => void;
  toggleFormCheckbox: (location: DocxFormFieldLocation) => void;
  updateFormFieldWidget: (
    location: DocxFormFieldLocation,
    patch: Partial<NonNullable<FormFieldRunNode["widget"]>>
  ) => void;
  updateSelectedFormFieldWidget: (
    patch: Partial<NonNullable<FormFieldRunNode["widget"]>>
  ) => void;
}

export interface UseDocxTrackChangesResult {
  trackedChanges: DocxTrackedChange[];
  showTrackedChanges: boolean;
  setShowTrackedChanges: (showTrackedChanges: boolean) => void;
  toggleShowTrackedChanges: () => void;
  changesByLocation: Map<string, DocxTrackedChange[]>;
  getChangesForLocation: (
    location: DocxTextRangeLocation
  ) => DocxTrackedChange[];
}

export interface UseDocxCommentsResult {
  comments: DocxComment[];
  showComments: boolean;
  setShowComments: (showComments: boolean) => void;
  toggleShowComments: () => void;
  commentsByLocation: Map<string, DocxComment[]>;
  getCommentsForLocation: (location: DocxTextRangeLocation) => DocxComment[];
}

export interface DocxSectionColumnLayout {
  count: number;
  gapPx: number;
}

export interface DocxPageLayoutInfo {
  pageWidthPx: number;
  pageHeightPx: number;
  contentWidthPx: number;
  contentHeightPx: number;
  marginsPx: {
    top: number;
    right: number;
    bottom: number;
    left: number;
  };
  headerDistancePx: number;
  footerDistancePx: number;
  pageNumberStart: number;
  columns?: DocxSectionColumnLayout;
  viewportDefaults: {
    zoomPercent: number;
    pageGapPx: number;
  };
}

export interface UseDocxPageLayoutResult {
  layout: DocxPageLayoutInfo;
}

export interface DocxPaginationInfo {
  currentPage: number;
  totalPages: number;
}

export interface UseDocxPaginationResult {
  pagination: DocxPaginationInfo;
}

export interface DocxPageThumbnailResolutionOptions {
  /** Source page width in CSS pixels. */
  sourceWidthPx: number;
  /** Source page height in CSS pixels. */
  sourceHeightPx: number;
  /**
   * Desired thumbnail bounds.
   *
   * A number means "fit within this max width and max height". An object lets
   * you constrain width and height independently.
   *
   * @defaultValue `160`
   */
  resolution?: DocxPageThumbnailBounds;
  /**
   * Maximum CSS pixel width for the thumbnail.
   *
   * Overrides `resolution.maxWidth` when provided.
   */
  maxWidthPx?: number;
  /**
   * Maximum CSS pixel height for the thumbnail.
   *
   * Overrides `resolution.maxHeight` when provided.
   */
  maxHeightPx?: number;
  /**
   * Canvas backing-store pixel ratio.
   *
   * Increase for sharper thumbnails on high-density displays.
   *
   * @defaultValue `window.devicePixelRatio`, capped internally.
   */
  pixelRatio?: number;
}

/**
 * Thumbnail size constraint.
 *
 * @example
 * ```tsx
 * useDocxPageThumbnails(editor, { resolution: 180 });
 * useDocxPageThumbnails(editor, { resolution: { maxWidth: 120 } });
 * ```
 */
export type DocxPageThumbnailBounds =
  | number
  | {
      /** Maximum thumbnail CSS height. */
      maxHeight?: number;
      /** Maximum thumbnail CSS width. */
      maxWidth?: number;
    };

export interface DocxPageThumbnailResolution {
  /** Thumbnail CSS width. */
  widthPx: number;
  /** Thumbnail CSS height. */
  heightPx: number;
  /** Canvas backing-store width. */
  pixelWidthPx: number;
  /** Canvas backing-store height. */
  pixelHeightPx: number;
  /** Scale from source page pixels to thumbnail CSS pixels. */
  scale: number;
}

export interface DocxPageThumbnailRenderWindow {
  /**
   * Page indexes whose attached thumbnail canvases should render first.
   *
   * Use this for the thumbnails currently visible in a virtualized sidebar.
   */
  visiblePageIndexes?: readonly number[];
  /**
   * Page indexes to rasterize into the thumbnail surface cache after visible
   * thumbnails. Prefetched pages paint quickly once their canvases mount.
   */
  prefetchPageIndexes?: readonly number[];
}

/**
 * Options for `useDocxPageThumbnails`.
 */
export interface UseDocxPageThumbnailsOptions {
  /**
   * Desired thumbnail bounds.
   *
   * @defaultValue `160`
   */
  resolution?: DocxPageThumbnailBounds;
  /** Maximum thumbnail CSS width. */
  maxWidthPx?: number;
  /** Maximum thumbnail CSS height. */
  maxHeightPx?: number;
  /**
   * Canvas backing-store pixel ratio.
   *
   * @defaultValue `window.devicePixelRatio`, capped internally.
   */
  pixelRatio?: number;
  /**
   * Minimum interval between repeat raster jobs for the same thumbnail canvas.
   *
   * Lower this when the consumer already limits work to a small visible
   * thumbnail window.
   *
   * @defaultValue `200`
   */
  minRasterIntervalMs?: number;
  /**
   * Prioritizes thumbnails for consumer-owned virtualized thumbnail rails.
   */
  renderWindow?: DocxPageThumbnailRenderWindow;
  /**
   * Prevents thumbnail rendering while keeping stable item metadata.
   *
   * @defaultValue `false`
   */
  disabled?: boolean;
}

export type DocxPageThumbnailStatus =
  | "idle"
  | "rendering"
  | "ready"
  | "unavailable"
  | "error";

export interface DocxPageThumbnailItem extends DocxPageThumbnailResolution {
  /** Source page aspect ratio. */
  aspectRatio: number;
  /** Alias for `sourceHeightPx` for compatibility with other preview APIs. */
  contentHeight: number;
  /** Alias for `sourceWidthPx` for compatibility with other preview APIs. */
  contentWidth: number;
  /** Alias for `heightPx`. */
  height: number;
  /** Zero-based page index. */
  pageIndex: number;
  /** One-based page number for display. */
  pageNumber: number;
  /** Source page width in CSS pixels. */
  sourceWidthPx: number;
  /** Source page height in CSS pixels. */
  sourceHeightPx: number;
  /** True when the source page DOM is currently mounted. */
  isMounted: boolean;
  /** Current thumbnail render status. */
  status: DocxPageThumbnailStatus;
  /** Last thumbnail rendering error, if any. */
  error?: Error;
  /** Paints this thumbnail into a canvas. Returns false when unavailable. */
  paint: (canvas: HTMLCanvasElement | null) => boolean;
  /** Ref callback that keeps an attached canvas rendered as the page changes. */
  canvasRef: (canvas: HTMLCanvasElement | null) => void;
  /** Asynchronously renders this thumbnail into a canvas. */
  renderToCanvas: (canvas: HTMLCanvasElement) => Promise<void>;
  /** Alias for `widthPx`. */
  width: number;
}

export interface UseDocxPageThumbnailsResult {
  /** Paints the requested page thumbnail into a canvas. */
  paintThumbnail: (
    pageIndex: number,
    canvas: HTMLCanvasElement | null
  ) => boolean;
  /** Thumbnail metadata and paint helpers for each known page. */
  thumbnails: DocxPageThumbnailItem[];
  /** Re-renders every thumbnail canvas currently attached through `canvasRef`. */
  rerenderAttachedThumbnails: () => Promise<void>;
}

export type UseDocxViewerThumbnailsOptions = UseDocxPageThumbnailsOptions;
export type DocxViewerThumbnails = UseDocxPageThumbnailsResult;

export const defaultStarterModel: DocModel = {
  nodes: [
    {
      type: "paragraph",
      children: [{ type: "text", text: "" }],
    },
  ],
  metadata: {
    sourceParts: 1,
    warnings: [],
    headerSections: [],
    footerSections: [],
    paragraphStyles: [
      {
        id: "Normal",
        name: "Body",
        isDefault: true,
        // Word's blank-document default body typeface. Without this the body
        // default fell through to a heading font / Times New Roman and did not
        // match the "Calibri" shown in the toolbar.
        runStyle: { fontFamily: "Calibri", fontSizePt: 11 },
      },
      {
        id: "Heading1",
        name: "Heading 1",
        headingLevel: 1,
        isPrimary: true,
        runStyle: {
          fontFamily: "Calibri Light",
          fontSizePt: 16,
          bold: true,
          color: "#2f5496",
        },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        headingLevel: 2,
        isPrimary: true,
        runStyle: {
          fontFamily: "Calibri Light",
          fontSizePt: 13,
          bold: true,
          color: "#2f5496",
        },
      },
      {
        id: "Heading3",
        name: "Heading 3",
        headingLevel: 3,
        isPrimary: true,
        runStyle: {
          fontFamily: "Calibri",
          fontSizePt: 12,
          bold: true,
          color: "#1f3763",
        },
      },
      {
        id: "Heading4",
        name: "Heading 4",
        headingLevel: 4,
        runStyle: {
          fontFamily: "Calibri",
          fontSizePt: 11,
          bold: true,
          color: "#1f3763",
        },
      },
      {
        id: "Heading5",
        name: "Heading 5",
        headingLevel: 5,
        runStyle: {
          fontFamily: "Calibri",
          fontSizePt: 11,
          bold: true,
          color: "#1f3763",
        },
      },
      {
        id: "Heading6",
        name: "Heading 6",
        headingLevel: 6,
        runStyle: {
          fontFamily: "Calibri",
          fontSizePt: 11,
          bold: true,
          color: "#1f3763",
        },
      },
    ],
    defaultParagraphStyleId: "Normal",
  },
};

function createBlankDocumentModel(): DocModel {
  return cloneDocModel(defaultStarterModel);
}

function textRuns(paragraph: ParagraphNode): TextRunNode[] {
  return paragraph.children.filter(
    (child): child is TextRunNode => child.type === "text"
  );
}

function paragraphText(paragraph: ParagraphNode): string {
  return textRuns(paragraph)
    .map((run) => run.text)
    .join("");
}

function nodeTreeContainsExplicitFontFamily(
  nodes: DocModel["nodes"] | HeaderSection["nodes"] | FooterSection["nodes"]
): boolean {
  return nodes.some((node) => {
    if (node.type === "paragraph") {
      return node.children.some((child) => {
        if (child.type === "text" || child.type === "form-field") {
          return Boolean(child.style?.fontFamily?.trim());
        }
        return false;
      });
    }

    if (node.type === "table") {
      return node.rows.some((row) =>
        row.cells.some((cell) => nodeTreeContainsExplicitFontFamily(cell.nodes))
      );
    }

    return false;
  });
}

function firstExplicitFontFamilyInNodeTree(
  nodes: DocModel["nodes"] | HeaderSection["nodes"] | FooterSection["nodes"]
): string | undefined {
  for (const node of nodes) {
    if (node.type === "paragraph") {
      for (const child of node.children) {
        if (child.type !== "text" && child.type !== "form-field") {
          continue;
        }

        const fontFamily = child.style?.fontFamily?.trim();
        if (fontFamily) {
          return fontFamily;
        }
      }
      continue;
    }

    if (node.type === "table") {
      for (const row of node.rows) {
        for (const cell of row.cells) {
          const nestedFontFamily = firstExplicitFontFamilyInNodeTree(
            cell.nodes
          );
          if (nestedFontFamily) {
            return nestedFontFamily;
          }
        }
      }
    }
  }

  return undefined;
}

function resolveDocumentInheritedFontFamily(
  model: DocModel
): string | undefined {
  const paragraphStyles = model.metadata.paragraphStyles ?? [];
  const normalizedDefaultStyleId =
    model.metadata.defaultParagraphStyleId?.trim().toLowerCase() ?? "";
  const defaultParagraphStyle =
    paragraphStyles.find(
      (style) => style.id.trim().toLowerCase() === normalizedDefaultStyleId
    ) ??
    paragraphStyles.find((style) => style.isDefault) ??
    paragraphStyles.find((style) => style.id.trim().toLowerCase() === "normal");
  const defaultStyleFontFamily =
    defaultParagraphStyle?.runStyle?.fontFamily?.trim();
  if (defaultStyleFontFamily) {
    return cssFontFamily(defaultStyleFontFamily);
  }

  // Never adopt a heading style's font as the document body default. Heading
  // styles (e.g. "Calibri Light") are not the body typeface, and using one here
  // made freshly typed body text render in the wrong font until it committed.
  const paragraphStyleFontFamily = paragraphStyles.find(
    (style) =>
      style.headingLevel === undefined &&
      Boolean(style.runStyle?.fontFamily?.trim())
  )?.runStyle?.fontFamily;
  if (paragraphStyleFontFamily) {
    return cssFontFamily(paragraphStyleFontFamily);
  }

  const explicitBodyFontFamily = firstExplicitFontFamilyInNodeTree(model.nodes);
  if (explicitBodyFontFamily) {
    return cssFontFamily(explicitBodyFontFamily);
  }

  for (const section of model.metadata.headerSections ?? []) {
    const explicitHeaderFontFamily = firstExplicitFontFamilyInNodeTree(
      section.nodes
    );
    if (explicitHeaderFontFamily) {
      return cssFontFamily(explicitHeaderFontFamily);
    }
  }

  for (const section of model.metadata.footerSections ?? []) {
    const explicitFooterFontFamily = firstExplicitFontFamilyInNodeTree(
      section.nodes
    );
    if (explicitFooterFontFamily) {
      return cssFontFamily(explicitFooterFontFamily);
    }
  }

  return nodeTreeContainsExplicitFontFamily(model.nodes) ||
    (model.metadata.headerSections ?? []).some((section) =>
      nodeTreeContainsExplicitFontFamily(section.nodes)
    ) ||
    (model.metadata.footerSections ?? []).some((section) =>
      nodeTreeContainsExplicitFontFamily(section.nodes)
    )
    ? undefined
    : cssFontFamily(DEFAULT_UNSPECIFIED_DOCX_FONT_FAMILY);
}

function replaceTabLayoutMarkersWithTabText(root: HTMLElement): void {
  const centerLayouts = Array.from(
    root.querySelectorAll<HTMLElement>("[data-docx-tab-layout='center']")
  );
  centerLayouts.forEach((layout) => {
    const left =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='left']")
        ?.textContent ?? "";
    const center =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='center']")
        ?.textContent ?? "";
    layout.replaceWith(`${left}\t${center}`);
  });

  const centerRightLayouts = Array.from(
    root.querySelectorAll<HTMLElement>("[data-docx-tab-layout='center-right']")
  );
  centerRightLayouts.forEach((layout) => {
    const first =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='0']")
        ?.textContent ?? "";
    const second =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='1']")
        ?.textContent ?? "";
    const third =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='2']")
        ?.textContent ?? "";
    layout.replaceWith(`${first}\t${second}\t${third}`);
  });

  const leaderLayouts = Array.from(
    root.querySelectorAll<HTMLElement>("[data-docx-tab-layout='leader']")
  );
  leaderLayouts.forEach((layout) => {
    const left =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='left']")
        ?.textContent ?? "";
    const right =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='right']")
        ?.textContent ?? "";
    layout.replaceWith(`${left}\t${right}`);
  });

  const rightLayouts = Array.from(
    root.querySelectorAll<HTMLElement>("[data-docx-tab-layout='right']")
  );
  rightLayouts.forEach((layout) => {
    const left =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='left']")
        ?.textContent ?? "";
    const right =
      layout.querySelector<HTMLElement>("[data-docx-tab-zone='right']")
        ?.textContent ?? "";
    layout.replaceWith(`${left}\t${right}`);
  });

  const explicitTabMarkers = Array.from(
    root.querySelectorAll<HTMLElement>("[data-docx-tab-char='true']")
  );
  explicitTabMarkers.forEach((marker) => {
    marker.replaceWith("\t");
  });
}

function editableTextFromElement(element: HTMLElement): string {
  const clone = element.cloneNode(true) as HTMLElement;
  replaceTabLayoutMarkersWithTabText(clone);
  clone
    .querySelectorAll("[data-docx-numbering-label='true']")
    .forEach((label) => {
      label.remove();
    });
  clone.querySelectorAll("br").forEach((lineBreak) => {
    lineBreak.replaceWith("\n");
  });
  return clone.textContent ?? "";
}

function editableTextFromTableCellElement(element: HTMLElement): string {
  const paragraphLikeChildren = Array.from(element.children).filter(
    (child): child is HTMLElement =>
      child instanceof HTMLElement &&
      child.getAttribute("contenteditable") !== "false"
  );
  if (paragraphLikeChildren.length === 0) {
    return editableTextFromElement(element);
  }

  return paragraphLikeChildren
    .map((paragraphLikeChild) => {
      if (paragraphLikeChild.tagName.toUpperCase() === "BR") {
        return "";
      }
      const text = editableTextFromElement(paragraphLikeChild);
      return text === "\n" ? "" : text;
    })
    .join("\n");
}

function editableTextFromDraftHtml(html: string): string {
  if (typeof document === "undefined") {
    return html
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<[^>]+>/g, "")
      .replace(/&nbsp;/gi, " ");
  }

  const container = document.createElement("div");
  container.innerHTML = html;
  return editableTextFromElement(container);
}

function editableTextFromTableCellDraftHtml(html: string): string {
  if (typeof document === "undefined") {
    return editableTextFromDraftHtml(html);
  }

  const container = document.createElement("div");
  container.innerHTML = html;
  return editableTextFromTableCellElement(container);
}

function formFieldDisplayValue(field: FormFieldRunNode): string {
  switch (field.fieldType) {
    case "checkbox":
      return field.checked ?? field.widget?.checkbox?.defaultChecked
        ? field.checkedSymbol ?? "☒"
        : field.uncheckedSymbol ?? "☐";
    case "dropdown":
      return field.value ?? field.options?.[0]?.displayText ?? "";
    case "date":
      return field.value ?? "";
    case "text":
      return field.value ?? field.widget?.text?.defaultText ?? "";
    default:
      return field.value ?? "";
  }
}

function firstRunStyle(paragraph?: ParagraphNode): TextRunNode["style"] {
  return textRuns(
    paragraph ?? ({ type: "paragraph", children: [] } as ParagraphNode)
  )[0]?.style;
}

function ensureTextRunNode(paragraph: ParagraphNode): TextRunNode {
  const existing = paragraph.children.find(
    (child): child is TextRunNode => child.type === "text"
  );
  if (existing) {
    return existing;
  }

  const created: TextRunNode = {
    type: "text",
    text: "",
    style: {},
  };
  paragraph.children.unshift(created);
  return created;
}

function isParagraphCellContentNode(
  node: TableCellContentNode
): node is ParagraphNode {
  return node.type === "paragraph";
}

function isTableCellTableContentNode(
  node: TableCellContentNode
): node is TableNode {
  return node.type === "table";
}

function tableCellParagraphs(
  nodeContent: TableCellContentNode[]
): ParagraphNode[] {
  return nodeContent.filter(isParagraphCellContentNode);
}

function tableCellParagraphsRecursively(
  nodeContent: TableCellContentNode[]
): ParagraphNode[] {
  const paragraphs: ParagraphNode[] = [];

  const walk = (entries: TableCellContentNode[]): void => {
    for (const entry of entries) {
      if (isParagraphCellContentNode(entry)) {
        paragraphs.push(entry);
        continue;
      }

      for (const row of entry.rows) {
        for (const nestedCell of row.cells) {
          walk(nestedCell.nodes);
        }
      }
    }
  };

  walk(nodeContent);
  return paragraphs;
}

function tableCellHasImage(nodeContent: TableCellContentNode[]): boolean {
  for (const entry of nodeContent) {
    if (isParagraphCellContentNode(entry) && paragraphHasImage(entry)) {
      return true;
    }

    if (isTableCellTableContentNode(entry)) {
      for (const row of entry.rows) {
        for (const nestedCell of row.cells) {
          if (tableCellHasImage(nestedCell.nodes)) {
            return true;
          }
        }
      }
    }
  }

  return false;
}

function paragraphHasImage(paragraph: ParagraphNode): boolean {
  return paragraph.children.some((child) => child.type === "image");
}

function paragraphHasInFlowImage(paragraph: ParagraphNode): boolean {
  return paragraph.children.some((child) => {
    if (child.type !== "image") {
      return false;
    }

    return (
      !shouldRenderAbsoluteFloatingImage(child) &&
      !shouldRenderWrappedFloatingImage(child)
    );
  });
}

function paragraphHasTextBearingAbsoluteFloatingTextBox(
  paragraph: ParagraphNode
): boolean {
  return paragraph.children.some(
    (child) =>
      child.type === "image" &&
      shouldRenderAbsoluteFloatingImage(child) &&
      child.syntheticTextBox === true &&
      !syntheticTextBoxContainsPictureLayer(child) &&
      Boolean(floatingTextBoxVisibleTextFromImage(child))
  );
}

function paragraphIsFloatingImageAnchorOnly(paragraph: ParagraphNode): boolean {
  if (
    paragraphHasVisibleText(paragraph) ||
    paragraphHasFormField(paragraph) ||
    paragraphHasTextBearingAbsoluteFloatingTextBox(paragraph)
  ) {
    return false;
  }

  let hasFloatingImage = false;
  for (const child of paragraph.children) {
    if (child.type === "text") {
      if (child.text.trim().length > 0) {
        return false;
      }
      continue;
    }

    if (child.type !== "image") {
      return false;
    }

    if (
      !shouldRenderWrappedFloatingImage(child) &&
      !shouldRenderAbsoluteFloatingImage(child)
    ) {
      return false;
    }

    hasFloatingImage = true;
  }

  return hasFloatingImage;
}

function paragraphIsAbsoluteFloatingImageAnchorOnly(
  paragraph: ParagraphNode
): boolean {
  const allowSectionBreakTextBearingAnchor =
    paragraphContainsSectionBreakProperties(paragraph) &&
    !paragraphHasVisibleText(paragraph) &&
    !paragraphHasFormField(paragraph);

  if (
    paragraphHasVisibleText(paragraph) ||
    paragraphHasFormField(paragraph) ||
    (paragraphHasTextBearingAbsoluteFloatingTextBox(paragraph) &&
      !allowSectionBreakTextBearingAnchor)
  ) {
    return false;
  }

  let hasAbsoluteFloatingImage = false;
  for (const child of paragraph.children) {
    if (child.type === "text") {
      if (child.text.trim().length > 0) {
        return false;
      }
      continue;
    }

    if (child.type !== "image" || !shouldRenderAbsoluteFloatingImage(child)) {
      return false;
    }

    hasAbsoluteFloatingImage = true;
  }

  return hasAbsoluteFloatingImage;
}

function paragraphContainsOnlyAbsoluteFloatingContent(
  paragraph: ParagraphNode
): boolean {
  if (paragraphHasFormField(paragraph)) {
    return false;
  }

  let hasAbsoluteFloatingImage = false;
  for (const child of paragraph.children) {
    if (child.type === "text") {
      continue;
    }

    if (child.type !== "image" || !shouldRenderAbsoluteFloatingImage(child)) {
      return false;
    }

    hasAbsoluteFloatingImage = true;
  }

  return hasAbsoluteFloatingImage;
}

function paragraphIsBehindTextAbsoluteFloatingImageAnchorOnly(
  paragraph: ParagraphNode
): boolean {
  if (!paragraphIsAbsoluteFloatingImageAnchorOnly(paragraph)) {
    return false;
  }

  return paragraph.children.every(
    (child) =>
      child.type === "text" ||
      (child.type === "image" && child.floating?.behindDocument === true)
  );
}

function imageBehavesAsDecorativeBehindTextBackground(
  image: ImageRunNode,
  paragraph: ParagraphNode
): boolean {
  return (
    image.floating?.behindDocument === true &&
    shouldRenderAbsoluteFloatingImage(image) &&
    paragraphIsBehindTextAbsoluteFloatingImageAnchorOnly(paragraph)
  );
}

function paragraphActsAsDecorativeBehindTextBackgroundOverlay(
  paragraph: ParagraphNode
): boolean {
  return (
    paragraphIsBehindTextAbsoluteFloatingImageAnchorOnly(paragraph) &&
    !paragraphHasVisibleText(paragraph) &&
    !paragraphHasFormField(paragraph) &&
    paragraph.children.every(
      (child) => child.type !== "image" || child.syntheticTextBox !== true
    )
  );
}

function paragraphNeedsPageWidthAnchorHost(paragraph: ParagraphNode): boolean {
  if (!paragraphIsFloatingImageAnchorOnly(paragraph)) {
    return false;
  }

  return paragraph.children.some((child) => {
    if (
      child.type !== "image" ||
      !shouldRenderAbsoluteFloatingImage(child) ||
      !child.floating
    ) {
      return false;
    }

    const horizontalRelativeTo = child.floating.horizontalRelativeTo
      ?.trim()
      .toLowerCase();
    return horizontalRelativeTo === "page" || horizontalRelativeTo === "margin";
  });
}

function paragraphHasAbsoluteFloatingImage(paragraph: ParagraphNode): boolean {
  return paragraph.children.some(
    (child) =>
      child.type === "image" && shouldRenderAbsoluteFloatingImage(child)
  );
}

function collectHeadingTextColorByLevel(
  model: DocModel
): Partial<Record<number, string>> {
  const colorsByLevel: Partial<Record<number, string>> = {};
  const paragraphStyleById = new Map(
    (model.metadata.paragraphStyles ?? []).map((styleDefinition) => [
      styleDefinition.id,
      styleDefinition,
    ])
  );

  const resolveParagraphHeadingLevel = (
    paragraph: ParagraphNode
  ): number | undefined => {
    if (
      Number.isFinite(paragraph.style?.headingLevel) &&
      (paragraph.style?.headingLevel as number) > 0
    ) {
      return Math.round(paragraph.style?.headingLevel as number);
    }

    const styleDefinition = paragraph.style?.styleId
      ? paragraphStyleById.get(paragraph.style.styleId)
      : undefined;
    const fromStyleDefinition = styleDefinition
      ? resolveParagraphStyleHeadingLevel(styleDefinition)
      : undefined;
    if (fromStyleDefinition) {
      return fromStyleDefinition;
    }

    return (
      headingLevelFromStyleLabel(paragraph.style?.styleId) ??
      headingLevelFromStyleLabel(paragraph.style?.styleName)
    );
  };

  const resolveParagraphTextColor = (
    paragraph: ParagraphNode
  ): string | undefined => {
    for (const child of paragraph.children) {
      if (
        (child.type === "text" || child.type === "form-field") &&
        child.style?.color
      ) {
        return child.style.color;
      }
    }

    const styleDefinition = paragraph.style?.styleId
      ? paragraphStyleById.get(paragraph.style.styleId)
      : undefined;
    return styleDefinition?.runStyle?.color;
  };

  const registerParagraph = (paragraph: ParagraphNode): void => {
    const headingLevel = resolveParagraphHeadingLevel(paragraph);
    if (!headingLevel || colorsByLevel[headingLevel]) {
      return;
    }

    const color = resolveParagraphTextColor(paragraph);
    if (!color) {
      return;
    }

    colorsByLevel[headingLevel] = color;
  };

  const walkCellNodes = (nodes: TableCellContentNode[]): void => {
    nodes.forEach((node) => {
      if (node.type === "paragraph") {
        registerParagraph(node);
        return;
      }

      node.rows.forEach((row) => {
        row.cells.forEach((cell) => {
          walkCellNodes(cell.nodes);
        });
      });
    });
  };

  model.nodes.forEach((node) => {
    if (node.type === "paragraph") {
      registerParagraph(node);
      return;
    }

    node.rows.forEach((row) => {
      row.cells.forEach((cell) => {
        walkCellNodes(cell.nodes);
      });
    });
  });

  return colorsByLevel;
}

function sectionNodesNeedPageWideLayout(
  nodes: DocModel["nodes"],
  pageWidthPx: number,
  contentWidthPx: number
): boolean {
  const safePageWidthPx = Math.max(1, Math.round(pageWidthPx));
  const safeContentWidthPx = Math.max(1, Math.round(contentWidthPx));
  const requiresWideImageLayout = (image: ImageRunNode): boolean => {
    if (!image.floating) {
      return false;
    }

    const horizontalRelativeTo =
      image.floating.horizontalRelativeTo?.toLowerCase();
    const usesPageWidthAnchorOrigin =
      horizontalRelativeTo === "page" || horizontalRelativeTo === "margin";
    if (image.floating.behindDocument && usesPageWidthAnchorOrigin) {
      return true;
    }

    if (
      usesPageWidthAnchorOrigin &&
      (shouldRenderAbsoluteFloatingImage(image) ||
        shouldRenderWrappedFloatingImage(image))
    ) {
      return true;
    }

    if (
      Number.isFinite(image.widthPx) &&
      Number.isFinite(image.floating.xPx) &&
      (image.widthPx as number) >= safeContentWidthPx - 12 &&
      (image.floating.xPx as number) <= 12
    ) {
      return true;
    }

    if (
      Number.isFinite(image.widthPx) &&
      Number.isFinite(image.floating.xPx) &&
      (image.widthPx as number) >= safePageWidthPx - 12 &&
      (image.floating.xPx as number) <= 12
    ) {
      return true;
    }

    return false;
  };

  const paragraphNeedsWideLayout = (paragraph: ParagraphNode): boolean =>
    paragraph.children.some(
      (child) => child.type === "image" && requiresWideImageLayout(child)
    );
  const tableNeedsWideLayout = (table: TableNode): boolean =>
    table.rows.some((row) =>
      row.cells.some((cell) =>
        cell.nodes.some((cellNode) =>
          cellNode.type === "paragraph"
            ? paragraphNeedsWideLayout(cellNode)
            : tableNeedsWideLayout(cellNode)
        )
      )
    );

  return nodes.some((node) =>
    node.type === "paragraph"
      ? paragraphNeedsWideLayout(node)
      : tableNeedsWideLayout(node)
  );
}

function sectionNodesNeedFullPageFooterOverlay(
  nodes: DocModel["nodes"]
): boolean {
  const paragraphNeedsFullPageFooterOverlay = (
    paragraph: ParagraphNode
  ): boolean => paragraphHasPageAnchoredAbsoluteFloatingImage(paragraph);
  const tableNeedsFullPageFooterOverlay = (table: TableNode): boolean =>
    table.rows.some((row) =>
      row.cells.some((cell) =>
        cell.nodes.some((cellNode) =>
          cellNode.type === "paragraph"
            ? paragraphNeedsFullPageFooterOverlay(cellNode)
            : tableNeedsFullPageFooterOverlay(cellNode)
        )
      )
    );

  return nodes.some((node) =>
    node.type === "paragraph"
      ? paragraphNeedsFullPageFooterOverlay(node)
      : tableNeedsFullPageFooterOverlay(node)
  );
}

function paragraphHasFormField(paragraph: ParagraphNode): boolean {
  return paragraph.children.some((child) => child.type === "form-field");
}

function paragraphHasCheckboxFormField(paragraph: ParagraphNode): boolean {
  return paragraph.children.some(
    (child) => child.type === "form-field" && child.fieldType === "checkbox"
  );
}

function paragraphHasVisibleText(paragraph: ParagraphNode): boolean {
  return paragraph.children.some(
    (child) =>
      (child.type === "text" && child.text.trim().length > 0) ||
      (child.type === "form-field" &&
        formFieldDisplayValue(child).trim().length > 0)
  );
}

function normalizeFloatingTextBoxComparisonText(text: string): string {
  return text
    .replace(/\r/g, "")
    .split("\n")
    .map((line) => line.replace(/[\s\u00a0]+/g, " ").trim())
    .filter((line) => line.length > 0)
    .join("\n")
    .trim();
}

function floatingTextBoxVisibleTextFromImage(
  image: ImageRunNode
): string | undefined {
  if (
    !shouldRenderAbsoluteFloatingImage(image) ||
    !image.sourceXml ||
    !/<w:txbxContent\b/i.test(image.sourceXml)
  ) {
    return undefined;
  }

  const paragraphs = syntheticTextBoxParagraphsFromRunXml(image.sourceXml);
  if (paragraphs.length === 0) {
    return undefined;
  }

  const text = paragraphs
    .map((paragraph) =>
      paragraph.segments.map((segment) => segment.text).join("")
    )
    .join("\n");
  const normalized = normalizeFloatingTextBoxComparisonText(text);
  return normalized.length > 0 ? normalized : undefined;
}

function paragraphVisibleTextIsOnlyAbsoluteFloatingTextBoxContent(
  paragraph: ParagraphNode
): boolean {
  if (!paragraphHasVisibleText(paragraph) || paragraphHasFormField(paragraph)) {
    return false;
  }

  const paragraphVisibleText = normalizeFloatingTextBoxComparisonText(
    paragraph.children
      .filter((child): child is TextRunNode => child.type === "text")
      .map((child) => child.text)
      .join("")
  );
  if (!paragraphVisibleText) {
    return false;
  }

  const floatingTextBoxTexts = paragraph.children
    .filter(
      (child): child is ImageRunNode =>
        child.type === "image" && shouldRenderAbsoluteFloatingImage(child)
    )
    .map((child) => floatingTextBoxVisibleTextFromImage(child))
    .filter((text): text is string => Boolean(text));

  if (floatingTextBoxTexts.length === 0) {
    return false;
  }

  if (floatingTextBoxTexts.includes(paragraphVisibleText)) {
    return true;
  }

  return (
    normalizeFloatingTextBoxComparisonText(floatingTextBoxTexts.join("\n")) ===
    paragraphVisibleText
  );
}

function paragraphHasOnlyWhitespaceText(paragraph: ParagraphNode): boolean {
  if (paragraphHasImage(paragraph) || paragraphHasFormField(paragraph)) {
    return false;
  }

  return paragraph.children.every((child) => {
    if (child.type !== "text") {
      return false;
    }

    return child.text.replace(/[\s\u00a0]+/g, "").length === 0;
  });
}

function paragraphHasActiveNumbering(paragraph: ParagraphNode): boolean {
  const numbering = paragraph.style?.numbering;
  return Boolean(
    numbering &&
      Number.isFinite(numbering.numId) &&
      Math.round(numbering.numId) > 0
  );
}

function paragraphContainsSectionBreakProperties(
  paragraph: ParagraphNode
): boolean {
  return /<w:sectPr\b/i.test(paragraph.sourceXml ?? "");
}

function paragraphAbsoluteFloatingAnchorsDependOnParagraphFlow(
  paragraph: ParagraphNode
): boolean {
  return paragraph.children.some((child) => {
    if (
      child.type !== "image" ||
      !shouldRenderAbsoluteFloatingImage(child) ||
      child.syntheticTextBox !== true ||
      !floatingTextBoxVisibleTextFromImage(child) ||
      child.floating?.behindDocument !== true
    ) {
      return false;
    }

    const verticalRelativeTo = child.floating?.verticalRelativeTo
      ?.trim()
      .toLowerCase();
    return (
      verticalRelativeTo === undefined ||
      verticalRelativeTo === "" ||
      verticalRelativeTo === "paragraph" ||
      verticalRelativeTo === "line"
    );
  });
}

function likelyFullPageCoverImageRelativeToContentBox(
  image: ImageRunNode,
  pageContentWidthPx: number,
  pageContentHeightPx: number
): boolean {
  if (!shouldRenderAbsoluteFloatingImage(image) || !image.floating) {
    return false;
  }

  const floating = image.floating;
  if (
    floating.behindDocument !== true ||
    (floating.wrapType ?? "none") !== "none"
  ) {
    return false;
  }

  const widthPx = Math.max(0, Math.round(image.widthPx ?? 0));
  const heightPx = Math.max(0, Math.round(image.heightPx ?? 0));
  const safeContentWidthPx = Math.max(1, Math.round(pageContentWidthPx));
  const safeContentHeightPx = Math.max(1, Math.round(pageContentHeightPx));
  const widthLooksCoverSized = widthPx >= safeContentWidthPx * 0.8;
  const heightLooksCoverSized = heightPx >= safeContentHeightPx * 0.7;
  const verticalRelativeTo = floating.verticalRelativeTo?.trim().toLowerCase();

  return (
    widthLooksCoverSized &&
    heightLooksCoverSized &&
    (verticalRelativeTo === undefined ||
      verticalRelativeTo === "" ||
      verticalRelativeTo === "paragraph" ||
      verticalRelativeTo === "line")
  );
}

export function isLikelyFullPageCoverFloatingImage(
  image: ImageRunNode,
  pageWidthPx: number,
  pageHeightPx: number
): boolean {
  return likelyFullPageCoverImageRelativeToContentBox(
    image,
    pageWidthPx,
    pageHeightPx
  );
}

export function resolveLinkedImageWrapperStyle(params: {
  baseStyle?: Record<string, string | number | undefined>;
  cursor?: Record<string, string | number | undefined>["cursor"];
}): Record<string, string | number | undefined> {
  const { baseStyle, cursor } = params;
  const hasPositionedBaseStyle =
    baseStyle !== undefined && Object.keys(baseStyle).length > 0;

  return {
    // Floating images carry their positioning (float/position/offsets) on the
    // wrapper so the hyperlink anchor occupies the same box as the image.
    ...(hasPositionedBaseStyle ? baseStyle : { display: "inline-block" }),
    lineHeight: 0,
    verticalAlign: "middle",
    ...(cursor !== undefined ? { cursor } : undefined),
  };
}

function paragraphIsLikelyFullPageCoverArtAnchor(
  paragraph: ParagraphNode,
  pageContentWidthPx: number,
  pageContentHeightPx: number
): boolean {
  return (
    paragraphContainsOnlyAbsoluteFloatingContent(paragraph) &&
    paragraph.children.some(
      (child) =>
        child.type === "image" &&
        likelyFullPageCoverImageRelativeToContentBox(
          child,
          pageContentWidthPx,
          pageContentHeightPx
        )
    )
  );
}

function pageAnchoredImageLikelyStartsNearPageTop(
  image: ImageRunNode,
  layout: DocumentLayoutMetrics
): boolean {
  if (!shouldRenderAbsoluteFloatingImage(image) || !image.floating) {
    return false;
  }

  const verticalRelativeTo = image.floating.verticalRelativeTo
    ?.trim()
    .toLowerCase();
  if (verticalRelativeTo !== "page" && verticalRelativeTo !== "margin") {
    return false;
  }

  const imageHeightPx = Math.max(0, Math.round(image.heightPx ?? 0));
  const imageWidthPx = Math.max(0, Math.round(image.widthPx ?? 0));
  const coverSized =
    imageWidthPx >= Math.round(layout.pageWidthPx * 0.8) &&
    imageHeightPx >= Math.round(layout.pageHeightPx * 0.7);
  if (!coverSized) {
    return false;
  }

  const topOffsetPx = image.floating.yPx ?? 0;
  return topOffsetPx <= layout.marginsPx.top;
}

function paragraphActsAsPageAnchoredCoverOverlayHost(
  paragraph: ParagraphNode,
  layout: DocumentLayoutMetrics
): boolean {
  if (!paragraphNeedsPageWidthAnchorHost(paragraph)) {
    return false;
  }

  return paragraph.children.some(
    (child) =>
      child.type === "image" &&
      pageAnchoredImageLikelyStartsNearPageTop(child, layout)
  );
}

function paragraphStartsNormalFlowContent(paragraph: ParagraphNode): boolean {
  if (!paragraphHasVisibleText(paragraph)) {
    return false;
  }

  if (paragraphVisibleTextIsOnlyAbsoluteFloatingTextBoxContent(paragraph)) {
    return false;
  }

  return !paragraph.children.some(
    (child) =>
      child.type === "image" && shouldRenderAbsoluteFloatingImage(child)
  );
}

function paragraphParticipatesInLeadingCoverLayout(
  model: DocModel,
  nodeIndex: number,
  pageContentWidthPx: number,
  pageContentHeightPx: number
): boolean {
  let sawLikelyCoverArtAnchor = false;
  for (let probeIndex = 0; probeIndex <= nodeIndex; probeIndex += 1) {
    const probeNode = model.nodes[probeIndex];
    if (!probeNode || probeNode.type !== "paragraph") {
      return false;
    }

    if (paragraphStartsNormalFlowContent(probeNode)) {
      return false;
    }

    if (
      paragraphIsLikelyFullPageCoverArtAnchor(
        probeNode,
        pageContentWidthPx,
        pageContentHeightPx
      )
    ) {
      sawLikelyCoverArtAnchor = true;
    }
  }

  return sawLikelyCoverArtAnchor;
}

function fullPageCoverImageRenderKey(
  nodeIndex: number,
  childIndex: number
): string {
  return `${nodeIndex}:${childIndex}`;
}

function paragraphActsAsLeadingCoverLayoutOverlay(
  model: DocModel,
  nodeIndex: number,
  paragraph: ParagraphNode,
  pageContentWidthPx: number,
  pageContentHeightPx: number
): boolean {
  if (
    !paragraphParticipatesInLeadingCoverLayout(
      model,
      nodeIndex,
      pageContentWidthPx,
      pageContentHeightPx
    ) ||
    paragraphIsLikelyFullPageCoverArtAnchor(
      paragraph,
      pageContentWidthPx,
      pageContentHeightPx
    ) ||
    (paragraphHasVisibleText(paragraph) &&
      !paragraphVisibleTextIsOnlyAbsoluteFloatingTextBoxContent(paragraph)) ||
    paragraphHasFormField(paragraph)
  ) {
    return false;
  }

  return paragraph.children.every(
    (child) =>
      child.type === "text" ||
      (child.type === "image" && shouldRenderAbsoluteFloatingImage(child))
  );
}

function fullPageCoverAbsoluteFloatingImageStyle(
  image: ImageRunNode,
  layout: DocumentLayoutMetrics,
  options?: {
    deltaX?: number;
    deltaY?: number;
    anchorToPageSurface?: boolean;
  }
): Record<string, string | number | undefined> {
  const floating = image.floating;
  const normalizedZIndex = Number.isFinite(floating?.zIndex)
    ? Math.max(
        1,
        Math.min(
          65535,
          Math.round((floating?.zIndex as number) / WORD_IMAGE_Z_INDEX_STEP)
        )
      )
    : 1;
  return {
    position: "absolute",
    left:
      (options?.anchorToPageSurface ? 0 : -layout.marginsPx.left) +
      Math.round(options?.deltaX ?? 0),
    top:
      (options?.anchorToPageSurface ? 0 : -layout.marginsPx.top) +
      Math.round(options?.deltaY ?? 0),
    width: layout.pageWidthPx,
    height: layout.pageHeightPx,
    zIndex: floating?.behindDocument === true ? 0 : normalizedZIndex,
  };
}

function paragraphLooksLikeCheckboxChoiceRow(
  paragraph: ParagraphNode
): boolean {
  if (paragraph.children.some((child) => child.type === "image")) {
    return false;
  }

  const checkboxCount = paragraph.children.filter(
    (child) => child.type === "form-field" && child.fieldType === "checkbox"
  ).length;
  if (checkboxCount < 2) {
    return false;
  }

  const combinedText = paragraph.children
    .filter((child): child is TextRunNode => child.type === "text")
    .map((child) => child.text)
    .join("");
  if (!combinedText.includes("\t")) {
    return false;
  }

  const normalized = combinedText.replace(/\s+/g, " ").trim().toLowerCase();
  return normalized.includes("yes") && normalized.includes("no");
}

interface ParagraphPretextLayoutRun {
  kind: "text" | "image" | "tab";
  key: string;
  text: string;
  startOffset: number;
  endOffset: number;
  style?: TextRunNode["style"];
  link?: string;
  image?: ImageRunNode;
  tabWidthPx?: number;
}

interface ParagraphPretextLayoutSource {
  text: string;
  runs: ParagraphPretextLayoutRun[];
}

const KEEP_ALL_SCRIPT_RE =
  /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u;

interface WrappedParagraphEditingSession {
  location: ParagraphLocation;
  locationKey: string;
  text: string;
  selectionStart: number;
  selectionEnd: number;
  anchorOffset: number;
  isComposing: boolean;
  preferredCaretX?: number;
}

interface WrappedParagraphSelectionDragState {
  pointerId: number;
  locationKey: string;
  anchorOffset: number;
  anchorBoundary: DocxTextRangeBoundary;
  startX: number;
  startY: number;
}

interface WrappedParagraphSurfaceRegistration {
  location: ParagraphLocation;
  element: HTMLElement;
  layout: PretextVariableWidthLayout;
  textLength: number;
}

interface DualWrappedFloatingImageGeometry {
  image: ImageRunNode;
  imageIndex: number;
  containerWidthPx: number;
  imageLeftPx: number;
  imageTopPx: number;
  imageWidthPx: number;
  imageHeightPx: number;
  exclusion: PretextExclusionRect;
}

const MIN_DUAL_WRAPPED_INTERIOR_BAND_PX = 72;

interface ParagraphDualWrappedTextLayout {
  source: ParagraphPretextLayoutSource;
  geometries: DualWrappedFloatingImageGeometry[];
  lineHeightPx: number;
  layout: PretextVariableWidthLayout;
}

interface PageFlowFloatingWrapObstacle extends PretextExclusionRect {
  sourceNodeIndex: number;
}

function resolveDualWrapParagraphRenderBlockHeightPx(
  layout: PretextVariableWidthLayout,
  geometries: DualWrappedFloatingImageGeometry[]
): number {
  const fullHeightPx = wrappedPretextParagraphBlockHeightPx(layout);
  const textBottomPx = pretextLayoutContentBottomPx(layout);
  const minImageTopPx = geometries.reduce(
    (smallest, geometry) => Math.min(smallest, geometry.exclusion.top),
    0
  );
  const maxImageBottomPx = geometries.reduce(
    (largest, geometry) => Math.max(largest, geometry.exclusion.bottom),
    0
  );
  if (minImageTopPx < -4 || maxImageBottomPx > textBottomPx + 4) {
    return Math.max(MIN_PARAGRAPH_LINE_HEIGHT_PX, textBottomPx);
  }

  return fullHeightPx;
}

export function resolveForeignWrapExclusionsForFlowRange(
  obstacles: PageFlowFloatingWrapObstacle[],
  sourceNodeIndex: number,
  pageFlowTopPx: number,
  pageFlowBottomPx: number
): PretextExclusionRect[] {
  const foreignExclusions: PretextExclusionRect[] = [];

  for (const obstacle of obstacles) {
    if (obstacle.sourceNodeIndex === sourceNodeIndex) {
      continue;
    }
    if (obstacle.bottom <= pageFlowTopPx || obstacle.top >= pageFlowBottomPx) {
      continue;
    }

    foreignExclusions.push({
      left: obstacle.left,
      right: obstacle.right,
      top: Math.max(0, Math.round(obstacle.top - pageFlowTopPx)),
      bottom: Math.max(0, Math.round(obstacle.bottom - pageFlowTopPx)),
    });
  }

  return foreignExclusions;
}

function resolveWrappedImageGeometryForPageFlow(
  image: ImageRunNode,
  childIndex: number,
  containerWidthPx: number,
  options?: {
    paragraphTopPx?: number;
    pageMarginTopPx?: number;
    widthPx?: number;
    heightPx?: number;
    movePreview?: {
      deltaX: number;
      deltaY: number;
      baseLeftPx?: number;
      baseTopPx?: number;
    };
    allowNegativeImageTop?: boolean;
  }
): DualWrappedFloatingImageGeometry | undefined {
  return resolveDualWrappedFloatingImageGeometry(image, containerWidthPx, {
    imageIndex: childIndex,
    widthPx: options?.widthPx,
    heightPx: options?.heightPx,
    paragraphTopPx: options?.paragraphTopPx,
    pageMarginTopPx: options?.pageMarginTopPx,
    deltaX: options?.movePreview?.deltaX,
    deltaY: options?.movePreview?.deltaY,
    baseLeftPx: options?.movePreview?.baseLeftPx,
    baseTopPx: options?.movePreview?.baseTopPx,
    allowNegativeImageTop:
      options?.allowNegativeImageTop ?? Boolean(options?.movePreview),
  });
}

export function collectPageFlowWrapObstaclesForParagraph(
  paragraph: ParagraphNode,
  nodeIndex: number,
  pageFlowTopPx: number,
  containerWidthPx: number,
  lineHeightPx: number,
  options?: {
    paragraphTopPx?: number;
    pageMarginTopPx?: number;
    location?: ParagraphLocation;
    floatingMovePreview?: {
      imageKey: string;
      deltaX: number;
      deltaY: number;
      baseLeftPx?: number;
      baseTopPx?: number;
    };
    resizePreview?: {
      imageKey: string;
      widthPx: number;
      heightPx: number;
    };
  }
): PageFlowFloatingWrapObstacle[] {
  const obstacles: PageFlowFloatingWrapObstacle[] = [];
  const movePreviewByImageIndex = new Map<
    number,
    {
      deltaX: number;
      deltaY: number;
      baseLeftPx?: number;
      baseTopPx?: number;
    }
  >();

  paragraph.children.forEach((child, childIndex) => {
    if (child.type !== "image" || !options?.location) {
      return;
    }

    const imageKey = imageLocationKey({ ...options.location, childIndex });
    const movePreview =
      options.floatingMovePreview?.imageKey === imageKey
        ? options.floatingMovePreview
        : undefined;
    if (!movePreview) {
      return;
    }

    movePreviewByImageIndex.set(childIndex, {
      deltaX: Math.round(movePreview.deltaX),
      deltaY: Math.round(movePreview.deltaY),
      ...(Number.isFinite(movePreview.baseLeftPx)
        ? { baseLeftPx: Math.round(movePreview.baseLeftPx as number) }
        : undefined),
      ...(Number.isFinite(movePreview.baseTopPx)
        ? { baseTopPx: Math.round(movePreview.baseTopPx as number) }
        : undefined),
    });
  });

  const dualWrappedLayout = resolveParagraphDualWrappedTextLayout(
    paragraph,
    containerWidthPx,
    lineHeightPx,
    {
      paragraphTopPx: options?.paragraphTopPx,
      pageMarginTopPx: options?.pageMarginTopPx,
      movePreviewByImageIndex,
      allowNegativeImageTop: movePreviewByImageIndex.size > 0,
    }
  );
  const coveredImageIndexes = new Set<number>();

  if (dualWrappedLayout) {
    for (const geometry of dualWrappedLayout.geometries) {
      coveredImageIndexes.add(geometry.imageIndex);
      obstacles.push({
        sourceNodeIndex: nodeIndex,
        left: geometry.exclusion.left,
        right: geometry.exclusion.right,
        top: pageFlowTopPx + geometry.exclusion.top,
        bottom: pageFlowTopPx + geometry.exclusion.bottom,
      });
    }
  }

  paragraph.children.forEach((child, childIndex) => {
    if (child.type !== "image" || !shouldRenderWrappedFloatingImage(child)) {
      return;
    }
    if (coveredImageIndexes.has(childIndex)) {
      return;
    }

    const imageKey = options?.location
      ? imageLocationKey({ ...options.location, childIndex })
      : undefined;
    const resizePreview =
      imageKey && options?.resizePreview?.imageKey === imageKey
        ? options.resizePreview
        : undefined;
    const movePreview = movePreviewByImageIndex.get(childIndex);
    const geometry = resolveWrappedImageGeometryForPageFlow(
      child,
      childIndex,
      containerWidthPx,
      {
        paragraphTopPx: options?.paragraphTopPx,
        pageMarginTopPx: options?.pageMarginTopPx,
        widthPx: resizePreview?.widthPx ?? child.widthPx,
        heightPx: resizePreview?.heightPx ?? child.heightPx,
        movePreview,
        allowNegativeImageTop: Boolean(movePreview),
      }
    );
    if (!geometry) {
      return;
    }

    obstacles.push({
      sourceNodeIndex: nodeIndex,
      left: geometry.exclusion.left,
      right: geometry.exclusion.right,
      top: pageFlowTopPx + geometry.exclusion.top,
      bottom: pageFlowTopPx + geometry.exclusion.bottom,
    });
  });

  return obstacles;
}

export function resolveParagraphForeignOnlyWrappedTextLayout(
  paragraph: ParagraphNode,
  containerWidthPx: number,
  lineHeightPx: number,
  foreignExclusions: PretextExclusionRect[]
): ParagraphDualWrappedTextLayout | undefined {
  if (foreignExclusions.length === 0) {
    return undefined;
  }

  const source = buildParagraphPretextLayoutSource(paragraph);
  if (!source) {
    return undefined;
  }

  const layout = resolveParagraphPretextExclusionLayout(
    paragraph,
    source,
    containerWidthPx,
    lineHeightPx,
    foreignExclusions
  );
  if (!layout) {
    return undefined;
  }

  return {
    source,
    geometries: [],
    lineHeightPx,
    layout,
  };
}

function precomputePageSegmentForeignWrapExclusions(
  segments: DocumentPageNodeSegment[],
  model: DocModel,
  availableWidthPx: number,
  numberingDefinitions: NumberingDefinitionSet | undefined,
  docGridLinePitchPxByNodeIndex: Map<number, number | undefined>,
  pageContentWidthPxByNodeIndex: Map<number, number>,
  pageLayout: DocumentLayoutMetrics,
  interaction?: {
    floatingMovePreview?: {
      imageKey: string;
      deltaX: number;
      deltaY: number;
      baseLeftPx?: number;
      baseTopPx?: number;
    };
    resizePreview?: {
      imageKey: string;
      widthPx: number;
      heightPx: number;
    };
  }
): PretextExclusionRect[][] {
  const flowTopPxBySegmentIndex: number[] = [];
  const flowHeightPxBySegmentIndex: number[] = [];
  let pageFlowTopPx = 0;

  for (const segment of segments) {
    flowTopPxBySegmentIndex.push(pageFlowTopPx);
    const node = model.nodes[segment.nodeIndex];
    const segmentWidthPx =
      pageContentWidthPxByNodeIndex.get(segment.nodeIndex) ?? availableWidthPx;
    const segmentHeightPx = estimateRenderedPageSegmentHeightPx(
      node,
      segment,
      model,
      segmentWidthPx,
      numberingDefinitions,
      docGridLinePitchPxByNodeIndex.get(segment.nodeIndex),
      // For the wrap-exclusion flow simulation, a square/tight-wrapped float
      // must not inflate its own paragraph's flow height. Following
      // paragraphs flow beside the float (Word behavior), so they need to
      // start right after this paragraph's text — not below the image — for
      // the float's exclusion rect to overlap and indent them.
      { excludeWrappedFloatingImageFootprint: true }
    );
    flowHeightPxBySegmentIndex.push(segmentHeightPx);
    pageFlowTopPx += segmentHeightPx;
  }

  const obstacles: PageFlowFloatingWrapObstacle[] = [];
  for (let segmentIndex = 0; segmentIndex < segments.length; segmentIndex += 1) {
    const segment = segments[segmentIndex];
    const node = model.nodes[segment.nodeIndex];
    if (node?.type !== "paragraph") {
      continue;
    }

    const segmentWidthPx =
      pageContentWidthPxByNodeIndex.get(segment.nodeIndex) ?? availableWidthPx;
    const paragraphRenderTextWidthPx = paragraphAvailableTextWidthPx(
      node,
      segmentWidthPx,
      numberingDefinitions
    );
    obstacles.push(
      ...collectPageFlowWrapObstaclesForParagraph(
        node,
        segment.nodeIndex,
        flowTopPxBySegmentIndex[segmentIndex] ?? 0,
        paragraphRenderTextWidthPx,
        estimateParagraphLineHeightPx(
          node,
          docGridLinePitchPxByNodeIndex.get(segment.nodeIndex)
        ),
        {
          paragraphTopPx: flowTopPxBySegmentIndex[segmentIndex] ?? 0,
          pageMarginTopPx: pageLayout.marginsPx.top,
          location: {
            kind: "paragraph",
            nodeIndex: segment.nodeIndex,
          },
          floatingMovePreview: interaction?.floatingMovePreview,
          resizePreview: interaction?.resizePreview,
        }
      )
    );
  }

  const result = segments.map((segment, segmentIndex) =>
    resolveForeignWrapExclusionsForFlowRange(
      obstacles,
      segment.nodeIndex,
      flowTopPxBySegmentIndex[segmentIndex] ?? 0,
      (flowTopPxBySegmentIndex[segmentIndex] ?? 0) +
        (flowHeightPxBySegmentIndex[segmentIndex] ?? 0)
    )
  );
  return result;
}

function applyWrappedFloatingInteractionPreviewToParagraph(
  paragraph: ParagraphNode,
  location: ParagraphLocation,
  paragraphRenderTextWidthPx: number,
  options?: {
    floatingMovePreview?: {
      imageKey: string;
      deltaX: number;
      deltaY: number;
      baseLeftPx?: number;
      baseTopPx?: number;
    };
    resizePreview?: {
      imageKey: string;
      widthPx: number;
      heightPx: number;
    };
  }
): ParagraphNode {
  const hasInteractionPreview = Boolean(
    options?.floatingMovePreview || options?.resizePreview
  );
  if (!hasInteractionPreview) {
    return paragraph;
  }

  return {
    ...paragraph,
    children: paragraph.children.map((child, childIndex) => {
      if (child.type !== "image") {
        return child;
      }

      const imageLocation: DocxImageLocation = { ...location, childIndex };
      const imageKey = imageLocationKey(imageLocation);
      const resizePreview =
        options?.resizePreview?.imageKey === imageKey
          ? options.resizePreview
          : undefined;
      const movePreview =
        options?.floatingMovePreview?.imageKey === imageKey
          ? options.floatingMovePreview
          : undefined;
      if (!movePreview || !shouldRenderWrappedFloatingImage(child)) {
        return child;
      }

      const widthPx =
        resizePreview?.widthPx ??
        child.widthPx ??
        MIN_PARAGRAPH_LINE_HEIGHT_PX;
      const heightPx =
        resizePreview?.heightPx ??
        child.heightPx ??
        widthPx;
      const baseGeometry = resolveDualWrappedFloatingImageGeometry(
        child,
        paragraphRenderTextWidthPx,
        {
          widthPx,
          heightPx,
          ...(Number.isFinite(movePreview.baseLeftPx)
            ? { baseLeftPx: Math.round(movePreview.baseLeftPx as number) }
            : undefined),
          ...(Number.isFinite(movePreview.baseTopPx)
            ? { baseTopPx: Math.round(movePreview.baseTopPx as number) }
            : undefined),
        }
      );
      const movedLeft = clampNumber(
        Math.round((baseGeometry?.imageLeftPx ?? 0) + movePreview.deltaX),
        0,
        Math.max(0, paragraphRenderTextWidthPx - widthPx)
      );
      const movedTop = Math.round(
        (baseGeometry?.imageTopPx ?? 0) + movePreview.deltaY
      );
      const patch = resolveWrappedFloatingImageDropPatch(
        child,
        paragraphRenderTextWidthPx,
        movedLeft,
        movedTop,
        {
          widthPx,
          heightPx,
        }
      );

      return {
        ...child,
        floating: {
          ...(child.floating ?? {}),
          ...patch,
        },
      };
    }),
  };
}

const paragraphPretextLayoutSourceCache = new WeakMap<
  ParagraphNode,
  Map<number, ParagraphPretextLayoutSource | null>
>();
const paragraphPretextLayoutItemsBySource = new WeakMap<
  ParagraphPretextLayoutSource,
  PretextLayoutItem[]
>();
const paragraphPretextUniformFontBySource = new WeakMap<
  ParagraphPretextLayoutSource,
  string | null
>();

function imageWrapModeFromFloating(
  floating?: ImageRunNode["floating"]
): DocxImageWrapMode {
  if (!floating) {
    return "inline";
  }

  const wrapType = floating.wrapType;
  if (!wrapType || wrapType === "none") {
    return floating.behindDocument ? "behindText" : "inFrontOfText";
  }

  if (
    wrapType === "square" ||
    wrapType === "tight" ||
    wrapType === "through" ||
    wrapType === "topAndBottom"
  ) {
    return wrapType;
  }

  return "square";
}

function floatingImageMovesWithText(
  floating?: ImageRunNode["floating"]
): boolean {
  if (!floating) {
    return true;
  }

  const verticalRelativeTo = floating.verticalRelativeTo?.trim().toLowerCase();
  return verticalRelativeTo !== "page" && verticalRelativeTo !== "margin";
}

export function resolveDocxImageWrapState(
  floating?: ImageRunNode["floating"]
): DocxImageWrapState {
  const moveWithText = floatingImageMovesWithText(floating);
  return {
    mode: imageWrapModeFromFloating(floating),
    moveWithText,
    fixedPositionOnPage: !moveWithText,
  };
}

function imageWrapModeActionId(
  mode: DocxImageWrapMode
): DocxContextMenuActionId {
  switch (mode) {
    case "inline":
      return "image-wrap-inline";
    case "square":
      return "image-wrap-square";
    case "tight":
      return "image-wrap-tight";
    case "through":
      return "image-wrap-through";
    case "topAndBottom":
      return "image-wrap-top-and-bottom";
    case "behindText":
      return "image-wrap-behind-text";
    case "inFrontOfText":
      return "image-wrap-in-front-of-text";
    default:
      return "image-wrap-square";
  }
}

function imageWrapModeFromActionId(
  actionId: DocxContextMenuActionId | (string & {})
): DocxImageWrapMode | undefined {
  switch (actionId) {
    case "image-wrap-inline":
      return "inline";
    case "image-wrap-square":
      return "square";
    case "image-wrap-tight":
      return "tight";
    case "image-wrap-through":
      return "through";
    case "image-wrap-top-and-bottom":
      return "topAndBottom";
    case "image-wrap-behind-text":
    case "image-behind-text":
      return "behindText";
    case "image-wrap-in-front-of-text":
    case "image-in-front-of-text":
      return "inFrontOfText";
    default:
      return undefined;
  }
}

function resolveFloatingForImageWrapMode(
  mode: DocxImageWrapMode,
  currentFloating?: NonNullable<ImageRunNode["floating"]>,
  seedFloating?: Partial<NonNullable<ImageRunNode["floating"]>>
): NonNullable<ImageRunNode["floating"]> | undefined {
  if (mode === "inline") {
    return undefined;
  }

  const currentZIndex = Number.isFinite(currentFloating?.zIndex)
    ? clampNumber(
        Math.round(currentFloating?.zIndex as number),
        WORD_IMAGE_Z_INDEX_MIN,
        WORD_IMAGE_Z_INDEX_MAX
      )
    : WORD_IMAGE_Z_INDEX_DEFAULT;

  const base: NonNullable<ImageRunNode["floating"]> = {
    ...(currentFloating ?? {}),
    ...(seedFloating ?? {}),
    distLPx: Math.max(
      0,
      Math.round(seedFloating?.distLPx ?? currentFloating?.distLPx ?? 8)
    ),
    distRPx: Math.max(
      0,
      Math.round(seedFloating?.distRPx ?? currentFloating?.distRPx ?? 8)
    ),
    distTPx: Math.max(
      0,
      Math.round(seedFloating?.distTPx ?? currentFloating?.distTPx ?? 2)
    ),
    distBPx: Math.max(
      0,
      Math.round(seedFloating?.distBPx ?? currentFloating?.distBPx ?? 4)
    ),
    horizontalRelativeTo:
      seedFloating?.horizontalRelativeTo ??
      currentFloating?.horizontalRelativeTo ??
      "column",
    verticalRelativeTo:
      seedFloating?.verticalRelativeTo ??
      currentFloating?.verticalRelativeTo ??
      "paragraph",
    wrapText:
      seedFloating?.wrapText ?? currentFloating?.wrapText ?? "bothSides",
    behindDocument:
      seedFloating?.behindDocument ?? currentFloating?.behindDocument ?? false,
    zIndex: Number.isFinite(seedFloating?.zIndex)
      ? clampNumber(
          Math.round(seedFloating?.zIndex as number),
          WORD_IMAGE_Z_INDEX_MIN,
          WORD_IMAGE_Z_INDEX_MAX
        )
      : currentZIndex,
  };

  switch (mode) {
    case "behindText":
      return {
        ...base,
        wrapType: "none",
        behindDocument: true,
        zIndex: WORD_IMAGE_Z_INDEX_MIN,
      };
    case "inFrontOfText":
      return {
        ...base,
        wrapType: "none",
        behindDocument: false,
        zIndex: Math.max(WORD_IMAGE_Z_INDEX_STEP, currentZIndex),
      };
    case "topAndBottom":
      return {
        ...base,
        wrapType: "topAndBottom",
        behindDocument: false,
      };
    case "square":
    case "tight":
    case "through":
      return {
        ...base,
        wrapType: mode,
        behindDocument: false,
      };
    default:
      return {
        ...base,
        wrapType: "square",
        behindDocument: false,
      };
  }
}

function buildParagraphPretextTabSpacerText(
  widthPx: number,
  style: TextRunNode["style"] | undefined,
  paragraphBaseFontPx: number
): string {
  const safeWidthPx = Math.max(8, Math.round(widthPx));
  const spacerCharacter = "\u00a0";
  const spacerAdvancePx = Math.max(
    1,
    measureTextWidthPx(spacerCharacter, style, paragraphBaseFontPx)
  );
  let spacerCount = Math.max(1, Math.round(safeWidthPx / spacerAdvancePx));
  let spacerText = spacerCharacter.repeat(spacerCount);
  let measuredWidthPx = measureTextWidthPx(
    spacerText,
    style,
    paragraphBaseFontPx
  );

  while (
    measuredWidthPx + spacerAdvancePx * 0.5 < safeWidthPx &&
    spacerCount < 64
  ) {
    spacerCount += 1;
    spacerText = spacerCharacter.repeat(spacerCount);
    measuredWidthPx = measureTextWidthPx(
      spacerText,
      style,
      paragraphBaseFontPx
    );
  }

  while (spacerCount > 1) {
    const nextText = spacerCharacter.repeat(spacerCount - 1);
    const nextWidthPx = measureTextWidthPx(
      nextText,
      style,
      paragraphBaseFontPx
    );
    if (
      Math.abs(nextWidthPx - safeWidthPx) >=
      Math.abs(measuredWidthPx - safeWidthPx)
    ) {
      break;
    }
    spacerCount -= 1;
    spacerText = nextText;
    measuredWidthPx = nextWidthPx;
  }

  return spacerText;
}

export function buildParagraphPretextLayoutSource(
  paragraph: ParagraphNode,
  options?: {
    allowExplicitLineBreakText?: boolean;
    expandTabsForLayout?: boolean;
  }
): ParagraphPretextLayoutSource | undefined {
  const cacheKey =
    (options?.allowExplicitLineBreakText ? 1 : 0) |
    (options?.expandTabsForLayout ? 2 : 0);
  const cachedVariants = paragraphPretextLayoutSourceCache.get(paragraph);
  const cachedSource = cachedVariants?.get(cacheKey);
  if (cachedSource !== undefined) {
    return cachedSource ?? undefined;
  }
  const storeCachedSource = (
    source: ParagraphPretextLayoutSource | undefined
  ): ParagraphPretextLayoutSource | undefined => {
    if (!cachedVariants) {
      paragraphPretextLayoutSourceCache.set(
        paragraph,
        new Map([[cacheKey, source ?? null]])
      );
    } else {
      cachedVariants.set(cacheKey, source ?? null);
    }
    return source;
  };
  if (paragraphHasFormField(paragraph)) {
    return storeCachedSource(undefined);
  }
  if (
    !options?.allowExplicitLineBreakText &&
    paragraphContainsExplicitLineBreakText(paragraph)
  ) {
    return storeCachedSource(undefined);
  }

  const runs: ParagraphPretextLayoutRun[] = [];
  let combinedText = "";
  const paragraphBaseFontPx = paragraphBaseFontSizePx(paragraph);
  const tabStopPositionsPx = options?.expandTabsForLayout
    ? resolveParagraphTabStopsPx(paragraph)
    : [];
  const usesCheckboxRowTabFallback =
    paragraphLooksLikeCheckboxChoiceRow(paragraph);
  const fallbackTabWidthPx = usesCheckboxRowTabFallback
    ? checkboxChoiceRowTabWidthPx(paragraph)
    : DEFAULT_TAB_STOP_PX;
  let approximateLineWidthPx = 0;
  let lastTextStyle: TextRunNode["style"] | undefined =
    firstRunStyle(paragraph);

  for (
    let childIndex = 0;
    childIndex < paragraph.children.length;
    childIndex += 1
  ) {
    const child = paragraph.children[childIndex];
    if (!child) {
      continue;
    }
    if (child.type === "image") {
      if (child.floating) {
        continue;
      }

      const placeholderText = buildInlineImagePlaceholderText(
        child,
        lastTextStyle,
        paragraphBaseFontPx
      );
      const startOffset = combinedText.length;
      combinedText += placeholderText;
      runs.push({
        kind: "image",
        key: `run-${childIndex}`,
        text: placeholderText,
        startOffset,
        endOffset: combinedText.length,
        style: lastTextStyle,
        image: child,
      });
      continue;
    }
    if (child.type !== "text") {
      return storeCachedSource(undefined);
    }
    if (child.noteReference) {
      return storeCachedSource(undefined);
    }
    if (child.text.includes("\t")) {
      if (!options?.expandTabsForLayout) {
        return storeCachedSource(undefined);
      }

      const tabSegments = child.text.split("\t");
      tabSegments.forEach((segmentText, segmentIndex) => {
        if (segmentText.length > 0) {
          const startOffset = combinedText.length;
          combinedText += segmentText;
          runs.push({
            kind: "text",
            key: `run-${childIndex}-${segmentIndex}`,
            text: segmentText,
            startOffset,
            endOffset: combinedText.length,
            style: child.style,
            link: child.link,
          });
          approximateLineWidthPx = updateEstimatedLineWidthPxForText(
            approximateLineWidthPx,
            segmentText,
            child.style
          );
        }

        if (segmentIndex >= tabSegments.length - 1) {
          return;
        }

        const tabWidthPx = resolveTabSpacerWidthPx(
          tabStopPositionsPx,
          approximateLineWidthPx,
          fallbackTabWidthPx,
          usesCheckboxRowTabFallback
        );
        const spacerText = buildParagraphPretextTabSpacerText(
          tabWidthPx,
          child.style,
          paragraphBaseFontPx
        );
        const startOffset = combinedText.length;
        combinedText += spacerText;
        runs.push({
          kind: "tab",
          key: `run-${childIndex}-${segmentIndex}-tab`,
          text: spacerText,
          startOffset,
          endOffset: combinedText.length,
          style: child.style,
          tabWidthPx,
        });
        approximateLineWidthPx += tabWidthPx;
      });
      lastTextStyle = child.style ?? lastTextStyle;
      continue;
    }

    const startOffset = combinedText.length;
    combinedText += child.text;
    runs.push({
      kind: "text",
      key: `run-${childIndex}`,
      text: child.text,
      startOffset,
      endOffset: combinedText.length,
      style: child.style,
      link: child.link,
    });
    approximateLineWidthPx = updateEstimatedLineWidthPxForText(
      approximateLineWidthPx,
      child.text,
      child.style
    );
    lastTextStyle = child.style ?? lastTextStyle;
  }

  if (combinedText.length === 0) {
    return storeCachedSource(undefined);
  }

  return storeCachedSource({
    text: combinedText,
    runs,
  });
}

function splitParagraphAtExplicitColumnBreaks(
  paragraph: ParagraphNode
): ParagraphNode[] | undefined {
  if (!paragraphHasExplicitColumnBreak(paragraph)) {
    return undefined;
  }

  const paragraphChildren: ParagraphNode["children"][] = [];
  let currentChildren: ParagraphNode["children"] = [];
  let sawExplicitColumnBreak = false;
  const appendCurrentSegment = (): void => {
    paragraphChildren.push(currentChildren);
    currentChildren = [];
  };

  for (
    let childIndex = 0;
    childIndex < paragraph.children.length;
    childIndex += 1
  ) {
    const child = paragraph.children[childIndex];
    if (!child) {
      continue;
    }

    if (child.type === "text") {
      if (/^[\r\n]+$/.test(child.text)) {
        const explicitBreakCount = child.text.replace(/[^\n]/g, "").length;
        if (explicitBreakCount > 0) {
          for (
            let breakIndex = 0;
            breakIndex < explicitBreakCount;
            breakIndex += 1
          ) {
            sawExplicitColumnBreak = true;
            appendCurrentSegment();
          }
          continue;
        }
      }

      currentChildren.push(cloneTextRunWithMetadata(child));
      continue;
    }

    if (child.type === "form-field") {
      currentChildren.push(cloneFormFieldRun(child));
      continue;
    }

    return undefined;
  }

  if (!sawExplicitColumnBreak) {
    return undefined;
  }

  appendCurrentSegment();

  return paragraphChildren.map((children, segmentIndex) => {
    const nextStyle = cloneParagraphStyle(paragraph.style);
    if (segmentIndex > 0 && nextStyle?.numbering) {
      nextStyle.numbering = undefined;
    }

    return {
      ...paragraph,
      style: nextStyle,
      sourceXml: undefined,
      children:
        children.length > 0
          ? children
          : [
              {
                type: "text",
                text: "",
              },
            ],
    };
  });
}

function estimateParagraphContentHeightPx(
  paragraph: ParagraphNode,
  availableWidthPx: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number
): number {
  if (!paragraphHasImage(paragraph) && !paragraphHasFormField(paragraph)) {
    const lineHeightPx = estimateParagraphLineHeightPx(
      paragraph,
      docGridLinePitchPx
    );
    const lineCount = paragraphLineCountWithinWidth(
      paragraph,
      availableWidthPx,
      numberingDefinitions
    );
    return Math.max(1, lineHeightPx * Math.max(1, lineCount));
  }

  return Math.max(
    1,
    estimateParagraphHeightPx(
      paragraph,
      availableWidthPx,
      numberingDefinitions,
      docGridLinePitchPx
    ) -
      paragraphBeforeSpacingPx(paragraph) -
      paragraphAfterSpacingPx(paragraph)
  );
}

function projectParagraphConsumedHeightWithExplicitColumnBreaks(
  paragraphSegments: ParagraphNode[],
  pageConsumedHeightPx: number,
  pageContentHeightPx: number,
  sectionFlowOriginPx: number,
  columnCount: number,
  availableWidthPx: number,
  beforeSpacingPx: number,
  afterSpacingPx: number,
  collapsedMarginPx: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number
): number | undefined {
  if (paragraphSegments.length <= 1 || columnCount <= 1) {
    return undefined;
  }

  const normalizedParagraphStartPx = Math.max(
    0,
    Math.max(pageConsumedHeightPx, sectionFlowOriginPx)
  );
  const firstSegmentTopSpacingPx =
    pageConsumedHeightPx > 0
      ? Math.max(0, beforeSpacingPx - collapsedMarginPx)
      : beforeSpacingPx;
  const tallestSegmentHeightPx = paragraphSegments.reduce(
    (tallest, segment) => {
      return Math.max(
        tallest,
        estimateParagraphContentHeightPx(
          segment,
          availableWidthPx,
          numberingDefinitions,
          docGridLinePitchPx
        )
      );
    },
    0
  );
  const projectedConsumedHeightPx =
    normalizedParagraphStartPx +
    firstSegmentTopSpacingPx +
    tallestSegmentHeightPx +
    afterSpacingPx;
  const sectionFlowCapacityPx = Math.max(
    0,
    pageContentHeightPx - sectionFlowOriginPx
  );
  const sectionConsumedPx = Math.max(
    0,
    pageConsumedHeightPx - sectionFlowOriginPx
  );
  const remainingColumnHeightPx = Math.max(
    0,
    sectionFlowCapacityPx / columnCount - sectionConsumedPx
  );

  if (
    projectedConsumedHeightPx - pageConsumedHeightPx >
    remainingColumnHeightPx + PAGE_OVERFLOW_TOLERANCE_PX
  ) {
    return undefined;
  }

  return projectedConsumedHeightPx;
}

function buildSyntheticPretextLayoutSource(
  text: string,
  style?: TextRunNode["style"]
): ParagraphPretextLayoutSource {
  return {
    text,
    runs: [
      {
        kind: "text",
        key: "synthetic-0",
        text,
        startOffset: 0,
        endOffset: text.length,
        style,
      },
    ],
  };
}

function pretextWordBreakModeForText(text: string): "normal" | "keep-all" {
  const cached = pretextWordBreakModeByText.get(text);
  if (cached) {
    return cached;
  }

  const mode = KEEP_ALL_SCRIPT_RE.test(text) ? "keep-all" : "normal";
  setCacheEntry(pretextWordBreakModeByText, text, mode);
  while (pretextWordBreakModeByText.size > TEXT_MEASURE_CACHE_MAX_ENTRIES) {
    const firstKey = pretextWordBreakModeByText.keys().next().value as
      | string
      | undefined;
    if (!firstKey) {
      break;
    }
    pretextWordBreakModeByText.delete(firstKey);
  }
  return mode;
}

export function sanitizeRenderedPretextFragmentText(text: string): string {
  return text.replace(/\r\n?|\n/g, "");
}

function buildParagraphPretextLayoutItems(
  paragraph: ParagraphNode,
  source: ParagraphPretextLayoutSource
): PretextLayoutItem[] {
  const cachedItems = paragraphPretextLayoutItemsBySource.get(source);
  if (cachedItems) {
    return cachedItems;
  }

  const paragraphBaseFontPx = paragraphBaseFontSizePx(paragraph);
  const wordBreak = pretextWordBreakModeForText(source.text);

  const items: PretextLayoutItem[] = source.runs
    .filter((run) => run.endOffset > run.startOffset)
    .map((run) => ({
      text: run.text,
      font: resolveMeasureFont(run.style, paragraphBaseFontPx),
      startOffset: run.startOffset,
      endOffset: run.endOffset,
      break: run.kind === "image" || run.kind === "tab" ? "never" : "normal",
      wordBreak,
    }));
  paragraphPretextLayoutItemsBySource.set(source, items);
  return items;
}

/**
 * Returns the single measurement font shared by every non-empty run of this
 * pretext source, or `undefined` if there are atomic items (images, tabs)
 * or the runs vary in font. Callers can use this to take a fast
 * `measureLineStats`-based line count path when it returns a value.
 */
function resolveUniformPretextSourceFont(
  paragraph: ParagraphNode,
  source: ParagraphPretextLayoutSource
): string | undefined {
  const cachedUniformFont = paragraphPretextUniformFontBySource.get(source);
  if (cachedUniformFont !== undefined) {
    return cachedUniformFont ?? undefined;
  }

  const paragraphBaseFontPx = paragraphBaseFontSizePx(paragraph);
  let uniformFont: string | undefined;

  for (const run of source.runs) {
    if (run.endOffset <= run.startOffset) {
      continue;
    }
    if (run.kind !== "text") {
      paragraphPretextUniformFontBySource.set(source, null);
      return undefined;
    }

    const runFont = resolveMeasureFont(run.style, paragraphBaseFontPx);
    if (uniformFont === undefined) {
      uniformFont = runFont;
      continue;
    }
    if (runFont !== uniformFont) {
      paragraphPretextUniformFontBySource.set(source, null);
      return undefined;
    }
  }

  paragraphPretextUniformFontBySource.set(source, uniformFont ?? null);
  return uniformFont;
}

function buildMeasureSegmentsPretextLayoutItems(
  segments: ParagraphMeasureSegment[],
  paragraphBaseFontPx: number,
  text: string
): PretextLayoutItem[] {
  const wordBreak = pretextWordBreakModeForText(text);
  let nextOffset = 0;

  return segments
    .map((segment) => {
      const startOffset = nextOffset;
      nextOffset += segment.text.length;
      return {
        text: segment.text,
        font: resolveMeasureFont(segment.style, paragraphBaseFontPx),
        startOffset,
        endOffset: nextOffset,
        break: "normal" as const,
        wordBreak,
      };
    })
    .filter((item) => item.endOffset > item.startOffset);
}

function layoutParagraphPretextSource(
  paragraph: ParagraphNode,
  source: ParagraphPretextLayoutSource,
  containerWidthPx: number,
  lineHeightPx: number,
  exclusions?: PretextExclusionRect[]
): PretextVariableWidthLayout | undefined {
  const fallbackFont = resolveMeasureFont(
    firstRunStyle(paragraph),
    paragraphBaseFontSizePx(paragraph)
  );
  const wordBreak = pretextWordBreakModeForText(source.text);
  const items = buildParagraphPretextLayoutItems(paragraph, source);
  const richLayout =
    items.length > 0
      ? layoutItemsWithPretextAroundExclusions(
          source.text,
          items,
          containerWidthPx,
          lineHeightPx,
          exclusions,
          fallbackFont
        )
      : undefined;
  if (richLayout) {
    return richLayout;
  }

  return layoutTextWithPretextAroundExclusions(
    source.text,
    fallbackFont,
    containerWidthPx,
    lineHeightPx,
    exclusions,
    {
      wordBreak,
    }
  );
}

function wrappedParagraphSessionText(paragraph: ParagraphNode): string {
  return (
    buildParagraphPretextLayoutSource(paragraph)?.text ??
    paragraphText(paragraph)
  );
}

function buildInlineImagePlaceholderText(
  image: ImageRunNode,
  style: TextRunNode["style"] | undefined,
  paragraphBaseFontPx: number
): string {
  const targetWidthPx = Math.max(
    1,
    Math.round(image.widthPx ?? image.heightPx ?? MIN_PARAGRAPH_LINE_HEIGHT_PX)
  );
  let placeholderText = "\u00a0";
  while (
    measureTextWidthPx(placeholderText, style, paragraphBaseFontPx) <
      targetWidthPx &&
    placeholderText.length < 64
  ) {
    placeholderText += "\u00a0";
  }
  return placeholderText;
}

function paragraphChildAnchorOffset(
  paragraph: ParagraphNode,
  childIndex: number
): number {
  const paragraphBaseFontPx = paragraphBaseFontSizePx(paragraph);
  let lastTextStyle: TextRunNode["style"] | undefined =
    firstRunStyle(paragraph);
  let offset = 0;

  for (let currentIndex = 0; currentIndex < childIndex; currentIndex += 1) {
    const child = paragraph.children[currentIndex];
    if (!child) {
      continue;
    }
    if (child.type === "image") {
      if (child.floating) {
        continue;
      }
      offset += buildInlineImagePlaceholderText(
        child,
        lastTextStyle,
        paragraphBaseFontPx
      ).length;
      continue;
    }
    if (child.type !== "text") {
      continue;
    }
    offset += child.text.length;
    lastTextStyle = child.style ?? lastTextStyle;
  }

  return offset;
}

function expandOffsetToWord(
  text: string,
  offset: number
): {
  start: number;
  end: number;
} {
  const safeOffset = Math.max(0, Math.min(Math.round(offset), text.length));
  if (!text) {
    return { start: 0, end: 0 };
  }

  const characterAt = (index: number): string => text.slice(index, index + 1);
  const isWordCharacter = (value: string): boolean =>
    /[0-9A-Za-z_]/.test(value);
  let probeIndex = safeOffset;
  if (probeIndex > 0 && probeIndex === text.length) {
    probeIndex -= 1;
  }
  if (
    !isWordCharacter(characterAt(probeIndex)) &&
    probeIndex > 0 &&
    isWordCharacter(characterAt(probeIndex - 1))
  ) {
    probeIndex -= 1;
  }

  if (!isWordCharacter(characterAt(probeIndex))) {
    return {
      start: safeOffset,
      end: safeOffset,
    };
  }

  let start = probeIndex;
  let end = probeIndex + 1;
  while (start > 0 && isWordCharacter(characterAt(start - 1))) {
    start -= 1;
  }
  while (end < text.length && isWordCharacter(characterAt(end))) {
    end += 1;
  }

  return { start, end };
}

function pixelsToTwips(valuePx: number): number {
  return Math.round(valuePx * TWIPS_PER_PIXEL);
}

export function resolveDualWrappedFloatingImageGeometry(
  image: ImageRunNode,
  containerWidthPx: number,
  options?: {
    imageIndex?: number;
    deltaX?: number;
    deltaY?: number;
    widthPx?: number;
    heightPx?: number;
    baseLeftPx?: number;
    baseTopPx?: number;
    paragraphTopPx?: number;
    pageMarginTopPx?: number;
    allowNegativeImageTop?: boolean;
  }
): DualWrappedFloatingImageGeometry | undefined {
  const floating = image.floating;
  const isWrappedFloatingImage = shouldRenderWrappedFloatingImage(image);
  const hasExplicitBaseLeftPx = Number.isFinite(options?.baseLeftPx);
  const hasExplicitBaseTopPx = Number.isFinite(options?.baseTopPx);
  const isInlinePreview =
    !floating && hasExplicitBaseLeftPx && hasExplicitBaseTopPx;
  if (!isWrappedFloatingImage && !isInlinePreview) {
    return undefined;
  }

  const safeContainerWidthPx = Math.max(1, Math.round(containerWidthPx));
  const imageWidthPx = Math.max(
    1,
    Math.round(
      (options?.widthPx ??
        image.widthPx ??
        image.heightPx ??
        MIN_PARAGRAPH_LINE_HEIGHT_PX) as number
    )
  );
  const imageHeightPx = Math.max(
    1,
    Math.round(
      (options?.heightPx ??
        image.heightPx ??
        image.widthPx ??
        MIN_PARAGRAPH_LINE_HEIGHT_PX) as number
    )
  );
  const deltaX = Number.isFinite(options?.deltaX)
    ? Math.round(options?.deltaX as number)
    : 0;
  const deltaY = Number.isFinite(options?.deltaY)
    ? Math.round(options?.deltaY as number)
    : 0;
  const wrapType = floating?.wrapType;
  const horizontalAlign = floating?.horizontalAlign?.trim().toLowerCase();
  const distLPx = Math.max(0, Math.round(floating?.distLPx ?? 0));
  const distRPx = Math.max(0, Math.round(floating?.distRPx ?? 0));
  const distTPx = Math.max(0, Math.round(floating?.distTPx ?? 0));
  const distBPx = Math.max(0, Math.round(floating?.distBPx ?? 0));
  const baseLeftPx = hasExplicitBaseLeftPx
    ? Math.round(options?.baseLeftPx as number)
    : Number.isFinite(floating?.xPx)
    ? Math.round(floating?.xPx as number)
    : horizontalAlign === "right" || horizontalAlign === "outside"
    ? safeContainerWidthPx - distRPx - imageWidthPx
    : horizontalAlign === "center"
    ? Math.round((safeContainerWidthPx - imageWidthPx) / 2)
    : distLPx;
  const baseTopPx = hasExplicitBaseTopPx
    ? Math.round(options?.baseTopPx as number)
    : Number.isFinite(floating?.yPx)
    ? Math.round(floating?.yPx as number)
    : 0;
  const normalizedVerticalRelativeTo = floating?.verticalRelativeTo
    ?.trim()
    .toLowerCase();
  const paragraphTopPx = Number.isFinite(options?.paragraphTopPx)
    ? Math.max(0, Math.round(options?.paragraphTopPx as number))
    : undefined;
  const pageMarginTopPx = Number.isFinite(options?.pageMarginTopPx)
    ? Math.max(0, Math.round(options?.pageMarginTopPx as number))
    : 0;
  const paragraphLocalBaseTopPx =
    !hasExplicitBaseTopPx &&
    paragraphTopPx !== undefined &&
    (normalizedVerticalRelativeTo === "margin" ||
      normalizedVerticalRelativeTo === "page")
      ? Math.round(
          (normalizedVerticalRelativeTo === "page"
            ? baseTopPx - pageMarginTopPx
            : baseTopPx) - paragraphTopPx
        )
      : baseTopPx;
  const imageLeftPx = clampNumber(
    baseLeftPx + deltaX,
    0,
    Math.max(0, safeContainerWidthPx - imageWidthPx)
  );
  const rawImageTopPx =
    (hasExplicitBaseTopPx
      ? baseTopPx
      : paragraphLocalBaseTopPx + (isInlinePreview ? 0 : distTPx)) + deltaY;
  const modelAllowsNegativeTop =
    Number.isFinite(floating?.yPx) && (floating?.yPx as number) < 0;
  const imageTopPx =
    options?.allowNegativeImageTop || modelAllowsNegativeTop
      ? Math.round(rawImageTopPx)
      : Math.max(0, Math.round(rawImageTopPx));

  let exclusionLeftPx = Math.max(0, imageLeftPx - distLPx);
  let exclusionRightPx = Math.min(
    safeContainerWidthPx,
    imageLeftPx + imageWidthPx + distRPx
  );
  if (wrapType === "topAndBottom") {
    exclusionLeftPx = 0;
    exclusionRightPx = safeContainerWidthPx;
  } else if (horizontalAlign === "left" || horizontalAlign === "inside") {
    exclusionLeftPx = 0;
  } else if (horizontalAlign === "right" || horizontalAlign === "outside") {
    exclusionRightPx = safeContainerWidthPx;
  }

  if (!isInlinePreview && wrapType !== "topAndBottom") {
    const leftBandWidthPx = exclusionLeftPx;
    const rightBandWidthPx = Math.max(
      0,
      safeContainerWidthPx - exclusionRightPx
    );
    const spansInteriorGap =
      exclusionLeftPx > 0 &&
      exclusionRightPx < safeContainerWidthPx &&
      leftBandWidthPx >= MIN_DUAL_WRAPPED_INTERIOR_BAND_PX &&
      rightBandWidthPx >= MIN_DUAL_WRAPPED_INTERIOR_BAND_PX;
    const spansSideFloat =
      exclusionLeftPx === 0 || exclusionRightPx === safeContainerWidthPx;
    if (!spansInteriorGap && !spansSideFloat) {
      return undefined;
    }
  }

  return {
    image,
    imageIndex: options?.imageIndex ?? 0,
    containerWidthPx: safeContainerWidthPx,
    imageLeftPx,
    imageTopPx,
    imageWidthPx,
    imageHeightPx,
    exclusion: {
      left: exclusionLeftPx,
      right: exclusionRightPx,
      top: imageTopPx,
      bottom: imageTopPx + imageHeightPx + distBPx,
    },
  };
}

export function resolveParagraphDualWrappedTextLayout(
  paragraph: ParagraphNode,
  containerWidthPx: number,
  lineHeightPx: number,
  options?: {
    deltaX?: number;
    deltaY?: number;
    widthPxByImageIndex?: Map<number, number>;
    heightPxByImageIndex?: Map<number, number>;
    paragraphTopPx?: number;
    pageMarginTopPx?: number;
    movePreviewByImageIndex?: Map<
      number,
      {
        deltaX: number;
        deltaY: number;
        baseLeftPx?: number;
        baseTopPx?: number;
      }
    >;
    foreignExclusions?: PretextExclusionRect[];
    allowNegativeImageTop?: boolean;
  }
): ParagraphDualWrappedTextLayout | undefined {
  const source = buildParagraphPretextLayoutSource(paragraph);
  if (!source) {
    return undefined;
  }

  const geometries = paragraph.children
    .map((child, childIndex) => {
      if (child.type !== "image" || !child.floating) {
        return undefined;
      }

      const movePreview = options?.movePreviewByImageIndex?.get(childIndex);
      return resolveDualWrappedFloatingImageGeometry(child, containerWidthPx, {
        imageIndex: childIndex,
        deltaX: movePreview?.deltaX ?? options?.deltaX,
        deltaY: movePreview?.deltaY ?? options?.deltaY,
        widthPx: options?.widthPxByImageIndex?.get(childIndex),
        heightPx: options?.heightPxByImageIndex?.get(childIndex),
        paragraphTopPx: options?.paragraphTopPx,
        pageMarginTopPx: options?.pageMarginTopPx,
        baseLeftPx: movePreview?.baseLeftPx,
        baseTopPx: movePreview?.baseTopPx,
        allowNegativeImageTop:
          options?.allowNegativeImageTop ?? Boolean(movePreview),
      });
    })
    .filter((candidate): candidate is DualWrappedFloatingImageGeometry =>
      Boolean(candidate)
    );
  if (geometries.length === 0) {
    return undefined;
  }

  const unexcludedLayout = layoutParagraphPretextSource(
    paragraph,
    source,
    Math.max(...geometries.map((geometry) => geometry.containerWidthPx)),
    lineHeightPx,
    []
  );
  const anchorAdjustedGeometries = geometries.map((geometry) => {
    const wrapType = geometry.image.floating?.wrapType?.trim().toLowerCase();
    const verticalRelativeTo = geometry.image.floating?.verticalRelativeTo
      ?.trim()
      .toLowerCase();
    const anchorOffset = paragraphChildAnchorOffset(
      paragraph,
      geometry.imageIndex
    );
    if (
      wrapType === "topandbottom" &&
      anchorOffset <= 0 &&
      syntheticTextBoxActsAsTopAndBottomMasthead(geometry.image) &&
      (verticalRelativeTo === undefined ||
        verticalRelativeTo === "" ||
        verticalRelativeTo === "paragraph" ||
        verticalRelativeTo === "line")
    ) {
      const clampedTopPx = 0;
      if (geometry.imageTopPx <= clampedTopPx) {
        return geometry;
      }

      const deltaTopPx = clampedTopPx - geometry.imageTopPx;
      return {
        ...geometry,
        imageTopPx: geometry.imageTopPx + deltaTopPx,
        exclusion: {
          ...geometry.exclusion,
          top: geometry.exclusion.top + deltaTopPx,
          bottom: geometry.exclusion.bottom + deltaTopPx,
        },
      };
    }

    if (
      !unexcludedLayout ||
      wrapType === "topandbottom" ||
      (verticalRelativeTo !== "margin" && verticalRelativeTo !== "page")
    ) {
      return geometry;
    }

    if (anchorOffset <= 0) {
      return geometry;
    }

    const anchorCaretRect = resolveCaretRectAtOffset(
      unexcludedLayout,
      anchorOffset
    );
    const anchorTopPx = Math.max(0, Math.round(anchorCaretRect?.top ?? 0));
    if (anchorTopPx <= geometry.imageTopPx) {
      return geometry;
    }

    const deltaTopPx = anchorTopPx - geometry.imageTopPx;
    return {
      ...geometry,
      imageTopPx: geometry.imageTopPx + deltaTopPx,
      exclusion: {
        ...geometry.exclusion,
        top: geometry.exclusion.top + deltaTopPx,
        bottom: geometry.exclusion.bottom + deltaTopPx,
      },
    };
  });

  const mergedExclusions = [
    ...(options?.foreignExclusions ?? []),
    ...anchorAdjustedGeometries.map((geometry) => geometry.exclusion),
  ];
  const layout = resolveParagraphPretextExclusionLayout(
    paragraph,
    source,
    Math.max(
      ...anchorAdjustedGeometries.map((geometry) => geometry.containerWidthPx)
    ),
    lineHeightPx,
    mergedExclusions
  );
  if (!layout) {
    return undefined;
  }

  return {
    source,
    geometries: anchorAdjustedGeometries,
    lineHeightPx,
    layout,
  };
}

function resolveParagraphPretextExclusionLayout(
  paragraph: ParagraphNode,
  source: ParagraphPretextLayoutSource,
  containerWidthPx: number,
  lineHeightPx: number,
  exclusions: PretextExclusionRect[]
): PretextVariableWidthLayout | undefined {
  const layout = layoutParagraphPretextSource(
    paragraph,
    source,
    containerWidthPx,
    lineHeightPx,
    exclusions
  );
  if (!layout || layout.lineCount <= 0) {
    return undefined;
  }

  const normalizedAlignment = (paragraph.style?.align ?? "left")
    .trim()
    .toLowerCase();
  const supportsSingleIntervalAlignment =
    normalizedAlignment === "center" || normalizedAlignment === "right";
  const alignedLines = supportsSingleIntervalAlignment
    ? layout.lines.map((line) => {
        if (line.fragments.length === 0) {
          return line;
        }

        const firstFragment = line.fragments[0];
        const lastFragment = line.fragments[line.fragments.length - 1];
        if (!firstFragment || !lastFragment) {
          return line;
        }
        const usesSingleInterval = line.fragments.every(
          (fragment) =>
            fragment.intervalX === firstFragment.intervalX &&
            fragment.intervalWidth === firstFragment.intervalWidth
        );
        if (!usesSingleInterval) {
          return line;
        }

        const occupiedWidthPx =
          lastFragment.x + lastFragment.width - firstFragment.x;

        const freeSpacePx = Math.max(
          0,
          Math.round(
            (firstFragment.intervalWidth ?? occupiedWidthPx) - occupiedWidthPx
          )
        );
        if (freeSpacePx <= 0) {
          return line;
        }

        const shiftPx =
          normalizedAlignment === "center"
            ? Math.round(freeSpacePx / 2)
            : freeSpacePx;
        if (shiftPx <= 0) {
          return line;
        }

        return {
          ...line,
          fragments: line.fragments.map((fragment) => ({
            ...fragment,
            x: fragment.x + shiftPx,
          })),
        };
      })
    : layout.lines;

  return {
    ...layout,
    lines: alignedLines,
  };
}

function checkboxChoiceRowTabWidthPx(paragraph: ParagraphNode): number {
  const checkboxCount = paragraph.children.filter(
    (child) => child.type === "form-field" && child.fieldType === "checkbox"
  ).length;
  const labelCharCount = paragraph.children
    .filter((child): child is TextRunNode => child.type === "text")
    .map((child) => child.text.replace(/\t/g, ""))
    .join("")
    .replace(/\s+/g, "").length;

  if (checkboxCount >= 3) {
    return 6;
  }

  if (labelCharCount <= 8) {
    return 8;
  }

  return 10;
}

function attachTextToPreviousCheckbox(
  paragraph: ParagraphNode,
  childIndex: number,
  text: string
): string {
  if (!text || text[0] !== " ") {
    return text;
  }

  const previousChild =
    childIndex > 0 ? paragraph.children[childIndex - 1] : undefined;
  if (
    previousChild?.type !== "form-field" ||
    previousChild.fieldType !== "checkbox"
  ) {
    return text;
  }

  return `\u00a0${text.slice(1)}`;
}

function runFontSizePx(
  style?: TextRunNode["style"] | FormFieldRunNode["style"]
): number {
  if (
    style?.fontSizePt &&
    Number.isFinite(style.fontSizePt) &&
    style.fontSizePt > 0
  ) {
    return Math.max(9, (style.fontSizePt * 96) / 72);
  }

  return Math.max(9, (DEFAULT_PARAGRAPH_FONT_SIZE_PT * 96) / 72);
}

function explicitRunFontSizePx(
  style?: TextRunNode["style"] | FormFieldRunNode["style"]
): number | undefined {
  if (
    style?.fontSizePt &&
    Number.isFinite(style.fontSizePt) &&
    style.fontSizePt > 0
  ) {
    return Math.max(9, (style.fontSizePt * 96) / 72);
  }

  return undefined;
}

export function resolveDropCapFontSizePx(
  style: TextRunNode["style"] | undefined,
  lineHeightPx: number,
  lineCount: number,
  previewFontSizePx?: number
): number {
  if (Number.isFinite(previewFontSizePx) && (previewFontSizePx as number) > 0) {
    return Math.max(12, Math.round(previewFontSizePx as number));
  }

  const explicitFontSizePx = explicitRunFontSizePx(style);
  if (
    Number.isFinite(explicitFontSizePx) &&
    (explicitFontSizePx as number) > 0
  ) {
    return Math.max(12, Math.round(explicitFontSizePx as number));
  }

  const baseFontSizePx = runFontSizePx(style);
  const derivedFromLinesPx = Math.max(
    18,
    Math.round(lineHeightPx * Math.max(1.6, lineCount * 0.92))
  );
  return Math.max(derivedFromLinesPx, Math.round(baseFontSizePx * 1.8));
}

export function resolveDropCapVisualHeightPx(
  style: TextRunNode["style"] | undefined,
  lineHeightPx: number,
  lineCount: number,
  currentFontSizePx?: number
): number {
  const baseHeightPx = Math.max(18, Math.round(lineHeightPx * lineCount));
  const baselineFontSizePx = resolveDropCapFontSizePx(
    style,
    lineHeightPx,
    lineCount
  );
  const effectiveFontSizePx =
    Number.isFinite(currentFontSizePx) && (currentFontSizePx as number) > 0
      ? Math.max(12, Math.round(currentFontSizePx as number))
      : baselineFontSizePx;
  const scale =
    baselineFontSizePx > 0 ? effectiveFontSizePx / baselineFontSizePx : 1;

  return Math.max(
    18,
    Math.round(baseHeightPx * scale),
    Math.round(effectiveFontSizePx * 1.05)
  );
}

function estimateTextAdvanceWidthPx(
  text: string,
  style?: TextRunNode["style"] | FormFieldRunNode["style"]
): number {
  if (!text) {
    return 0;
  }

  const fontSizePx = runFontSizePx(style);
  const normalized = text.includes("\u00a0")
    ? text.replace(/\u00a0/g, " ")
    : text;
  const cacheKey = `${fontSizePx}\u0000${normalized}`;
  const cached = estimatedTextAdvanceWidthByFontAndValue.get(cacheKey);
  if (cached !== undefined) {
    return cached;
  }

  let total = 0;

  for (let index = 0; index < normalized.length; index += 1) {
    const code = normalized.charCodeAt(index);
    if (code === 10 || code === 13) {
      continue;
    }
    if (code === 9) {
      total += fontSizePx * 2;
      continue;
    }
    if (code === 32) {
      total += fontSizePx * 0.33;
      continue;
    }
    if (
      code === 33 ||
      code === 39 ||
      code === 44 ||
      code === 46 ||
      code === 49 ||
      code === 58 ||
      code === 59 ||
      code === 73 ||
      code === 96 ||
      code === 105 ||
      code === 108 ||
      code === 124
    ) {
      total += fontSizePx * 0.29;
      continue;
    }
    if ((code >= 65 && code <= 90) || (code >= 48 && code <= 57)) {
      total += fontSizePx * 0.6;
      continue;
    }
    if (code >= 0x3000 && code <= 0x9fff) {
      total += fontSizePx * 0.95;
      continue;
    }
    total += fontSizePx * 0.54;
  }

  const estimatedWidthPx = Math.max(0, Math.round(total));
  setCacheEntry(
    estimatedTextAdvanceWidthByFontAndValue,
    cacheKey,
    estimatedWidthPx
  );
  while (
    estimatedTextAdvanceWidthByFontAndValue.size >
    TEXT_MEASURE_CACHE_MAX_ENTRIES
  ) {
    const firstKey = estimatedTextAdvanceWidthByFontAndValue.keys().next()
      .value as string | undefined;
    if (!firstKey) {
      break;
    }
    estimatedTextAdvanceWidthByFontAndValue.delete(firstKey);
  }

  return estimatedWidthPx;
}

function updateEstimatedLineWidthPxForText(
  currentLineWidthPx: number,
  text: string,
  style?: TextRunNode["style"] | FormFieldRunNode["style"]
): number {
  if (!text) {
    return currentLineWidthPx;
  }

  if (!text.includes("\n")) {
    return currentLineWidthPx + estimateTextAdvanceWidthPx(text, style);
  }

  const segments = text.split("\n");
  const trailingSegment = segments[segments.length - 1] ?? "";
  return estimateTextAdvanceWidthPx(trailingSegment, style);
}

function resolveTabSpacerWidthPx(
  tabStopPositionsPx: number[],
  currentLineWidthPx: number,
  fallbackWidthPx: number,
  fixedFallback = false
): number {
  const safeFallback = Math.max(12, Math.round(fallbackWidthPx));
  if (tabStopPositionsPx.length === 0) {
    if (fixedFallback) {
      return safeFallback;
    }
    // Word's default tab behavior: advance to the next multiple of the
    // default tab stop from the current position (not a fixed-width gap).
    const nextStop =
      (Math.floor((currentLineWidthPx + 0.5) / safeFallback) + 1) *
      safeFallback;
    return Math.max(2, Math.round(nextStop - currentLineWidthPx));
  }

  const nextStop = tabStopPositionsPx.find(
    (stop) => stop > currentLineWidthPx + 0.5
  );
  if (nextStop !== undefined) {
    return Math.max(8, Math.round(nextStop - currentLineWidthPx));
  }

  const lastStop = tabStopPositionsPx[tabStopPositionsPx.length - 1] ?? 0;
  const overflow = Math.max(0, currentLineWidthPx - lastStop);
  const stepCount = Math.floor(overflow / safeFallback) + 1;
  const projectedStop = lastStop + stepCount * safeFallback;
  return Math.max(8, Math.round(projectedStop - currentLineWidthPx));
}

function estimateInteractiveFieldWidthPx(field: FormFieldRunNode): number {
  if (field.fieldType === "checkbox") {
    return resolveCheckboxFieldWidthPx(field);
  }

  if (field.fieldType === "date") {
    const text = field.value?.trim() || "MM/DD/YYYY";
    return Math.max(
      68,
      Math.min(220, estimateTextAdvanceWidthPx(text, field.style) + 12)
    );
  }

  if (field.fieldType === "dropdown") {
    const text =
      field.value?.trim() ||
      field.options?.[0]?.displayText?.trim() ||
      field.title?.trim() ||
      "Select";
    return Math.max(
      58,
      Math.min(240, estimateTextAdvanceWidthPx(text, field.style) + 16)
    );
  }

  const defaultText = field.widget?.text?.defaultText?.trim();
  const textValue =
    field.value?.trim() || defaultText || field.title?.trim() || "Click here.";
  return Math.max(
    42,
    Math.min(280, estimateTextAdvanceWidthPx(textValue, field.style) + 12)
  );
}

export function resolveCheckboxFieldWidthPx(field: FormFieldRunNode): number {
  const exactWidgetWidthPx =
    field.widget?.checkbox?.sizeMode === "exact" &&
    Number.isFinite(field.widget.checkbox.sizePt)
      ? Math.max(
          14,
          Math.round(((field.widget.checkbox.sizePt as number) * 96) / 72 + 4)
        )
      : undefined;
  if (Number.isFinite(exactWidgetWidthPx)) {
    return exactWidgetWidthPx as number;
  }

  const checkedSymbol = field.checkedSymbol ?? "☒";
  const uncheckedSymbol = field.uncheckedSymbol ?? "☐";
  return Math.max(
    14,
    estimateTextAdvanceWidthPx(checkedSymbol, field.style),
    estimateTextAdvanceWidthPx(uncheckedSymbol, field.style)
  );
}

function paragraphIsEffectivelyEmpty(paragraph: ParagraphNode): boolean {
  if (paragraphHasImage(paragraph) || paragraphHasFormField(paragraph)) {
    return false;
  }

  return paragraph.children.every(
    (child) => child.type === "text" && child.text.length === 0
  );
}

function paragraphHasDeletedParagraphMark(paragraph: ParagraphNode): boolean {
  if (paragraph.paragraphMarkDeleted === true) {
    return true;
  }

  const sourceXml = paragraph.sourceXml ?? "";
  return /<w:pPr\b[\s\S]*?<w:rPr\b[\s\S]*?<w:del\b/i.test(sourceXml);
}

function paragraphCollapsesIntoPreviousParagraph(
  paragraph: ParagraphNode,
  previousNode?: DocModel["nodes"][number]
): boolean {
  return (
    paragraphIsEffectivelyEmpty(paragraph) &&
    previousNode?.type === "paragraph" &&
    paragraphHasDeletedParagraphMark(previousNode)
  );
}

function paragraphIsStructuralSectionBreakSpacer(
  paragraph: ParagraphNode
): boolean {
  const sourceXml = paragraph.sourceXml ?? "";
  if (!sourceXml || !SECTION_PROPERTIES_XML_PATTERN.test(sourceXml)) {
    return false;
  }

  if (
    !paragraphIsEffectivelyEmpty(paragraph) ||
    paragraphHasExplicitPageBreak(paragraph) ||
    paragraphHasPageBreakBefore(paragraph) ||
    paragraphStartsWithLastRenderedPageBreak(paragraph)
  ) {
    return false;
  }

  if (paragraphLetterheadSideFromIndent(paragraph)) {
    return false;
  }

  return true;
}

function paragraphActsAsSectionBreakCarryoverSpacer(
  model: DocModel,
  nodeIndex: number,
  paragraph: ParagraphNode
): boolean {
  if (
    !paragraphIsEffectivelyEmpty(paragraph) ||
    paragraphHasExplicitPageBreak(paragraph) ||
    paragraphHasPageBreakBefore(paragraph) ||
    paragraphStartsWithLastRenderedPageBreak(paragraph)
  ) {
    return false;
  }

  const nextNode = model.nodes[nodeIndex + 1];
  if (!nextNode || nextNode.type !== "paragraph") {
    return false;
  }

  return paragraphIsSectionBreakAnchorCarryover(nextNode);
}

function paragraphActsAsTrailingRenderedPageBreakSpacer(
  model: DocModel,
  nodeIndex: number,
  paragraph: ParagraphNode
): boolean {
  if (
    !paragraphIsEffectivelyEmpty(paragraph) ||
    paragraphHasExplicitPageBreak(paragraph) ||
    paragraphHasPageBreakBefore(paragraph) ||
    paragraphStartsWithLastRenderedPageBreak(paragraph) ||
    paragraphHasImage(paragraph) ||
    paragraphHasFormField(paragraph)
  ) {
    return false;
  }

  let lookaheadIndex = nodeIndex + 1;
  while (lookaheadIndex < model.nodes.length) {
    const nextNode = model.nodes[lookaheadIndex];
    if (nextNode?.type !== "paragraph") {
      return false;
    }
    if (!paragraphIsEffectivelyEmpty(nextNode)) {
      return (
        paragraphHasPageBreakBefore(nextNode) ||
        paragraphStartsWithLastRenderedPageBreak(nextNode)
      );
    }
    if (
      paragraphHasExplicitPageBreak(nextNode) ||
      paragraphHasPageBreakBefore(nextNode) ||
      paragraphStartsWithLastRenderedPageBreak(nextNode) ||
      paragraphHasImage(nextNode) ||
      paragraphHasFormField(nextNode)
    ) {
      return false;
    }
    lookaheadIndex += 1;
  }

  return false;
}

function paragraphContextualSpacingStyleKey(
  paragraph: ParagraphNode
): string | undefined {
  const styleId = paragraph.style?.styleId?.trim().toLowerCase();
  if (styleId) {
    return `id:${styleId}`;
  }

  const styleName = paragraph.style?.styleName?.trim().toLowerCase();
  if (styleName) {
    return `name:${styleName}`;
  }

  return undefined;
}

function paragraphsSuppressInterParagraphSpacing(
  previousParagraph: ParagraphNode | undefined,
  currentParagraph: ParagraphNode
): boolean {
  if (!previousParagraph) {
    return false;
  }

  const previousStyleKey =
    paragraphContextualSpacingStyleKey(previousParagraph);
  const currentStyleKey = paragraphContextualSpacingStyleKey(currentParagraph);
  if (!previousStyleKey || previousStyleKey !== currentStyleKey) {
    return false;
  }

  return (
    previousParagraph.style?.contextualSpacing === true ||
    currentParagraph.style?.contextualSpacing === true
  );
}

function effectiveParagraphAfterSpacingPx(
  model: DocModel,
  nodeIndex: number,
  paragraph: ParagraphNode
): number {
  const afterSpacingPx = paragraphAfterSpacingPx(paragraph);
  const nextNode = model.nodes[nodeIndex + 1];
  if (nextNode?.type !== "paragraph") {
    return afterSpacingPx;
  }

  return paragraphsSuppressInterParagraphSpacing(paragraph, nextNode)
    ? 0
    : afterSpacingPx;
}

function effectiveParagraphBeforeSpacingPx(
  model: DocModel,
  nodeIndex: number,
  paragraph: ParagraphNode,
  pageConsumedHeightPx: number,
  suppressSpacingBeforeAfterPageBreak: boolean
): number {
  let beforeSpacingPx = resolveParagraphBeforeSpacingPx(
    model,
    nodeIndex,
    paragraph,
    pageConsumedHeightPx,
    suppressSpacingBeforeAfterPageBreak
  );

  if (pageConsumedHeightPx <= 0) {
    return beforeSpacingPx;
  }

  const previousNode = model.nodes[nodeIndex - 1];
  if (previousNode?.type !== "paragraph") {
    return beforeSpacingPx;
  }

  return paragraphsSuppressInterParagraphSpacing(previousNode, paragraph)
    ? 0
    : beforeSpacingPx;
}

function nodeHasSubstantiveContentForPagination(
  node: DocModel["nodes"][number]
): boolean {
  if (node.type === "table") {
    return true;
  }

  if (
    paragraphHasVisibleText(node) ||
    paragraphHasImage(node) ||
    paragraphHasFormField(node)
  ) {
    return true;
  }

  return (
    paragraphHasPageBreakBefore(node) ||
    paragraphHasExplicitPageBreak(node) ||
    sectionBreakAfterParagraphStartsNewPage(node)
  );
}

function paragraphBookmarkNames(paragraph: ParagraphNode): string[] {
  const sourceXml = paragraph.sourceXml ?? "";
  if (!sourceXml) {
    return [];
  }

  const names = [
    ...sourceXml.matchAll(new RegExp(BOOKMARK_START_XML_PATTERN.source, "gi")),
  ]
    .map((match) => match[1]?.trim())
    .filter((name): name is string =>
      Boolean(name && name.length > 0 && name !== "_GoBack")
    );
  return [...new Set(names)];
}

function paragraphReferencedNoteIds(
  paragraph: ParagraphNode,
  noteType: "footnote" | "endnote"
): number[] {
  const sourceXml = paragraph.sourceXml ?? "";
  if (!sourceXml) {
    return [];
  }

  const pattern =
    noteType === "footnote"
      ? new RegExp(FOOTNOTE_REFERENCE_XML_PATTERN.source, "gi")
      : new RegExp(ENDNOTE_REFERENCE_XML_PATTERN.source, "gi");
  const references: number[] = [];

  for (const match of sourceXml.matchAll(pattern)) {
    const rawId = Number(match[1]);
    if (!Number.isFinite(rawId) || rawId < 0) {
      continue;
    }
    references.push(Math.round(rawId));
  }

  return references;
}

function nodeReferencedNoteIds(
  node: DocModel["nodes"][number],
  noteType: "footnote" | "endnote",
  tableRowRange?: TableRowRange,
  paragraphLineRange?: ParagraphLineRange
): number[] {
  if (node.type === "paragraph") {
    if (paragraphLineRange && paragraphLineRange.startLineIndex > 0) {
      return [];
    }
    return paragraphReferencedNoteIds(node, noteType);
  }

  const references: number[] = [];
  const startRowIndex = Math.max(0, tableRowRange?.startRowIndex ?? 0);
  const endRowIndex = Math.min(
    node.rows.length,
    tableRowRange?.endRowIndex ?? node.rows.length
  );
  for (let rowIndex = startRowIndex; rowIndex < endRowIndex; rowIndex += 1) {
    const row = node.rows[rowIndex];
    row?.cells.forEach((cell) => {
      tableCellParagraphsRecursively(cell.nodes).forEach((paragraph) => {
        references.push(...paragraphReferencedNoteIds(paragraph, noteType));
      });
    });
  }
  return references;
}

function eventTargetIsInteractiveControl(target: EventTarget | null): boolean {
  if (!(target instanceof Element)) {
    return false;
  }

  return Boolean(
    target.closest("a[href],button,input,select,textarea,[role='checkbox']")
  );
}

function eventTargetIsNestedTableParagraphEditor(
  target: EventTarget | null
): boolean {
  if (!(target instanceof Element)) {
    return false;
  }

  return Boolean(
    target.closest("[data-docx-table-cell-paragraph-host='true']")
  );
}

function sectionBreakPropertiesStartNewPage(
  sectionPropertiesXml: string
): boolean {
  const sectionType =
    sectionPropertiesXml
      .match(SECTION_TYPE_XML_PATTERN)?.[1]
      ?.trim()
      .toLowerCase() ?? "nextpage";

  if (sectionType === "continuous") {
    return false;
  }

  if (sectionType === "nextcolumn") {
    const columnsTag = sectionPropertiesXml.match(/<w:cols\b[^>]*\/?>/i)?.[0];
    const columnsCount = Number.parseInt(
      columnsTag?.match(/\bw:num="(\d+)"/i)?.[1] ?? "",
      10
    );
    // In single-column sections, "nextColumn" effectively behaves like "nextPage".
    return !Number.isFinite(columnsCount) || columnsCount <= 1;
  }

  return true;
}

function paragraphHasExplicitPageBreak(paragraph: ParagraphNode): boolean {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return false;
  }

  const cached = paragraphBreakFlagsBySourceXml.get(xml);
  if (cached) {
    return cached.explicitPageBreak;
  }

  const flags = {
    explicitPageBreak: PAGE_BREAK_XML_PATTERN.test(xml),
    explicitColumnBreak: COLUMN_BREAK_XML_PATTERN.test(xml),
    lastRenderedPageBreak: LAST_RENDERED_PAGE_BREAK_XML_PATTERN.test(xml),
    pageBreakBefore: isOnOffTagEnabled(
      xml.match(PAGE_BREAK_BEFORE_XML_PATTERN)?.[0]
    ),
    sectionBreakStartsNewPage: (() => {
      const sectionProperties = xml.match(SECTION_PROPERTIES_XML_PATTERN)?.[0];
      if (!sectionProperties) {
        return false;
      }
      // ECMA-376 §2.6.22: omitted <w:type> defaults to nextPage.
      return sectionBreakPropertiesStartNewPage(sectionProperties);
    })(),
  };
  setCacheEntry(paragraphBreakFlagsBySourceXml, xml, flags);
  return flags.explicitPageBreak;
}

function paragraphHasExplicitColumnBreak(paragraph: ParagraphNode): boolean {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return false;
  }

  const cached = paragraphBreakFlagsBySourceXml.get(xml);
  if (cached) {
    return cached.explicitColumnBreak;
  }

  paragraphHasExplicitPageBreak(paragraph);
  return paragraphBreakFlagsBySourceXml.get(xml)?.explicitColumnBreak ?? false;
}

function paragraphHasLastRenderedPageBreak(paragraph: ParagraphNode): boolean {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return false;
  }

  const cached = paragraphBreakFlagsBySourceXml.get(xml);
  if (cached) {
    return cached.lastRenderedPageBreak;
  }

  paragraphHasExplicitPageBreak(paragraph);
  return (
    paragraphBreakFlagsBySourceXml.get(xml)?.lastRenderedPageBreak ?? false
  );
}

function paragraphStartsWithLastRenderedPageBreak(
  paragraph: ParagraphNode
): boolean {
  const xml = paragraph.sourceXml ?? "";
  if (!xml || !paragraphHasLastRenderedPageBreak(paragraph)) {
    return false;
  }

  const breakMatch = xml.match(LAST_RENDERED_PAGE_BREAK_XML_PATTERN);
  if (!breakMatch || breakMatch.index === undefined) {
    return false;
  }

  const leadingXml = xml
    .slice(0, breakMatch.index)
    .replace(/^<w:p\b[^>]*>/i, "")
    .replace(/<w:pPr\b(?:[^/>]*\/>|[\s\S]*?<\/w:pPr>)/i, "")
    .replace(/<w:rPr\b[\s\S]*?<\/w:rPr>/gi, "")
    .replace(/<\/?w:r\b[^>]*>/gi, "")
    .replace(
      /<w:(?:proofErr|bookmarkStart|bookmarkEnd|permStart|permEnd)\b[^>]*\/?>/gi,
      ""
    )
    .replace(/<\/?w:(?:ins|smartTag)\b[^>]*>/gi, "")
    .replace(/\s+/g, "");

  return leadingXml.length === 0;
}

function shouldHonorParagraphStartLastRenderedPageBreak(params: {
  pageConsumedHeightPx: number;
  pageContentHeightPx: number;
}): boolean {
  const pageConsumedHeightPx = Math.max(
    0,
    Math.round(params.pageConsumedHeightPx)
  );
  const pageContentHeightPx = Math.max(
    0,
    Math.round(params.pageContentHeightPx)
  );
  if (pageConsumedHeightPx <= 0 || pageContentHeightPx <= 0) {
    return false;
  }

  const remainingHeightPx = Math.max(
    0,
    pageContentHeightPx - pageConsumedHeightPx
  );
  const maxAllowedRemainingHeightPx = clampNumber(
    Math.round(
      pageContentHeightPx *
        LAST_RENDERED_PAGE_BREAK_HINT_MAX_REMAINING_SPACE_RATIO
    ),
    LAST_RENDERED_PAGE_BREAK_HINT_MIN_REMAINING_SPACE_PX,
    LAST_RENDERED_PAGE_BREAK_HINT_MAX_REMAINING_SPACE_PX
  );
  return remainingHeightPx <= maxAllowedRemainingHeightPx;
}

function isOnOffTagEnabled(tagXml: string | undefined): boolean {
  if (!tagXml) {
    return false;
  }

  const valueMatch = tagXml
    .match(/\bw:val="([^"]+)"/i)?.[1]
    ?.trim()
    .toLowerCase();
  if (!valueMatch) {
    return true;
  }

  return !["0", "false", "off", "no"].includes(valueMatch);
}

function sectionTitlePageEnabled(sectionPropertiesXml?: string): boolean {
  if (!sectionPropertiesXml) {
    return false;
  }

  const titlePageTag = sectionPropertiesXml.match(
    /<w:titlePg\b[^>]*\/?>/i
  )?.[0];
  return isOnOffTagEnabled(titlePageTag);
}

function selectSectionVariantForPage<T extends HeaderSection | FooterSection>(
  sections: T[],
  sectionPropertiesXml: string | undefined,
  pageIndex: number,
  options?: {
    evenAndOddHeaders?: boolean;
  }
): T | undefined {
  if (sections.length === 0) {
    return undefined;
  }

  const titlePage = sectionTitlePageEnabled(sectionPropertiesXml);
  const normalizeType = (value: string | undefined): string =>
    value?.trim().toLowerCase() ?? "";
  const first = sections.find(
    (section) => normalizeType(section.referenceType) === "first"
  );
  const defaultSection = sections.find((section) => {
    const referenceType = normalizeType(section.referenceType);
    return referenceType === "default" || referenceType === "";
  });
  const even = sections.find(
    (section) => normalizeType(section.referenceType) === "even"
  );
  const evenAndOddHeadersEnabled = options?.evenAndOddHeaders ?? true;

  const safePageIndex = Number.isFinite(pageIndex)
    ? Math.max(0, Math.round(pageIndex))
    : 0;
  const oddPageNumber = safePageIndex % 2 === 0;

  if (safePageIndex === 0 && titlePage) {
    return first;
  }

  if (evenAndOddHeadersEnabled && !oddPageNumber && even) {
    return even;
  }

  if (defaultSection) {
    return defaultSection;
  }

  return first ?? even ?? sections[0];
}

function resolveSectionIndexForNodeIndex(
  sections: ResolvedDocumentSection[],
  nodeIndex: number,
  previousSectionIndex: number
): number {
  if (sections.length === 0) {
    return 0;
  }

  const safePrevious = Math.max(
    0,
    Math.min(previousSectionIndex, sections.length - 1)
  );
  let sectionIndex = safePrevious;

  if (nodeIndex < sections[sectionIndex].startNodeIndex) {
    sectionIndex = 0;
  }

  while (
    sectionIndex + 1 < sections.length &&
    sections[sectionIndex + 1].startNodeIndex <= nodeIndex
  ) {
    sectionIndex += 1;
  }

  return sectionIndex;
}

function paragraphHasVisibleBorder(paragraph: ParagraphNode): boolean {
  return (
    paragraphBorderVisible(paragraph.style?.borders?.top) ||
    paragraphBorderVisible(paragraph.style?.borders?.right) ||
    paragraphBorderVisible(paragraph.style?.borders?.bottom) ||
    paragraphBorderVisible(paragraph.style?.borders?.left) ||
    paragraphBorderVisible(paragraph.style?.borders?.between) ||
    paragraphBorderVisible(paragraph.style?.borders?.bar)
  );
}

function paragraphIsSectionBreakAnchorCarryover(
  paragraph: ParagraphNode
): boolean {
  if (!paragraph.sourceXml?.includes("<w:sectPr")) {
    return false;
  }

  // Section-break anchor paragraphs can legitimately carry a text-bearing
  // synthetic textbox duplicate used by Word's floating-shape anchoring.
  // Treat these as carryover anchors so they don't consume a standalone page.
  if (!paragraphIsAbsoluteFloatingImageAnchorOnly(paragraph)) {
    return false;
  }

  if (paragraphHasVisibleText(paragraph) || paragraphHasFormField(paragraph)) {
    return false;
  }

  return !paragraphHasVisibleBorder(paragraph);
}

function resolveSectionIndexForPageSegments(
  sections: ResolvedDocumentSection[],
  nodes: DocModel["nodes"],
  pageSegments: DocumentPageNodeSegment[],
  previousSectionIndex: number
): number {
  if (sections.length === 0 || pageSegments.length === 0) {
    return 0;
  }

  const firstNodeIndex = pageSegments[0]?.nodeIndex;
  let sectionIndex = Number.isFinite(firstNodeIndex)
    ? resolveSectionIndexForNodeIndex(
        sections,
        firstNodeIndex as number,
        previousSectionIndex
      )
    : Math.max(0, Math.min(previousSectionIndex, sections.length - 1));
  sectionIndex = Math.max(
    sectionIndex,
    Math.max(0, Math.min(previousSectionIndex, sections.length - 1))
  );

  let walkingSectionIndex = sectionIndex;
  for (const segment of pageSegments) {
    walkingSectionIndex = resolveSectionIndexForNodeIndex(
      sections,
      segment.nodeIndex,
      walkingSectionIndex
    );
    if (walkingSectionIndex <= sectionIndex) {
      continue;
    }

    const nextSection = sections[walkingSectionIndex];
    if (!nextSection) {
      continue;
    }

    if (
      parseSectionStartType(nextSection.sectionPropertiesXml) === "continuous"
    ) {
      continue;
    }

    const candidateNode = nodes[segment.nodeIndex];
    if (
      candidateNode?.type === "paragraph" &&
      paragraphIsSectionBreakAnchorCarryover(candidateNode)
    ) {
      continue;
    }

    sectionIndex = walkingSectionIndex;
  }

  return sectionIndex;
}

function paragraphHasPageBreakBefore(paragraph: ParagraphNode): boolean {
  if (paragraph.style?.pageBreakBefore === true) {
    return true;
  }

  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return false;
  }

  const cached = paragraphBreakFlagsBySourceXml.get(xml);
  if (cached) {
    return cached.pageBreakBefore;
  }

  paragraphHasExplicitPageBreak(paragraph);
  return paragraphBreakFlagsBySourceXml.get(xml)?.pageBreakBefore ?? false;
}

function sectionBreakAfterParagraphStartsNewPage(
  paragraph: ParagraphNode
): boolean {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return false;
  }

  const cached = paragraphBreakFlagsBySourceXml.get(xml);
  if (cached) {
    return cached.sectionBreakStartsNewPage;
  }

  // `paragraphHasExplicitPageBreak` populates the shared cache with every
  // break-related flag, including `sectionBreakStartsNewPage`. The cache
  // read below has to happen unconditionally — a section break paragraph
  // is allowed to start a new page (per ECMA-376 §2.6.22) even when it
  // carries no explicit `<w:br w:type="page"/>` run, and an earlier
  // version of this helper returned `false` in exactly that case.
  paragraphHasExplicitPageBreak(paragraph);
  return (
    paragraphBreakFlagsBySourceXml.get(xml)?.sectionBreakStartsNewPage ?? false
  );
}

function nodeAlreadyEndsAtExplicitPageBoundary(
  node: DocModel["nodes"][number] | undefined
): boolean {
  if (!node || node.type !== "paragraph") {
    return false;
  }

  return (
    paragraphHasExplicitPageBreak(node) ||
    paragraphHasPageBreakBefore(node) ||
    sectionBreakAfterParagraphStartsNewPage(node)
  );
}

const docxHardPageBreakStartNodeIndexesByModel = new WeakMap<
  DocModel,
  Set<number>
>();

function collectDocxHardPageBreakStartNodeIndexes(
  model: DocModel
): Set<number> {
  const cached = docxHardPageBreakStartNodeIndexesByModel.get(model);
  if (cached) {
    return cached;
  }
  const result = computeDocxHardPageBreakStartNodeIndexes(model);
  docxHardPageBreakStartNodeIndexesByModel.set(model, result);
  return result;
}

function computeDocxHardPageBreakStartNodeIndexes(
  model: DocModel
): Set<number> {
  const breaks = collectTopLevelExplicitPageBreakStartNodeIndexes(model.nodes);

  const sections = resolveDocumentSectionsFromMetadata(model.metadata);
  for (
    let sectionIndex = 1;
    sectionIndex < sections.length;
    sectionIndex += 1
  ) {
    const section = sections[sectionIndex];
    const startNodeIndex = Math.max(0, Math.round(section.startNodeIndex));
    if (startNodeIndex <= 0 || startNodeIndex >= model.nodes.length) {
      continue;
    }

    const sectionPropertiesXml = section.sectionPropertiesXml;
    if (!sectionPropertiesXml) {
      continue;
    }

    // Section break pagination should follow the section that starts at this
    // node index, which is represented by metadata.sections.
    if (sectionBreakPropertiesStartNewPage(sectionPropertiesXml)) {
      breaks.add(startNodeIndex);
    }
  }

  for (const breakIndex of [...breaks]) {
    if (breakIndex <= 0 || breakIndex >= model.nodes.length) {
      breaks.delete(breakIndex);
    }
  }

  return breaks;
}

const docxSectionStartPageBreakNodeIndexesByModel = new WeakMap<
  DocModel,
  Set<number>
>();

function collectDocxSectionStartPageBreakNodeIndexes(
  model: DocModel
): Set<number> {
  const cached = docxSectionStartPageBreakNodeIndexesByModel.get(model);
  if (cached) {
    return cached;
  }
  const result = computeDocxSectionStartPageBreakNodeIndexes(model);
  docxSectionStartPageBreakNodeIndexesByModel.set(model, result);
  return result;
}

function computeDocxSectionStartPageBreakNodeIndexes(
  model: DocModel
): Set<number> {
  const breaks = new Set<number>();
  const sections = resolveDocumentSectionsFromMetadata(model.metadata);
  for (
    let sectionIndex = 1;
    sectionIndex < sections.length;
    sectionIndex += 1
  ) {
    const section = sections[sectionIndex];
    const startNodeIndex = Math.max(0, Math.round(section.startNodeIndex));
    if (startNodeIndex <= 0 || startNodeIndex >= model.nodes.length) {
      continue;
    }

    const sectionPropertiesXml = section.sectionPropertiesXml;
    if (!sectionPropertiesXml) {
      continue;
    }

    if (sectionBreakPropertiesStartNewPage(sectionPropertiesXml)) {
      breaks.add(startNodeIndex);
    }
  }

  return breaks;
}

function buildNextHardBreakStartNodeIndexLookup(
  nodeCount: number,
  hardBreakStartNodeIndexes: Set<number>
): number[] {
  const nextBreakStartNodeIndexes = new Array<number>(nodeCount).fill(-1);
  let nextBreakStartNodeIndex = -1;

  for (let nodeIndex = nodeCount - 1; nodeIndex >= 0; nodeIndex -= 1) {
    nextBreakStartNodeIndexes[nodeIndex] = nextBreakStartNodeIndex;
    if (hardBreakStartNodeIndexes.has(nodeIndex)) {
      nextBreakStartNodeIndex = nodeIndex;
    }
  }

  return nextBreakStartNodeIndexes;
}

function paragraphIsSimpleTrailingSectionTailCandidate(
  paragraph: ParagraphNode
): boolean {
  if (
    !paragraphHasVisibleText(paragraph) ||
    paragraphHasImage(paragraph) ||
    paragraphHasFormField(paragraph) ||
    paragraphHasExplicitPageBreak(paragraph) ||
    paragraphHasPageBreakBefore(paragraph) ||
    paragraphStartsWithLastRenderedPageBreak(paragraph) ||
    sectionBreakAfterParagraphStartsNewPage(paragraph)
  ) {
    return false;
  }

  return true;
}

function shouldKeepTrailingSectionTailOnCurrentPage(
  model: DocModel,
  startNodeIndex: number,
  pageConsumedHeightPx: number,
  previousParagraphAfterPx: number,
  pageContentWidthPx: number,
  pageContentHeightPx: number,
  hardBreakStartNodeIndexes: Set<number>,
  sectionStartPageBreakNodeIndexes: Set<number>,
  nextHardBreakStartNodeIndexByNodeIndex: number[],
  estimateNodeHeightPx: (
    nodeIndex: number,
    node: DocModel["nodes"][number],
    pageContentWidthPx: number
  ) => number,
  resolveNodeBeforeSpacingPx: (
    nodeIndex: number,
    paragraph: ParagraphNode,
    consumedHeightPx: number
  ) => number,
  resolveNodeAfterSpacingPx: (
    nodeIndex: number,
    paragraph: ParagraphNode
  ) => number
): boolean {
  if (pageConsumedHeightPx <= 0 || startNodeIndex <= 0) {
    return false;
  }

  const nextHardBreakStartNodeIndex =
    nextHardBreakStartNodeIndexByNodeIndex[startNodeIndex] ?? -1;
  if (
    nextHardBreakStartNodeIndex <= startNodeIndex ||
    !hardBreakStartNodeIndexes.has(nextHardBreakStartNodeIndex) ||
    !sectionStartPageBreakNodeIndexes.has(nextHardBreakStartNodeIndex)
  ) {
    return false;
  }

  const remainingHeightPx = pageContentHeightPx - pageConsumedHeightPx;

  let tailConsumedHeightPx = 0;
  let tailPreviousParagraphAfterPx = previousParagraphAfterPx;
  let substantiveParagraphCount = 0;

  for (
    let nodeIndex = startNodeIndex;
    nodeIndex < nextHardBreakStartNodeIndex;
    nodeIndex += 1
  ) {
    const node = model.nodes[nodeIndex];
    if (
      node.type === "paragraph" &&
      paragraphIsStructuralSectionBreakSpacer(node)
    ) {
      tailPreviousParagraphAfterPx = 0;
      continue;
    }
    if (
      node.type === "paragraph" &&
      paragraphActsAsLeadingCoverLayoutOverlay(
        model,
        nodeIndex,
        node,
        pageContentWidthPx,
        pageContentHeightPx
      )
    ) {
      return false;
    }
    if (
      node.type === "paragraph" &&
      paragraphActsAsDecorativeBehindTextBackgroundOverlay(node)
    ) {
      return false;
    }
    if (
      node.type === "paragraph" &&
      paragraphCollapsesIntoPreviousParagraph(node, model.nodes[nodeIndex - 1])
    ) {
      continue;
    }
    if (node.type !== "paragraph") {
      return false;
    }
    if (!paragraphIsSimpleTrailingSectionTailCandidate(node)) {
      return false;
    }

    substantiveParagraphCount += 1;
    if (substantiveParagraphCount > MAX_TRAILING_SECTION_TAIL_PARAGRAPHS) {
      return false;
    }

    const directBeforeSpacingPx = paragraphBeforeSpacingPx(node);
    const directAfterSpacingPx = paragraphAfterSpacingPx(node);
    const nodeBeforeSpacingPx = resolveNodeBeforeSpacingPx(
      nodeIndex,
      node,
      pageConsumedHeightPx + tailConsumedHeightPx
    );
    const nodeAfterSpacingPx = resolveNodeAfterSpacingPx(nodeIndex, node);
    const rawNodeHeightPx = Math.max(
      1,
      estimateNodeHeightPx(nodeIndex, node, pageContentWidthPx) -
        directBeforeSpacingPx -
        directAfterSpacingPx +
        nodeBeforeSpacingPx +
        nodeAfterSpacingPx
    );
    const collapsedMarginPx =
      pageConsumedHeightPx + tailConsumedHeightPx > 0
        ? Math.min(tailPreviousParagraphAfterPx, nodeBeforeSpacingPx)
        : 0;
    const effectiveNodeHeightPx = Math.max(
      1,
      rawNodeHeightPx - collapsedMarginPx
    );
    tailConsumedHeightPx += effectiveNodeHeightPx;
    tailPreviousParagraphAfterPx = nodeAfterSpacingPx;
  }

  if (substantiveParagraphCount === 0) {
    return false;
  }

  const overflowPx = tailConsumedHeightPx - remainingHeightPx;
  return (
    overflowPx > PAGE_OVERFLOW_TOLERANCE_PX &&
    overflowPx <= MAX_TRAILING_SECTION_TAIL_OVERFLOW_PX
  );
}

function paragraphDominantFontSizePt(
  paragraph: ParagraphNode
): number | undefined {
  const weightByFontSizePt = new Map<number, number>();

  const addWeight = (fontSizePt: number | undefined, weight: number): void => {
    if (
      !Number.isFinite(fontSizePt) ||
      (fontSizePt as number) <= 0 ||
      weight <= 0
    ) {
      return;
    }

    const key = Number((fontSizePt as number).toFixed(2));
    weightByFontSizePt.set(key, (weightByFontSizePt.get(key) ?? 0) + weight);
  };

  paragraph.children.forEach((child) => {
    if (child.type !== "text" && child.type !== "form-field") {
      return;
    }

    const text =
      child.type === "text"
        ? child.text.replace(/\u2063/g, "")
        : formFieldDisplayValue(child);
    const textWeight = Math.max(1, text.length);
    addWeight(child.style?.fontSizePt, textWeight);
  });

  if (weightByFontSizePt.size === 0) {
    return undefined;
  }

  let dominantFontSizePt: number | undefined;
  let dominantWeight = -1;

  for (const [fontSizePt, weight] of weightByFontSizePt) {
    if (
      weight > dominantWeight ||
      (weight === dominantWeight &&
        (dominantFontSizePt === undefined || fontSizePt > dominantFontSizePt))
    ) {
      dominantWeight = weight;
      dominantFontSizePt = fontSizePt;
    }
  }

  return dominantFontSizePt;
}

function paragraphBaseFontSizePx(paragraph: ParagraphNode): number {
  const cached = paragraphBaseFontSizePxByParagraph.get(paragraph);
  if (cached !== undefined) {
    return cached;
  }

  const dominantRunFontSizePt = paragraphDominantFontSizePt(paragraph);
  const fontSizePt =
    dominantRunFontSizePt && dominantRunFontSizePt > 0
      ? dominantRunFontSizePt
      : Number.isFinite(paragraph.style?.headingLevel)
      ? DEFAULT_PARAGRAPH_FONT_SIZE_PT +
        Math.max(0, 6 - (paragraph.style?.headingLevel ?? 6))
      : DEFAULT_PARAGRAPH_FONT_SIZE_PT;

  const baseFontSizePx = Math.max(10, Math.round((fontSizePt * 96) / 72));
  paragraphBaseFontSizePxByParagraph.set(paragraph, baseFontSizePx);
  return baseFontSizePx;
}

function paragraphMaxFontSizePx(paragraph: ParagraphNode): number {
  const paragraphBaseFontPx = paragraphBaseFontSizePx(paragraph);
  let maxFontSizePx = paragraphBaseFontPx;

  paragraph.children.forEach((child) => {
    if (child.type !== "text" && child.type !== "form-field") {
      return;
    }
    maxFontSizePx = Math.max(
      maxFontSizePx,
      resolveMeasureFontSizePx(child.style, paragraphBaseFontPx)
    );
  });

  return Math.max(1, Math.round(maxFontSizePx));
}

function normalizeFontFamilyToken(fontFamily?: string): string | undefined {
  if (!fontFamily) {
    return undefined;
  }

  const [firstToken] = fontFamily.split(",");
  const normalized = firstToken
    ?.trim()
    .replace(/^['"]+|['"]+$/g, "")
    .toLowerCase();
  return normalized || undefined;
}

function paragraphDominantFontFamily(
  paragraph: ParagraphNode
): string | undefined {
  const cached = paragraphDominantFontFamilyByParagraph.get(paragraph);
  if (cached !== undefined) {
    return cached ?? undefined;
  }

  const weightByFamily = new Map<string, number>();
  const addWeight = (fontFamily: string | undefined, weight: number): void => {
    const normalizedFamily = normalizeFontFamilyToken(fontFamily);
    if (!normalizedFamily) {
      return;
    }

    const safeWeight = Math.max(1, Math.round(weight));
    weightByFamily.set(
      normalizedFamily,
      (weightByFamily.get(normalizedFamily) ?? 0) + safeWeight
    );
  };

  paragraph.children.forEach((child) => {
    if (child.type !== "text" && child.type !== "form-field") {
      return;
    }

    const text =
      child.type === "text"
        ? child.text.replace(/\u2063/g, "")
        : formFieldDisplayValue(child);
    addWeight(child.style?.fontFamily, text.length);
  });

  if (weightByFamily.size === 0) {
    paragraphDominantFontFamilyByParagraph.set(paragraph, null);
    return undefined;
  }

  let dominantFamily: string | undefined;
  let dominantWeight = -1;
  for (const [family, weight] of weightByFamily) {
    if (weight > dominantWeight) {
      dominantFamily = family;
      dominantWeight = weight;
    }
  }

  paragraphDominantFontFamilyByParagraph.set(paragraph, dominantFamily ?? null);
  return dominantFamily;
}

function singleLineAutoScaleForFontFamily(fontFamily?: string): number {
  const normalized =
    normalizeFontFamilyToken(fontFamily) ?? fontFamily?.toLowerCase();
  if (!normalized) {
    return WORD_SINGLE_LINE_AUTO_SCALE;
  }

  if (
    normalized === "times roman" ||
    normalized === "times new roman" ||
    normalized === "cambria" ||
    normalized === "garamond" ||
    normalized === "georgia" ||
    normalized === "book antiqua" ||
    normalized === "palatino linotype"
  ) {
    return WORD_SINGLE_LINE_AUTO_SCALE_SERIF;
  }

  if (normalized === "arial") {
    return WORD_SINGLE_LINE_AUTO_SCALE_SANS;
  }

  return WORD_SINGLE_LINE_AUTO_SCALE;
}

function emptyParagraphLineScaleForFontFamily(fontFamily?: string): number {
  const normalized =
    normalizeFontFamilyToken(fontFamily) ?? fontFamily?.toLowerCase();
  if (!normalized) {
    return WORD_EMPTY_PARAGRAPH_LINE_SCALE;
  }

  if (
    normalized === "times roman" ||
    normalized === "times new roman" ||
    normalized === "cambria" ||
    normalized === "garamond" ||
    normalized === "georgia" ||
    normalized === "book antiqua" ||
    normalized === "palatino linotype"
  ) {
    return WORD_EMPTY_PARAGRAPH_LINE_SCALE_SERIF;
  }

  if (normalized === "arial") {
    return WORD_EMPTY_PARAGRAPH_LINE_SCALE_SANS;
  }

  return WORD_EMPTY_PARAGRAPH_LINE_SCALE;
}

// A paragraph whose only content is whitespace and/or floating anchors lays
// out as one empty line at the mark font's natural metrics — floating
// objects never contribute to the line box.
function paragraphRendersTextFreeLine(paragraph: ParagraphNode): boolean {
  return (
    paragraphHasOnlyWhitespaceText(paragraph) ||
    paragraphIsFloatingImageAnchorOnly(paragraph)
  );
}

function resolveParagraphSingleLineAutoScale(
  paragraph: ParagraphNode,
  fontFamily?: string
): number {
  if (paragraphRendersTextFreeLine(paragraph)) {
    return emptyParagraphLineScaleForFontFamily(fontFamily);
  }

  const baseScale = singleLineAutoScaleForFontFamily(fontFamily);
  return paragraphHasCheckboxFormField(paragraph)
    ? Math.max(1.08, baseScale)
    : baseScale;
}

function paragraphLineCount(paragraph: ParagraphNode): number {
  return paragraphLineCountWithinWidth(paragraph);
}

function paragraphContainsExplicitLineBreakText(
  paragraph: ParagraphNode
): boolean {
  return paragraph.children.some(
    (child) => child.type === "text" && /[\r\n]/.test(child.text)
  );
}

function estimatedGlyphWidthPx(character: string, fontSizePx: number): number {
  if (/\s/.test(character)) {
    return fontSizePx * 0.34;
  }
  if (/[A-Z]/.test(character)) {
    return fontSizePx * 0.66;
  }
  if (/[a-z]/.test(character)) {
    return fontSizePx * 0.55;
  }
  if (/[0-9]/.test(character)) {
    return fontSizePx * 0.57;
  }
  if (/[\u2e80-\u9fff\u3040-\u30ff\uac00-\ud7af]/.test(character)) {
    return fontSizePx;
  }

  return fontSizePx * 0.48;
}

function fallbackMeasureTextWidthPx(text: string, fontSizePx: number): number {
  let widthPx = 0;
  for (const character of text) {
    widthPx += estimatedGlyphWidthPx(character, fontSizePx);
  }
  return widthPx;
}

function resolveMeasureFontSizePx(
  style: TextRunNode["style"] | FormFieldRunNode["style"] | undefined,
  paragraphBaseFontPx: number
): number {
  const runFontSizePx =
    Number.isFinite(style?.fontSizePt) && (style?.fontSizePt as number) > 0
      ? ((style?.fontSizePt as number) * 96) / 72
      : paragraphBaseFontPx;
  const verticalAlignScale =
    style?.verticalAlign === "superscript" ||
    style?.verticalAlign === "subscript"
      ? SCRIPT_FONT_SCALE
      : 1;
  return Math.max(8, Math.round(runFontSizePx * verticalAlignScale));
}

function resolveMeasureFont(
  style: TextRunNode["style"] | FormFieldRunNode["style"] | undefined,
  paragraphBaseFontPx: number
): string {
  const fontStyle = style?.italic ? "italic" : "normal";
  const fontWeight = style?.bold ? "700" : "400";
  const fontSizePx = resolveMeasureFontSizePx(style, paragraphBaseFontPx);
  const fontFamily =
    cssFontFamily(style?.fontFamily) ??
    cssFontFamily(DEFAULT_UNSPECIFIED_DOCX_FONT_FAMILY) ??
    '"Times New Roman", serif';
  return `${fontStyle} normal ${fontWeight} ${fontSizePx}px ${fontFamily}`;
}

function measureTextWidthPx(
  text: string,
  style: TextRunNode["style"] | FormFieldRunNode["style"] | undefined,
  paragraphBaseFontPx: number
): number {
  if (!text) {
    return 0;
  }

  const font = resolveMeasureFont(style, paragraphBaseFontPx);
  const characterSpacingPx = Number.isFinite(style?.characterSpacingTwips)
    ? (style?.characterSpacingTwips as number) / TWIPS_PER_PIXEL
    : 0;
  const cacheKey = `${font}\u0000${characterSpacingPx}\u0000${text}`;
  const cached = textWidthByFontAndValue.get(cacheKey);
  if (cached !== undefined) {
    return cached;
  }

  const fontSizePx = resolveMeasureFontSizePx(style, paragraphBaseFontPx);
  let measuredWidthPx = fallbackMeasureTextWidthPx(text, fontSizePx);

  if (typeof document !== "undefined") {
    try {
      if (!paragraphMeasureCanvasContext) {
        const canvas = document.createElement("canvas");
        paragraphMeasureCanvasContext = canvas.getContext("2d") ?? undefined;
      }
      const context = paragraphMeasureCanvasContext;
      if (context) {
        context.font = font;
        measuredWidthPx = context.measureText(text).width;
      }
    } catch {
      // Keep deterministic fallback width when canvas measurement is unavailable.
    }
  }

  if (characterSpacingPx !== 0 && text.length > 1) {
    measuredWidthPx += characterSpacingPx * Math.max(0, text.length - 1);
  }

  setCacheEntry(textWidthByFontAndValue, cacheKey, measuredWidthPx);
  while (textWidthByFontAndValue.size > TEXT_MEASURE_CACHE_MAX_ENTRIES) {
    const firstKey = textWidthByFontAndValue.keys().next().value as
      | string
      | undefined;
    if (!firstKey) {
      break;
    }
    textWidthByFontAndValue.delete(firstKey);
  }

  return measuredWidthPx;
}

function resolveParagraphTabStopsPx(paragraph: ParagraphNode): number[] {
  const stopsPx = (paragraph.style?.tabStops ?? [])
    .map((tabStop) => twipsToPixels(tabStop.positionTwips))
    .filter(
      (value): value is number =>
        Number.isFinite(value) && (value as number) > 0
    )
    .map((value) => Math.round(value))
    .sort((left, right) => left - right);

  return stopsPx;
}

function resolveParagraphFirstLineOriginPx(paragraph: ParagraphNode): number {
  const leftIndentPx = twipsToSignedPixels(paragraph.style?.indent?.leftTwips);
  const firstLineIndentPx = twipsToSignedPixels(
    paragraph.style?.indent?.firstLineTwips
  );
  const hangingIndentPx = twipsToSignedPixels(
    paragraph.style?.indent?.hangingTwips
  );
  const textIndentPx =
    firstLineIndentPx ??
    (Number.isFinite(hangingIndentPx) ? -(hangingIndentPx as number) : 0);

  return (
    (Number.isFinite(leftIndentPx) ? (leftIndentPx as number) : 0) +
    (Number.isFinite(textIndentPx) ? (textIndentPx as number) : 0)
  );
}

function resolveParagraphFirstLineLeftTabStopsPx(
  paragraph: ParagraphNode
): number[] {
  const firstLineOriginPx = resolveParagraphFirstLineOriginPx(paragraph);
  return (paragraph.style?.tabStops ?? [])
    .filter((tabStop) => tabStop.alignment !== "right")
    .map((tabStop) => twipsToPixels(tabStop.positionTwips))
    .filter(
      (value): value is number =>
        Number.isFinite(value) && (value as number) > firstLineOriginPx + 0.5
    )
    .map((value) => Math.round(value - firstLineOriginPx))
    .sort((left, right) => left - right);
}

function resolveNextTabStopPx(
  currentLineWidthPx: number,
  tabStopsPx: number[]
): number {
  const nextExplicit = tabStopsPx.find(
    (stopPx) => stopPx > currentLineWidthPx + 0.5
  );
  if (nextExplicit !== undefined) {
    return nextExplicit;
  }

  const tabStepPx = DEFAULT_TAB_STOP_PX;
  const nextMultiple =
    Math.floor(Math.max(0, currentLineWidthPx) / tabStepPx + 1) * tabStepPx;
  return Math.max(tabStepPx, nextMultiple);
}

interface ParagraphMeasureSegment {
  text: string;
  style?: TextRunNode["style"] | FormFieldRunNode["style"];
}

interface ParagraphMeasureToken extends ParagraphMeasureSegment {
  isTab?: boolean;
}

function normalizeTextForTabLeaderMeasurement(text: string): string {
  return text.replace(/\t/g, "        ");
}

function measureParagraphSegmentTextWidthPx(
  segments: ParagraphMeasureSegment[],
  paragraphBaseFontPx: number
): number {
  return segments.reduce((totalWidthPx, segment) => {
    const normalizedText = normalizeTextForTabLeaderMeasurement(segment.text);
    if (!normalizedText) {
      return totalWidthPx;
    }
    return (
      totalWidthPx +
      measureTextWidthPx(normalizedText, segment.style, paragraphBaseFontPx)
    );
  }, 0);
}

function estimateWrappedLineCountForSegments(
  segments: ParagraphMeasureSegment[],
  maxLineWidthPx: number,
  paragraphBaseFontPx: number
): number {
  let lineCount = 1;
  let currentLineWidthPx = 0;
  let hasVisibleContent = false;

  const advanceByTextToken = (
    token: string,
    style: TextRunNode["style"] | FormFieldRunNode["style"] | undefined
  ): void => {
    const tokenWidthPx = measureTextWidthPx(token, style, paragraphBaseFontPx);
    if (
      token.trim().length > 0 &&
      currentLineWidthPx > 0 &&
      currentLineWidthPx + tokenWidthPx > maxLineWidthPx
    ) {
      lineCount += 1;
      currentLineWidthPx = 0;
    }

    if (tokenWidthPx <= maxLineWidthPx) {
      currentLineWidthPx = Math.min(
        maxLineWidthPx,
        currentLineWidthPx + tokenWidthPx
      );
      return;
    }

    for (const character of token) {
      const characterWidthPx = measureTextWidthPx(
        character,
        style,
        paragraphBaseFontPx
      );
      if (
        currentLineWidthPx > 0 &&
        currentLineWidthPx + characterWidthPx > maxLineWidthPx
      ) {
        lineCount += 1;
        currentLineWidthPx = 0;
      }
      currentLineWidthPx = Math.min(
        maxLineWidthPx,
        currentLineWidthPx + characterWidthPx
      );
    }
  };

  for (const segment of segments) {
    const normalizedText = normalizeTextForTabLeaderMeasurement(segment.text);
    if (!normalizedText) {
      continue;
    }

    const tokens =
      normalizedText.match(/(\r\n|\n|[^\S\r\n]+|[^\s\r\n]+)/g) ?? [];
    for (const token of tokens) {
      if (token.length === 0 || token === "\r") {
        continue;
      }

      hasVisibleContent = true;
      if (token === "\n" || token === "\r\n") {
        lineCount += 1;
        currentLineWidthPx = 0;
        continue;
      }

      advanceByTextToken(token, segment.style);
    }
  }

  return hasVisibleContent ? Math.max(1, lineCount) : 1;
}

function collectParagraphMeasureTokens(
  paragraph: ParagraphNode
): ParagraphMeasureToken[] {
  const tokens: ParagraphMeasureToken[] = [];

  const appendText = (
    text: string,
    style?: TextRunNode["style"] | FormFieldRunNode["style"]
  ): void => {
    if (!text || text === "\r") {
      return;
    }

    const parts = text.split("\t");
    parts.forEach((part, partIndex) => {
      if (part.length > 0) {
        tokens.push({
          text: part,
          style,
        });
      }

      if (partIndex < parts.length - 1) {
        tokens.push({
          text: "\t",
          style,
          isTab: true,
        });
      }
    });
  };

  paragraph.children.forEach((child) => {
    if (child.type !== "text" && child.type !== "form-field") {
      return;
    }

    appendText(
      child.type === "text" ? child.text : formFieldDisplayValue(child),
      child.style
    );
  });

  return tokens;
}

function collectParagraphTabLeaderMeasureSegments(paragraph: ParagraphNode): {
  leadingSegments: ParagraphMeasureSegment[];
  contentSegments: ParagraphMeasureSegment[];
  rightSegments: ParagraphMeasureSegment[];
  hasTabSplit: boolean;
} {
  const tokens = collectParagraphMeasureTokens(paragraph);
  const tabIndexes = tokens.reduce<number[]>((indexes, token, tokenIndex) => {
    if (token.isTab) {
      indexes.push(tokenIndex);
    }
    return indexes;
  }, []);

  if (tabIndexes.length === 0) {
    return {
      leadingSegments: [],
      contentSegments: tokens,
      rightSegments: [],
      hasTabSplit: false,
    };
  }

  const firstTabIndex = tabIndexes[0];
  const lastTabIndex = tabIndexes[tabIndexes.length - 1];

  const toSegments = (
    slice: ParagraphMeasureToken[]
  ): ParagraphMeasureSegment[] =>
    slice
      .filter((token) => !token.isTab && token.text.length > 0)
      .map((token) => ({
        text: token.text,
        style: token.style,
      }));

  if (firstTabIndex === lastTabIndex) {
    return {
      leadingSegments: [],
      contentSegments: toSegments(tokens.slice(0, firstTabIndex)),
      rightSegments: toSegments(tokens.slice(lastTabIndex + 1)),
      hasTabSplit: true,
    };
  }

  return {
    leadingSegments: toSegments(tokens.slice(0, firstTabIndex)),
    contentSegments: toSegments(tokens.slice(firstTabIndex + 1, lastTabIndex)),
    rightSegments: toSegments(tokens.slice(lastTabIndex + 1)),
    hasTabSplit: true,
  };
}

function estimateTabLeaderWrappedLineCountForParagraph(
  paragraph: ParagraphNode,
  maxLineWidthPx: number,
  paragraphBaseFontPx: number
): number | undefined {
  const { leadingSegments, contentSegments, rightSegments, hasTabSplit } =
    collectParagraphTabLeaderMeasureSegments(paragraph);
  if (!hasTabSplit) {
    return undefined;
  }

  const reservedRightZoneWidthPx =
    measureParagraphSegmentTextWidthPx(rightSegments, paragraphBaseFontPx) +
    TAB_LEADER_ZONE_GAP_PX;
  const leftZoneWidthPx = Math.max(
    paragraphBaseFontPx * 2,
    Math.round(maxLineWidthPx - reservedRightZoneWidthPx)
  );
  const leadingTextWidthPx = measureParagraphSegmentTextWidthPx(
    leadingSegments,
    paragraphBaseFontPx
  );
  const tabStopPositionsPx = resolveParagraphFirstLineLeftTabStopsPx(paragraph);
  const explicitLeadingTabStopPx =
    tableOfContentsLeadingLeftTabStopPx(paragraph);
  const leadingReservationWidthPx =
    leadingSegments.length === 0
      ? 0
      : Number.isFinite(explicitLeadingTabStopPx) &&
        (explicitLeadingTabStopPx as number) > 0
      ? Math.max(
          leadingTextWidthPx,
          Math.round(explicitLeadingTabStopPx as number)
        )
      : Math.max(
          leadingTextWidthPx,
          Math.round(
            leadingTextWidthPx +
              resolveTabSpacerWidthPx(
                tabStopPositionsPx,
                leadingTextWidthPx,
                DEFAULT_TAB_STOP_PX
              )
          )
        );
  const contentZoneWidthPx = Math.max(
    paragraphBaseFontPx * 2,
    Math.round(leftZoneWidthPx - leadingReservationWidthPx)
  );
  const leftZoneText = contentSegments
    .map((segment) => normalizeTextForTabLeaderMeasurement(segment.text))
    .join("");
  if (leftZoneText.trim().length === 0) {
    return 1;
  }

  const leftZoneFont = resolveMeasureFont(
    contentSegments.find((segment) => segment.text.trim().length > 0)?.style,
    paragraphBaseFontPx
  );
  const lineHeightPx = estimateParagraphLineHeightPx(paragraph);
  const pretextLayout = layoutItemsWithPretextAroundExclusions(
    leftZoneText,
    buildMeasureSegmentsPretextLayoutItems(
      contentSegments,
      paragraphBaseFontPx,
      leftZoneText
    ),
    contentZoneWidthPx,
    lineHeightPx,
    [],
    leftZoneFont
  );
  if (pretextLayout?.lineCount) {
    return pretextLayout.lineCount;
  }

  return estimateWrappedLineCountForSegments(
    contentSegments,
    contentZoneWidthPx,
    paragraphBaseFontPx
  );
}

const wrappedLineCountByParagraph = new WeakMap<
  ParagraphNode,
  Map<number | string, number>
>();

function cachedWrappedLineCountForParagraph(
  paragraph: ParagraphNode,
  widthKey: number | string
): number | undefined {
  return wrappedLineCountByParagraph.get(paragraph)?.get(widthKey);
}

function rememberWrappedLineCountForParagraph(
  paragraph: ParagraphNode,
  widthKey: number | string,
  lineCount: number
): number {
  let countsByWidth = wrappedLineCountByParagraph.get(paragraph);
  if (!countsByWidth) {
    countsByWidth = new Map<number | string, number>();
    wrappedLineCountByParagraph.set(paragraph, countsByWidth);
  }
  countsByWidth.set(widthKey, lineCount);
  return lineCount;
}

function estimateWrappedLineCountForParagraph(
  paragraph: ParagraphNode,
  availableWidthPx: number,
  firstLineAvailableWidthPx?: number
): number {
  const paragraphBaseFontPx = paragraphBaseFontSizePx(paragraph);
  const maxLineWidthPx = Math.max(
    paragraphBaseFontPx * 2,
    Math.round(availableWidthPx)
  );
  const firstLineMaxWidthPx =
    Number.isFinite(firstLineAvailableWidthPx) &&
    (firstLineAvailableWidthPx as number) > 0
      ? Math.max(
          paragraphBaseFontPx * 2,
          Math.round(firstLineAvailableWidthPx as number)
        )
      : maxLineWidthPx;
  const widthCacheKey: number | string =
    firstLineMaxWidthPx === maxLineWidthPx
      ? maxLineWidthPx
      : `${maxLineWidthPx}|${firstLineMaxWidthPx}`;
  const cachedLineCount = cachedWrappedLineCountForParagraph(
    paragraph,
    widthCacheKey
  );
  if (cachedLineCount !== undefined) {
    return cachedLineCount;
  }

  const rememberLineCount = (lineCount: number): number =>
    rememberWrappedLineCountForParagraph(
      paragraph,
      widthCacheKey,
      Math.max(1, Math.round(lineCount))
    );
  const tabStopsPx = resolveParagraphTabStopsPx(paragraph);
  const useTabLeaderLayout = paragraphUsesTabLeaders(paragraph);
  const anchoredTabLayout = paragraphAnchoredTabLayout(paragraph);
  const useCenterRightTabLayout =
    !useTabLeaderLayout && anchoredTabLayout === "center-right";
  const useCenterTabLayout =
    !useTabLeaderLayout && anchoredTabLayout === "center";
  const useRightTabLayout =
    !useTabLeaderLayout && anchoredTabLayout === "right";
  const useAnchoredTabLayout =
    useCenterRightTabLayout || useCenterTabLayout || useRightTabLayout;
  const paragraphEligibleForPlainPretextLineCount =
    !paragraph.style?.indent &&
    firstLineMaxWidthPx === maxLineWidthPx &&
    !paragraphContainsExplicitLineBreakText(paragraph);

  if (
    paragraphEligibleForPlainPretextLineCount &&
    !useTabLeaderLayout &&
    !useAnchoredTabLayout
  ) {
    const pretextPlainLineCount = (() => {
      const pretextSource = buildParagraphPretextLayoutSource(paragraph);
      if (!pretextSource) {
        return undefined;
      }

      // Fast path (pretext 0.0.5+): when the paragraph's measurement runs all
      // share a single font and there are no atomic items (images, tabs),
      // `measureLineStats` wraps the text and returns the line count using
      // pure arithmetic over cached segment widths — no per-line string
      // allocations, no fragment objects. This is the hot path during
      // pagination of large documents.
      const uniformFont = resolveUniformPretextSourceFont(
        paragraph,
        pretextSource
      );
      if (uniformFont !== undefined) {
        const fastLineCount = measurePretextPlainTextLineCount(
          pretextSource.text,
          uniformFont,
          maxLineWidthPx,
          { wordBreak: pretextWordBreakModeForText(pretextSource.text) }
        );
        if (fastLineCount !== undefined) {
          return fastLineCount;
        }
      }

      const lineHeightPx = estimateParagraphLineHeightPx(paragraph);
      const layout = layoutParagraphPretextSource(
        paragraph,
        pretextSource,
        maxLineWidthPx,
        lineHeightPx,
        []
      );
      return layout?.lineCount;
    })();
    if (pretextPlainLineCount) {
      return rememberLineCount(pretextPlainLineCount);
    }
  }

  if (useTabLeaderLayout) {
    const tabLeaderLineCount = estimateTabLeaderWrappedLineCountForParagraph(
      paragraph,
      maxLineWidthPx,
      paragraphBaseFontPx
    );
    if (tabLeaderLineCount !== undefined) {
      return rememberLineCount(tabLeaderLineCount);
    }
  }
  let lineCount = 1;
  let currentLineWidthPx = 0;
  let hasVisibleContent = false;
  let tabLeaderRightZoneActive = false;
  let tabLeaderRightZoneWidthPx = 0;

  // The first line keeps a hanging indent's extra room (minus any list marker
  // box); subsequent lines wrap within the regular indented width.
  const currentMaxLineWidthPx = (): number =>
    lineCount === 1 ? firstLineMaxWidthPx : maxLineWidthPx;

  const advanceByTextToken = (
    token: string,
    style: TextRunNode["style"] | FormFieldRunNode["style"] | undefined
  ): void => {
    const tokenWidthPx = measureTextWidthPx(token, style, paragraphBaseFontPx);
    if (
      token.trim().length > 0 &&
      currentLineWidthPx > 0 &&
      currentLineWidthPx + tokenWidthPx > currentMaxLineWidthPx()
    ) {
      lineCount += 1;
      currentLineWidthPx = 0;
    }

    if (tokenWidthPx <= currentMaxLineWidthPx()) {
      currentLineWidthPx = Math.min(
        currentMaxLineWidthPx(),
        currentLineWidthPx + tokenWidthPx
      );
      return;
    }

    for (const character of token) {
      const characterWidthPx = measureTextWidthPx(
        character,
        style,
        paragraphBaseFontPx
      );
      if (
        currentLineWidthPx > 0 &&
        currentLineWidthPx + characterWidthPx > currentMaxLineWidthPx()
      ) {
        lineCount += 1;
        currentLineWidthPx = 0;
      }
      currentLineWidthPx = Math.min(
        currentMaxLineWidthPx(),
        currentLineWidthPx + characterWidthPx
      );
    }
  };

  const commitToken = (
    token: string,
    style: TextRunNode["style"] | FormFieldRunNode["style"] | undefined
  ): void => {
    if (token.length === 0 || token === "\r") {
      return;
    }

    hasVisibleContent = true;
    if (token === "\n" || token === "\r\n") {
      tabLeaderRightZoneActive = false;
      tabLeaderRightZoneWidthPx = 0;
      lineCount += 1;
      currentLineWidthPx = 0;
      return;
    }

    if (useTabLeaderLayout && tabLeaderRightZoneActive) {
      tabLeaderRightZoneWidthPx += measureTextWidthPx(
        token,
        style,
        paragraphBaseFontPx
      );
      if (
        currentLineWidthPx > 0 &&
        currentLineWidthPx +
          tabLeaderRightZoneWidthPx +
          TAB_LEADER_ZONE_GAP_PX >
          maxLineWidthPx + PAGE_OVERFLOW_TOLERANCE_PX
      ) {
        lineCount += 1;
        currentLineWidthPx = 0;
        tabLeaderRightZoneWidthPx = 0;
      }
      return;
    }

    if (token === "\t") {
      if (useTabLeaderLayout) {
        // TOC leader layout renders the right-side page-number zone independently
        // from left text width while still sharing the same visual line.
        tabLeaderRightZoneActive = true;
        tabLeaderRightZoneWidthPx = 0;
        return;
      }

      if (useAnchoredTabLayout) {
        currentLineWidthPx = Math.min(
          maxLineWidthPx,
          Math.max(currentLineWidthPx, maxLineWidthPx - 1)
        );
        return;
      }

      const nextTabStopPx = resolveNextTabStopPx(
        currentLineWidthPx,
        tabStopsPx
      );
      if (
        nextTabStopPx > maxLineWidthPx + PAGE_OVERFLOW_TOLERANCE_PX &&
        currentLineWidthPx > 0
      ) {
        lineCount += 1;
        currentLineWidthPx = 0;
      } else {
        currentLineWidthPx = Math.min(maxLineWidthPx, nextTabStopPx);
      }
      return;
    }

    advanceByTextToken(token, style);
  };

  for (const child of paragraph.children) {
    if (child.type === "image") {
      if (
        shouldRenderAbsoluteFloatingImage(child) ||
        shouldRenderWrappedFloatingImage(child)
      ) {
        continue;
      }
      const inlineImageWidthPx = Math.max(
        1,
        Math.round(child.widthPx ?? child.heightPx ?? 24)
      );
      hasVisibleContent = true;
      if (
        currentLineWidthPx > 0 &&
        currentLineWidthPx + inlineImageWidthPx > maxLineWidthPx
      ) {
        lineCount += 1;
        currentLineWidthPx = 0;
      }
      currentLineWidthPx = Math.min(
        maxLineWidthPx,
        currentLineWidthPx + inlineImageWidthPx
      );
      continue;
    }

    const text =
      child.type === "text" ? child.text : formFieldDisplayValue(child);
    if (!text) {
      continue;
    }

    const tokens = text.match(/(\r\n|\n|\t|[^\S\r\n\t]+|[^\s\r\n\t]+)/g) ?? [];
    for (const token of tokens) {
      commitToken(token, child.style);
    }
  }

  return rememberLineCount(hasVisibleContent ? lineCount : 1);
}

function paragraphAvailableTextWidthPx(
  paragraph: ParagraphNode,
  availableWidthPx: number,
  numberingDefinitions?: NumberingDefinitionSet
): number {
  const safeAvailableWidthPx = Math.max(24, Math.round(availableWidthPx));
  const resolvedIndent = resolveListParagraphIndent(
    paragraph,
    numberingDefinitions
  );
  const leftIndentPx = Math.max(
    0,
    twipsToSignedPixels(resolvedIndent?.leftTwips) ?? 0
  );
  const rightIndentPx = Math.max(
    0,
    twipsToSignedPixels(paragraph.style?.indent?.rightTwips) ?? 0
  );
  const firstLineIndentPx = twipsToSignedPixels(resolvedIndent?.firstLineTwips);
  const hangingIndentPx = twipsToSignedPixels(resolvedIndent?.hangingTwips);
  const firstLineDeltaPx =
    firstLineIndentPx ?? (hangingIndentPx ? -hangingIndentPx : 0);
  const textIndentReductionPx =
    Number.isFinite(firstLineDeltaPx) && (firstLineDeltaPx as number) > 0
      ? (firstLineDeltaPx as number)
      : 0;
  const leftBorderInsetPx = paragraphBorderInsetPx(
    paragraph.style?.borders?.left
  );
  const rightBorderInsetPx = paragraphBorderInsetPx(
    paragraph.style?.borders?.right
  );

  return Math.max(
    24,
    Math.round(
      safeAvailableWidthPx -
        leftIndentPx -
        rightIndentPx -
        textIndentReductionPx -
        leftBorderInsetPx -
        rightBorderInsetPx
    )
  );
}

export function paragraphLineCountWithinWidth(
  paragraph: ParagraphNode,
  availableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet,
  numberingLabel?: ParagraphNumberingLabel
): number {
  const textContent = paragraph.children
    .map((child) => {
      if (child.type === "text") {
        return child.text;
      }
      if (child.type === "form-field") {
        return formFieldDisplayValue(child);
      }
      return "";
    })
    .join("");
  if (!textContent) {
    return 1;
  }

  if (!Number.isFinite(availableWidthPx) || (availableWidthPx as number) <= 0) {
    return Math.max(1, textContent.split(/\r?\n/).length);
  }

  const effectiveWidthPx = paragraphAvailableTextWidthPx(
    paragraph,
    availableWidthPx as number,
    numberingDefinitions
  );
  const dualWrappedLayout = resolveParagraphDualWrappedTextLayout(
    paragraph,
    effectiveWidthPx,
    estimateParagraphLineHeightPx(paragraph)
  );
  if (dualWrappedLayout) {
    return dualWrappedLayout.layout.lineCount;
  }

  // A hanging indent only narrows lines after the first, so the first line
  // keeps the hanging offset as extra room. Numbered list paragraphs render an
  // inline marker box at the start of that first line, which consumes part (or
  // all) of that extra room before the paragraph text begins.
  const resolvedIndent = resolveListParagraphIndent(
    paragraph,
    numberingDefinitions
  );
  const hangingIndentPx = twipsToSignedPixels(resolvedIndent?.hangingTwips);
  const firstLineHangingBonusPx =
    Number.isFinite(hangingIndentPx) && (hangingIndentPx as number) > 0
      ? (hangingIndentPx as number)
      : 0;
  const numberingMarkerReservePx =
    (paragraph.style?.numbering?.numId ?? 0) > 0
      ? resolveNumberingMarkerBoxWidthPx(
          paragraph,
          numberingDefinitions,
          numberingLabel
        ) ?? 0
      : 0;
  const firstLineWidthPx = Math.max(
    24,
    Math.round(
      effectiveWidthPx + firstLineHangingBonusPx - numberingMarkerReservePx
    )
  );

  return estimateWrappedLineCountForParagraph(
    paragraph,
    effectiveWidthPx,
    firstLineWidthPx
  );
}

function estimateWrappedFloatingImageFootprintPx(
  paragraph: ParagraphNode,
  image: ImageRunNode
): number {
  if (!shouldRenderWrappedFloatingImage(image)) {
    return 0;
  }

  const wrapType = image.floating?.wrapType;
  const isTopAndBottomWrap = wrapType === "topAndBottom";
  const isImageOnlyAnchorParagraph = !paragraphHasVisibleText(paragraph);
  if (!isTopAndBottomWrap && !isImageOnlyAnchorParagraph) {
    return 0;
  }

  // For paragraphs that already have visible text, top-and-bottom wrapped
  // objects should not force the anchor paragraph box to extend all the way to
  // the image bottom. Word allows following text to occupy the gap above a
  // lowered object, so keep that reserve logic only for image-only anchors.
  if (isTopAndBottomWrap && !isImageOnlyAnchorParagraph) {
    return 0;
  }

  const floating = image.floating;
  const imageHeightPx =
    Number.isFinite(image.heightPx) && (image.heightPx as number) > 0
      ? Math.round(image.heightPx as number)
      : Number.isFinite(image.widthPx) && (image.widthPx as number) > 0
      ? Math.round(image.widthPx as number)
      : MIN_PARAGRAPH_LINE_HEIGHT_PX;
  const distTPx = Math.max(0, Math.round(floating?.distTPx ?? 0));
  const distBPx = Math.max(0, Math.round(floating?.distBPx ?? 0));
  const verticalOffsetPx = Math.max(0, Math.round(floating?.yPx ?? 0));

  return Math.max(
    MIN_PARAGRAPH_LINE_HEIGHT_PX,
    imageHeightPx + distTPx + distBPx + verticalOffsetPx
  );
}

function pretextLayoutContentBottomPx(
  layout: PretextVariableWidthLayout
): number {
  if (layout.lines.length === 0) {
    return 0;
  }

  const lineHeightPx = Math.max(1, Math.round(layout.lineHeightPx ?? 1));
  return (layout.lines[layout.lines.length - 1]?.y ?? 0) + lineHeightPx;
}

function topAndBottomExclusionCanOverflowParagraphBox(
  layout: PretextVariableWidthLayout
): boolean {
  const containerWidthPx = Math.max(
    1,
    Math.round(layout.containerWidthPx ?? 0)
  );
  if (containerWidthPx <= 0 || layout.lines.length === 0) {
    return false;
  }

  return (layout.exclusions ?? []).some((exclusion) => {
    const left = Math.round(exclusion.left);
    const right = Math.round(exclusion.right);
    return left <= 0 && right >= containerWidthPx;
  });
}

export function wrappedPretextParagraphBlockHeightPx(
  layout: PretextVariableWidthLayout
): number {
  const contentBottomPx = pretextLayoutContentBottomPx(layout);
  if (topAndBottomExclusionCanOverflowParagraphBox(layout)) {
    return Math.max(1, contentBottomPx);
  }

  return Math.max(1, Math.round(layout.height));
}

export function resolvePretextLineRangeContentHeightPx(
  layout: PretextVariableWidthLayout,
  startLineIndex: number,
  endLineIndex: number
): number {
  if (layout.lines.length === 0) {
    return 0;
  }

  const lineHeightPx = Math.max(1, Math.round(layout.lineHeightPx ?? 1));
  const safeStart = Math.max(
    0,
    Math.min(Math.round(startLineIndex), layout.lines.length)
  );
  const safeEnd = Math.max(
    safeStart,
    Math.min(Math.round(endLineIndex), layout.lines.length)
  );
  if (safeEnd <= safeStart) {
    return 0;
  }

  const firstLineTopPx = layout.lines[safeStart]?.y ?? safeStart * lineHeightPx;
  const lastLineTopPx =
    layout.lines[safeEnd - 1]?.y ?? (safeEnd - 1) * lineHeightPx;
  return Math.max(1, Math.round(lastLineTopPx - firstLineTopPx + lineHeightPx));
}

export function resolveMaxPretextLineRangeEndIndexThatFits(
  layout: PretextVariableWidthLayout,
  startLineIndex: number,
  maxEndLineIndex: number,
  availableHeightPx: number
): number {
  const safeStart = Math.max(
    0,
    Math.min(Math.round(startLineIndex), layout.lines.length)
  );
  const safeMaxEnd = Math.max(
    safeStart,
    Math.min(Math.round(maxEndLineIndex), layout.lines.length)
  );
  const safeAvailableHeightPx = Math.max(0, Math.round(availableHeightPx));
  if (safeAvailableHeightPx <= 0 || safeMaxEnd <= safeStart) {
    return safeStart;
  }

  let low = safeStart;
  let high = safeMaxEnd;
  let bestEnd = safeStart;
  while (low <= high) {
    const mid = Math.floor((low + high) / 2);
    const candidateHeightPx = resolvePretextLineRangeContentHeightPx(
      layout,
      safeStart,
      mid
    );
    if (candidateHeightPx <= safeAvailableHeightPx) {
      bestEnd = mid;
      low = mid + 1;
      continue;
    }
    high = mid - 1;
  }

  return bestEnd;
}

function resolveAutoLineSpacingMultiple(
  lineTwips: number | undefined,
  fallbackMultiple: number
): number {
  if (!Number.isFinite(lineTwips)) {
    return Math.max(MIN_AUTO_LINE_MULTIPLE, fallbackMultiple);
  }

  return Math.max(MIN_AUTO_LINE_MULTIPLE, (lineTwips as number) / 240);
}

function autoLineHeightScaleForMultiple(
  multiple: number,
  singleLineScale: number
): number {
  const safeSingleLineScale = Math.max(
    MIN_AUTO_LINE_MULTIPLE,
    Number.isFinite(singleLineScale)
      ? singleLineScale
      : WORD_SINGLE_LINE_AUTO_SCALE
  );
  if (!Number.isFinite(multiple)) {
    return safeSingleLineScale;
  }

  if (multiple <= 1) {
    return safeSingleLineScale;
  }

  if (multiple >= WORD_AUTO_LINE_SCALE_BLEND_END_MULTIPLE) {
    return 1;
  }

  const blendProgress =
    (multiple - 1) / (WORD_AUTO_LINE_SCALE_BLEND_END_MULTIPLE - 1);
  return Number(
    (safeSingleLineScale + (1 - safeSingleLineScale) * blendProgress).toFixed(4)
  );
}

function calibrateAutoLineSpacingMultiple(
  multiple: number,
  fontFamily?: string,
  singleLineScaleOverride?: number
): number {
  const normalizedMultiple = Math.max(MIN_AUTO_LINE_MULTIPLE, multiple);
  const singleLineScale =
    singleLineScaleOverride ?? singleLineAutoScaleForFontFamily(fontFamily);
  return Math.max(
    MIN_AUTO_LINE_MULTIPLE,
    Number(
      (
        normalizedMultiple *
        autoLineHeightScaleForMultiple(normalizedMultiple, singleLineScale)
      ).toFixed(3)
    )
  );
}

function paragraphDocGridSnapState(
  paragraph: ParagraphNode
): "inherit" | "snap" | "disable" {
  const sourceXml = paragraph.sourceXml ?? "";
  if (!sourceXml) {
    return "inherit";
  }

  const paragraphPropertiesXml =
    sourceXml.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/i)?.[0] ??
    sourceXml.match(/<w:pPr\b[^>]*\/>/i)?.[0] ??
    "";
  if (!paragraphPropertiesXml) {
    return "inherit";
  }

  const snapToGridTag =
    paragraphPropertiesXml.match(/<w:snapToGrid\b[^>]*\/?>/i)?.[0] ?? "";
  if (!snapToGridTag) {
    return "inherit";
  }

  const valueMatch = snapToGridTag
    .match(/\bw:val="([^"]+)"/i)?.[1]
    ?.trim()
    .toLowerCase();
  if (!valueMatch) {
    return "snap";
  }

  return valueMatch === "0" || valueMatch === "false" || valueMatch === "off"
    ? "disable"
    : "snap";
}

function resolveParagraphDocGridLinePitchPx(
  paragraph: ParagraphNode,
  docGridLinePitchPx?: number,
  disableDocGridSnap = false
): number | undefined {
  const docGridSnapState = paragraphDocGridSnapState(paragraph);
  if (
    disableDocGridSnap ||
    docGridSnapState === "disable" ||
    !Number.isFinite(docGridLinePitchPx) ||
    (docGridLinePitchPx as number) <= 0
  ) {
    return undefined;
  }

  return Math.max(1, Math.round(docGridLinePitchPx as number));
}

export function estimateParagraphLineHeightPx(
  paragraph: ParagraphNode,
  docGridLinePitchPx?: number,
  disableDocGridSnap = false
): number {
  const lineTwips = paragraph.style?.spacing?.lineTwips;
  const lineRule = paragraph.style?.spacing?.lineRule ?? "auto";
  const docGridMinimumLineHeightPx = resolveParagraphDocGridLinePitchPx(
    paragraph,
    docGridLinePitchPx,
    disableDocGridSnap
  );
  const baseFontPx = paragraphBaseFontSizePx(paragraph);
  const baseFontFamily = paragraphDominantFontFamily(paragraph);
  const singleLineScale = resolveParagraphSingleLineAutoScale(
    paragraph,
    baseFontFamily
  );
  const defaultLineMultiple = isTableOfContentsParagraph(paragraph)
    ? 1.05
    : DEFAULT_PARAGRAPH_LINE_MULTIPLE;
  const normalLineHeightPx = Math.max(
    1,
    Math.round(
      baseFontPx *
        calibrateAutoLineSpacingMultiple(
          DEFAULT_PARAGRAPH_LINE_MULTIPLE,
          baseFontFamily,
          singleLineScale
        )
    )
  );

  if (lineRule !== "auto" && Number.isFinite(lineTwips)) {
    const explicitLineHeightPx = Math.max(1, twipsToPixels(lineTwips) ?? 1);
    if (lineRule === "exact") {
      return explicitLineHeightPx;
    }

    return Math.max(
      explicitLineHeightPx,
      normalLineHeightPx,
      docGridMinimumLineHeightPx ?? 0
    );
  }

  const resolvedAutoMultiple = resolveAutoLineSpacingMultiple(
    lineTwips,
    defaultLineMultiple
  );
  // A text-free paragraph renders exactly one line at the mark font's natural
  // metrics, and an explicit auto multiple scales that natural line (Word
  // semantics). The wrapped-text blend toward bare font-size lines exists to
  // offset wrapped-line overcounting, which cannot happen here.
  const multiple = paragraphRendersTextFreeLine(paragraph)
    ? Math.max(
        MIN_AUTO_LINE_MULTIPLE,
        Number((resolvedAutoMultiple * singleLineScale).toFixed(3))
      )
    : calibrateAutoLineSpacingMultiple(
        resolvedAutoMultiple,
        baseFontFamily,
        singleLineScale
      );
  const autoLineHeightPx = Math.max(1, Math.round(baseFontPx * multiple));
  const minimumReadableAutoLineHeightPx = paragraph.style?.numbering
    ? Math.ceil(baseFontPx)
    : 0;
  return Math.max(
    autoLineHeightPx,
    minimumReadableAutoLineHeightPx,
    docGridMinimumLineHeightPx ?? 0
  );
}

function estimateParagraphHeightPx(
  paragraph: ParagraphNode,
  availableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number,
  disableDocGridSnap = false,
  numberingLabel?: ParagraphNumberingLabel,
  options?: {
    // Word-like flow: square/tight wrapped floats do not consume flow height;
    // following paragraphs flow beside them. The foreign-wrap exclusion
    // simulation needs the text-flow height, not the float footprint.
    excludeWrappedFloatingImageFootprint?: boolean;
  }
): number {
  const excludeWrappedFloatingImageFootprint =
    options?.excludeWrappedFloatingImageFootprint === true;
  const sourceXml = paragraph.sourceXml;
  const baseWidthKey = heightEstimateCacheKeyPx(
    availableWidthPx,
    docGridLinePitchPx,
    disableDocGridSnap
  );
  let widthKey: number | string =
    numberingLabel?.text !== undefined
      ? `${baseWidthKey}|${numberingLabel.text}`
      : baseWidthKey;
  if (excludeWrappedFloatingImageFootprint) {
    widthKey = `${widthKey}|no-wrap-footprint`;
  }
  if (sourceXml) {
    const cachedByWidth = paragraphEstimatedHeightBySourceXml.get(sourceXml);
    const cached = cachedByWidth?.get(widthKey);
    if (cached !== undefined) {
      return cached;
    }
  }

  const beforeSpacing =
    twipsToPixels(paragraph.style?.spacing?.beforeTwips) ?? 0;
  const afterSpacing = twipsToPixels(paragraph.style?.spacing?.afterTwips) ?? 0;
  const lineHeightPx = estimateParagraphLineHeightPx(
    paragraph,
    docGridLinePitchPx,
    disableDocGridSnap
  );
  const effectiveWidthPx =
    Number.isFinite(availableWidthPx) && (availableWidthPx as number) > 0
      ? paragraphAvailableTextWidthPx(
          paragraph,
          availableWidthPx as number,
          numberingDefinitions
        )
      : undefined;
  const dualWrappedLayout =
    effectiveWidthPx !== undefined
      ? resolveParagraphDualWrappedTextLayout(
          paragraph,
          effectiveWidthPx,
          lineHeightPx
        )
      : undefined;
  const lineCount = paragraphLineCountWithinWidth(
    paragraph,
    availableWidthPx,
    numberingDefinitions,
    numberingLabel
  );
  const absoluteFloatingAnchorOnlyParagraph =
    paragraphIsAbsoluteFloatingImageAnchorOnly(paragraph);
  // Word still lays out an anchor-only paragraph as one empty line at its
  // mark/run height; only section-break carryover anchors collapse (their
  // paragraph mark belongs to the swallowed section boundary).
  const collapsibleAbsoluteFloatingAnchorOnlyParagraph =
    absoluteFloatingAnchorOnlyParagraph &&
    paragraphIsSectionBreakAnchorCarryover(paragraph);
  const inlineImageHeightPx = paragraph.children.reduce((largest, child) => {
    if (child.type !== "image") {
      return largest;
    }
    if (
      shouldRenderAbsoluteFloatingImage(child) ||
      shouldRenderWrappedFloatingImage(child)
    ) {
      return largest;
    }

    return Math.max(largest, child.heightPx ?? 0);
  }, 0);
  const wrappedFloatingImageHeightPx = excludeWrappedFloatingImageFootprint
    ? 0
    : paragraph.children.reduce((largest, child) => {
        if (child.type !== "image") {
          return largest;
        }

        return Math.max(
          largest,
          estimateWrappedFloatingImageFootprintPx(paragraph, child)
        );
      }, 0);
  const emptyParagraphHeightPx = paragraphIsEffectivelyEmpty(paragraph)
    ? lineHeightPx + EMPTY_PARAGRAPH_EXTRA_HEIGHT_PX
    : 0;
  const topBorderInsetPx = paragraphBorderInsetPx(
    paragraph.style?.borders?.top
  );
  const bottomBorderInsetPx = paragraphBorderInsetPx(
    paragraph.style?.borders?.bottom
  );
  const textFlowHeightPx = collapsibleAbsoluteFloatingAnchorOnlyParagraph
    ? 0
    : // When excluding the wrapped-float footprint, the dual-wrapped block
    // height spans the image; but the rendered paragraph only occupies its
    // text lines while the float overhangs. Use the text-line height so the
    // float's exclusion can overlap following paragraphs.
    dualWrappedLayout && !excludeWrappedFloatingImageFootprint
    ? wrappedPretextParagraphBlockHeightPx(dualWrappedLayout.layout)
    : lineHeightPx * lineCount;

  const contentHeightPx = Math.max(
    collapsibleAbsoluteFloatingAnchorOnlyParagraph ? 0 : lineHeightPx,
    textFlowHeightPx,
    inlineImageHeightPx,
    wrappedFloatingImageHeightPx,
    emptyParagraphHeightPx
  );
  if (excludeWrappedFloatingImageFootprint && paragraph.children.some((c) => c.type === "image")) {
  }
  const estimatedHeightPx = Math.max(
    1,
    beforeSpacing +
      afterSpacing +
      topBorderInsetPx +
      bottomBorderInsetPx +
      contentHeightPx
  );
  if (sourceXml) {
    const cachedByWidth =
      paragraphEstimatedHeightBySourceXml.get(sourceXml) ??
      new Map<number, number>();
    cachedByWidth.set(widthKey, estimatedHeightPx);
    setCacheEntry(
      paragraphEstimatedHeightBySourceXml,
      sourceXml,
      cachedByWidth
    );
  }
  return estimatedHeightPx;
}

function paragraphHasExplicitBeforeSpacing(paragraph: ParagraphNode): boolean {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return false;
  }

  const spacingTag = xml.match(/<w:spacing\b[^>]*\/?>/i)?.[0];
  if (!spacingTag) {
    return false;
  }

  return /\bw:before(?:\s*=|Lines\s*=|Autospacing\s*=)/i.test(spacingTag);
}

function paragraphHasExplicitSpacing(paragraph: ParagraphNode): boolean {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return false;
  }

  const paragraphPropertiesXml =
    xml.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/i)?.[0] ??
    xml.match(/<w:pPr\b[^>]*\/>/i)?.[0] ??
    "";
  if (!paragraphPropertiesXml) {
    return false;
  }

  return /<w:spacing\b[^>]*\/?>/i.test(paragraphPropertiesXml);
}

function wordLikeTableCellParagraph(
  paragraph: ParagraphNode,
  applyWordTableDefaults: boolean
): ParagraphNode {
  if (
    !applyWordTableDefaults ||
    !paragraph.sourceXml ||
    paragraphHasExplicitSpacing(paragraph)
  ) {
    return paragraph;
  }

  return {
    ...paragraph,
    sourceXml: undefined,
    style: {
      ...(paragraph.style ?? {}),
      spacing: {
        ...(paragraph.style?.spacing ?? {}),
        beforeTwips: WORD_TABLE_CELL_PARAGRAPH_BEFORE_TWIPS,
        afterTwips: WORD_TABLE_CELL_PARAGRAPH_AFTER_TWIPS,
        lineTwips: WORD_TABLE_CELL_PARAGRAPH_AUTO_LINE_TWIPS,
        lineRule: "auto",
      },
    },
  };
}

function suppressFirstTableCellParagraphTopSpacing(
  paragraph: ParagraphNode
): boolean {
  // Keep explicit paragraph top spacing from DOCX; only suppress implicit style spacing.
  if (!paragraph.sourceXml) {
    const beforeTwips = paragraph.style?.spacing?.beforeTwips;
    return !(Number.isFinite(beforeTwips) && (beforeTwips as number) > 0);
  }

  return !paragraphHasExplicitBeforeSpacing(paragraph);
}

function estimateTableCellContentHeightPx(
  nodeContent: TableCellContentNode[],
  availableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet,
  applyWordTableDefaults = false,
  docGridLinePitchPx?: number
): number {
  let paragraphIndex = 0;
  let expandedWithPretextLayout = false;
  let totalHeightPx = 0;

  for (const contentNode of nodeContent) {
    if (!isParagraphCellContentNode(contentNode)) {
      totalHeightPx += estimateTableHeightPx(
        contentNode,
        availableWidthPx,
        numberingDefinitions,
        docGridLinePitchPx
      );
      continue;
    }

    const disableDocGridSnap =
      paragraphDocGridSnapState(contentNode) !== "snap";
    const paragraphForLayout = wordLikeTableCellParagraph(
      contentNode,
      applyWordTableDefaults
    );
    const baseHeight = estimateParagraphHeightPx(
      paragraphForLayout,
      availableWidthPx,
      numberingDefinitions,
      docGridLinePitchPx,
      disableDocGridSnap
    );
    const lineHeightPx = Math.max(
      MIN_PARAGRAPH_LINE_HEIGHT_PX,
      estimateParagraphLineHeightPx(
        paragraphForLayout,
        docGridLinePitchPx,
        disableDocGridSnap
      )
    );
    const pretextSource = buildParagraphPretextLayoutSource(
      paragraphForLayout,
      {
        allowExplicitLineBreakText: true,
        expandTabsForLayout: true,
      }
    );
    const paragraphTextWidthPx =
      typeof availableWidthPx === "number" && availableWidthPx > 0
        ? paragraphAvailableTextWidthPx(
            paragraphForLayout,
            availableWidthPx,
            numberingDefinitions
          )
        : undefined;
    const pretextLayout =
      pretextSource &&
      typeof paragraphTextWidthPx === "number" &&
      paragraphTextWidthPx > 0
        ? layoutParagraphPretextSource(
            paragraphForLayout,
            pretextSource,
            paragraphTextWidthPx,
            lineHeightPx,
            []
          )
        : undefined;
    const suppressTopSpacing =
      paragraphIndex === 0 &&
      suppressFirstTableCellParagraphTopSpacing(contentNode);
    paragraphIndex += 1;
    const beforeSpacing = suppressTopSpacing
      ? 0
      : twipsToPixels(paragraphForLayout.style?.spacing?.beforeTwips) ?? 0;
    const afterSpacing =
      twipsToPixels(paragraphForLayout.style?.spacing?.afterTwips) ?? 0;
    const topBorderInsetPx = paragraphBorderInsetPx(
      paragraphForLayout.style?.borders?.top
    );
    const bottomBorderInsetPx = paragraphBorderInsetPx(
      paragraphForLayout.style?.borders?.bottom
    );
    const pretextHeightPx = pretextLayout
      ? beforeSpacing +
        afterSpacing +
        topBorderInsetPx +
        bottomBorderInsetPx +
        wrappedPretextParagraphBlockHeightPx(pretextLayout)
      : 0;
    const resolvedBaseHeight =
      pretextHeightPx > 0 ? Math.max(baseHeight, pretextHeightPx) : baseHeight;
    // Only treat the pretext layout as a genuine multi-line expansion when it
    // exceeds the base estimate by at least half a line. A tiny (rounding)
    // difference between the two height calcs must not trigger the extra-line
    // safety pad below, which otherwise added a phantom ~14px to every
    // single-line cell and roughly doubled table heights during pagination.
    if (pretextHeightPx > baseHeight + lineHeightPx / 2) {
      expandedWithPretextLayout = true;
    }

    if (!suppressTopSpacing) {
      totalHeightPx += resolvedBaseHeight;
      continue;
    }

    totalHeightPx += Math.max(1, resolvedBaseHeight - beforeSpacing);
  }

  return (
    totalHeightPx +
    (expandedWithPretextLayout ? Math.max(1, MIN_PARAGRAPH_LINE_HEIGHT_PX) : 0)
  );
}

function rowAllowsPageSplit(row: TableNode["rows"][number]): boolean {
  return row.style?.cantSplit !== true && row.style?.heightRule !== "exact";
}

function rowHasDeepFlowContent(row: TableNode["rows"][number]): boolean {
  let blockNodeCount = 0;
  let nestedTableCount = 0;

  for (const cell of row.cells) {
    blockNodeCount += cell.nodes.length;
    for (const contentNode of cell.nodes) {
      if (contentNode.type === "table") {
        nestedTableCount += 1;
      }
    }
  }

  return (
    nestedTableCount > 0 ||
    blockNodeCount >= SPLITTABLE_TABLE_ROW_DEEP_CONTENT_NODE_THRESHOLD
  );
}

function rowHasNestedTableContent(row: TableNode["rows"][number]): boolean {
  return row.cells.some((cell) =>
    cell.nodes.some((contentNode) => contentNode.type === "table")
  );
}

function capSplitFriendlyTableRowEstimatePx(
  row: TableNode["rows"][number],
  estimatedRowHeightPx: number,
  explicitHeightPx?: number,
  pageContentHeightPx?: number
): number {
  if (!rowAllowsPageSplit(row)) {
    return estimatedRowHeightPx;
  }

  if (!Number.isFinite(explicitHeightPx) || (explicitHeightPx as number) <= 0) {
    return estimatedRowHeightPx;
  }

  if (!rowHasDeepFlowContent(row)) {
    return estimatedRowHeightPx;
  }

  const safeExplicitHeightPx = Math.max(
    MIN_PARAGRAPH_LINE_HEIGHT_PX * 2,
    Math.round(explicitHeightPx as number)
  );
  const safePageContentHeightPx =
    Number.isFinite(pageContentHeightPx) && (pageContentHeightPx as number) > 0
      ? Math.max(
          MIN_PARAGRAPH_LINE_HEIGHT_PX * 4,
          Math.round(pageContentHeightPx as number)
        )
      : undefined;
  if (
    safePageContentHeightPx !== undefined &&
    estimatedRowHeightPx > safePageContentHeightPx + PAGE_OVERFLOW_TOLERANCE_PX
  ) {
    return estimatedRowHeightPx;
  }

  const cappedHeightPx =
    safeExplicitHeightPx +
    MIN_PARAGRAPH_LINE_HEIGHT_PX *
      SPLITTABLE_TABLE_ROW_ESTIMATE_EXTRA_LINE_COUNT;

  return Math.min(estimatedRowHeightPx, cappedHeightPx);
}

function tableStyleIdFromSourceXml(table: TableNode): string | undefined {
  const sourceXml = table.sourceXml ?? "";
  if (!sourceXml) {
    return undefined;
  }

  const styleMatch = sourceXml.match(/<w:tblStyle\b[^>]*w:val="([^"]+)"/i);
  const styleId = styleMatch?.[1]?.trim();
  return styleId ? styleId : undefined;
}

function tableHasVisibleBorders(table: TableNode): boolean {
  const borders = table.style?.borders;
  if (!borders) {
    return false;
  }

  return Object.values(borders).some(
    (border) => border && border.type !== "none" && border.type !== "nil"
  );
}

function tableContainsParagraphsWithoutExplicitSpacing(
  table: TableNode
): boolean {
  return table.rows.some((row) =>
    row.cells.some((cell) =>
      cell.nodes.some(
        (node) =>
          node.type === "paragraph" && !paragraphHasExplicitSpacing(node)
      )
    )
  );
}

function tableUsesWordLikeParagraphDefaults(table: TableNode): boolean {
  const styleId = tableStyleIdFromSourceXml(table)?.toLowerCase();
  if (styleId === "tablegrid") {
    return true;
  }

  return (
    table.style?.layout === "fixed" &&
    tableHasVisibleBorders(table) &&
    tableContainsParagraphsWithoutExplicitSpacing(table)
  );
}

export function estimateTableRowHeightsPx(
  table: TableNode,
  maxAvailableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number,
  pageContentHeightPx?: number
): number[] {
  const baseCacheKey = heightEstimateCacheKeyPx(
    maxAvailableWidthPx,
    docGridLinePitchPx
  );
  const cachedByKey = tableEstimatedRowHeightsByNode.get(table);
  let baseRowHeights = cachedByKey?.get(baseCacheKey);

  if (!baseRowHeights) {
    baseRowHeights = computeTableCellDerivedRowHeightsPx(
      table,
      maxAvailableWidthPx,
      numberingDefinitions,
      docGridLinePitchPx
    );
    const cacheByKey =
      cachedByKey ?? new Map<number, number[]>();
    cacheByKey.set(baseCacheKey, baseRowHeights);
    tableEstimatedRowHeightsByNode.set(table, cacheByKey);
  }

  return table.rows.map((row, rowIndex) => {
    let rowHeightPx = baseRowHeights[rowIndex] ?? 0;

    const explicitHeightPx = twipsToPixels(row.style?.heightTwips);
    if (explicitHeightPx && explicitHeightPx > 0) {
      rowHeightPx =
        row.style?.heightRule === "exact"
          ? explicitHeightPx
          : Math.max(rowHeightPx, explicitHeightPx);
    }
    rowHeightPx = capSplitFriendlyTableRowEstimatePx(
      row,
      rowHeightPx,
      explicitHeightPx,
      pageContentHeightPx
    );

    const paginationPaddingRatio =
      table.rows.length >= 35
        ? 1.32
        : table.rows.length >= TABLE_ROW_HEIGHT_PAGINATION_ESTIMATE_PADDING_MIN_ROWS
        ? TABLE_ROW_HEIGHT_PAGINATION_ESTIMATE_PADDING_RATIO
        : 1;
    const paddedRowHeightPx =
      paginationPaddingRatio > 1
        ? Math.round(rowHeightPx * paginationPaddingRatio)
        : rowHeightPx;

    return Math.max(MIN_PARAGRAPH_LINE_HEIGHT_PX, paddedRowHeightPx);
  });
}

function computeTableCellDerivedRowHeightsPx(
  table: TableNode,
  maxAvailableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number
): number[] {
  const defaultCellMargin = table.style?.cellMarginTwips;
  const columnCount = tableColumnCount(table);
  const tableWidthPx = twipsToPixels(table.style?.widthTwips);
  const rawTableColumnWidthsPx = (() => {
    const definedWidthsTwips = columnWidthsFromTableDefinition(
      table,
      columnCount
    );
    if (!definedWidthsTwips || definedWidthsTwips.length === 0) {
      return defaultColumnWidthsPx(columnCount, tableWidthPx);
    }

    const widthsPx = definedWidthsTwips.map(
      (widthTwips) => twipsToPixels(widthTwips) ?? 0
    );
    return normalizeColumnWidthsPx(widthsPx, columnCount, tableWidthPx, 1);
  })();
  const rawResolvedTableWidthPx =
    tableWidthPx ??
    rawTableColumnWidthsPx.reduce((sum, widthPx) => sum + widthPx, 0);
  const collapsedHorizontalBorderBleedPx =
    resolveCollapsedTableHorizontalOuterBleedPx(table, columnCount);
  const maxTableWidthPx =
    Number.isFinite(maxAvailableWidthPx) && (maxAvailableWidthPx as number) > 0
      ? Math.max(
          120,
          (maxAvailableWidthPx as number) - collapsedHorizontalBorderBleedPx
        )
      : undefined;
  const resolvedTableWidthPx = clampTableWidthPx(
    rawResolvedTableWidthPx,
    maxTableWidthPx
  );
  const tableColumnWidthsPx = fitColumnWidthsToWidth(
    rawTableColumnWidthsPx,
    resolvedTableWidthPx
  );
  const applyWordTableDefaults = tableUsesWordLikeParagraphDefaults(table);

  return table.rows.map((row) => {
    let columnCursor = 0;
    const rowHeightPx = row.cells.reduce((largest, cell) => {
      const columnSpan =
        cell.style?.gridSpan && cell.style.gridSpan > 1
          ? cell.style.gridSpan
          : 1;
      const startColumnIndex = columnCursor;
      const endColumnIndex = Math.min(
        columnCount - 1,
        startColumnIndex + columnSpan - 1
      );
      columnCursor += columnSpan;
      const spanWidthPx = tableColumnWidthsPx
        .slice(startColumnIndex, endColumnIndex + 1)
        .reduce((sum, widthPx) => sum + widthPx, 0);
      const fallbackCellWidthPx =
        (resolvedTableWidthPx / Math.max(1, columnCount)) * columnSpan;
      const cellWidthPx = spanWidthPx > 0 ? spanWidthPx : fallbackCellWidthPx;
      const margin = cell.style?.marginTwips ?? defaultCellMargin;
      const resolvedPaddingPx = resolveTableSpacingPaddingPx(margin);
      const verticalPaddingPx =
        resolvedPaddingPx.top + resolvedPaddingPx.bottom;
      const horizontalPaddingPx =
        resolvedPaddingPx.left + resolvedPaddingPx.right;
      const contentWidthPx = Math.max(
        24,
        Math.round(cellWidthPx - horizontalPaddingPx)
      );
      const paragraphHeightPx = estimateTableCellContentHeightPx(
        cell.nodes,
        contentWidthPx,
        numberingDefinitions,
        applyWordTableDefaults,
        docGridLinePitchPx
      );
      return Math.max(largest, paragraphHeightPx + verticalPaddingPx);
    }, 0);

    return rowHeightPx;
  });
}

function resolveTableRowHeightCss(
  row: TableNode["rows"][number],
  rowHeightPx?: number
): Record<string, string | number | undefined> | undefined {
  if (!Number.isFinite(rowHeightPx) || (rowHeightPx as number) <= 0) {
    return undefined;
  }

  const resolvedHeightPx = Math.max(
    MIN_PARAGRAPH_LINE_HEIGHT_PX,
    Math.round(rowHeightPx as number)
  );
  if (row.style?.heightRule === "exact") {
    return { height: `${resolvedHeightPx}px` };
  }

  // Use `height` rather than `min-height`: browsers ignore `min-height` on
  // table rows/cells, but on a table row `height` is treated as a MINIMUM
  // (the row still grows when content is taller). This is what makes an
  // explicit "at least" row height (e.g. a docx trHeight on a header band)
  // actually apply instead of collapsing to the text line height.
  return { height: `${resolvedHeightPx}px` };
}

interface TableCellSliceBoundaryLayout {
  safeBoundariesPx: number[];
  contentBottomPx: number;
}

function uniqueSortedPixelBoundaries(values: number[]): number[] {
  const sorted = values
    .filter((value) => Number.isFinite(value))
    .map((value) => Math.max(0, Math.round(value)))
    .sort((left, right) => left - right);
  const unique: number[] = [];
  for (const value of sorted) {
    const previous = unique[unique.length - 1];
    if (
      previous === undefined ||
      Math.abs(previous - value) > TABLE_ROW_SLICE_BOUNDARY_TOLERANCE_PX
    ) {
      unique.push(value);
    }
  }
  return unique;
}

function estimateParagraphBoundaryOffsetsPx(
  paragraph: ParagraphNode,
  availableWidthPx: number,
  numberingDefinitions: NumberingDefinitionSet | undefined,
  applyWordTableDefaults: boolean,
  docGridLinePitchPx: number | undefined,
  paragraphIndex: number
): {
  heightPx: number;
  safeBoundariesPx: number[];
} {
  const paragraphForLayout = wordLikeTableCellParagraph(
    paragraph,
    applyWordTableDefaults
  );
  const disableDocGridSnap = paragraphDocGridSnapState(paragraph) === "disable";
  const paragraphHeightPx = estimateParagraphHeightPx(
    paragraphForLayout,
    availableWidthPx,
    numberingDefinitions,
    docGridLinePitchPx,
    disableDocGridSnap
  );
  const suppressTopSpacing =
    paragraphIndex === 0 &&
    suppressFirstTableCellParagraphTopSpacing(paragraph);
  const beforeSpacingPx = suppressTopSpacing
    ? 0
    : twipsToPixels(paragraphForLayout.style?.spacing?.beforeTwips) ?? 0;
  const afterSpacingPx =
    twipsToPixels(paragraphForLayout.style?.spacing?.afterTwips) ?? 0;
  const topBorderInsetPx = paragraphBorderInsetPx(
    paragraphForLayout.style?.borders?.top
  );
  const bottomBorderInsetPx = paragraphBorderInsetPx(
    paragraphForLayout.style?.borders?.bottom
  );
  const lineHeightPx = Math.max(
    MIN_PARAGRAPH_LINE_HEIGHT_PX,
    estimateParagraphLineHeightPx(
      paragraphForLayout,
      docGridLinePitchPx,
      disableDocGridSnap
    )
  );
  const pretextSource = buildParagraphPretextLayoutSource(paragraphForLayout, {
    allowExplicitLineBreakText: true,
    expandTabsForLayout: true,
  });
  const paragraphTextWidthPx = paragraphAvailableTextWidthPx(
    paragraphForLayout,
    availableWidthPx,
    numberingDefinitions
  );
  const pretextLayout = pretextSource
    ? layoutParagraphPretextSource(
        paragraphForLayout,
        pretextSource,
        paragraphTextWidthPx,
        lineHeightPx,
        []
      )
    : undefined;
  const lineTopOffsetsPx = pretextLayout
    ? pretextLayout.lines.map((line) => Math.max(0, Math.round(line.y)))
    : Array.from(
        {
          length: Math.max(
            1,
            paragraphLineCountWithinWidth(
              paragraphForLayout,
              availableWidthPx,
              numberingDefinitions
            )
          ),
        },
        (_, lineIndex) => lineIndex * lineHeightPx
      );
  const textTopPx = beforeSpacingPx + topBorderInsetPx;
  const textHeightPx = pretextLayout
    ? wrappedPretextParagraphBlockHeightPx(pretextLayout)
    : lineTopOffsetsPx.length * lineHeightPx;
  const visualHeightPx = Math.max(
    1,
    beforeSpacingPx +
      topBorderInsetPx +
      textHeightPx +
      bottomBorderInsetPx +
      afterSpacingPx
  );
  const heightPx = Math.max(1, paragraphHeightPx, visualHeightPx);
  const lineBoundariesPx = lineTopOffsetsPx.map(
    (lineTopPx) => textTopPx + lineTopPx + lineHeightPx
  );

  return {
    heightPx,
    safeBoundariesPx: uniqueSortedPixelBoundaries([
      ...lineBoundariesPx,
      heightPx,
    ]),
  };
}

function estimateNestedTableBoundaryOffsetsPx(
  table: TableNode,
  availableWidthPx: number,
  numberingDefinitions: NumberingDefinitionSet | undefined,
  docGridLinePitchPx: number | undefined
): {
  heightPx: number;
  safeBoundariesPx: number[];
} {
  const rowHeightsPx = estimateTableRowHeightsPx(
    table,
    availableWidthPx,
    numberingDefinitions,
    docGridLinePitchPx
  );
  const boundariesPx: number[] = [];
  let cursorPx = 0;
  for (const rowHeightPx of rowHeightsPx) {
    cursorPx += Math.max(MIN_PARAGRAPH_LINE_HEIGHT_PX, rowHeightPx);
    boundariesPx.push(cursorPx);
  }

  return {
    heightPx: Math.max(MIN_PARAGRAPH_LINE_HEIGHT_PX, cursorPx),
    safeBoundariesPx: uniqueSortedPixelBoundaries(boundariesPx),
  };
}

function estimateTableCellSliceBoundaryLayoutPx(params: {
  cell: TableNode["rows"][number]["cells"][number];
  rowHeightPx: number;
  contentWidthPx: number;
  tableCellMarginTwips?: TableSpacingTwips;
  numberingDefinitions?: NumberingDefinitionSet;
  applyWordTableDefaults: boolean;
  docGridLinePitchPx?: number;
}): TableCellSliceBoundaryLayout {
  const {
    cell,
    rowHeightPx,
    contentWidthPx,
    tableCellMarginTwips,
    numberingDefinitions,
    applyWordTableDefaults,
    docGridLinePitchPx,
  } = params;
  const paddingPx = resolveTableSpacingPaddingPx(
    mergeTableSpacing(tableCellMarginTwips, cell.style?.marginTwips)
  );
  const localBoundariesPx = [0, paddingPx.top];
  let contentCursorPx = paddingPx.top;
  let paragraphIndex = 0;

  for (const contentNode of cell.nodes) {
    const layout =
      contentNode.type === "paragraph"
        ? estimateParagraphBoundaryOffsetsPx(
            contentNode,
            contentWidthPx,
            numberingDefinitions,
            applyWordTableDefaults,
            docGridLinePitchPx,
            paragraphIndex++
          )
        : estimateNestedTableBoundaryOffsetsPx(
            contentNode,
            contentWidthPx,
            numberingDefinitions,
            docGridLinePitchPx
          );

    localBoundariesPx.push(
      ...layout.safeBoundariesPx.map(
        (boundaryPx) => contentCursorPx + boundaryPx
      )
    );
    contentCursorPx += layout.heightPx;
  }

  const contentBottomPx = contentCursorPx + paddingPx.bottom;
  const contentFlowHeightPx = Math.max(0, contentCursorPx - paddingPx.top);
  const availableContentHeightPx = Math.max(
    0,
    rowHeightPx - paddingPx.top - paddingPx.bottom
  );
  const extraVerticalSpacePx = Math.max(
    0,
    availableContentHeightPx - contentFlowHeightPx
  );
  const verticalOffsetPx =
    cell.style?.verticalAlign === "center"
      ? Math.round(extraVerticalSpacePx / 2)
      : cell.style?.verticalAlign === "bottom"
      ? extraVerticalSpacePx
      : 0;

  return {
    safeBoundariesPx: uniqueSortedPixelBoundaries(
      localBoundariesPx.map((boundaryPx) =>
        Math.min(rowHeightPx, boundaryPx + verticalOffsetPx)
      )
    ),
    contentBottomPx: Math.min(rowHeightPx, contentBottomPx + verticalOffsetPx),
  };
}

function tableCellSliceBoundaryIsSafe(
  layout: TableCellSliceBoundaryLayout,
  boundaryPx: number
): boolean {
  if (boundaryPx <= TABLE_ROW_SLICE_BOUNDARY_TOLERANCE_PX) {
    return true;
  }

  if (
    boundaryPx >=
    layout.contentBottomPx - TABLE_ROW_SLICE_BOUNDARY_TOLERANCE_PX
  ) {
    return true;
  }

  return layout.safeBoundariesPx.some(
    (safeBoundaryPx) =>
      Math.abs(safeBoundaryPx - boundaryPx) <=
      TABLE_ROW_SLICE_BOUNDARY_TOLERANCE_PX
  );
}

function resolveTableRowSliceHeightOnSafeBoundaryPx(params: {
  table: TableNode;
  rowIndex: number;
  rowHeightPx: number;
  rowSliceOffsetPx: number;
  preferredSliceHeightPx: number;
  maxAvailableWidthPx?: number;
  numberingDefinitions?: NumberingDefinitionSet;
  docGridLinePitchPx?: number;
}): number | undefined {
  const {
    table,
    rowIndex,
    rowHeightPx,
    rowSliceOffsetPx,
    preferredSliceHeightPx,
    maxAvailableWidthPx,
    numberingDefinitions,
    docGridLinePitchPx,
  } = params;
  const row = table.rows[rowIndex];
  if (!row || !rowHasNestedTableContent(row)) {
    return preferredSliceHeightPx;
  }

  const sliceStartPx = Math.max(0, Math.round(rowSliceOffsetPx));
  const preferredSliceEndPx = Math.min(
    rowHeightPx,
    sliceStartPx + Math.max(0, Math.round(preferredSliceHeightPx))
  );
  if (
    preferredSliceEndPx >=
    rowHeightPx - TABLE_ROW_SLICE_BOUNDARY_TOLERANCE_PX
  ) {
    return Math.max(0, rowHeightPx - sliceStartPx);
  }

  const columnCount = tableColumnCount(table);
  const tableWidthPx = twipsToPixels(table.style?.widthTwips);
  const rawTableColumnWidthsPx = (() => {
    const definedWidthsTwips = columnWidthsFromTableDefinition(
      table,
      columnCount
    );
    if (!definedWidthsTwips || definedWidthsTwips.length === 0) {
      return defaultColumnWidthsPx(columnCount, tableWidthPx);
    }

    const widthsPx = definedWidthsTwips.map(
      (widthTwips) => twipsToPixels(widthTwips) ?? 0
    );
    return normalizeColumnWidthsPx(widthsPx, columnCount, tableWidthPx, 1);
  })();
  const rawResolvedTableWidthPx =
    tableWidthPx ??
    rawTableColumnWidthsPx.reduce((sum, widthPx) => sum + widthPx, 0);
  const collapsedHorizontalBorderBleedPx =
    resolveCollapsedTableHorizontalOuterBleedPx(table, columnCount);
  const maxTableWidthPx =
    Number.isFinite(maxAvailableWidthPx) && (maxAvailableWidthPx as number) > 0
      ? Math.max(
          120,
          (maxAvailableWidthPx as number) - collapsedHorizontalBorderBleedPx
        )
      : undefined;
  const resolvedTableWidthPx = clampTableWidthPx(
    rawResolvedTableWidthPx,
    maxTableWidthPx
  );
  const tableColumnWidthsPx = fitColumnWidthsToWidth(
    rawTableColumnWidthsPx,
    resolvedTableWidthPx
  );
  const applyWordTableDefaults = tableUsesWordLikeParagraphDefaults(table);
  const tableCellMarginTwips = table.style?.cellMarginTwips;
  const cellLayouts: TableCellSliceBoundaryLayout[] = [];
  const candidateBoundariesPx = [preferredSliceEndPx];
  let columnCursor = 0;

  for (const cell of row.cells) {
    const colSpanValue =
      cell.style?.gridSpan && cell.style.gridSpan > 1 ? cell.style.gridSpan : 1;
    const startColumnIndex = columnCursor;
    const endColumnIndex = Math.min(
      columnCount - 1,
      startColumnIndex + colSpanValue - 1
    );
    columnCursor += colSpanValue;
    const spannedWidthPx = tableColumnWidthsPx
      .slice(startColumnIndex, endColumnIndex + 1)
      .reduce((sum, widthPx) => sum + widthPx, 0);
    const fallbackCellWidthPx =
      (resolvedTableWidthPx / Math.max(1, columnCount)) * colSpanValue;
    const cellRenderedWidthPx =
      twipsToPixels(cell.style?.widthTwips) ??
      (spannedWidthPx > 0 ? spannedWidthPx : fallbackCellWidthPx);
    const cellPaddingPx = resolveTableSpacingPaddingPx(
      mergeTableSpacing(tableCellMarginTwips, cell.style?.marginTwips)
    );
    const cellContentWidthPx = Math.max(
      1,
      cellRenderedWidthPx - cellPaddingPx.left - cellPaddingPx.right
    );
    const cellLayout = estimateTableCellSliceBoundaryLayoutPx({
      cell,
      rowHeightPx,
      contentWidthPx: cellContentWidthPx,
      tableCellMarginTwips,
      numberingDefinitions,
      applyWordTableDefaults,
      docGridLinePitchPx,
    });
    cellLayouts.push(cellLayout);
    candidateBoundariesPx.push(...cellLayout.safeBoundariesPx);
  }

  const minimumSliceEndPx =
    sliceStartPx + Math.max(1, MIN_TABLE_ROW_SLICE_REMAINING_HEIGHT_PX);
  const candidatesPx = uniqueSortedPixelBoundaries(candidateBoundariesPx)
    .filter(
      (boundaryPx) =>
        boundaryPx >= minimumSliceEndPx &&
        boundaryPx <=
          preferredSliceEndPx + TABLE_ROW_SLICE_BOUNDARY_TOLERANCE_PX
    )
    .sort((left, right) => right - left);

  for (const candidatePx of candidatesPx) {
    if (
      cellLayouts.every((layout) =>
        tableCellSliceBoundaryIsSafe(layout, candidatePx)
      )
    ) {
      return Math.max(0, candidatePx - sliceStartPx);
    }
  }

  return undefined;
}

function estimateTableHeightPx(
  table: TableNode,
  maxAvailableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number
): number {
  const sourceXml = table.sourceXml;
  const widthKey = heightEstimateCacheKeyPx(
    maxAvailableWidthPx,
    docGridLinePitchPx
  );
  if (sourceXml) {
    const cachedByWidth = tableEstimatedHeightBySourceXml.get(sourceXml);
    const cached = cachedByWidth?.get(widthKey);
    if (cached !== undefined) {
      return cached;
    }
  }

  const estimatedRowsHeightPx = estimateTableRowHeightsPx(
    table,
    maxAvailableWidthPx,
    numberingDefinitions,
    docGridLinePitchPx
  ).reduce(
    (sum, rowHeightPx) =>
      sum + Math.max(MIN_PARAGRAPH_LINE_HEIGHT_PX, rowHeightPx),
    0
  );

  const estimatedHeightPx = Math.max(
    MIN_PARAGRAPH_LINE_HEIGHT_PX * 2,
    estimatedRowsHeightPx
  );
  if (sourceXml) {
    const cachedByWidth =
      tableEstimatedHeightBySourceXml.get(sourceXml) ??
      new Map<number, number>();
    cachedByWidth.set(widthKey, estimatedHeightPx);
    setCacheEntry(tableEstimatedHeightBySourceXml, sourceXml, cachedByWidth);
  }
  return estimatedHeightPx;
}

function estimateDocNodeHeightPx(
  node: DocModel["nodes"][number],
  availableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number
): number {
  return node.type === "paragraph"
    ? estimateParagraphHeightPx(
        node,
        availableWidthPx,
        numberingDefinitions,
        docGridLinePitchPx
      )
    : estimateTableHeightPx(
        node,
        availableWidthPx,
        numberingDefinitions,
        docGridLinePitchPx
      );
}

function paragraphBeforeSpacingPx(paragraph: ParagraphNode): number {
  return twipsToPixels(paragraph.style?.spacing?.beforeTwips) ?? 0;
}

function paragraphAfterSpacingPx(paragraph: ParagraphNode): number {
  return twipsToPixels(paragraph.style?.spacing?.afterTwips) ?? 0;
}

function paragraphWidowControlEnabled(paragraph: ParagraphNode): boolean {
  return paragraph.style?.widowControl !== false;
}

function paragraphIsOnlyExplicitPageBreak(paragraph: ParagraphNode): boolean {
  if (!paragraphHasExplicitPageBreak(paragraph)) {
    return false;
  }

  return (
    !paragraphHasVisibleText(paragraph) &&
    !paragraphHasImage(paragraph) &&
    !paragraphHasFormField(paragraph)
  );
}

function resolveParagraphBeforeSpacingPx(
  model: DocModel,
  nodeIndex: number,
  paragraph: ParagraphNode,
  pageConsumedHeightPx: number,
  suppressSpacingBeforeAfterPageBreak: boolean
): number {
  const beforeSpacingPx = paragraphBeforeSpacingPx(paragraph);
  if (
    !suppressSpacingBeforeAfterPageBreak ||
    pageConsumedHeightPx > 0 ||
    nodeIndex <= 0
  ) {
    return beforeSpacingPx;
  }

  const previousNode = model.nodes[nodeIndex - 1];
  if (previousNode?.type !== "paragraph") {
    return beforeSpacingPx;
  }

  // ECMA-376 §2.15.3.49: when enabled, suppress before-spacing on first content
  // line following a page-break-only paragraph.
  return paragraphIsOnlyExplicitPageBreak(previousNode) ? 0 : beforeSpacingPx;
}

function paragraphCanSplitAcrossPages(
  paragraph: ParagraphNode,
  lineCount: number,
  options?: {
    allowKeepLinesOverflow?: boolean;
    allowKeepNextOverflow?: boolean;
    allowImageParagraphSplit?: boolean;
  }
): boolean {
  if (lineCount < 2) {
    return false;
  }

  if (paragraph.style?.keepLines === true && !options?.allowKeepLinesOverflow) {
    return false;
  }

  if (paragraph.style?.keepNext === true && !options?.allowKeepNextOverflow) {
    return false;
  }

  if (paragraphHasImage(paragraph) && !options?.allowImageParagraphSplit) {
    return false;
  }

  if (paragraphHasFormField(paragraph)) {
    return false;
  }

  return true;
}

function collectDocxEstimatedOverflowBreakStartNodeIndexes(
  model: DocModel,
  hardBreakStartNodeIndexes: Set<number>,
  pageContentHeightPx: number,
  pageContentWidthPx: number,
  numberingDefinitions?: NumberingDefinitionSet,
  paginationMetricsBySection?: PaginationSectionMetrics[],
  options?: {
    suppressSpacingBeforeAfterPageBreak?: boolean;
  }
): Set<number> {
  const breaks = new Set<number>();
  if (!Number.isFinite(pageContentHeightPx) || pageContentHeightPx <= 0) {
    return breaks;
  }

  const fallbackMetrics: PaginationSectionMetrics = {
    startNodeIndex: 0,
    pageContentWidthPx: Math.max(120, Math.round(pageContentWidthPx)),
    pageContentHeightPx: Math.max(120, Math.round(pageContentHeightPx)),
    pageContentHeightMultiplier: 1,
    docGridLinePitchPx: undefined,
  };
  const metricsBySection = paginationMetricsBySection?.length
    ? paginationMetricsBySection
    : [fallbackMetrics];
  const sectionStartPageBreakNodeIndexes =
    collectDocxSectionStartPageBreakNodeIndexes(model);
  const nextHardBreakStartNodeIndexByNodeIndex =
    buildNextHardBreakStartNodeIndexLookup(
      model.nodes.length,
      hardBreakStartNodeIndexes
    );

  let pageConsumedHeightPx = 0;
  let previousParagraphAfterPx = 0;
  let currentMetricsIndex = 0;
  // Mirrors buildDocumentPageNodeSegments: once a keepNext chain head starts a
  // page, later chain members must not re-trigger a keep-induced break — that
  // would strand the head on a near-empty page.
  let committedKeepNextChainEndNodeIndex = -1;
  const suppressSpacingBeforeAfterPageBreak =
    options?.suppressSpacingBeforeAfterPageBreak ?? false;
  let currentPageContentHeightPx =
    metricsBySection[0]?.pageContentHeightPx ??
    fallbackMetrics.pageContentHeightPx;
  const projectConsumedHeightAcrossSectionMultipliers = (
    consumedHeightPx: number,
    fromMetrics: PaginationSectionMetrics | undefined,
    toMetrics: PaginationSectionMetrics | undefined
  ): number => {
    const safeConsumedHeightPx = Math.max(0, Math.round(consumedHeightPx));
    if (safeConsumedHeightPx <= 0) {
      return 0;
    }

    const fromMultiplier = Math.max(
      1,
      Math.round(fromMetrics?.pageContentHeightMultiplier ?? 1)
    );
    const toMultiplier = Math.max(
      1,
      Math.round(toMetrics?.pageContentHeightMultiplier ?? 1)
    );
    if (fromMultiplier === toMultiplier) {
      return safeConsumedHeightPx;
    }

    const approximateVisualDepthPx = safeConsumedHeightPx / fromMultiplier;
    return Math.max(0, Math.round(approximateVisualDepthPx * toMultiplier));
  };
  for (let nodeIndex = 0; nodeIndex < model.nodes.length; nodeIndex += 1) {
    const previousMetricsIndex = currentMetricsIndex;
    currentMetricsIndex = resolvePaginationSectionMetricsIndexForNodeIndex(
      metricsBySection,
      nodeIndex,
      currentMetricsIndex
    );
    const nodeMetrics =
      metricsBySection[currentMetricsIndex] ?? fallbackMetrics;
    if (nodeIndex > 0 && currentMetricsIndex !== previousMetricsIndex) {
      pageConsumedHeightPx = projectConsumedHeightAcrossSectionMultipliers(
        pageConsumedHeightPx,
        metricsBySection[previousMetricsIndex],
        nodeMetrics
      );
      currentPageContentHeightPx = nodeMetrics.pageContentHeightPx;
    }

    if (hardBreakStartNodeIndexes.has(nodeIndex)) {
      pageConsumedHeightPx = 0;
      previousParagraphAfterPx = 0;
      currentPageContentHeightPx = nodeMetrics.pageContentHeightPx;
    }

    const node = model.nodes[nodeIndex];
    if (
      node.type === "paragraph" &&
      (paragraphIsStructuralSectionBreakSpacer(node) ||
        paragraphActsAsSectionBreakCarryoverSpacer(model, nodeIndex, node))
    ) {
      previousParagraphAfterPx = 0;
      continue;
    }
    if (
      node.type === "paragraph" &&
      paragraphActsAsTrailingRenderedPageBreakSpacer(model, nodeIndex, node)
    ) {
      previousParagraphAfterPx = 0;
      continue;
    }
    if (
      node.type === "paragraph" &&
      paragraphCollapsesIntoPreviousParagraph(node, model.nodes[nodeIndex - 1])
    ) {
      continue;
    }
    const directNodeBeforeSpacingPx =
      node.type === "paragraph" ? paragraphBeforeSpacingPx(node) : 0;
    const directNodeAfterSpacingPx =
      node.type === "paragraph" ? paragraphAfterSpacingPx(node) : 0;
    const nodeBeforeSpacingPx =
      node.type === "paragraph"
        ? effectiveParagraphBeforeSpacingPx(
            model,
            nodeIndex,
            node,
            pageConsumedHeightPx,
            suppressSpacingBeforeAfterPageBreak
          )
        : 0;
    const nodeAfterSpacingPx =
      node.type === "paragraph"
        ? effectiveParagraphAfterSpacingPx(model, nodeIndex, node)
        : 0;
    const rawNodeHeightPx = Math.max(
      1,
      estimateDocNodeHeightPx(
        node,
        nodeMetrics.pageContentWidthPx,
        numberingDefinitions,
        nodeMetrics.docGridLinePitchPx
      ) -
        directNodeBeforeSpacingPx -
        directNodeAfterSpacingPx +
        nodeBeforeSpacingPx +
        nodeAfterSpacingPx
    );
    const collapsedMarginPx =
      node.type === "paragraph" && pageConsumedHeightPx > 0
        ? Math.min(previousParagraphAfterPx, nodeBeforeSpacingPx)
        : 0;
    const collapsedNodeHeightPx = Math.max(
      1,
      rawNodeHeightPx - collapsedMarginPx
    );

    let requiredHeightPx = collapsedNodeHeightPx;
    let keepNextChainEndNodeIndex = -1;

    if (
      node.type === "paragraph" &&
      node.style?.keepNext === true &&
      nodeIndex > committedKeepNextChainEndNodeIndex &&
      paragraphHasVisibleText(node)
    ) {
      let chainCursor = nodeIndex;
      let chainPreviousParagraphAfterPx = nodeAfterSpacingPx;
      while (chainCursor < model.nodes.length - 1) {
        const currentChainNode = model.nodes[chainCursor];
        if (
          currentChainNode.type !== "paragraph" ||
          currentChainNode.style?.keepNext !== true ||
          !paragraphHasVisibleText(currentChainNode)
        ) {
          break;
        }
        if (hardBreakStartNodeIndexes.has(chainCursor + 1)) {
          break;
        }
        const nextChainNode = model.nodes[chainCursor + 1];
        if (nextChainNode.type !== "paragraph") {
          break;
        }

        chainCursor += 1;
        const chainMetricsIndex =
          resolvePaginationSectionMetricsIndexForNodeIndex(
            metricsBySection,
            chainCursor,
            currentMetricsIndex
          );
        const chainMetrics =
          metricsBySection[chainMetricsIndex] ?? fallbackMetrics;
        const nextDirectBeforeSpacingPx =
          nextChainNode.type === "paragraph"
            ? paragraphBeforeSpacingPx(nextChainNode)
            : 0;
        const nextDirectAfterSpacingPx =
          nextChainNode.type === "paragraph"
            ? paragraphAfterSpacingPx(nextChainNode)
            : 0;
        const nextBeforeSpacingPx =
          nextChainNode.type === "paragraph"
            ? effectiveParagraphBeforeSpacingPx(
                model,
                chainCursor,
                nextChainNode,
                1,
                suppressSpacingBeforeAfterPageBreak
              )
            : 0;
        const nextAfterSpacingPx =
          nextChainNode.type === "paragraph"
            ? effectiveParagraphAfterSpacingPx(
                model,
                chainCursor,
                nextChainNode
              )
            : 0;
        const nextRawHeightPx = Math.max(
          1,
          estimateDocNodeHeightPx(
            nextChainNode,
            chainMetrics.pageContentWidthPx,
            numberingDefinitions,
            chainMetrics.docGridLinePitchPx
          ) -
            nextDirectBeforeSpacingPx -
            nextDirectAfterSpacingPx +
            nextBeforeSpacingPx +
            nextAfterSpacingPx
        );
        const collapsedChainMarginPx =
          nextChainNode.type === "paragraph"
            ? Math.min(chainPreviousParagraphAfterPx, nextBeforeSpacingPx)
            : 0;
        requiredHeightPx += Math.max(
          1,
          nextRawHeightPx - collapsedChainMarginPx
        );
        requiredHeightPx += keepNextPaginationReservePx(
          currentChainNode,
          nextChainNode,
          chainMetrics.docGridLinePitchPx
        );
        chainPreviousParagraphAfterPx = nextAfterSpacingPx;
      }
      keepNextChainEndNodeIndex = chainCursor;
    }

    const remainingHeightPx = currentPageContentHeightPx - pageConsumedHeightPx;
    const canKeepTrailingSectionTailOnCurrentPage =
      shouldKeepTrailingSectionTailOnCurrentPage(
        model,
        nodeIndex,
        pageConsumedHeightPx,
        previousParagraphAfterPx,
        nodeMetrics.pageContentWidthPx,
        currentPageContentHeightPx,
        hardBreakStartNodeIndexes,
        sectionStartPageBreakNodeIndexes,
        nextHardBreakStartNodeIndexByNodeIndex,
        (_candidateNodeIndex, candidateNode, candidatePageContentWidthPx) =>
          estimateDocNodeHeightPx(
            candidateNode,
            candidatePageContentWidthPx,
            numberingDefinitions,
            nodeMetrics.docGridLinePitchPx
          ),
        (candidateNodeIndex, candidateParagraph, consumedHeightPx) =>
          effectiveParagraphBeforeSpacingPx(
            model,
            candidateNodeIndex,
            candidateParagraph,
            consumedHeightPx,
            suppressSpacingBeforeAfterPageBreak
          ),
        (candidateNodeIndex, candidateParagraph) =>
          effectiveParagraphAfterSpacingPx(
            model,
            candidateNodeIndex,
            candidateParagraph
          )
      );
    if (
      pageConsumedHeightPx > 0 &&
      requiredHeightPx > remainingHeightPx + PAGE_OVERFLOW_TOLERANCE_PX &&
      !canKeepTrailingSectionTailOnCurrentPage
    ) {
      breaks.add(nodeIndex);
      pageConsumedHeightPx = 0;
      previousParagraphAfterPx = 0;
      currentPageContentHeightPx = nodeMetrics.pageContentHeightPx;
    }

    if (pageConsumedHeightPx === 0 && keepNextChainEndNodeIndex > nodeIndex) {
      committedKeepNextChainEndNodeIndex = keepNextChainEndNodeIndex;
    }
    const effectiveNodeHeightPx =
      pageConsumedHeightPx > 0 ? collapsedNodeHeightPx : rawNodeHeightPx;
    pageConsumedHeightPx += effectiveNodeHeightPx;
    previousParagraphAfterPx =
      node.type === "paragraph" ? nodeAfterSpacingPx : 0;
  }

  for (const breakIndex of [...breaks]) {
    if (
      breakIndex <= 0 ||
      breakIndex >= model.nodes.length ||
      hardBreakStartNodeIndexes.has(breakIndex)
    ) {
      breaks.delete(breakIndex);
    }
  }

  return breaks;
}

interface DocumentPageRange {
  startNodeIndex: number;
  endNodeIndex: number;
}

function buildDocumentPageRanges(
  nodeCount: number,
  pageBreakStartNodeIndexes: Set<number>
): DocumentPageRange[] {
  if (nodeCount <= 0) {
    return [];
  }

  const sortedBreakStartIndexes = [...pageBreakStartNodeIndexes]
    .filter((index) => index > 0 && index < nodeCount)
    .sort((left, right) => left - right);

  const ranges: DocumentPageRange[] = [];
  let startNodeIndex = 0;

  for (const breakStartIndex of sortedBreakStartIndexes) {
    if (breakStartIndex <= startNodeIndex) {
      continue;
    }

    ranges.push({
      startNodeIndex,
      endNodeIndex: breakStartIndex,
    });
    startNodeIndex = breakStartIndex;
  }

  ranges.push({
    startNodeIndex,
    endNodeIndex: nodeCount,
  });

  return ranges;
}

function collectDocxPageBreakStartNodeIndexes(
  model: DocModel,
  pageContentHeightPx: number,
  pageContentWidthPx: number,
  numberingDefinitions?: NumberingDefinitionSet,
  paginationMetricsBySection?: PaginationSectionMetrics[],
  options?: {
    suppressSpacingBeforeAfterPageBreak?: boolean;
  }
): Set<number> {
  const hardBreaks = collectDocxHardPageBreakStartNodeIndexes(model);
  const overflowBreaks = collectDocxEstimatedOverflowBreakStartNodeIndexes(
    model,
    hardBreaks,
    pageContentHeightPx,
    pageContentWidthPx,
    numberingDefinitions,
    paginationMetricsBySection,
    options
  );
  return new Set<number>([...hardBreaks, ...overflowBreaks]);
}

interface TableRowRange {
  startRowIndex: number;
  endRowIndex: number;
}

interface TableRowSlice {
  rowIndex: number;
  startOffsetPx: number;
  sliceHeightPx: number;
  totalRowHeightPx: number;
}

interface ParagraphLineRange {
  startLineIndex: number;
  endLineIndex: number;
  totalLineCount: number;
  lineHeightPx: number;
}

interface ParagraphSegmentIdentity {
  nodeIndex: number;
  startLineIndex: number;
  endLineIndex: number;
}

interface DocumentPageNodeSegment {
  nodeIndex: number;
  tableRowRange?: TableRowRange;
  tableRowSlice?: TableRowSlice;
  paragraphLineRange?: ParagraphLineRange;
}

function paragraphSegmentHasPartialLineRange(
  paragraphLineRange?: ParagraphLineRange
): boolean {
  if (!paragraphLineRange) {
    return false;
  }

  return (
    paragraphLineRange.startLineIndex > 0 ||
    paragraphLineRange.endLineIndex < paragraphLineRange.totalLineCount
  );
}

export function resolveLineRangeWithinVerticalSlice(
  lineTopOffsetsPx: number[],
  lineHeightPx: number,
  sliceTopPx: number,
  sliceBottomPx: number
): ParagraphLineRange | undefined {
  if (
    lineTopOffsetsPx.length === 0 ||
    !Number.isFinite(lineHeightPx) ||
    lineHeightPx <= 0
  ) {
    return undefined;
  }

  const safeSliceTopPx = Math.max(0, sliceTopPx);
  const safeSliceBottomPx = Math.max(safeSliceTopPx, sliceBottomPx);
  const sliceHasHeight = safeSliceBottomPx > safeSliceTopPx;
  let startLineIndex: number | undefined;
  let endLineIndex: number | undefined;

  for (let lineIndex = 0; lineIndex < lineTopOffsetsPx.length; lineIndex += 1) {
    const lineTopPx = lineTopOffsetsPx[lineIndex] ?? lineIndex * lineHeightPx;
    const lineBottomPx = lineTopPx + lineHeightPx;
    const lineBelongsToSlice =
      sliceHasHeight &&
      lineBottomPx > safeSliceTopPx + PAGE_OVERFLOW_TOLERANCE_PX &&
      lineBottomPx <= safeSliceBottomPx + PAGE_OVERFLOW_TOLERANCE_PX;
    if (lineBelongsToSlice) {
      if (startLineIndex === undefined) {
        startLineIndex = lineIndex;
      }
      endLineIndex = lineIndex + 1;
    }
  }

  if (
    startLineIndex === undefined ||
    endLineIndex === undefined ||
    endLineIndex <= startLineIndex
  ) {
    return undefined;
  }

  return {
    startLineIndex,
    endLineIndex,
    totalLineCount: lineTopOffsetsPx.length,
    lineHeightPx,
  };
}

export function resolveTableCellParagraphVisualBottomPx(params: {
  paragraphTopPx: number;
  paragraphHeightPx: number;
  textBottomPx: number;
}): number {
  return Math.max(
    Math.round(params.paragraphTopPx + params.paragraphHeightPx),
    Math.round(params.textBottomPx)
  );
}

export function tableCellParagraphFitsFullyWithinSlice(params: {
  sliceStartPx: number;
  sliceBottomPx: number;
  paragraphTopPx: number;
  paragraphBottomPx: number;
}): boolean {
  return (
    params.sliceStartPx <= params.paragraphTopPx + PAGE_OVERFLOW_TOLERANCE_PX &&
    params.sliceBottomPx >=
      params.paragraphBottomPx +
        TABLE_CELL_SLICE_FULLY_VISIBLE_BOTTOM_BUFFER_PX -
        PAGE_OVERFLOW_TOLERANCE_PX
  );
}

export function resolveParagraphSegmentClipBleedPx(
  paragraphLineRange?: ParagraphLineRange
): {
  topPx: number;
  bottomPx: number;
} {
  if (!paragraphSegmentHasPartialLineRange(paragraphLineRange)) {
    return {
      topPx: 0,
      bottomPx: 0,
    };
  }

  return {
    topPx:
      paragraphLineRange && paragraphLineRange.startLineIndex > 0
        ? Math.max(0, PARAGRAPH_SEGMENT_TOP_BLEED_PX)
        : 0,
    bottomPx:
      paragraphLineRange &&
      paragraphLineRange.endLineIndex < paragraphLineRange.totalLineCount
        ? Math.max(0, PARAGRAPH_SEGMENT_DESCENDER_BLEED_PX)
        : 0,
  };
}

export function resolveFallbackParagraphSegmentClipBleedPx(
  paragraph: ParagraphNode,
  paragraphLineRange?: ParagraphLineRange
): {
  topPx: number;
  bottomPx: number;
} {
  if (!paragraphSegmentHasPartialLineRange(paragraphLineRange)) {
    return {
      topPx: 0,
      bottomPx: 0,
    };
  }

  const lineHeightPx = Math.max(1, paragraphLineRange?.lineHeightPx ?? 0);
  const maxFontSizePx = paragraphMaxFontSizePx(paragraph);
  const glyphOvershootPx = Math.max(
    0,
    Math.ceil((maxFontSizePx - lineHeightPx) / 2)
  );
  const ascenderSafetyPx = Math.max(
    glyphOvershootPx,
    Math.ceil(lineHeightPx * 0.22)
  );

  return {
    topPx:
      paragraphLineRange && paragraphLineRange.startLineIndex > 0
        ? Math.min(
            PARAGRAPH_SEGMENT_FALLBACK_TOP_BLEED_MAX_PX,
            ascenderSafetyPx
          )
        : 0,
    bottomPx:
      paragraphLineRange &&
      paragraphLineRange.endLineIndex < paragraphLineRange.totalLineCount
        ? Math.min(
            PARAGRAPH_SEGMENT_FALLBACK_BOTTOM_BLEED_MAX_PX,
            glyphOvershootPx
          )
        : 0,
  };
}

export function resolveParagraphSegmentNonFlowReservePx(
  paragraphLineRange?: ParagraphLineRange
): number {
  const bleed = resolveParagraphSegmentClipBleedPx(paragraphLineRange);
  if (bleed.topPx <= 0 && bleed.bottomPx <= 0) {
    return 0;
  }

  const lineHeightSafetyPx = Math.max(
    0,
    Math.ceil((paragraphLineRange?.lineHeightPx ?? 0) * 0.9)
  );
  return (
    Math.max(0, bleed.topPx) +
    Math.max(0, bleed.bottomPx) +
    Math.max(0, PARAGRAPH_SEGMENT_VISUAL_SAFETY_PX, lineHeightSafetyPx)
  );
}

function resolveFallbackParagraphSegmentNonFlowReservePx(
  paragraph: ParagraphNode,
  paragraphLineRange?: ParagraphLineRange
): number {
  const bleed = resolveFallbackParagraphSegmentClipBleedPx(
    paragraph,
    paragraphLineRange
  );
  if (bleed.topPx <= 0 && bleed.bottomPx <= 0) {
    return 0;
  }

  const lineHeightSafetyPx = Math.max(
    0,
    Math.ceil((paragraphLineRange?.lineHeightPx ?? 0) * 0.15)
  );
  return (
    Math.max(0, bleed.topPx) +
    Math.max(0, bleed.bottomPx) +
    Math.max(1, PARAGRAPH_SEGMENT_FALLBACK_VISUAL_SAFETY_PX, lineHeightSafetyPx)
  );
}

function paragraphSegmentIdentityMatches(
  segment: ParagraphSegmentIdentity | undefined,
  nodeIndex: number,
  paragraphLineRange?: ParagraphLineRange
): boolean {
  if (!segment || !paragraphLineRange) {
    return false;
  }

  return (
    segment.nodeIndex === nodeIndex &&
    segment.startLineIndex === paragraphLineRange.startLineIndex &&
    segment.endLineIndex === paragraphLineRange.endLineIndex
  );
}

function estimateRenderedPageSegmentHeightPx(
  node: DocModel["nodes"][number],
  segment: DocumentPageNodeSegment,
  model: DocModel,
  availableWidthPx: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number,
  options?: {
    excludeWrappedFloatingImageFootprint?: boolean;
  }
): number {
  if (node.type === "paragraph") {
    const paragraphLineRange = segment.paragraphLineRange;
    if (paragraphLineRange) {
      const beforeSpacingPx =
        paragraphLineRange.startLineIndex === 0
          ? effectiveParagraphBeforeSpacingPx(
              model,
              segment.nodeIndex,
              node,
              segment.nodeIndex > 0 ? 1 : 0,
              false
            )
          : 0;
      const afterSpacingPx =
        paragraphLineRange.endLineIndex >= paragraphLineRange.totalLineCount
          ? effectiveParagraphAfterSpacingPx(model, segment.nodeIndex, node)
          : 0;
      const paragraphPretextSource = buildParagraphPretextLayoutSource(node, {
        allowExplicitLineBreakText: true,
        expandTabsForLayout: true,
      });
      const paragraphPretextLayout = paragraphPretextSource
        ? layoutParagraphPretextSource(
            node,
            paragraphPretextSource,
            paragraphAvailableTextWidthPx(
              node,
              availableWidthPx,
              numberingDefinitions
            ),
            Math.max(1, paragraphLineRange.lineHeightPx),
            []
          )
        : undefined;
      const segmentContentHeightPx =
        paragraphPretextLayout && paragraphPretextLayout.lineCount > 0
          ? resolvePretextLineRangeContentHeightPx(
              paragraphPretextLayout,
              paragraphLineRange.startLineIndex,
              paragraphLineRange.endLineIndex
            )
          : Math.max(
              1,
              paragraphLineRange.endLineIndex -
                paragraphLineRange.startLineIndex
            ) * Math.max(1, paragraphLineRange.lineHeightPx);
      return Math.max(
        1,
        beforeSpacingPx + segmentContentHeightPx + afterSpacingPx
      );
    }

    return Math.max(
      1,
      estimateParagraphHeightPx(
        node,
        availableWidthPx,
        numberingDefinitions,
        docGridLinePitchPx,
        false,
        undefined,
        options?.excludeWrappedFloatingImageFootprint
          ? { excludeWrappedFloatingImageFootprint: true }
          : undefined
      )
    );
  }

  if (segment.tableRowRange) {
    if (segment.tableRowSlice) {
      return Math.max(
        MIN_PARAGRAPH_LINE_HEIGHT_PX,
        Math.round(segment.tableRowSlice.sliceHeightPx)
      );
    }
    const rowHeightsPx = estimateTableRowHeightsPx(
      node,
      availableWidthPx,
      numberingDefinitions,
      docGridLinePitchPx
    );
    return Math.max(
      MIN_PARAGRAPH_LINE_HEIGHT_PX,
      sumEstimatedTableRowHeightsPx(
        rowHeightsPx,
        segment.tableRowRange.startRowIndex,
        segment.tableRowRange.endRowIndex
      )
    );
  }

  return Math.max(
    MIN_PARAGRAPH_LINE_HEIGHT_PX,
    estimateTableHeightPx(
      node,
      availableWidthPx,
      numberingDefinitions,
      docGridLinePitchPx
    )
  );
}

function resolveParagraphColumnRenderLineRange(
  paragraph: ParagraphNode,
  segment: DocumentPageNodeSegment,
  availableWidthPx: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPx?: number
): ParagraphLineRange {
  const lineHeightPx = Math.max(
    1,
    segment.paragraphLineRange?.lineHeightPx ??
      estimateParagraphLineHeightPx(paragraph, docGridLinePitchPx)
  );
  if (segment.paragraphLineRange) {
    return {
      ...segment.paragraphLineRange,
      lineHeightPx,
    };
  }

  const paragraphTextWidthPx = paragraphAvailableTextWidthPx(
    paragraph,
    availableWidthPx,
    numberingDefinitions
  );
  const pretextSource = buildParagraphPretextLayoutSource(paragraph, {
    allowExplicitLineBreakText: true,
    expandTabsForLayout: true,
  });
  const pretextLayout = pretextSource
    ? layoutParagraphPretextSource(
        paragraph,
        pretextSource,
        paragraphTextWidthPx,
        lineHeightPx,
        []
      )
    : undefined;
  const totalLineCount =
    pretextLayout && pretextLayout.lineCount > 0
      ? pretextLayout.lineCount
      : paragraphLineCountWithinWidth(
          paragraph,
          availableWidthPx,
          numberingDefinitions
        );

  return {
    startLineIndex: 0,
    endLineIndex: Math.max(1, totalLineCount),
    totalLineCount: Math.max(1, totalLineCount),
    lineHeightPx,
  };
}

function splitParagraphSegmentForColumnRender(params: {
  paragraph: ParagraphNode;
  segment: DocumentPageNodeSegment;
  model: DocModel;
  availableWidthPx: number;
  availableHeightPx: number;
  numberingDefinitions?: NumberingDefinitionSet;
  docGridLinePitchPx?: number;
}):
  | {
      currentSegment: DocumentPageNodeSegment;
      currentHeightPx: number;
      remainderSegment: DocumentPageNodeSegment;
    }
  | undefined {
  const {
    paragraph,
    segment,
    model,
    availableWidthPx,
    availableHeightPx,
    numberingDefinitions,
    docGridLinePitchPx,
  } = params;
  if (
    segment.tableRowRange ||
    segment.tableRowSlice ||
    paragraphHasExplicitColumnBreak(paragraph)
  ) {
    return undefined;
  }

  const fullLineRange = resolveParagraphColumnRenderLineRange(
    paragraph,
    segment,
    availableWidthPx,
    numberingDefinitions,
    docGridLinePitchPx
  );
  const startLineIndex = Math.max(0, fullLineRange.startLineIndex);
  const endLineIndex = Math.max(startLineIndex, fullLineRange.endLineIndex);
  if (
    endLineIndex - startLineIndex < 2 ||
    !paragraphCanSplitAcrossPages(paragraph, fullLineRange.totalLineCount)
  ) {
    return undefined;
  }

  const safeAvailableHeightPx = Math.max(0, Math.round(availableHeightPx));
  let bestSegment: DocumentPageNodeSegment | undefined;
  let bestHeightPx = 0;

  for (
    let candidateEndLineIndex = startLineIndex + 1;
    candidateEndLineIndex < endLineIndex;
    candidateEndLineIndex += 1
  ) {
    const candidateSegment: DocumentPageNodeSegment = {
      ...segment,
      paragraphLineRange: {
        startLineIndex,
        endLineIndex: candidateEndLineIndex,
        totalLineCount: fullLineRange.totalLineCount,
        lineHeightPx: fullLineRange.lineHeightPx,
      },
    };
    const candidateHeightPx = estimateRenderedPageSegmentHeightPx(
      paragraph,
      candidateSegment,
      model,
      availableWidthPx,
      numberingDefinitions,
      docGridLinePitchPx
    );
    if (
      candidateHeightPx >
      safeAvailableHeightPx + PAGE_OVERFLOW_TOLERANCE_PX
    ) {
      break;
    }

    bestSegment = candidateSegment;
    bestHeightPx = candidateHeightPx;
  }

  if (!bestSegment?.paragraphLineRange) {
    return undefined;
  }

  return {
    currentSegment: bestSegment,
    currentHeightPx: bestHeightPx,
    remainderSegment: {
      ...segment,
      paragraphLineRange: {
        startLineIndex: bestSegment.paragraphLineRange.endLineIndex,
        endLineIndex,
        totalLineCount: fullLineRange.totalLineCount,
        lineHeightPx: fullLineRange.lineHeightPx,
      },
    },
  };
}

export function buildRenderColumnSegmentsForPageSection(
  model: DocModel,
  flowSegments: DocumentPageNodeSegment[],
  columnWidthsPx: number[],
  columnHeightPx: number,
  numberingDefinitions?: NumberingDefinitionSet,
  docGridLinePitchPxByNodeIndex?: Map<number, number | undefined>,
  measuredParagraphOuterHeightsPxByNodeIndex?: Map<number, number>,
  balanceColumns = false,
  forceColumnBreakNodeIndexes?: Set<number>
): DocumentPageNodeSegment[][] {
  const columnCount = Math.max(1, columnWidthsPx.length);
  const columns = Array.from(
    { length: columnCount },
    () => [] as DocumentPageNodeSegment[]
  );
  const maxColumnHeightPx = Math.max(120, Math.round(columnHeightPx));
  const resolveSegmentHeightPx = (
    segment: DocumentPageNodeSegment,
    columnWidthPx: number
  ): number => {
    const segmentNode = model.nodes[segment.nodeIndex];
    if (!segmentNode) {
      return MIN_PARAGRAPH_LINE_HEIGHT_PX;
    }

    const docGridLinePitchPx = docGridLinePitchPxByNodeIndex?.get(
      segment.nodeIndex
    );
    const measuredSegmentHeightPx =
      segmentNode.type === "paragraph" &&
      !segment.paragraphLineRange &&
      !segment.tableRowRange &&
      !segment.tableRowSlice
        ? measuredParagraphOuterHeightsPxByNodeIndex?.get(segment.nodeIndex)
        : undefined;
    return Number.isFinite(measuredSegmentHeightPx) &&
      (measuredSegmentHeightPx as number) > 0
      ? Math.max(1, Math.round(measuredSegmentHeightPx as number))
      : estimateRenderedPageSegmentHeightPx(
          segmentNode,
          segment,
          model,
          columnWidthPx,
          numberingDefinitions,
          docGridLinePitchPx
        );
  };
  const safeColumnHeightPx =
    balanceColumns && columnCount > 1
      ? Math.min(
          maxColumnHeightPx,
          Math.max(
            MIN_PARAGRAPH_LINE_HEIGHT_PX * 4,
            Math.ceil(
              flowSegments.reduce((totalHeightPx, segment) => {
                const columnWidthPx = Math.max(
                  120,
                  Math.round(columnWidthsPx[0] ?? 120)
                );
                return (
                  totalHeightPx + resolveSegmentHeightPx(segment, columnWidthPx)
                );
              }, 0) / columnCount
            ) + PAGE_OVERFLOW_TOLERANCE_PX
          )
        )
      : maxColumnHeightPx;
  let columnIndex = 0;
  let consumedHeightPx = 0;

  const moveToNextColumn = (): boolean => {
    if (columnIndex + 1 >= columnCount) {
      return false;
    }

    columnIndex += 1;
    consumedHeightPx = 0;
    return true;
  };

  const pushSegment = (
    segment: DocumentPageNodeSegment,
    heightPx: number
  ): void => {
    columns[columnIndex]?.push(segment);
    consumedHeightPx += Math.max(1, Math.round(heightPx));
  };

  for (const flowSegment of flowSegments) {
    // Explicit "nextColumn" section breaks force a column advance regardless
    // of how much room is left in the current column.
    const isNodeStartSegment =
      (flowSegment.paragraphLineRange?.startLineIndex ?? 0) === 0 &&
      (flowSegment.tableRowRange?.startRowIndex ?? 0) === 0 &&
      !flowSegment.tableRowSlice;
    if (
      isNodeStartSegment &&
      forceColumnBreakNodeIndexes?.has(flowSegment.nodeIndex) &&
      consumedHeightPx > 0
    ) {
      moveToNextColumn();
    }
    let pendingSegment: DocumentPageNodeSegment | undefined = flowSegment;
    let splitGuard = 0;

    while (pendingSegment && splitGuard < 256) {
      splitGuard += 1;
      const currentSegment: DocumentPageNodeSegment = pendingSegment;
      const segmentNode: DocModel["nodes"][number] | undefined =
        model.nodes[currentSegment.nodeIndex];
      if (!segmentNode) {
        columns[columnIndex]?.push(currentSegment);
        break;
      }

      const columnWidthPx = Math.max(
        120,
        Math.round(columnWidthsPx[columnIndex] ?? columnWidthsPx[0] ?? 120)
      );
      const docGridLinePitchPx = docGridLinePitchPxByNodeIndex?.get(
        currentSegment.nodeIndex
      );
      const segmentHeightPx = resolveSegmentHeightPx(
        currentSegment,
        columnWidthPx
      );
      const remainingHeightPx = Math.max(
        0,
        safeColumnHeightPx - consumedHeightPx
      );

      if (
        segmentHeightPx <= remainingHeightPx + PAGE_OVERFLOW_TOLERANCE_PX ||
        columnIndex + 1 >= columnCount
      ) {
        pushSegment(currentSegment, segmentHeightPx);
        pendingSegment = undefined;
        break;
      }

      const splitSegment:
        | {
            currentSegment: DocumentPageNodeSegment;
            currentHeightPx: number;
            remainderSegment: DocumentPageNodeSegment;
          }
        | undefined =
        segmentNode.type === "paragraph"
          ? splitParagraphSegmentForColumnRender({
              paragraph: segmentNode,
              segment: currentSegment,
              model,
              availableWidthPx: columnWidthPx,
              availableHeightPx: remainingHeightPx,
              numberingDefinitions,
              docGridLinePitchPx,
            })
          : undefined;
      if (splitSegment) {
        pushSegment(splitSegment.currentSegment, splitSegment.currentHeightPx);
        pendingSegment = splitSegment.remainderSegment;
        if (!moveToNextColumn()) {
          const remainderSegment = splitSegment.remainderSegment;
          const remainderHeightPx = estimateRenderedPageSegmentHeightPx(
            segmentNode,
            remainderSegment,
            model,
            columnWidthPx,
            numberingDefinitions,
            docGridLinePitchPx
          );
          pushSegment(remainderSegment, remainderHeightPx);
          pendingSegment = undefined;
        }
        continue;
      }

      if (!moveToNextColumn()) {
        pushSegment(currentSegment, segmentHeightPx);
        pendingSegment = undefined;
      }
    }
  }

  return columns;
}

function sumEstimatedTableRowHeightsPx(
  rowHeightsPx: number[],
  startRowIndex: number,
  endRowIndex: number
): number {
  let total = 0;
  const clampedStart = Math.max(0, startRowIndex);
  const clampedEnd = Math.max(
    clampedStart,
    Math.min(endRowIndex, rowHeightsPx.length)
  );
  for (let rowIndex = clampedStart; rowIndex < clampedEnd; rowIndex += 1) {
    total += Math.max(
      1,
      rowHeightsPx[rowIndex] ?? MIN_PARAGRAPH_LINE_HEIGHT_PX
    );
  }
  return total;
}

function fitTableRowsWithinHeightPx(
  rowHeightsPx: number[],
  startRowIndex: number,
  availableHeightPx: number,
  forceAtLeastOneRow: boolean
): number {
  if (startRowIndex >= rowHeightsPx.length) {
    return startRowIndex;
  }

  const safeAvailableHeightPx =
    Number.isFinite(availableHeightPx) && availableHeightPx > 0
      ? availableHeightPx
      : 0;
  let consumedHeightPx = 0;
  let rowCursor = startRowIndex;

  while (rowCursor < rowHeightsPx.length) {
    const rowHeightPx = Math.max(
      1,
      rowHeightsPx[rowCursor] ?? MIN_PARAGRAPH_LINE_HEIGHT_PX
    );
    if (
      consumedHeightPx + rowHeightPx >
      safeAvailableHeightPx + PAGE_OVERFLOW_TOLERANCE_PX
    ) {
      break;
    }

    consumedHeightPx += rowHeightPx;
    rowCursor += 1;
  }

  if (rowCursor === startRowIndex && forceAtLeastOneRow) {
    return Math.min(rowHeightsPx.length, startRowIndex + 1);
  }

  return rowCursor;
}

export function buildDocumentPageNodeSegmentsFromLastRenderedPageBreakHints(
  model: DocModel
): DocumentPageNodeSegment[][] {
  if (model.nodes.length === 0) {
    return [];
  }

  const pageStartNodeIndexes: number[] = [0];
  model.nodes.forEach((node, nodeIndex) => {
    if (nodeIndex === 0 || node.type !== "paragraph") {
      return;
    }

    if (
      paragraphHasLastRenderedPageBreak(node) ||
      paragraphHasPageBreakBefore(node)
    ) {
      pageStartNodeIndexes.push(nodeIndex);
    }
  });

  const uniquePageStartNodeIndexes = [...new Set(pageStartNodeIndexes)].sort(
    (left, right) => left - right
  );
  const pages: DocumentPageNodeSegment[][] = [];

  uniquePageStartNodeIndexes.forEach((startNodeIndex, pageIndex) => {
    const nextStartNodeIndex =
      uniquePageStartNodeIndexes[pageIndex + 1] ?? model.nodes.length;
    const pageSegments: DocumentPageNodeSegment[] = [];
    for (
      let nodeIndex = startNodeIndex;
      nodeIndex < nextStartNodeIndex;
      nodeIndex += 1
    ) {
      pageSegments.push({ nodeIndex });
    }
    if (pageSegments.length > 0) {
      pages.push(pageSegments);
    }
  });

  return pages;
}

export function buildDocumentPageNodeSegments(
  model: DocModel,
  pageContentHeightPx: number,
  pageContentWidthPx: number,
  numberingDefinitions?: NumberingDefinitionSet,
  paginationMetricsBySection?: PaginationSectionMetrics[],
  options?: {
    allowParagraphLineSplitting?: boolean;
    suppressSpacingBeforeAfterPageBreak?: boolean;
    measuredTableRowHeightsByNodeIndex?: Record<number, number[]>;
    measuredPageContentHeightsPxByPageIndex?: number[];
    measuredParagraphOuterHeightsPxByNodeIndex?: Map<number, number>;
    preferLastRenderedParagraphStartBreaks?: boolean;
    strictLastRenderedParagraphStartBreaks?: boolean;
  }
): DocumentPageNodeSegment[][] {
  if (model.nodes.length === 0) {
    return [];
  }

  const fallbackMetrics: PaginationSectionMetrics = {
    startNodeIndex: 0,
    pageContentWidthPx: Math.max(120, Math.round(pageContentWidthPx)),
    pageContentHeightPx: Math.max(120, Math.round(pageContentHeightPx)),
    docGridLinePitchPx: undefined,
  };
  const metricsBySection = paginationMetricsBySection?.length
    ? paginationMetricsBySection
    : [fallbackMetrics];

  const pages: DocumentPageNodeSegment[][] = [];
  let currentPageSegments: DocumentPageNodeSegment[] = [];
  let currentPageIndex = 0;
  const allowParagraphLineSplitting =
    options?.allowParagraphLineSplitting ?? true;
  const shouldSuppressSpacingAtPageBreaks =
    options?.suppressSpacingBeforeAfterPageBreak ?? false;
  const preferLastRenderedParagraphStartBreaks =
    options?.preferLastRenderedParagraphStartBreaks ?? false;
  const strictLastRenderedParagraphStartBreaks =
    options?.strictLastRenderedParagraphStartBreaks ?? false;
  const hardBreakStartNodeIndexes =
    collectDocxHardPageBreakStartNodeIndexes(model);
  // Numbered-list markers consume horizontal text room, so pagination needs the
  // rendered label text (e.g. "10.") to estimate line wrapping accurately.
  const paginationNumberingLabels = buildParagraphNumberingLabels(model);
  const sectionStartPageBreakNodeIndexes =
    collectDocxSectionStartPageBreakNodeIndexes(model);
  const nextHardBreakStartNodeIndexByNodeIndex =
    buildNextHardBreakStartNodeIndexLookup(
      model.nodes.length,
      hardBreakStartNodeIndexes
    );
  const estimatedRowHeightsByTableNodeIndex = new Map<number, number[]>();
  const measuredPageContentHeightsPxByPageIndex =
    options?.measuredPageContentHeightsPxByPageIndex;
  const resolvePageContentHeightPx = (
    pageIndex: number,
    fallbackHeightPx: number,
    heightMultiplier = 1
  ): number => {
    const overrideHeightPx =
      measuredPageContentHeightsPxByPageIndex?.[pageIndex];
    if (Number.isFinite(overrideHeightPx) && (overrideHeightPx as number) > 0) {
      return Math.max(
        24,
        Math.min(
          Math.round(fallbackHeightPx),
          Math.round(
            (overrideHeightPx as number) * Math.max(1, heightMultiplier)
          )
        )
      );
    }
    return Math.max(24, Math.round(fallbackHeightPx));
  };
  const resolveMetricsPageContentHeightPx = (
    pageIndex: number,
    metrics: PaginationSectionMetrics | undefined
  ): number =>
    resolvePageContentHeightPx(
      pageIndex,
      metrics?.pageContentHeightPx ?? fallbackMetrics.pageContentHeightPx,
      metrics?.pageContentHeightMultiplier ??
        fallbackMetrics.pageContentHeightMultiplier ??
        1
    );

  const startNextPage = (): void => {
    if (currentPageSegments.length > 0) {
      pages.push(currentPageSegments);
    }
    currentPageSegments = [];
    currentPageIndex += 1;
  };
  if (!Number.isFinite(pageContentHeightPx) || pageContentHeightPx <= 0) {
    return [model.nodes.map((_, nodeIndex) => ({ nodeIndex }))];
  }

  let pageConsumedHeightPx = 0;
  let previousParagraphAfterPx = 0;
  let currentMetricsIndex = 0;
  let currentSectionPageFlowOriginPx = 0;
  // Once a keepNext chain head starts a page, every page up to the chain end
  // begins with a chain member, so a keep-induced push for a later member can
  // never reunite it with its predecessors — it would only strand the chain
  // head on a near-empty page. Track the chain end to suppress those pushes.
  let committedKeepNextChainEndNodeIndex = -1;
  let currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
    0,
    metricsBySection[0]
  );
  const projectConsumedHeightAcrossSectionMultipliers = (
    consumedHeightPx: number,
    fromMetrics: PaginationSectionMetrics | undefined,
    toMetrics: PaginationSectionMetrics | undefined
  ): number => {
    const safeConsumedHeightPx = Math.max(0, Math.round(consumedHeightPx));
    if (safeConsumedHeightPx <= 0) {
      return 0;
    }

    const fromMultiplier = Math.max(
      1,
      Math.round(fromMetrics?.pageContentHeightMultiplier ?? 1)
    );
    const toMultiplier = Math.max(
      1,
      Math.round(toMetrics?.pageContentHeightMultiplier ?? 1)
    );
    if (fromMultiplier === toMultiplier) {
      return safeConsumedHeightPx;
    }

    const approximateVisualDepthPx = safeConsumedHeightPx / fromMultiplier;
    return Math.max(0, Math.round(approximateVisualDepthPx * toMultiplier));
  };

  for (let nodeIndex = 0; nodeIndex < model.nodes.length; nodeIndex += 1) {
    const previousMetricsIndex = currentMetricsIndex;
    currentMetricsIndex = resolvePaginationSectionMetricsIndexForNodeIndex(
      metricsBySection,
      nodeIndex,
      currentMetricsIndex
    );
    const nodeMetrics =
      metricsBySection[currentMetricsIndex] ?? fallbackMetrics;
    if (nodeIndex > 0 && currentMetricsIndex !== previousMetricsIndex) {
      pageConsumedHeightPx = projectConsumedHeightAcrossSectionMultipliers(
        pageConsumedHeightPx,
        metricsBySection[previousMetricsIndex],
        nodeMetrics
      );
      currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
        currentPageIndex,
        nodeMetrics
      );
      currentSectionPageFlowOriginPx = pageConsumedHeightPx;
    }

    if (
      hardBreakStartNodeIndexes.has(nodeIndex) &&
      currentPageSegments.length > 0
    ) {
      startNextPage();
      pageConsumedHeightPx = 0;
      previousParagraphAfterPx = 0;
      currentSectionPageFlowOriginPx = 0;
      currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
        currentPageIndex,
        nodeMetrics
      );
    }

    const node = model.nodes[nodeIndex];
    if (
      node.type === "paragraph" &&
      (paragraphIsStructuralSectionBreakSpacer(node) ||
        paragraphActsAsSectionBreakCarryoverSpacer(model, nodeIndex, node))
    ) {
      previousParagraphAfterPx = 0;
      continue;
    }
    if (
      node.type === "paragraph" &&
      paragraphActsAsTrailingRenderedPageBreakSpacer(model, nodeIndex, node)
    ) {
      previousParagraphAfterPx = 0;
      continue;
    }
    if (
      node.type === "paragraph" &&
      paragraphCollapsesIntoPreviousParagraph(node, model.nodes[nodeIndex - 1])
    ) {
      continue;
    }
    if (node.type === "paragraph") {
      if (paragraphHasPageBreakBefore(node) && currentPageSegments.length > 0) {
        startNextPage();
        pageConsumedHeightPx = 0;
        previousParagraphAfterPx = 0;
        currentSectionPageFlowOriginPx = 0;
        currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
          currentPageIndex,
          nodeMetrics
        );
      }

      if (
        preferLastRenderedParagraphStartBreaks &&
        paragraphStartsWithLastRenderedPageBreak(node) &&
        (strictLastRenderedParagraphStartBreaks ||
          shouldHonorParagraphStartLastRenderedPageBreak({
            pageConsumedHeightPx,
            pageContentHeightPx: currentPageContentHeightPx,
          })) &&
        !nodeAlreadyEndsAtExplicitPageBoundary(model.nodes[nodeIndex - 1]) &&
        currentPageSegments.length > 0
      ) {
        startNextPage();
        pageConsumedHeightPx = 0;
        previousParagraphAfterPx = 0;
        currentSectionPageFlowOriginPx = 0;
        currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
          currentPageIndex,
          nodeMetrics
        );
      }

      const directBeforeSpacingPx = paragraphBeforeSpacingPx(node);
      const directAfterSpacingPx = paragraphAfterSpacingPx(node);
      const beforeSpacingPx = effectiveParagraphBeforeSpacingPx(
        model,
        nodeIndex,
        node,
        pageConsumedHeightPx,
        shouldSuppressSpacingAtPageBreaks
      );
      const afterSpacingPx = effectiveParagraphAfterSpacingPx(
        model,
        nodeIndex,
        node
      );
      const paragraphLineHeightPx = estimateParagraphLineHeightPx(
        node,
        nodeMetrics.docGridLinePitchPx
      );
      const paragraphTextWidthPx = paragraphAvailableTextWidthPx(
        node,
        nodeMetrics.pageContentWidthPx,
        numberingDefinitions
      );
      const paragraphPretextSourceForSegmentRendering =
        buildParagraphPretextLayoutSource(node, {
          allowExplicitLineBreakText: true,
          expandTabsForLayout: true,
        });
      let paragraphPretextLayoutForSegmentRendering:
        | PretextVariableWidthLayout
        | undefined;
      let paragraphPretextLayoutForSegmentRenderingResolved = false;
      const resolveParagraphPretextLayoutForSegmentRendering = ():
        | PretextVariableWidthLayout
        | undefined => {
        if (paragraphPretextLayoutForSegmentRenderingResolved) {
          return paragraphPretextLayoutForSegmentRendering;
        }

        paragraphPretextLayoutForSegmentRenderingResolved = true;
        paragraphPretextLayoutForSegmentRendering =
          paragraphPretextSourceForSegmentRendering
            ? layoutParagraphPretextSource(
                node,
                paragraphPretextSourceForSegmentRendering,
                paragraphTextWidthPx,
                paragraphLineHeightPx,
                []
              )
            : undefined;
        return paragraphPretextLayoutForSegmentRendering;
      };
      const nodeNumberingLabel = paginationNumberingLabels.get(
        `p:${nodeIndex}`
      );
      // Column-flow sections have no measured-page-height convergence, so
      // estimate drift (custom fonts, dense wrapping) accumulates unchecked;
      // prefer the actually rendered paragraph height there. Single-column
      // pagination keeps its tuned estimate + reconciliation behavior.
      const sectionUsesColumnFlow =
        (nodeMetrics.pageContentHeightMultiplier ?? 1) > 1;
      const measuredOuterHeightPx = sectionUsesColumnFlow
        ? options?.measuredParagraphOuterHeightsPxByNodeIndex?.get(nodeIndex)
        : undefined;
      const estimatedOrMeasuredHeightPx =
        Number.isFinite(measuredOuterHeightPx) &&
        (measuredOuterHeightPx as number) > 0
          ? (measuredOuterHeightPx as number)
          : estimateParagraphHeightPx(
              node,
              nodeMetrics.pageContentWidthPx,
              numberingDefinitions,
              nodeMetrics.docGridLinePitchPx,
              false,
              nodeNumberingLabel
            );
      let rawNodeHeightPx = Math.max(
        1,
        estimatedOrMeasuredHeightPx -
          directBeforeSpacingPx -
          directAfterSpacingPx +
          beforeSpacingPx +
          afterSpacingPx
      );
      const paragraphTooTallForSinglePage =
        rawNodeHeightPx >
        nodeMetrics.pageContentHeightPx + PAGE_OVERFLOW_TOLERANCE_PX;
      const keepLinesOverflowSplit =
        node.style?.keepLines === true && paragraphTooTallForSinglePage;
      const keepNextOverflowSplit =
        node.style?.keepNext === true && paragraphTooTallForSinglePage;
      const forceOverflowSplit =
        keepLinesOverflowSplit || keepNextOverflowSplit;
      const nodeIsWithinCommittedKeepNextChain =
        nodeIndex <= committedKeepNextChainEndNodeIndex;
      if (
        forceOverflowSplit &&
        !nodeIsWithinCommittedKeepNextChain &&
        pageConsumedHeightPx > 0 &&
        currentPageSegments.length > 0
      ) {
        // ECMA-376 §2.3.1.14/§2.3.1.15: if "keep" constraints cannot fit on one
        // page, start at a new page and continue with page breaks as needed.
        startNextPage();
        pageConsumedHeightPx = 0;
        previousParagraphAfterPx = 0;
        currentSectionPageFlowOriginPx = 0;
        currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
          currentPageIndex,
          nodeMetrics
        );
      }

      const collapsedMarginPx =
        pageConsumedHeightPx > 0
          ? Math.min(previousParagraphAfterPx, beforeSpacingPx)
          : 0;
      const collapsedNodeHeightPx = Math.max(
        1,
        rawNodeHeightPx - collapsedMarginPx
      );
      const paragraphSupportsPretextSegmentRendering = Boolean(
        paragraphPretextSourceForSegmentRendering
      );
      const remainingHeightBeforeParagraphPx = Math.max(
        0,
        currentPageContentHeightPx - pageConsumedHeightPx
      );
      if (
        paragraphSupportsPretextSegmentRendering &&
        remainingHeightBeforeParagraphPx <=
          collapsedNodeHeightPx +
            paragraphLineHeightPx +
            PAGE_OVERFLOW_TOLERANCE_PX
      ) {
        const pretextLayout =
          resolveParagraphPretextLayoutForSegmentRendering();
        const pretextContentHeightPx = pretextLayout
          ? wrappedPretextParagraphBlockHeightPx(pretextLayout)
          : undefined;
        if (Number.isFinite(pretextContentHeightPx)) {
          rawNodeHeightPx = Math.max(
            rawNodeHeightPx,
            Math.max(
              1,
              Math.round(
                beforeSpacingPx +
                  (pretextContentHeightPx as number) +
                  afterSpacingPx
              )
            )
          );
        }
      }
      const collapsedNodeHeightPxAdjusted = Math.max(
        1,
        rawNodeHeightPx - collapsedMarginPx
      );
      const paragraphPretextLineCount =
        paragraphContainsExplicitLineBreakText(node) ||
        paragraphContainsTabCharacter(node)
          ? resolveParagraphPretextLayoutForSegmentRendering()?.lineCount
          : undefined;
      const supportsImageParagraphLineSplit =
        paragraphHasImage(node) && paragraphSupportsPretextSegmentRendering;
      const paragraphLineCount = paragraphLineCountWithinWidth(
        node,
        nodeMetrics.pageContentWidthPx,
        numberingDefinitions,
        nodeNumberingLabel
      );
      const resolvedParagraphLineCount =
        Number.isFinite(paragraphPretextLineCount) &&
        (paragraphPretextLineCount as number) > 0
          ? Math.max(1, Math.round(paragraphPretextLineCount as number))
          : paragraphLineCount;
      const explicitColumnBreakParagraphSegments =
        (nodeMetrics.pageContentHeightMultiplier ?? 1) > 1
          ? splitParagraphAtExplicitColumnBreaks(node)
          : undefined;
      const tryConsumeExplicitColumnBreakParagraph = (): boolean => {
        if (
          !explicitColumnBreakParagraphSegments ||
          explicitColumnBreakParagraphSegments.length <= 1
        ) {
          return false;
        }

        const projectedConsumedHeightPx =
          projectParagraphConsumedHeightWithExplicitColumnBreaks(
            explicitColumnBreakParagraphSegments,
            pageConsumedHeightPx,
            currentPageContentHeightPx,
            currentSectionPageFlowOriginPx,
            Math.max(
              1,
              Math.round(nodeMetrics.pageContentHeightMultiplier ?? 1)
            ),
            nodeMetrics.pageContentWidthPx,
            beforeSpacingPx,
            afterSpacingPx,
            collapsedMarginPx,
            numberingDefinitions,
            nodeMetrics.docGridLinePitchPx
          );
        if (!Number.isFinite(projectedConsumedHeightPx)) {
          return false;
        }

        currentPageSegments.push({ nodeIndex });
        pageConsumedHeightPx = Math.max(
          pageConsumedHeightPx,
          Math.round(projectedConsumedHeightPx as number)
        );
        previousParagraphAfterPx = afterSpacingPx;
        return true;
      };
      if (tryConsumeExplicitColumnBreakParagraph()) {
        continue;
      }
      if (
        explicitColumnBreakParagraphSegments &&
        explicitColumnBreakParagraphSegments.length > 1 &&
        pageConsumedHeightPx > 0 &&
        currentPageSegments.length > 0
      ) {
        startNextPage();
        pageConsumedHeightPx = 0;
        previousParagraphAfterPx = 0;
        currentSectionPageFlowOriginPx = 0;
        currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
          currentPageIndex,
          nodeMetrics
        );
        if (tryConsumeExplicitColumnBreakParagraph()) {
          continue;
        }
      }
      const widowControlEnabled = paragraphWidowControlEnabled(node);
      const minLinesPerSegment = widowControlEnabled ? 2 : 1;
      const canSplitParagraphAcrossPages =
        paragraphCanSplitAcrossPages(node, resolvedParagraphLineCount, {
          allowKeepLinesOverflow: keepLinesOverflowSplit,
          allowKeepNextOverflow: keepNextOverflowSplit,
          allowImageParagraphSplit: supportsImageParagraphLineSplit,
        }) &&
        (!widowControlEnabled || resolvedParagraphLineCount > 3);

      if (canSplitParagraphAcrossPages && allowParagraphLineSplitting) {
        const pretextLayoutForSegmentSplitting =
          resolveParagraphPretextLayoutForSegmentRendering();
        const resolveSegmentReservePx = (
          startLineIndex: number,
          endLineIndex: number
        ): number => {
          const paragraphSegmentRange: ParagraphLineRange = {
            startLineIndex,
            endLineIndex,
            totalLineCount: resolvedParagraphLineCount,
            lineHeightPx: paragraphLineHeightPx,
          };
          return paragraphSupportsPretextSegmentRendering
            ? resolveParagraphSegmentNonFlowReservePx(paragraphSegmentRange)
            : resolveFallbackParagraphSegmentNonFlowReservePx(
                node,
                paragraphSegmentRange
              );
        };
        const resolveSegmentContentHeightPx = (
          startLineIndex: number,
          endLineIndex: number
        ): number => {
          if (
            pretextLayoutForSegmentSplitting &&
            pretextLayoutForSegmentSplitting.lineCount > 0
          ) {
            return resolvePretextLineRangeContentHeightPx(
              pretextLayoutForSegmentSplitting,
              startLineIndex,
              endLineIndex
            );
          }

          return (
            Math.max(1, endLineIndex - startLineIndex) * paragraphLineHeightPx
          );
        };
        let lineCursor = 0;
        let isFirstSegment = true;
        while (lineCursor < resolvedParagraphLineCount) {
          const linesRemaining = resolvedParagraphLineCount - lineCursor;
          const topSpacingPx = isFirstSegment
            ? pageConsumedHeightPx > 0
              ? Math.max(0, beforeSpacingPx - collapsedMarginPx)
              : beforeSpacingPx
            : 0;
          const mustKeepBottomSpacing = linesRemaining <= minLinesPerSegment;
          const bottomSpacingPx = mustKeepBottomSpacing ? afterSpacingPx : 0;
          const remainingHeightPx = Math.max(
            0,
            currentPageContentHeightPx - pageConsumedHeightPx
          );
          const allRemainingSegmentReservePx = resolveSegmentReservePx(
            lineCursor,
            resolvedParagraphLineCount
          );
          const allRemainingHeightPx =
            topSpacingPx +
            resolveSegmentContentHeightPx(
              lineCursor,
              resolvedParagraphLineCount
            ) +
            bottomSpacingPx;

          if (
            allRemainingHeightPx + allRemainingSegmentReservePx <=
            remainingHeightPx
          ) {
            currentPageSegments.push({
              nodeIndex,
              paragraphLineRange: {
                startLineIndex: lineCursor,
                endLineIndex: resolvedParagraphLineCount,
                totalLineCount: resolvedParagraphLineCount,
                lineHeightPx: paragraphLineHeightPx,
              },
            });
            pageConsumedHeightPx += allRemainingHeightPx;
            previousParagraphAfterPx = afterSpacingPx;
            lineCursor = resolvedParagraphLineCount;
            break;
          }

          const maxLinesThisPage = Math.max(
            0,
            linesRemaining - minLinesPerSegment
          );
          const continuingSegmentReservePx = resolveSegmentReservePx(
            lineCursor,
            Math.min(resolvedParagraphLineCount, lineCursor + maxLinesThisPage)
          );
          const availableForLinesPx = Math.max(
            0,
            remainingHeightPx - topSpacingPx - continuingSegmentReservePx
          );
          let linesThatFit = Math.floor(
            availableForLinesPx / paragraphLineHeightPx
          );
          linesThatFit = Math.min(linesThatFit, maxLinesThisPage);
          if (
            pretextLayoutForSegmentSplitting &&
            pretextLayoutForSegmentSplitting.lineCount > 0 &&
            linesThatFit > 0
          ) {
            const exactSegmentEndLineIndex =
              resolveMaxPretextLineRangeEndIndexThatFits(
                pretextLayoutForSegmentSplitting,
                lineCursor,
                Math.min(resolvedParagraphLineCount, lineCursor + linesThatFit),
                availableForLinesPx
              );
            linesThatFit = Math.max(0, exactSegmentEndLineIndex - lineCursor);
          }

          if (linesThatFit < minLinesPerSegment) {
            if (currentPageSegments.length > 0) {
              startNextPage();
              pageConsumedHeightPx = 0;
              previousParagraphAfterPx = 0;
              currentSectionPageFlowOriginPx = 0;
              currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
                currentPageIndex,
                nodeMetrics
              );
              continue;
            }

            const fallbackLines = Math.max(
              1,
              Math.floor(
                Math.max(1, availableForLinesPx) / paragraphLineHeightPx
              )
            );
            linesThatFit = Math.max(
              1,
              Math.min(
                maxLinesThisPage > 0 ? maxLinesThisPage : linesRemaining,
                fallbackLines
              )
            );
          }

          let segmentEndLineIndex = Math.min(
            resolvedParagraphLineCount,
            lineCursor + linesThatFit
          );
          while (linesThatFit > minLinesPerSegment) {
            const segmentReservePx = resolveSegmentReservePx(
              lineCursor,
              segmentEndLineIndex
            );
            if (
              topSpacingPx +
                resolveSegmentContentHeightPx(lineCursor, segmentEndLineIndex) +
                segmentReservePx <=
              remainingHeightPx
            ) {
              break;
            }
            linesThatFit -= 1;
            segmentEndLineIndex = Math.min(
              resolvedParagraphLineCount,
              lineCursor + linesThatFit
            );
          }
          const safeSegmentEndLineIndex = Math.min(
            resolvedParagraphLineCount,
            lineCursor + linesThatFit
          );
          currentPageSegments.push({
            nodeIndex,
            paragraphLineRange: {
              startLineIndex: lineCursor,
              endLineIndex: safeSegmentEndLineIndex,
              totalLineCount: resolvedParagraphLineCount,
              lineHeightPx: paragraphLineHeightPx,
            },
          });

          pageConsumedHeightPx +=
            topSpacingPx +
            resolveSegmentContentHeightPx(lineCursor, safeSegmentEndLineIndex);
          previousParagraphAfterPx = 0;
          lineCursor = safeSegmentEndLineIndex;
          isFirstSegment = false;

          if (lineCursor < resolvedParagraphLineCount) {
            startNextPage();
            pageConsumedHeightPx = 0;
            previousParagraphAfterPx = 0;
            currentSectionPageFlowOriginPx = 0;
            currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
              currentPageIndex,
              nodeMetrics
            );
          }
        }
        continue;
      }

      let requiredHeightPx = collapsedNodeHeightPxAdjusted;
      let keepNextChainEndNodeIndex = -1;
      if (
        node.style?.keepNext === true &&
        !nodeIsWithinCommittedKeepNextChain &&
        paragraphHasVisibleText(node)
      ) {
        let chainCursor = nodeIndex;
        let chainPreviousParagraphAfterPx = afterSpacingPx;
        while (chainCursor < model.nodes.length - 1) {
          const currentChainNode = model.nodes[chainCursor];
          if (
            currentChainNode.type !== "paragraph" ||
            currentChainNode.style?.keepNext !== true ||
            !paragraphHasVisibleText(currentChainNode)
          ) {
            break;
          }
          if (hardBreakStartNodeIndexes.has(chainCursor + 1)) {
            break;
          }
          const nextChainNode = model.nodes[chainCursor + 1];
          if (nextChainNode.type !== "paragraph") {
            break;
          }

          chainCursor += 1;
          const chainMetricsIndex =
            resolvePaginationSectionMetricsIndexForNodeIndex(
              metricsBySection,
              chainCursor,
              currentMetricsIndex
            );
          const chainMetrics =
            metricsBySection[chainMetricsIndex] ?? fallbackMetrics;
          const nextDirectBeforeSpacingPx =
            paragraphBeforeSpacingPx(nextChainNode);
          const nextDirectAfterSpacingPx =
            paragraphAfterSpacingPx(nextChainNode);
          const nextBeforeSpacingPx = effectiveParagraphBeforeSpacingPx(
            model,
            chainCursor,
            nextChainNode,
            1,
            shouldSuppressSpacingAtPageBreaks
          );
          const nextAfterSpacingPx = effectiveParagraphAfterSpacingPx(
            model,
            chainCursor,
            nextChainNode
          );
          const nextRawHeightPx = Math.max(
            1,
            estimateParagraphHeightPx(
              nextChainNode,
              chainMetrics.pageContentWidthPx,
              numberingDefinitions,
              chainMetrics.docGridLinePitchPx
            ) -
              nextDirectBeforeSpacingPx -
              nextDirectAfterSpacingPx +
              nextBeforeSpacingPx +
              nextAfterSpacingPx
          );
          const collapsedChainMarginPx = Math.min(
            chainPreviousParagraphAfterPx,
            nextBeforeSpacingPx
          );
          requiredHeightPx += Math.max(
            1,
            nextRawHeightPx - collapsedChainMarginPx
          );
          requiredHeightPx += keepNextPaginationReservePx(
            currentChainNode,
            nextChainNode,
            chainMetrics.docGridLinePitchPx
          );
          chainPreviousParagraphAfterPx = nextAfterSpacingPx;
        }
        keepNextChainEndNodeIndex = chainCursor;
      }

      const remainingHeightPx =
        currentPageContentHeightPx - pageConsumedHeightPx;
      const canKeepTrailingSectionTailOnCurrentPage =
        shouldKeepTrailingSectionTailOnCurrentPage(
          model,
          nodeIndex,
          pageConsumedHeightPx,
          previousParagraphAfterPx,
          nodeMetrics.pageContentWidthPx,
          currentPageContentHeightPx,
          hardBreakStartNodeIndexes,
          sectionStartPageBreakNodeIndexes,
          nextHardBreakStartNodeIndexByNodeIndex,
          (_candidateNodeIndex, candidateNode, candidatePageContentWidthPx) =>
            estimateDocNodeHeightPx(
              candidateNode,
              candidatePageContentWidthPx,
              numberingDefinitions,
              nodeMetrics.docGridLinePitchPx
            ),
          (candidateNodeIndex, candidateParagraph, consumedHeightPx) =>
            effectiveParagraphBeforeSpacingPx(
              model,
              candidateNodeIndex,
              candidateParagraph,
              consumedHeightPx,
              shouldSuppressSpacingAtPageBreaks
            ),
          (candidateNodeIndex, candidateParagraph) =>
            effectiveParagraphAfterSpacingPx(
              model,
              candidateNodeIndex,
              candidateParagraph
            )
        );
      if (
        pageConsumedHeightPx > 0 &&
        requiredHeightPx > remainingHeightPx + PAGE_OVERFLOW_TOLERANCE_PX &&
        !canKeepTrailingSectionTailOnCurrentPage
      ) {
        startNextPage();
        pageConsumedHeightPx = 0;
        previousParagraphAfterPx = 0;
        currentSectionPageFlowOriginPx = 0;
        currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
          currentPageIndex,
          nodeMetrics
        );
      }

      if (pageConsumedHeightPx === 0 && keepNextChainEndNodeIndex > nodeIndex) {
        committedKeepNextChainEndNodeIndex = keepNextChainEndNodeIndex;
      }
      currentPageSegments.push({ nodeIndex });
      const effectiveNodeHeightPx =
        pageConsumedHeightPx > 0 ? collapsedNodeHeightPx : rawNodeHeightPx;
      pageConsumedHeightPx += effectiveNodeHeightPx;
      previousParagraphAfterPx = afterSpacingPx;
      continue;
    }

    const measuredRowHeightsPxRaw =
      options?.measuredTableRowHeightsByNodeIndex?.[nodeIndex];
    const measuredRowHeightsPx =
      measuredRowHeightsPxRaw &&
      measuredRowHeightsPxRaw.length === node.rows.length
        ? measuredRowHeightsPxRaw.map((heightPx) =>
            Math.max(
              MIN_PARAGRAPH_LINE_HEIGHT_PX,
              Number.isFinite(heightPx)
                ? Math.round(heightPx as number)
                : MIN_PARAGRAPH_LINE_HEIGHT_PX
            )
          )
        : undefined;
    const estimatedRowHeightsPx =
      measuredRowHeightsPx ??
      estimatedRowHeightsByTableNodeIndex.get(nodeIndex) ??
      estimateTableRowHeightsPx(
        node,
        nodeMetrics.pageContentWidthPx,
        numberingDefinitions,
        nodeMetrics.docGridLinePitchPx,
        nodeMetrics.pageContentHeightPx
      );
    if (
      !measuredRowHeightsPx &&
      !estimatedRowHeightsByTableNodeIndex.has(nodeIndex)
    ) {
      estimatedRowHeightsByTableNodeIndex.set(nodeIndex, estimatedRowHeightsPx);
    }

    if (estimatedRowHeightsPx.length === 0) {
      currentPageSegments.push({ nodeIndex });
      previousParagraphAfterPx = 0;
      continue;
    }

    const tableExplicitPageBreakInfo = collectTableExplicitPageBreakInfo(node);
    const tableBreakStartRows = tableExplicitPageBreakInfo.startRowIndexes;
    if (tableBreakStartRows.includes(0) && currentPageSegments.length > 0) {
      startNextPage();
      pageConsumedHeightPx = 0;
      previousParagraphAfterPx = 0;
      currentSectionPageFlowOriginPx = 0;
      currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
        currentPageIndex,
        nodeMetrics
      );
    }

    let headerRunLength = 0;
    for (const row of node.rows) {
      if (row.style?.isHeader === true) {
        headerRunLength += 1;
      } else {
        break;
      }
    }
    const headerRunHeightPx =
      headerRunLength > 0
        ? sumEstimatedTableRowHeightsPx(
            estimatedRowHeightsPx,
            0,
            headerRunLength
          )
        : 0;
    const canRepeatHeader =
      headerRunLength > 0 &&
      headerRunHeightPx > 0 &&
      headerRunHeightPx + MIN_PARAGRAPH_LINE_HEIGHT_PX * 2 <
        currentPageContentHeightPx;

    let rowStartIndex = 0;
    let rowSliceOffsetPx = 0;
    let repeatedHeaderHeightPxOnThisPage = 0;
    let tableBreakStartRowCursor = 0;
    while (rowStartIndex < estimatedRowHeightsPx.length) {
      if (currentPageSegments.length === 0) {
        repeatedHeaderHeightPxOnThisPage = 0;
      }

      if (
        canRepeatHeader &&
        rowStartIndex >= headerRunLength &&
        currentPageSegments.length === 0
      ) {
        currentPageSegments.push({
          nodeIndex,
          tableRowRange: {
            startRowIndex: 0,
            endRowIndex: headerRunLength,
          },
        });
        pageConsumedHeightPx += headerRunHeightPx;
        previousParagraphAfterPx = 0;
        repeatedHeaderHeightPxOnThisPage = headerRunHeightPx;
      }

      const remainingHeightPx = Math.max(
        0,
        currentPageContentHeightPx - pageConsumedHeightPx
      );
      const currentRow = node.rows[rowStartIndex];
      const currentRowTotalHeightPx = Math.max(
        MIN_PARAGRAPH_LINE_HEIGHT_PX,
        estimatedRowHeightsPx[rowStartIndex] ?? MIN_PARAGRAPH_LINE_HEIGHT_PX
      );
      const currentRowRemainingHeightPx = Math.max(
        0,
        currentRowTotalHeightPx - rowSliceOffsetPx
      );
      const freshPageAvailableHeightPx = Math.max(
        MIN_PARAGRAPH_LINE_HEIGHT_PX,
        currentPageContentHeightPx - repeatedHeaderHeightPxOnThisPage
      );
      const rowExceedsFreshPageHeightPx =
        currentRowTotalHeightPx >
        freshPageAvailableHeightPx + PAGE_OVERFLOW_TOLERANCE_PX;
      const hasUsableCurrentPageSliceSpace =
        pageConsumedHeightPx > 0 &&
        remainingHeightPx >= MIN_TABLE_ROW_SLICE_REMAINING_HEIGHT_PX;
      const canSplitCurrentRow =
        rowSliceOffsetPx > 0 ||
        rowAllowsPageSplit(currentRow) ||
        rowExceedsFreshPageHeightPx;
      const shouldContinueExistingRowSlice = rowSliceOffsetPx > 0;
      const rowNeedsSliceOnThisPage =
        ENABLE_TABLE_ROW_SLICING &&
        canSplitCurrentRow &&
        currentRowRemainingHeightPx >
          remainingHeightPx + PAGE_OVERFLOW_TOLERANCE_PX &&
        (shouldContinueExistingRowSlice ||
          rowExceedsFreshPageHeightPx ||
          hasUsableCurrentPageSliceSpace);

      if (
        canSplitCurrentRow &&
        currentRowRemainingHeightPx >
          remainingHeightPx + PAGE_OVERFLOW_TOLERANCE_PX &&
        !rowNeedsSliceOnThisPage &&
        currentPageSegments.length > 0
      ) {
        startNextPage();
        pageConsumedHeightPx = 0;
        previousParagraphAfterPx = 0;
        currentSectionPageFlowOriginPx = 0;
        currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
          currentPageIndex,
          nodeMetrics
        );
        continue;
      }

      if (rowNeedsSliceOnThisPage) {
        if (
          currentPageSegments.length > 0 &&
          remainingHeightPx <= MIN_PARAGRAPH_LINE_HEIGHT_PX
        ) {
          startNextPage();
          pageConsumedHeightPx = 0;
          previousParagraphAfterPx = 0;
          currentSectionPageFlowOriginPx = 0;
          currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
            currentPageIndex,
            nodeMetrics
          );
          continue;
        }

        const availableSliceHeightPx =
          pageConsumedHeightPx <= 0
            ? currentPageContentHeightPx
            : Math.max(
                MIN_PARAGRAPH_LINE_HEIGHT_PX,
                Math.round(remainingHeightPx)
              );
        const preferredSliceHeightPx = Math.max(
          MIN_PARAGRAPH_LINE_HEIGHT_PX,
          Math.min(currentRowRemainingHeightPx, availableSliceHeightPx)
        );
        const safeSliceHeightPx = resolveTableRowSliceHeightOnSafeBoundaryPx({
          table: node,
          rowIndex: rowStartIndex,
          rowHeightPx: currentRowTotalHeightPx,
          rowSliceOffsetPx,
          preferredSliceHeightPx,
          maxAvailableWidthPx: nodeMetrics.pageContentWidthPx,
          numberingDefinitions,
          docGridLinePitchPx: nodeMetrics.docGridLinePitchPx,
        });
        if (
          safeSliceHeightPx === undefined &&
          pageConsumedHeightPx > 0 &&
          currentPageSegments.length > 0
        ) {
          startNextPage();
          pageConsumedHeightPx = 0;
          previousParagraphAfterPx = 0;
          currentSectionPageFlowOriginPx = 0;
          currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
            currentPageIndex,
            nodeMetrics
          );
          continue;
        }

        const sliceHeightPx = Math.max(
          MIN_PARAGRAPH_LINE_HEIGHT_PX,
          Math.min(
            currentRowRemainingHeightPx,
            safeSliceHeightPx ?? preferredSliceHeightPx
          )
        );
        currentPageSegments.push({
          nodeIndex,
          tableRowRange: {
            startRowIndex: rowStartIndex,
            endRowIndex: Math.min(
              estimatedRowHeightsPx.length,
              rowStartIndex + 1
            ),
          },
          tableRowSlice: {
            rowIndex: rowStartIndex,
            startOffsetPx: rowSliceOffsetPx,
            sliceHeightPx,
            totalRowHeightPx: currentRowTotalHeightPx,
          },
        });
        pageConsumedHeightPx += sliceHeightPx;
        previousParagraphAfterPx = 0;
        rowSliceOffsetPx += sliceHeightPx;
        if (
          rowSliceOffsetPx >=
          currentRowTotalHeightPx - PAGE_OVERFLOW_TOLERANCE_PX
        ) {
          rowStartIndex += 1;
          rowSliceOffsetPx = 0;
        }

        if (rowStartIndex < estimatedRowHeightsPx.length) {
          startNextPage();
          pageConsumedHeightPx = 0;
          previousParagraphAfterPx = 0;
          currentSectionPageFlowOriginPx = 0;
          currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
            currentPageIndex,
            nodeMetrics
          );
        }
        continue;
      }

      if (rowSliceOffsetPx > 0) {
        currentPageSegments.push({
          nodeIndex,
          tableRowRange: {
            startRowIndex: rowStartIndex,
            endRowIndex: Math.min(
              estimatedRowHeightsPx.length,
              rowStartIndex + 1
            ),
          },
          tableRowSlice: {
            rowIndex: rowStartIndex,
            startOffsetPx: rowSliceOffsetPx,
            sliceHeightPx: currentRowRemainingHeightPx,
            totalRowHeightPx: currentRowTotalHeightPx,
          },
        });
        pageConsumedHeightPx += currentRowRemainingHeightPx;
        previousParagraphAfterPx = 0;
        rowStartIndex += 1;
        rowSliceOffsetPx = 0;
        continue;
      }

      const fittedRowEndIndex = fitTableRowsWithinHeightPx(
        estimatedRowHeightsPx,
        rowStartIndex,
        remainingHeightPx,
        pageConsumedHeightPx <= 0
      );
      let rowEndIndex = fittedRowEndIndex;
      while (
        tableBreakStartRowCursor < tableBreakStartRows.length &&
        tableBreakStartRows[tableBreakStartRowCursor] <= rowStartIndex
      ) {
        tableBreakStartRowCursor += 1;
      }
      const forcedBreakRowIndex = tableBreakStartRows[tableBreakStartRowCursor];
      if (forcedBreakRowIndex !== undefined) {
        rowEndIndex = Math.min(rowEndIndex, forcedBreakRowIndex);
      }

      if (rowEndIndex <= rowStartIndex) {
        if (currentPageSegments.length > 0) {
          startNextPage();
          pageConsumedHeightPx = 0;
          previousParagraphAfterPx = 0;
          currentSectionPageFlowOriginPx = 0;
          currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
            currentPageIndex,
            nodeMetrics
          );
          continue;
        }

        const forcedEndIndex = Math.min(
          estimatedRowHeightsPx.length,
          rowStartIndex + 1
        );
        const forcedHeightPx = sumEstimatedTableRowHeightsPx(
          estimatedRowHeightsPx,
          rowStartIndex,
          forcedEndIndex
        );
        currentPageSegments.push({
          nodeIndex,
          tableRowRange: {
            startRowIndex: rowStartIndex,
            endRowIndex: forcedEndIndex,
          },
        });
        pageConsumedHeightPx += forcedHeightPx;
        previousParagraphAfterPx = 0;
        rowStartIndex = forcedEndIndex;
        rowSliceOffsetPx = 0;

        if (rowStartIndex < estimatedRowHeightsPx.length) {
          startNextPage();
          pageConsumedHeightPx = 0;
          previousParagraphAfterPx = 0;
          currentSectionPageFlowOriginPx = 0;
          currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
            currentPageIndex,
            nodeMetrics
          );
        }
        continue;
      }

      const segmentHeightPx = sumEstimatedTableRowHeightsPx(
        estimatedRowHeightsPx,
        rowStartIndex,
        rowEndIndex
      );
      const coversWholeTable =
        rowStartIndex === 0 && rowEndIndex >= estimatedRowHeightsPx.length;
      currentPageSegments.push({
        nodeIndex,
        tableRowRange: coversWholeTable
          ? undefined
          : {
              startRowIndex: rowStartIndex,
              endRowIndex: rowEndIndex,
            },
      });
      pageConsumedHeightPx += segmentHeightPx;
      previousParagraphAfterPx = 0;
      rowStartIndex = rowEndIndex;
      rowSliceOffsetPx = 0;

      if (rowStartIndex < estimatedRowHeightsPx.length) {
        const nextRow = node.rows[rowStartIndex];
        const nextRowHeightPx = Math.max(
          MIN_PARAGRAPH_LINE_HEIGHT_PX,
          estimatedRowHeightsPx[rowStartIndex] ?? MIN_PARAGRAPH_LINE_HEIGHT_PX
        );
        const remainingHeightAfterSegmentPx = Math.max(
          0,
          currentPageContentHeightPx - pageConsumedHeightPx
        );
        const canSliceNextRowOnCurrentPage =
          ENABLE_TABLE_ROW_SLICING &&
          nextRow !== undefined &&
          rowAllowsPageSplit(nextRow) &&
          !tableBreakStartRows.includes(rowStartIndex) &&
          remainingHeightAfterSegmentPx >=
            MIN_TABLE_ROW_SLICE_REMAINING_HEIGHT_PX &&
          nextRowHeightPx >
            remainingHeightAfterSegmentPx + PAGE_OVERFLOW_TOLERANCE_PX;
        if (canSliceNextRowOnCurrentPage) {
          continue;
        }

        startNextPage();
        pageConsumedHeightPx = 0;
        previousParagraphAfterPx = 0;
        currentSectionPageFlowOriginPx = 0;
        currentPageContentHeightPx = resolveMetricsPageContentHeightPx(
          currentPageIndex,
          nodeMetrics
        );
      }
    }
  }

  if (currentPageSegments.length > 0 || pages.length === 0) {
    pages.push(currentPageSegments);
  }

  return pages;
}

function mergeTrailingPagesToTargetCount(
  pages: DocumentPageNodeSegment[][],
  targetPageCount: number
): DocumentPageNodeSegment[][] {
  const safeTargetPageCount = Math.max(1, Math.round(targetPageCount));
  if (pages.length <= safeTargetPageCount) {
    return pages;
  }

  const merged = pages.map((pageSegments) => [...pageSegments]);
  while (merged.length > safeTargetPageCount) {
    const trailingPage = merged.pop();
    if (!trailingPage || merged.length === 0) {
      break;
    }

    merged[merged.length - 1] = [...merged[merged.length - 1], ...trailingPage];
  }

  return merged;
}

function shouldRenderWrappedFloatingImage(image: ImageRunNode): boolean {
  const floating = image.floating;
  if (!floating) {
    return false;
  }

  if (image.syntheticTextBox && !syntheticTextBoxContainsPictureLayer(image)) {
    return false;
  }

  const wrapType = floating.wrapType;
  if (wrapType === undefined || wrapType === "none") {
    return false;
  }

  if (shouldRenderTopAnchoredMarginFloatAsAbsolute(image)) {
    return false;
  }

  if (!floating.behindDocument) {
    return true;
  }

  const horizontalRelativeTo = floating.horizontalRelativeTo?.toLowerCase();
  const verticalRelativeTo = floating.verticalRelativeTo?.toLowerCase();
  const hasAnchoredAlignment = Boolean(
    floating.horizontalAlign || floating.verticalAlign || floating.wrapText
  );

  // Some DOCX exports mark anchored wrapped images as behindDoc even when Word
  // still lays them out like wrapped content. Keep those in flow unless they
  // are explicitly page-anchored overlays.
  return (
    hasAnchoredAlignment &&
    horizontalRelativeTo !== "page" &&
    verticalRelativeTo !== "page"
  );
}

function isFixedPositionWrappedFloatingImage(image: ImageRunNode): boolean {
  return (
    shouldRenderWrappedFloatingImage(image) &&
    !floatingImageMovesWithText(image.floating)
  );
}

function shouldRenderTopAnchoredMarginFloatAsAbsolute(
  image: ImageRunNode
): boolean {
  const floating = image.floating;
  if (!floating) {
    return false;
  }

  if (
    !floating.wrapType ||
    floating.wrapType === "none" ||
    floating.behindDocument
  ) {
    return false;
  }

  const horizontalRelativeTo = floating.horizontalRelativeTo?.toLowerCase();
  const verticalRelativeTo = floating.verticalRelativeTo?.toLowerCase();
  const horizontalAlign = floating.horizontalAlign?.toLowerCase();
  const verticalAlign = floating.verticalAlign?.toLowerCase();

  const pageAnchoredHorizontally =
    horizontalRelativeTo === "margin" || horizontalRelativeTo === "page";
  const pageAnchoredVertically =
    verticalRelativeTo === "margin" || verticalRelativeTo === "page";
  const sideAligned =
    horizontalAlign === "left" ||
    horizontalAlign === "right" ||
    horizontalAlign === "inside" ||
    horizontalAlign === "outside";
  const topAligned = verticalAlign === "top" || verticalAlign === "inside";

  // Anchor it absolutely only when an explicit side/top ALIGNMENT is present.
  // Such corner-anchored floats stay fixed on the page. Dragging one keeps its
  // alignment and adds a posOffset (xPx/yPx) for the new position — the
  // absolute render honors that offset, so it lands exactly at the drop point.
  // Offset-only floats (no alignment) keep flowing as wrapped content so text
  // still wraps around them; those are handled by the wrapped layout path.
  return (
    pageAnchoredHorizontally &&
    pageAnchoredVertically &&
    sideAligned &&
    topAligned
  );
}

function resolveWrappedFloatingSide(
  image: ImageRunNode,
  options?: {
    containerWidthPx?: number;
    imageWidthPx?: number;
  }
): "left" | "right" {
  const floating = image.floating;
  const wrapText = floating?.wrapText;
  if (wrapText === "left") {
    return "right";
  }
  if (wrapText === "right") {
    return "left";
  }

  const horizontalAlign = floating?.horizontalAlign?.toLowerCase();
  if (horizontalAlign === "right" || horizontalAlign === "outside") {
    return "right";
  }
  if (horizontalAlign === "left" || horizontalAlign === "inside") {
    return "left";
  }

  const containerWidthPx =
    Number.isFinite(options?.containerWidthPx) &&
    (options?.containerWidthPx as number) > 0
      ? Math.max(1, Math.round(options?.containerWidthPx as number))
      : undefined;
  const imageWidthPx =
    Number.isFinite(options?.imageWidthPx) &&
    (options?.imageWidthPx as number) > 0
      ? Math.max(1, Math.round(options?.imageWidthPx as number))
      : undefined;
  if (
    Number.isFinite(containerWidthPx) &&
    Number.isFinite(imageWidthPx) &&
    Number.isFinite(floating?.xPx)
  ) {
    const centerX = (floating?.xPx as number) + (imageWidthPx as number) / 2;
    return centerX <= (containerWidthPx as number) / 2 ? "left" : "right";
  }

  if ((floating?.xPx ?? 0) >= 0) {
    return (floating?.xPx ?? 0) > 96 ? "right" : "left";
  }

  return "left";
}

export function wrappedFloatingImageStyle(
  image: ImageRunNode,
  options?: {
    containerWidthPx?: number;
    deltaX?: number;
    deltaY?: number;
    allowNegativeOffsets?: boolean;
  }
): Record<string, string | number | undefined> {
  const floating = image.floating;
  const wrapType = floating?.wrapType;
  if (!floating || !wrapType) {
    return {};
  }

  // Wrap distances should come from the DOCX anchor only; avoid synthetic gaps.
  const distL = floating.distLPx ?? 0;
  const distR = floating.distRPx ?? 0;
  const distT = floating.distTPx ?? 0;
  const distB = floating.distBPx ?? 0;
  const deltaX = Number.isFinite(options?.deltaX)
    ? Math.round(options?.deltaX as number)
    : 0;
  const deltaY = Number.isFinite(options?.deltaY)
    ? Math.round(options?.deltaY as number)
    : 0;
  const allowNegativeOffsets = options?.allowNegativeOffsets === true;
  const shiftedXPx = Number.isFinite(floating.xPx)
    ? Math.round((floating.xPx as number) + deltaX)
    : undefined;
  const shiftedYPx = Number.isFinite(floating.yPx)
    ? Math.round((floating.yPx as number) + deltaY)
    : undefined;
  const horizontalOffset = allowNegativeOffsets
    ? Math.round(shiftedXPx ?? 0)
    : Math.max(0, Math.round(shiftedXPx ?? 0));
  const verticalOffset = allowNegativeOffsets
    ? Math.round(shiftedYPx ?? 0)
    : Math.max(0, Math.round(shiftedYPx ?? 0));
  const horizontalAlign = floating.horizontalAlign?.toLowerCase();
  const hasExplicitHorizontalAlign =
    horizontalAlign === "left" ||
    horizontalAlign === "center" ||
    horizontalAlign === "right" ||
    horizontalAlign === "inside" ||
    horizontalAlign === "outside";
  const containerWidthPx = Number.isFinite(options?.containerWidthPx)
    ? Math.max(1, Math.round(options?.containerWidthPx as number))
    : undefined;
  const imageWidthPx = Number.isFinite(image.widthPx)
    ? Math.max(1, Math.round(image.widthPx as number))
    : undefined;
  const intrinsicBlockWidthStyle: Pick<Record<string, string | number | undefined>, "width"> =
    imageWidthPx ? { width: imageWidthPx } : { width: "fit-content" };
  const rightOffsetPx =
    Number.isFinite(shiftedXPx) &&
    Number.isFinite(shiftedXPx) &&
    Number.isFinite(containerWidthPx) &&
    Number.isFinite(imageWidthPx)
      ? allowNegativeOffsets
        ? Math.round(
            (containerWidthPx as number) -
              (shiftedXPx as number) -
              (imageWidthPx as number)
          )
        : Math.max(
            0,
            Math.round(
              (containerWidthPx as number) -
                (shiftedXPx as number) -
                (imageWidthPx as number)
            )
          )
      : 0;
  const hasExplicitHorizontalOffset = Number.isFinite(shiftedXPx);
  const leftOffsetPx =
    hasExplicitHorizontalAlign && !hasExplicitHorizontalOffset
      ? 0
      : horizontalOffset;
  const topOffsetPx = distT + verticalOffset;

  if (wrapType === "topAndBottom") {
    if (horizontalAlign === "center") {
      return {
        display: "block",
        ...intrinsicBlockWidthStyle,
        marginTop: topOffsetPx,
        marginBottom: distB,
        marginLeft: "auto",
        marginRight: "auto",
        clear: "both",
      };
    }
    if (horizontalAlign === "right" || horizontalAlign === "outside") {
      return {
        display: "block",
        ...intrinsicBlockWidthStyle,
        marginTop: topOffsetPx,
        marginBottom: distB,
        marginLeft: "auto",
        marginRight: hasExplicitHorizontalOffset ? rightOffsetPx : distR,
        clear: "both",
      };
    }
    return {
      display: "block",
      ...intrinsicBlockWidthStyle,
      marginTop: topOffsetPx,
      marginBottom: distB,
      marginLeft: hasExplicitHorizontalOffset
        ? leftOffsetPx
        : distL + leftOffsetPx,
      marginRight: distR,
      clear: "both",
    };
  }

  const side = resolveWrappedFloatingSide(image, {
    containerWidthPx,
    imageWidthPx,
  });
  const horizontalRelativeTo = floating.horizontalRelativeTo?.toLowerCase();
  const explicitHorizontalInsetPx =
    hasExplicitHorizontalOffset &&
    !hasExplicitHorizontalAlign &&
    (horizontalRelativeTo === "margin" ||
      horizontalRelativeTo === "page" ||
      horizontalRelativeTo === "column")
      ? side === "left"
        ? Math.max(0, leftOffsetPx)
        : Math.max(0, rightOffsetPx)
      : 0;
  const columnAnchoredExplicitInset =
    explicitHorizontalInsetPx > 0 && horizontalRelativeTo === "column";
  return {
    display: "block",
    ...intrinsicBlockWidthStyle,
    float: side,
    marginTop: topOffsetPx,
    marginBottom: distB,
    marginLeft:
      side === "left"
        ? explicitHorizontalInsetPx > 0
          ? columnAnchoredExplicitInset
            ? distL + explicitHorizontalInsetPx
            : 0
          : hasExplicitHorizontalOffset
          ? leftOffsetPx
          : distL + leftOffsetPx
        : distL,
    marginRight:
      side === "right"
        ? explicitHorizontalInsetPx > 0
          ? columnAnchoredExplicitInset
            ? distR + explicitHorizontalInsetPx
            : 0
          : hasExplicitHorizontalOffset
          ? rightOffsetPx
          : distR + rightOffsetPx
        : distR,
    paddingLeft:
      side === "left" &&
      explicitHorizontalInsetPx > 0 &&
      !columnAnchoredExplicitInset
        ? explicitHorizontalInsetPx
        : undefined,
    paddingRight:
      side === "right" &&
      explicitHorizontalInsetPx > 0 &&
      !columnAnchoredExplicitInset
        ? explicitHorizontalInsetPx
        : undefined,
    boxSizing:
      explicitHorizontalInsetPx > 0 && !columnAnchoredExplicitInset
        ? "content-box"
        : undefined,
  };
}

interface WrappedFloatingDualExclusionLayout {
  leftSpacerStyle: Record<string, string | number | undefined>;
  rightSpacerStyle: Record<string, string | number | undefined>;
  imageStyle: Record<string, string | number | undefined>;
}

function wrappedFloatingImageDualExclusionLayout(
  image: ImageRunNode,
  options?: {
    containerWidthPx?: number;
    deltaX?: number;
    deltaY?: number;
  }
): WrappedFloatingDualExclusionLayout | undefined {
  void image;
  void options;
  // Center-hole exclusion requires a real line layout engine. The CSS spacer approach
  // created incorrect middle-column wrapping and placement regressions across documents.
  // Keep the stable side-float path until a proper glyph-line layout implementation exists.
  return undefined;
}

export function absoluteFloatingImageStyle(
  image: ImageRunNode,
  options?: {
    pageOriginLeft?: number;
    pageOriginTop?: number;
    marginOriginLeft?: number;
    marginOriginTop?: number;
    columnOriginLeft?: number;
    columnOriginTop?: number;
    paragraphOriginLeft?: number;
    paragraphOriginTop?: number;
    deltaX?: number;
    deltaY?: number;
  }
): Record<string, string | number | undefined> {
  const floating = image.floating;
  if (!floating) {
    return {};
  }

  const horizontalRelativeTo = floating.horizontalRelativeTo?.toLowerCase();
  const verticalRelativeTo = floating.verticalRelativeTo?.toLowerCase();
  const horizontalAlign = floating.horizontalAlign?.toLowerCase();
  const verticalAlign = floating.verticalAlign?.toLowerCase();
  const usesWrapDistance = Boolean(
    floating.wrapType && floating.wrapType !== "none"
  );
  const distL = usesWrapDistance ? floating.distLPx ?? 0 : 0;
  const distR = usesWrapDistance ? floating.distRPx ?? 0 : 0;
  const distT = usesWrapDistance ? floating.distTPx ?? 0 : 0;
  const distB = usesWrapDistance ? floating.distBPx ?? 0 : 0;
  const deltaX = Number.isFinite(options?.deltaX)
    ? Math.round(options?.deltaX as number)
    : 0;
  const deltaY = Number.isFinite(options?.deltaY)
    ? Math.round(options?.deltaY as number)
    : 0;
  const normalizedZIndex = Number.isFinite(floating.zIndex)
    ? Math.max(
        1,
        Math.min(65535, Math.round((floating.zIndex as number) / 65536))
      )
    : undefined;
  const resolvedZIndex = floating.behindDocument
    ? -100000 + (normalizedZIndex ?? 1)
    : normalizedZIndex ?? 4;

  const resolvedLeft =
    floating.xPx !== undefined
      ? horizontalRelativeTo === "page"
        ? floating.xPx + (options?.pageOriginLeft ?? 0)
        : horizontalRelativeTo === "margin"
        ? floating.xPx +
          (options?.marginOriginLeft ?? options?.pageOriginLeft ?? 0)
        : horizontalRelativeTo === "column"
        ? floating.xPx + (options?.columnOriginLeft ?? 0)
        : horizontalRelativeTo === "paragraph" ||
          horizontalRelativeTo === "line"
        ? floating.xPx + (options?.paragraphOriginLeft ?? 0)
        : floating.xPx
      : undefined;
  const resolvedTop =
    floating.yPx !== undefined
      ? verticalRelativeTo === "page"
        ? floating.yPx + (options?.pageOriginTop ?? 0)
        : verticalRelativeTo === "margin"
        ? floating.yPx +
          (options?.marginOriginTop ?? options?.pageOriginTop ?? 0)
        : verticalRelativeTo === "column"
        ? floating.yPx + (options?.columnOriginTop ?? 0)
        : verticalRelativeTo === "paragraph" || verticalRelativeTo === "line"
        ? floating.yPx + (options?.paragraphOriginTop ?? 0)
        : floating.yPx
      : undefined;

  const style: Record<string, string | number | undefined> = {
    position: "absolute",
    zIndex: resolvedZIndex,
  };
  const transforms: string[] = [];

  if (resolvedLeft !== undefined) {
    style.left = resolvedLeft + deltaX;
  } else if (horizontalAlign === "right" || horizontalAlign === "outside") {
    style.right = distR - deltaX;
  } else if (horizontalAlign === "center") {
    style.left = "50%";
    transforms.push("translateX(-50%)");
  } else {
    style.left = distL + deltaX;
  }

  if (resolvedTop !== undefined) {
    style.top = resolvedTop + deltaY;
  } else if (verticalAlign === "bottom" || verticalAlign === "outside") {
    style.bottom = distB - deltaY;
  } else if (verticalAlign === "center") {
    style.top = "50%";
    transforms.push("translateY(-50%)");
  } else {
    style.top = distT + deltaY;
  }

  if (transforms.length > 0 || deltaX !== 0 || deltaY !== 0) {
    const applyDeltaTranslationX =
      resolvedLeft === undefined && horizontalAlign === "center";
    const applyDeltaTranslationY =
      resolvedTop === undefined && verticalAlign === "center";
    const translatePart =
      applyDeltaTranslationX || applyDeltaTranslationY
        ? `translate(${applyDeltaTranslationX ? deltaX : 0}px, ${
            applyDeltaTranslationY ? deltaY : 0
          }px)`
        : "";
    style.transform = [...transforms, translatePart].filter(Boolean).join(" ");
  }

  return style;
}

interface AbsoluteFloatingDropRect {
  left: number;
  top: number;
  width?: number;
  height?: number;
}

export function resolveAbsoluteFloatingImageDropPatch(
  floating: NonNullable<ImageRunNode["floating"]> | undefined,
  layout: Pick<
    DocumentLayoutMetrics,
    "marginsPx" | "pageWidthPx" | "pageHeightPx"
  >,
  options: {
    wrapperRect: AbsoluteFloatingDropRect;
    pageSurfaceRect?: AbsoluteFloatingDropRect;
    deltaX: number;
    deltaY: number;
  }
): Partial<NonNullable<ImageRunNode["floating"]>> {
  const nextLeftPx = options.wrapperRect.left + Math.round(options.deltaX);
  const nextTopPx = options.wrapperRect.top + Math.round(options.deltaY);
  const imageWidthPx = Number.isFinite(options.wrapperRect.width)
    ? Math.max(1, Math.round(options.wrapperRect.width as number))
    : undefined;
  const imageHeightPx = Number.isFinite(options.wrapperRect.height)
    ? Math.max(1, Math.round(options.wrapperRect.height as number))
    : undefined;

  // Preserve any existing side/top alignment so corner-anchored margin floats
  // keep rendering via the fixed-position (absolute) path after a drag instead
  // of flipping to in-flow wrapped rendering (which would re-clamp them to
  // their anchor paragraph and jump away from the drop point). The explicit
  // xPx/yPx below take precedence in the absolute renderer.
  const preservedHorizontalAlign = floating?.horizontalAlign;
  const preservedVerticalAlign = floating?.verticalAlign;

  if (!options.pageSurfaceRect) {
    return {
      xPx: Math.round((floating?.xPx ?? 0) + options.deltaX),
      yPx: Math.round((floating?.yPx ?? 0) + options.deltaY),
      horizontalAlign: preservedHorizontalAlign,
      verticalAlign: preservedVerticalAlign,
      horizontalRelativeTo: "margin",
      verticalRelativeTo: "margin",
    };
  }

  const minimumXPx = -Math.round(layout.marginsPx.left);
  const minimumYPx = -Math.round(layout.marginsPx.top);
  const maximumXPx = Number.isFinite(imageWidthPx)
    ? Math.round(
        layout.pageWidthPx - layout.marginsPx.left - (imageWidthPx as number)
      )
    : Math.round(layout.pageWidthPx - layout.marginsPx.left);
  const maximumYPx = Number.isFinite(imageHeightPx)
    ? Math.round(
        layout.pageHeightPx - layout.marginsPx.top - (imageHeightPx as number)
      )
    : Math.round(layout.pageHeightPx - layout.marginsPx.top);

  return {
    xPx: clampNumber(
      Math.round(
        nextLeftPx - options.pageSurfaceRect.left - layout.marginsPx.left
      ),
      minimumXPx,
      maximumXPx
    ),
    yPx: clampNumber(
      Math.round(
        nextTopPx - options.pageSurfaceRect.top - layout.marginsPx.top
      ),
      minimumYPx,
      maximumYPx
    ),
    horizontalAlign: preservedHorizontalAlign,
    verticalAlign: preservedVerticalAlign,
    horizontalRelativeTo: "margin",
    verticalRelativeTo: "margin",
  };
}

export function resolveWrappedFloatingImageDropPatch(
  image: ImageRunNode,
  hostWidth: number,
  movedLeft: number,
  movedTop: number,
  options?: {
    widthPx?: number;
    heightPx?: number;
  }
): Partial<NonNullable<ImageRunNode["floating"]>> {
  const baseFloating = image.floating ?? {};
  const imageWidth =
    options?.widthPx ?? image.widthPx ?? MIN_PARAGRAPH_LINE_HEIGHT_PX;
  const imageHeight = options?.heightPx ?? image.heightPx ?? imageWidth;
  const wrapText = baseFloating.wrapText ?? "bothSides";
  const wrapType = baseFloating.wrapType ?? "square";
  const convertFixedPositionToMoveWithText =
    !floatingImageMovesWithText(baseFloating);
  const preserveAlignedHorizontalPlacement =
    !convertFixedPositionToMoveWithText &&
    wrapType.trim().toLowerCase() === "topandbottom" &&
    Boolean(baseFloating.horizontalAlign);
  const side: "left" | "right" =
    movedLeft + imageWidth / 2 <= hostWidth / 2 ? "left" : "right";
  const previewGeometry = resolveDualWrappedFloatingImageGeometry(
    {
      ...image,
      widthPx: imageWidth,
      heightPx: imageHeight,
    },
    hostWidth,
    {
      widthPx: imageWidth,
      heightPx: imageHeight,
      baseLeftPx: movedLeft,
      baseTopPx: movedTop,
      allowNegativeImageTop: true,
    }
  );
  const previewUsesSideFloat =
    previewGeometry === undefined ||
    previewGeometry.exclusion.left <= 0 ||
    previewGeometry.exclusion.right >= hostWidth;
  const rawExplicitTopPx = Math.round(
    movedTop - Math.round(baseFloating.distTPx ?? 0)
  );
  const currentTopPx = Math.round(baseFloating.yPx ?? 0);
  const explicitTopPx =
    wrapType.trim().toLowerCase() === "topandbottom" &&
    Math.abs(rawExplicitTopPx - currentTopPx) <
      TOP_AND_BOTTOM_VERTICAL_DRAG_SNAP_PX
      ? currentTopPx
      : rawExplicitTopPx;

  return {
    wrapType,
    wrapText,
    horizontalAlign: preserveAlignedHorizontalPlacement
      ? baseFloating.horizontalAlign
      : previewUsesSideFloat
      ? side
      : wrapText === "bothSides" || wrapText === "largest"
      ? undefined
      : side,
    xPx: preserveAlignedHorizontalPlacement ? undefined : Math.round(movedLeft),
    yPx: explicitTopPx,
    distLPx: Math.max(4, Math.round(baseFloating.distLPx ?? 8)),
    distRPx: Math.max(4, Math.round(baseFloating.distRPx ?? 8)),
    distTPx: Math.max(0, Math.round(baseFloating.distTPx ?? 2)),
    distBPx: Math.max(0, Math.round(baseFloating.distBPx ?? 4)),
    horizontalRelativeTo: convertFixedPositionToMoveWithText
      ? "column"
      : baseFloating.horizontalRelativeTo ?? "column",
    verticalRelativeTo: convertFixedPositionToMoveWithText
      ? "paragraph"
      : baseFloating.verticalRelativeTo ?? "paragraph",
    behindDocument: false,
  };
}

function resolvePageSpanningAbsoluteFloatingDimensions(
  image: ImageRunNode,
  widthPx?: number,
  heightPx?: number,
  floatingPageOriginPx?: {
    left: number;
    top: number;
    marginLeft?: number;
    marginTop?: number;
    pageWidth?: number;
  }
): {
  widthPx?: number;
  heightPx?: number;
} {
  const floating = image.floating;
  const pageWidthPx = Number.isFinite(floatingPageOriginPx?.pageWidth)
    ? Math.max(1, Math.round(floatingPageOriginPx?.pageWidth as number))
    : undefined;
  if (
    !floating ||
    !shouldRenderAbsoluteFloatingImage(image) ||
    floating.behindDocument !== true ||
    (floating.wrapType !== undefined && floating.wrapType !== "none") ||
    !Number.isFinite(pageWidthPx)
  ) {
    return { widthPx, heightPx };
  }

  const horizontalRelativeTo = floating.horizontalRelativeTo?.toLowerCase();
  if (horizontalRelativeTo !== "page" && horizontalRelativeTo !== "margin") {
    return { widthPx, heightPx };
  }

  const currentWidthPx = Number.isFinite(widthPx)
    ? Math.max(1, Math.round(widthPx as number))
    : undefined;
  if (
    Number.isFinite(currentWidthPx) &&
    (currentWidthPx as number) < (pageWidthPx as number) * 0.55
  ) {
    return { widthPx, heightPx };
  }

  const resolvedLeftPx = Number.isFinite(floating.xPx)
    ? horizontalRelativeTo === "margin"
      ? Math.round(
          (floating.xPx as number) +
            (floatingPageOriginPx?.marginLeft ??
              floatingPageOriginPx?.left ??
              0)
        )
      : Math.round(floating.xPx as number)
    : horizontalRelativeTo === "margin"
    ? Math.round(
        floatingPageOriginPx?.marginLeft ?? floatingPageOriginPx?.left ?? 0
      )
    : 0;
  const minimumWidthPx = Math.max(
    1,
    Math.round((pageWidthPx as number) - Math.max(0, resolvedLeftPx))
  );
  if (
    Number.isFinite(currentWidthPx) &&
    (currentWidthPx as number) >= minimumWidthPx
  ) {
    return { widthPx, heightPx };
  }

  const nextWidthPx = minimumWidthPx;
  const nextHeightPx =
    Number.isFinite(widthPx) &&
    Number.isFinite(heightPx) &&
    (widthPx as number) > 0
      ? Math.max(
          1,
          Math.round(((heightPx as number) * nextWidthPx) / (widthPx as number))
        )
      : heightPx;

  return {
    widthPx: nextWidthPx,
    heightPx: nextHeightPx,
  };
}

function shouldRenderAbsoluteFloatingImage(image: ImageRunNode): boolean {
  const floating = image.floating;
  if (!floating) {
    return false;
  }

  if (shouldRenderWrappedFloatingImage(image)) {
    return false;
  }

  return (
    floating.xPx !== undefined ||
    floating.yPx !== undefined ||
    floating.horizontalAlign !== undefined ||
    floating.verticalAlign !== undefined ||
    floating.zIndex !== undefined ||
    floating.behindDocument === true
  );
}

function isPageOrMarginAnchoredAbsoluteFloatingImage(
  image: ImageRunNode
): boolean {
  if (!shouldRenderAbsoluteFloatingImage(image)) {
    return false;
  }

  const floating = image.floating;
  if (!floating) {
    return false;
  }

  const horizontalRelativeTo = floating.horizontalRelativeTo?.toLowerCase();
  const verticalRelativeTo = floating.verticalRelativeTo?.toLowerCase();
  const horizontalPageAnchored =
    horizontalRelativeTo === "page" || horizontalRelativeTo === "margin";
  const verticalPageAnchored =
    verticalRelativeTo === "page" || verticalRelativeTo === "margin";

  // Use page-level absolute positioning context only when both axes are page/margin
  // anchored. Mixed anchors (e.g. horizontal=page + vertical=line) must stay
  // paragraph-anchored so the line-relative axis remains stable.
  return horizontalPageAnchored && verticalPageAnchored;
}

function isPageOrMarginAnchoredWrappedFloatingImage(
  image: ImageRunNode
): boolean {
  if (!shouldRenderWrappedFloatingImage(image)) {
    return false;
  }

  const floating = image.floating;
  if (!floating) {
    return false;
  }

  const horizontalRelativeTo = floating.horizontalRelativeTo?.toLowerCase();
  const verticalRelativeTo = floating.verticalRelativeTo?.toLowerCase();
  const horizontalPageAnchored =
    horizontalRelativeTo === "page" || horizontalRelativeTo === "margin";
  const verticalPageAnchored =
    verticalRelativeTo === "page" || verticalRelativeTo === "margin";

  return horizontalPageAnchored && verticalPageAnchored;
}

function paragraphNeedsPageAnchoredAbsolutePositioningContext(
  paragraph: ParagraphNode
): boolean {
  return paragraph.children.some(
    (child) =>
      child.type === "image" &&
      (isPageOrMarginAnchoredAbsoluteFloatingImage(child) ||
        isPageOrMarginAnchoredWrappedFloatingImage(child))
  );
}

function paragraphNeedsLocalAbsolutePositioningContext(
  paragraph: ParagraphNode
): boolean {
  return paragraph.children.some(
    (child) =>
      child.type === "image" &&
      shouldRenderAbsoluteFloatingImage(child) &&
      !isPageOrMarginAnchoredAbsoluteFloatingImage(child)
  );
}

function paragraphHasPageAnchoredAbsoluteFloatingImage(
  paragraph: ParagraphNode
): boolean {
  return paragraph.children.some(
    (child) =>
      child.type === "image" &&
      isPageOrMarginAnchoredAbsoluteFloatingImage(child)
  );
}

function paragraphHasPageAnchoredForegroundAbsoluteFloatingImage(
  paragraph: ParagraphNode
): boolean {
  return paragraph.children.some(
    (child) =>
      child.type === "image" &&
      isPageOrMarginAnchoredAbsoluteFloatingImage(child) &&
      child.floating?.behindDocument !== true
  );
}

function imageCropLayout(
  image: ImageRunNode,
  widthPx?: number,
  heightPx?: number
):
  | {
      frameWidthPx: number;
      frameHeightPx: number;
      imageWidthPx: number;
      imageHeightPx: number;
      offsetXPx: number;
      offsetYPx: number;
    }
  | undefined {
  const crop = image.crop;
  if (
    !crop ||
    !Number.isFinite(widthPx) ||
    !Number.isFinite(heightPx) ||
    (widthPx as number) <= 0 ||
    (heightPx as number) <= 0
  ) {
    return undefined;
  }

  const left = Math.max(0, Math.min(1, crop.leftFraction ?? 0));
  const top = Math.max(0, Math.min(1, crop.topFraction ?? 0));
  const right = Math.max(0, Math.min(1, crop.rightFraction ?? 0));
  const bottom = Math.max(0, Math.min(1, crop.bottomFraction ?? 0));
  const visibleWidthFraction = 1 - left - right;
  const visibleHeightFraction = 1 - top - bottom;
  if (visibleWidthFraction <= 0 || visibleHeightFraction <= 0) {
    return undefined;
  }

  const frameWidthPx = Math.max(1, Math.round(widthPx as number));
  const frameHeightPx = Math.max(1, Math.round(heightPx as number));
  const imageWidthPx = Math.max(
    1,
    Math.round(frameWidthPx / visibleWidthFraction)
  );
  const imageHeightPx = Math.max(
    1,
    Math.round(frameHeightPx / visibleHeightFraction)
  );
  const offsetXPx = Math.round(imageWidthPx * left);
  const offsetYPx = Math.round(imageHeightPx * top);

  return {
    frameWidthPx,
    frameHeightPx,
    imageWidthPx,
    imageHeightPx,
    offsetXPx,
    offsetYPx,
  };
}

function resolveHighlightColor(value?: string): string | undefined {
  if (!value) {
    return undefined;
  }

  const normalized = value.trim().toLowerCase();
  if (!normalized) {
    return undefined;
  }

  if (normalized.startsWith("#")) {
    return normalized;
  }

  return HIGHLIGHT_TO_CSS[normalized] ?? normalized;
}

function paragraphLineHeight(
  paragraph: ParagraphNode,
  docGridLinePitchPx?: number,
  disableDocGridSnap = false
): number | string | undefined {
  const baseFontFamily = paragraphDominantFontFamily(paragraph);
  const singleLineScale = resolveParagraphSingleLineAutoScale(
    paragraph,
    baseFontFamily
  );
  const lineTwips = paragraph.style?.spacing?.lineTwips;
  const docGridMinimumLineHeightPx = resolveParagraphDocGridLinePitchPx(
    paragraph,
    docGridLinePitchPx,
    disableDocGridSnap
  );
  if (!Number.isFinite(lineTwips)) {
    if (docGridMinimumLineHeightPx) {
      return `${estimateParagraphLineHeightPx(
        paragraph,
        docGridLinePitchPx,
        disableDocGridSnap
      )}px`;
    }
    return calibrateAutoLineSpacingMultiple(
      DEFAULT_PARAGRAPH_LINE_MULTIPLE,
      baseFontFamily,
      singleLineScale
    );
  }

  const lineRule = paragraph.style?.spacing?.lineRule ?? "auto";
  if (lineRule === "auto") {
    if (docGridMinimumLineHeightPx) {
      return `${estimateParagraphLineHeightPx(
        paragraph,
        docGridLinePitchPx,
        disableDocGridSnap
      )}px`;
    }
    const resolvedAutoMultiple = resolveAutoLineSpacingMultiple(
      lineTwips as number,
      DEFAULT_PARAGRAPH_LINE_MULTIPLE
    );
    // Keep the rendered strut in lockstep with estimateParagraphLineHeightPx:
    // text-free paragraphs scale the natural single line by the multiple
    // instead of blending toward bare font-size lines.
    const lineMultiple = paragraphRendersTextFreeLine(paragraph)
      ? Math.max(
          MIN_AUTO_LINE_MULTIPLE,
          Number((resolvedAutoMultiple * singleLineScale).toFixed(3))
        )
      : calibrateAutoLineSpacingMultiple(
          resolvedAutoMultiple,
          baseFontFamily,
          singleLineScale
        );
    return Number(lineMultiple.toFixed(3));
  }

  const lineHeightPx = twipsToPixels(lineTwips);
  if (lineRule === "atLeast") {
    const normalLineHeightPx = Math.max(
      1,
      Math.round(
        paragraphBaseFontSizePx(paragraph) *
          calibrateAutoLineSpacingMultiple(
            DEFAULT_PARAGRAPH_LINE_MULTIPLE,
            baseFontFamily,
            singleLineScale
          )
      )
    );
    return `${Math.max(
      normalLineHeightPx,
      lineHeightPx ?? 0,
      docGridMinimumLineHeightPx ?? 0
    )}px`;
  }

  if (lineHeightPx && lineHeightPx > 0) {
    return `${lineHeightPx}px`;
  }

  if (lineRule === "exact") {
    return calibrateAutoLineSpacingMultiple(
      DEFAULT_PARAGRAPH_LINE_MULTIPLE,
      baseFontFamily,
      singleLineScale
    );
  }

  return calibrateAutoLineSpacingMultiple(
    DEFAULT_PARAGRAPH_LINE_MULTIPLE,
    baseFontFamily,
    singleLineScale
  );
}

function keepNextPaginationReservePx(
  paragraph: ParagraphNode,
  nextParagraph: ParagraphNode | undefined,
  docGridLinePitchPx?: number
): number {
  if (
    paragraph.style?.keepNext !== true ||
    !nextParagraph ||
    !Number.isFinite(paragraph.style?.headingLevel)
  ) {
    return 0;
  }

  const nextParagraphText = paragraphText(nextParagraph)
    .replace(/\s+/g, " ")
    .trim();
  if (
    nextParagraph.style?.numbering === undefined &&
    nextParagraphText.length < 80
  ) {
    return 0;
  }

  return Math.max(
    nextParagraph.style?.numbering ? 10 : 6,
    Math.round(
      estimateParagraphLineHeightPx(nextParagraph, docGridLinePitchPx) *
        (nextParagraph.style?.numbering ? 1 : 0.5)
    )
  );
}

function paragraphBorderToCss(
  border: ParagraphBorderStyle | undefined
): string | undefined {
  return tableBorderToCss(border);
}

function paragraphBorderPaddingPx(
  border: ParagraphBorderStyle | undefined
): number | undefined {
  const type = normalizeBorderType(border?.type);
  if (!type || type === "none" || type === "nil") {
    return undefined;
  }

  return pointsToPixels(border?.spacePt);
}

function paragraphBorderStrokeWidthPx(
  border: ParagraphBorderStyle | undefined
): number {
  const type = normalizeBorderType(border?.type);
  if (!type || type === "none" || type === "nil") {
    return 0;
  }

  const sizeEighthPt = border?.sizeEighthPt;
  return Number.isFinite(sizeEighthPt) && (sizeEighthPt as number) > 0
    ? Math.max(0.5, Number(((sizeEighthPt as number) / 6).toFixed(2)))
    : 1;
}

function paragraphBorderInsetPx(
  border: ParagraphBorderStyle | undefined
): number {
  const type = normalizeBorderType(border?.type);
  if (!type || type === "none" || type === "nil") {
    return 0;
  }

  return (
    paragraphBorderStrokeWidthPx(border) +
    (paragraphBorderPaddingPx(border) ?? 0)
  );
}

function paragraphExplicitIndentTwips(
  paragraph: ParagraphNode
): ParagraphIndent | undefined {
  const sourceXml = paragraph.sourceXml;
  if (!sourceXml) {
    return undefined;
  }

  const cached = paragraphExplicitIndentBySourceXml.get(sourceXml);
  if (cached !== undefined) {
    return cached === null ? undefined : cached;
  }

  const paragraphPropertiesXml =
    sourceXml.match(/<w:pPr\b[\s\S]*?<\/w:pPr>/i)?.[0] ??
    sourceXml.match(/<w:pPr\b[^>]*\/>/i)?.[0];
  if (!paragraphPropertiesXml) {
    paragraphExplicitIndentBySourceXml.set(sourceXml, null);
    return undefined;
  }

  const indentTag = paragraphPropertiesXml.match(/<w:ind\b[^>]*\/?>/i)?.[0];
  if (!indentTag) {
    paragraphExplicitIndentBySourceXml.set(sourceXml, null);
    return undefined;
  }

  const parseIndentTwips = (attribute: string): number | undefined => {
    const match = indentTag.match(
      new RegExp(`\\b${attribute}="(-?\\d+)"`, "i")
    );
    if (!match?.[1]) {
      return undefined;
    }

    const parsed = Number.parseInt(match[1], 10);
    return Number.isFinite(parsed) ? parsed : undefined;
  };

  let firstLineTwips = parseIndentTwips("w:firstLine");
  const hangingTwips = parseIndentTwips("w:hanging");
  if (Number.isFinite(firstLineTwips) && Number.isFinite(hangingTwips)) {
    // ECMA-376: when both firstLine and hanging are present, firstLine is ignored.
    firstLineTwips = undefined;
  }

  const explicitIndent: ParagraphIndent = {
    leftTwips: parseIndentTwips("w:left"),
    firstLineTwips,
    hangingTwips,
  };
  const hasAnyExplicitIndent =
    Number.isFinite(explicitIndent.leftTwips) ||
    Number.isFinite(explicitIndent.firstLineTwips) ||
    Number.isFinite(explicitIndent.hangingTwips);
  if (!hasAnyExplicitIndent) {
    paragraphExplicitIndentBySourceXml.set(sourceXml, null);
    return undefined;
  }

  paragraphExplicitIndentBySourceXml.set(sourceXml, explicitIndent);
  return explicitIndent;
}

function resolveListParagraphIndent(
  paragraph: ParagraphNode,
  numberingDefinitions?: NumberingDefinitionSet
): ParagraphIndent | undefined {
  const numbering = paragraph.style?.numbering;
  if (!numbering || !Number.isFinite(numbering.numId) || numbering.numId <= 0) {
    return paragraph.style?.indent;
  }

  if (!numberingDefinitions) {
    return paragraph.style?.indent;
  }

  const effectiveNumId =
    effectiveNumberingNumIdForParagraph(paragraph, numberingDefinitions) ??
    numbering.numId;
  const numberingRecoveryActive = effectiveNumId !== numbering.numId;
  const ilvl = Math.max(0, Math.round(numbering.ilvl ?? 0));
  const level = findNumberingLevelDefinition(
    numberingDefinitions,
    effectiveNumId,
    ilvl
  );
  const baseLevel = findNumberingLevelDefinition(
    numberingDefinitions,
    effectiveNumId,
    0
  );
  const levelIndent = level?.indent;
  const styleIndent = paragraph.style?.indent;
  const numberingHasVisibleMarker = Boolean(
    (level?.text && level.text.trim().length > 0) || level?.pictureBullet?.src
  );
  const numberingProvidesUsableIndent = Boolean(
    Number.isFinite(levelIndent?.leftTwips) ||
      Number.isFinite(levelIndent?.firstLineTwips) ||
      Number.isFinite(levelIndent?.hangingTwips)
  );
  const styleLeftTwips = styleIndent?.leftTwips;
  const explicitParagraphIndent = paragraphExplicitIndentTwips(paragraph);
  const explicitParagraphLeftTwips = explicitParagraphIndent?.leftTwips;
  const explicitParagraphFirstLineTwips =
    explicitParagraphIndent?.firstLineTwips;
  const explicitParagraphHangingTwips = explicitParagraphIndent?.hangingTwips;
  const hasExplicitParagraphFirstLineTwips = Number.isFinite(
    explicitParagraphFirstLineTwips
  );
  const hasExplicitParagraphHangingTwips = Number.isFinite(
    explicitParagraphHangingTwips
  );
  const preferRecoveredNumberingTextIndent = numberingRecoveryActive;
  if (!numberingHasVisibleMarker && !numberingProvidesUsableIndent) {
    return styleIndent;
  }
  const baseLevelLeftTwips = Number.isFinite(baseLevel?.indent?.leftTwips)
    ? baseLevel?.indent?.leftTwips ?? 0
    : Number.isFinite(styleLeftTwips)
    ? styleLeftTwips
    : undefined;
  const levelLeftTwips = levelIndent?.leftTwips;
  const hasExplicitLevelLeftTwips = Number.isFinite(levelLeftTwips);
  const hasExplicitStyleLeftTwips = Number.isFinite(styleLeftTwips);
  const hasExplicitParagraphLeftTwips = Number.isFinite(
    explicitParagraphLeftTwips
  );
  let nextLeftTwips = styleLeftTwips;

  if (hasExplicitParagraphLeftTwips) {
    nextLeftTwips = explicitParagraphLeftTwips;
  } else if (
    Number.isFinite(levelLeftTwips) &&
    Number.isFinite(baseLevelLeftTwips) &&
    Number.isFinite(styleLeftTwips)
  ) {
    const levelOffsetTwips = Math.max(
      0,
      (levelLeftTwips ?? 0) - (baseLevelLeftTwips ?? 0)
    );
    // When paragraph style indentation comes from defaults (often 0), it should not
    // suppress the numbering level indentation from DOCX.
    const styleUsesListBaseIndent =
      (styleLeftTwips as number) >= (baseLevelLeftTwips as number) - 120;
    nextLeftTwips = styleUsesListBaseIndent
      ? (styleLeftTwips ?? 0) + levelOffsetTwips
      : levelLeftTwips ?? 0;
  } else if (Number.isFinite(levelLeftTwips)) {
    nextLeftTwips = levelLeftTwips;
  } else if (ilvl > 0) {
    nextLeftTwips = ilvl * LIST_LEVEL_STEP_TWIPS;
  }

  // Some documents provide numbering but omit usable paragraph indents for list
  // paragraphs. Preserve explicit zero indents from OOXML and only synthesize
  // a fallback list indent when both style and numbering level omit left indents.
  if (!Number.isFinite(nextLeftTwips)) {
    if (hasExplicitParagraphLeftTwips) {
      nextLeftTwips = explicitParagraphLeftTwips;
    } else if (hasExplicitLevelLeftTwips) {
      nextLeftTwips = levelLeftTwips;
    } else if (hasExplicitStyleLeftTwips) {
      nextLeftTwips = styleLeftTwips;
    } else {
      nextLeftTwips = Math.max(
        LIST_LEVEL_STEP_TWIPS,
        (ilvl + 1) * LIST_LEVEL_STEP_TWIPS
      );
    }
  } else if ((nextLeftTwips as number) <= 0) {
    if (hasExplicitParagraphLeftTwips) {
      nextLeftTwips = explicitParagraphLeftTwips;
    } else if (hasExplicitLevelLeftTwips) {
      nextLeftTwips = levelLeftTwips;
    } else if (hasExplicitStyleLeftTwips) {
      nextLeftTwips = styleLeftTwips;
    } else if (
      Number.isFinite(levelLeftTwips) &&
      (levelLeftTwips as number) > 0
    ) {
      nextLeftTwips = levelLeftTwips;
    } else if (
      Number.isFinite(styleLeftTwips) &&
      (styleLeftTwips as number) > 0
    ) {
      nextLeftTwips = styleLeftTwips;
    } else {
      nextLeftTwips = Math.max(
        LIST_LEVEL_STEP_TWIPS,
        (ilvl + 1) * LIST_LEVEL_STEP_TWIPS
      );
    }
  }

  if (!Number.isFinite(nextLeftTwips)) {
    return styleIndent;
  }

  const nextLeftTwipsRounded = Math.max(0, Math.round(nextLeftTwips ?? 0));
  let nextFirstLineTwips: number | undefined;
  let nextHangingTwips: number | undefined;
  if (preferRecoveredNumberingTextIndent) {
    nextFirstLineTwips = levelIndent?.firstLineTwips;
    nextHangingTwips = levelIndent?.hangingTwips;
  } else if (
    hasExplicitParagraphFirstLineTwips ||
    hasExplicitParagraphHangingTwips
  ) {
    // Paragraph-level w:ind overrides numbering/style indentation semantics.
    // If firstLine is explicitly present without hanging, do not inherit hanging.
    // If hanging is explicitly present without firstLine, do not inherit firstLine.
    nextFirstLineTwips = hasExplicitParagraphFirstLineTwips
      ? explicitParagraphFirstLineTwips
      : undefined;
    nextHangingTwips = hasExplicitParagraphHangingTwips
      ? explicitParagraphHangingTwips
      : undefined;
  } else {
    nextFirstLineTwips = Number.isFinite(styleIndent?.firstLineTwips)
      ? styleIndent?.firstLineTwips
      : levelIndent?.firstLineTwips;
    nextHangingTwips = Number.isFinite(styleIndent?.hangingTwips)
      ? styleIndent?.hangingTwips
      : levelIndent?.hangingTwips;
  }

  if (
    Number.isFinite(nextFirstLineTwips) &&
    Number.isFinite(nextHangingTwips)
  ) {
    // ECMA-376: firstLine is ignored when both values are present.
    nextFirstLineTwips = undefined;
  }

  return {
    ...styleIndent,
    leftTwips: nextLeftTwipsRounded,
    firstLineTwips: nextFirstLineTwips,
    hangingTwips: nextHangingTwips,
  };
}

function resolveNumberingMarkerBoxWidthPx(
  paragraph: ParagraphNode,
  numberingDefinitions?: NumberingDefinitionSet,
  numberingLabel?: ParagraphNumberingLabel
): number | undefined {
  const resolvedIndent = resolveListParagraphIndent(
    paragraph,
    numberingDefinitions
  );
  const hangingIndentPx = twipsToSignedPixels(resolvedIndent?.hangingTwips);
  const firstLineIndentPx = twipsToSignedPixels(resolvedIndent?.firstLineTwips);
  const candidateWidthPx =
    Number.isFinite(hangingIndentPx) && Math.abs(hangingIndentPx as number) > 0
      ? Math.abs(hangingIndentPx as number)
      : Number.isFinite(firstLineIndentPx) && (firstLineIndentPx as number) < 0
      ? Math.abs(firstLineIndentPx as number)
      : undefined;

  const labelTextForWidth = numberingLabel?.imageSrc
    ? numberingLabel.trailingText ?? ""
    : numberingLabel?.text ?? "";
  const normalizedLabelTextForWidth = labelTextForWidth
    .replace(/\t/g, " ")
    .replace(/\u00a0/g, " ");
  const measuredLabelWidthPx =
    normalizedLabelTextForWidth.length > 0
      ? Math.ceil(
          measureTextWidthPx(
            normalizedLabelTextForWidth,
            numberingLabel?.style,
            paragraphBaseFontSizePx(paragraph)
          )
        ) + 4
      : 0;

  const imageWidthPx = numberingLabel?.imageWidthPx;
  const tocLeadingLeftTabStopPx =
    tableOfContentsLeadingLeftTabStopPx(paragraph);
  const tocLabelGapPx = isTableOfContentsParagraph(paragraph) ? 6 : 0;
  const tocMarkerTargetWidthPx =
    Number.isFinite(tocLeadingLeftTabStopPx) &&
    (tocLeadingLeftTabStopPx as number) > 0
      ? Math.min(
          Math.ceil(tocLeadingLeftTabStopPx as number),
          Math.max(
            measuredLabelWidthPx + tocLabelGapPx,
            measuredLabelWidthPx + 16
          )
        )
      : 0;
  const minimumVisualWidthPx = Math.max(
    measuredLabelWidthPx + tocLabelGapPx,
    Number.isFinite(imageWidthPx) && (imageWidthPx as number) > 0
      ? Math.ceil(imageWidthPx as number) + 2
      : 0,
    tocMarkerTargetWidthPx
  );

  if (!Number.isFinite(candidateWidthPx) || (candidateWidthPx as number) < 8) {
    return minimumVisualWidthPx > 0
      ? clampNumber(Math.round(minimumVisualWidthPx), 8, 220)
      : undefined;
  }

  return Math.max(
    minimumVisualWidthPx,
    clampNumber(Math.round(candidateWidthPx as number), 8, 220)
  );
}

function numberingMarkerStyle(
  paragraph: ParagraphNode,
  numberingDefinitions: NumberingDefinitionSet | undefined,
  numberingLabel: ParagraphNumberingLabel,
  baseStyle: Record<string, string | number | undefined> | undefined,
  documentTheme: DocxDocumentTheme
): Record<string, string | number | undefined> {
  const markerBoxWidthPx = resolveNumberingMarkerBoxWidthPx(
    paragraph,
    numberingDefinitions,
    numberingLabel
  );
  const markerGapPx = isTableOfContentsParagraph(paragraph)
    ? 6
    : markerBoxWidthPx
    ? 0
    : 2;

  return {
    ...(baseStyle ?? {}),
    display: "inline-flex",
    alignItems: "baseline",
    justifyContent: "flex-end",
    verticalAlign: "baseline",
    width: markerBoxWidthPx ? `${markerBoxWidthPx}px` : undefined,
    minWidth: markerBoxWidthPx ? `${markerBoxWidthPx}px` : "1.1em",
    marginRight: markerGapPx,
    whiteSpace: "pre",
    fontFamily: cssFontFamily(
      numberingLabel.fontFamily ?? numberingLabel.style?.fontFamily
    ),
    color: themedRunColor(
      numberingLabel.color ?? numberingLabel.style?.color,
      documentTheme
    ),
  };
}

function paragraphBlockStyle(
  paragraph: ParagraphNode,
  numberingDefinitions?: NumberingDefinitionSet,
  headingStyles?: DocxHeadingStyleMap,
  docGridLinePitchPx?: number,
  disableDocGridSnap = false
): Record<string, string | number | undefined> {
  const beforeSpacing =
    twipsToPixels(paragraph.style?.spacing?.beforeTwips) ?? 0;
  const afterSpacing = twipsToPixels(paragraph.style?.spacing?.afterTwips) ?? 0;
  const resolvedIndent = resolveListParagraphIndent(
    paragraph,
    numberingDefinitions
  );
  const leftIndent = twipsToSignedPixels(resolvedIndent?.leftTwips) ?? 0;
  const rightIndent =
    twipsToSignedPixels(paragraph.style?.indent?.rightTwips) ?? 0;
  const firstLineIndent = twipsToSignedPixels(resolvedIndent?.firstLineTwips);
  const hangingIndent = twipsToSignedPixels(resolvedIndent?.hangingTwips);
  const topBorder = paragraphBorderToCss(paragraph.style?.borders?.top);
  const rightBorder = paragraphBorderToCss(paragraph.style?.borders?.right);
  const bottomBorder = paragraphBorderToCss(paragraph.style?.borders?.bottom);
  const leftBorder = paragraphBorderToCss(paragraph.style?.borders?.left);
  const topPadding = paragraphBorderPaddingPx(paragraph.style?.borders?.top);
  const rightPadding = paragraphBorderPaddingPx(
    paragraph.style?.borders?.right
  );
  const bottomPadding = paragraphBorderPaddingPx(
    paragraph.style?.borders?.bottom
  );
  const leftPadding = paragraphBorderPaddingPx(paragraph.style?.borders?.left);
  const checkboxChoiceRow = paragraphLooksLikeCheckboxChoiceRow(paragraph);
  const suppressTocNumberingTextIndent =
    isTableOfContentsParagraph(paragraph) && paragraphHasNumbering(paragraph);
  const suppressIndentForFloatingAnchorOnlyParagraph =
    paragraphIsFloatingImageAnchorOnly(paragraph);
  const suppressStackingContextForBehindTextAnchorOnlyParagraph =
    paragraphIsBehindTextAbsoluteFloatingImageAnchorOnly(paragraph);
  const reservedMinHeightPx = paragraphIsEffectivelyEmpty(paragraph)
    ? estimateParagraphLineHeightPx(
        paragraph,
        docGridLinePitchPx,
        disableDocGridSnap
      ) + EMPTY_PARAGRAPH_EXTRA_HEIGHT_PX
    : undefined;
  const headingLevel = paragraph.style?.headingLevel;
  const applyWordLikeHeadingFallback = !paragraph.sourceXml;
  const hasSoftLineBreak = paragraphText(paragraph).includes("\n");
  const headingStyle =
    headingLevel && headingLevel >= 1 && headingLevel <= 6
      ? headingStyles?.[headingLevel] ??
        (applyWordLikeHeadingFallback
          ? DEFAULT_WORD_HEADING_STYLES[headingLevel]
          : undefined)
      : undefined;

  return {
    position: "relative",
    zIndex: suppressStackingContextForBehindTextAnchorOnlyParagraph
      ? undefined
      : 1,
    overflow: "visible",
    textAlign: paragraph.style?.align ?? "left",
    ...(paragraph.style?.align === "justify" && hasSoftLineBreak
      ? ({ textAlignLast: "justify" } as Record<string, string | number | undefined>)
      : undefined),
    // Pin the block font size to the paragraph's dominant run size so the
    // line-box strut tracks the actual content instead of the browser's
    // 16px default, which inflates lines for sub-12pt paragraphs.
    fontSize: `${paragraphBaseFontSizePx(paragraph)}px`,
    lineHeight: paragraphLineHeight(
      paragraph,
      docGridLinePitchPx,
      disableDocGridSnap
    ),
    marginTop: beforeSpacing,
    marginBottom: afterSpacing,
    marginLeft: suppressIndentForFloatingAnchorOnlyParagraph ? 0 : leftIndent,
    marginRight: suppressIndentForFloatingAnchorOnlyParagraph ? 0 : rightIndent,
    backgroundColor: paragraph.style?.backgroundColor,
    textIndent:
      suppressTocNumberingTextIndent ||
      suppressIndentForFloatingAnchorOnlyParagraph
        ? undefined
        : firstLineIndent ?? (hangingIndent ? -hangingIndent : undefined),
    minHeight: Number.isFinite(reservedMinHeightPx)
      ? `${reservedMinHeightPx}px`
      : undefined,
    ...(headingStyle ?? undefined),
    ...(topBorder !== undefined ? { borderTop: topBorder } : undefined),
    ...(rightBorder !== undefined ? { borderRight: rightBorder } : undefined),
    ...(bottomBorder !== undefined
      ? { borderBottom: bottomBorder }
      : undefined),
    ...(leftBorder !== undefined ? { borderLeft: leftBorder } : undefined),
    ...(topPadding !== undefined ? { paddingTop: topPadding } : undefined),
    ...(rightPadding !== undefined
      ? { paddingRight: rightPadding }
      : undefined),
    ...(bottomPadding !== undefined
      ? { paddingBottom: bottomPadding }
      : undefined),
    ...(leftPadding !== undefined ? { paddingLeft: leftPadding } : undefined),
    whiteSpace: checkboxChoiceRow ? "nowrap" : "pre-wrap",
    ...(checkboxChoiceRow
      ? ({ tabSize: 1 } as Record<string, string | number | undefined>)
      : undefined),
    wordWrap: checkboxChoiceRow ? "normal" : "break-word",
    overflowWrap: checkboxChoiceRow ? "normal" : "break-word",
    wordBreak: checkboxChoiceRow ? "normal" : "break-word",
  };
}

function tableCellParagraphBlockStyle(
  paragraph: ParagraphNode,
  numberingDefinitions: NumberingDefinitionSet | undefined,
  headingStyles: DocxHeadingStyleMap | undefined,
  paragraphIndex: number,
  applyWordTableDefaults: boolean,
  docGridLinePitchPx?: number
): Record<string, string | number | undefined> {
  const paragraphForLayout = wordLikeTableCellParagraph(
    paragraph,
    applyWordTableDefaults
  );
  const disableDocGridSnap = paragraphDocGridSnapState(paragraph) === "disable";
  const baseStyle = paragraphBlockStyle(
    paragraphForLayout,
    numberingDefinitions,
    headingStyles,
    docGridLinePitchPx,
    disableDocGridSnap
  );
  const suppressTopSpacing =
    paragraphIndex <= 0 && suppressFirstTableCellParagraphTopSpacing(paragraph);
  const normalizedMarginLeft =
    typeof baseStyle.marginLeft === "number"
      ? baseStyle.marginLeft
      : Number.parseFloat(String(baseStyle.marginLeft ?? ""));
  const normalizedMarginRight =
    typeof baseStyle.marginRight === "number"
      ? baseStyle.marginRight
      : Number.parseFloat(String(baseStyle.marginRight ?? ""));
  const normalizedPaddingLeft =
    typeof baseStyle.paddingLeft === "number"
      ? baseStyle.paddingLeft
      : Number.parseFloat(String(baseStyle.paddingLeft ?? ""));
  const normalizedPaddingRight =
    typeof baseStyle.paddingRight === "number"
      ? baseStyle.paddingRight
      : Number.parseFloat(String(baseStyle.paddingRight ?? ""));
  const convertPositiveCellIndentToPadding =
    (Number.isFinite(normalizedMarginLeft) &&
      (normalizedMarginLeft as number) > 0) ||
    (Number.isFinite(normalizedMarginRight) &&
      (normalizedMarginRight as number) > 0);

  return {
    ...baseStyle,
    ...(convertPositiveCellIndentToPadding
      ? {
          marginLeft:
            Number.isFinite(normalizedMarginLeft) &&
            (normalizedMarginLeft as number) > 0
              ? 0
              : baseStyle.marginLeft,
          marginRight:
            Number.isFinite(normalizedMarginRight) &&
            (normalizedMarginRight as number) > 0
              ? 0
              : baseStyle.marginRight,
          paddingLeft:
            Number.isFinite(normalizedMarginLeft) &&
            (normalizedMarginLeft as number) > 0
              ? Math.max(
                  0,
                  Math.round(
                    (Number.isFinite(normalizedPaddingLeft)
                      ? (normalizedPaddingLeft as number)
                      : 0) + (normalizedMarginLeft as number)
                  )
                )
              : baseStyle.paddingLeft,
          paddingRight:
            Number.isFinite(normalizedMarginRight) &&
            (normalizedMarginRight as number) > 0
              ? Math.max(
                  0,
                  Math.round(
                    (Number.isFinite(normalizedPaddingRight)
                      ? (normalizedPaddingRight as number)
                      : 0) + (normalizedMarginRight as number)
                  )
                )
              : baseStyle.paddingRight,
          boxSizing: "border-box",
          width: "100%",
        }
      : undefined),
    ...(suppressTopSpacing ? { marginTop: 0 } : undefined),
  };
}

function themedRunColor(
  color: string | undefined,
  documentTheme: DocxDocumentTheme
): string | undefined {
  if (documentTheme !== "dark") {
    return color;
  }

  if (!color) {
    return "#f3f4f6";
  }

  const normalized = color.trim().toLowerCase();
  if (
    normalized === "#000" ||
    normalized === "#000000" ||
    normalized === "#111111" ||
    normalized === "#111827" ||
    normalized === "black" ||
    normalized === "rgb(0,0,0)" ||
    normalized === "rgb(0, 0, 0)"
  ) {
    return "#f3f4f6";
  }

  return color;
}

function cssFontFamily(fontFamily?: string): string | undefined {
  if (!fontFamily) {
    return undefined;
  }

  const trimmed = fontFamily.trim();
  if (!trimmed) {
    return undefined;
  }

  if (trimmed.includes(",")) {
    return trimmed;
  }

  const normalized = trimmed.replace(/^['"]+|['"]+$/g, "");
  if (!normalized) {
    return undefined;
  }

  const escaped = normalized.replace(/"/g, '\\"');
  const familyToken = /\s/.test(escaped) ? `"${escaped}"` : escaped;
  const lower = normalized.toLowerCase();
  const genericFamily = /(mono|consolas|courier|menlo|code)/.test(lower)
    ? "monospace"
    : /(times|cambria|georgia|garamond|baskerville|serif)/.test(lower)
    ? "serif"
    : "sans-serif";

  return `${familyToken}, ${genericFamily}`;
}

function runStyleToCss(
  style?: TextRunNode["style"],
  documentTheme: DocxDocumentTheme = "light"
): Record<string, string | number | undefined> {
  const hasScriptVerticalAlign =
    style?.verticalAlign === "superscript" ||
    style?.verticalAlign === "subscript";
  const verticalAlign =
    style?.verticalAlign === "superscript"
      ? "super"
      : style?.verticalAlign === "subscript"
      ? "sub"
      : undefined;
  const textDecorationTokens = [
    style?.underline ? "underline" : "",
    style?.strike ? "line-through" : "",
  ].filter(Boolean);
  const textDecoration =
    textDecorationTokens.length > 0 ? textDecorationTokens.join(" ") : "none";
  const borderType = style?.runBorder?.type?.trim().toLowerCase();
  const borderStyle =
    borderType === "single"
      ? "solid"
      : borderType === "nil" || borderType === "none"
      ? undefined
      : borderType;
  const borderWidthPx =
    Number.isFinite(style?.runBorder?.sizeEighthPt) &&
    (style?.runBorder?.sizeEighthPt as number) > 0
      ? Math.max(
          1,
          Number(
            (
              ((style?.runBorder?.sizeEighthPt as number) / 8) *
              (96 / 72)
            ).toFixed(2)
          )
        )
      : borderStyle
      ? 1
      : undefined;
  const borderPaddingPt =
    Number.isFinite(style?.runBorder?.spacePt) &&
    (style?.runBorder?.spacePt as number) >= 0
      ? Math.max(0, Math.round(style?.runBorder?.spacePt as number))
      : undefined;

  return {
    fontWeight: style?.bold ? 700 : undefined,
    fontStyle: style?.italic ? "italic" : undefined,
    textDecoration,
    color: themedRunColor(style?.color, documentTheme),
    backgroundColor:
      style?.backgroundColor ?? resolveHighlightColor(style?.highlight),
    fontSize: style?.fontSizePt
      ? `${Number(
          (
            style.fontSizePt * (hasScriptVerticalAlign ? SCRIPT_FONT_SCALE : 1)
          ).toFixed(3)
        )}pt`
      : hasScriptVerticalAlign
      ? `${SCRIPT_FONT_SCALE}em`
      : undefined,
    fontFamily: cssFontFamily(style?.fontFamily),
    letterSpacing: Number.isFinite(style?.characterSpacingTwips)
      ? `${Number(
          ((style?.characterSpacingTwips as number) / 20).toFixed(3)
        )}pt`
      : undefined,
    verticalAlign,
    display: borderStyle ? "inline-block" : undefined,
    borderStyle,
    borderWidth: borderWidthPx ? `${borderWidthPx}px` : undefined,
    borderColor: borderStyle
      ? style?.runBorder?.color ?? "currentColor"
      : undefined,
    ...(borderPaddingPt !== undefined
      ? {
          paddingTop: `${borderPaddingPt}pt`,
          paddingRight: `${borderPaddingPt}pt`,
          paddingBottom: `${borderPaddingPt}pt`,
          paddingLeft: `${borderPaddingPt}pt`,
        }
      : undefined),
    boxDecorationBreak: borderStyle ? "clone" : undefined,
    whiteSpace: "pre-wrap",
  };
}

function linkStyleToCss(
  style?: TextRunNode["style"],
  documentTheme: DocxDocumentTheme = "light"
): Record<string, string | number | undefined> {
  const base = runStyleToCss(style, documentTheme);
  const resolvedTextDecoration =
    typeof base.textDecoration === "string" &&
    base.textDecoration.trim().length > 0
      ? base.textDecoration
      : "none";
  return {
    ...base,
    color: base.color ?? "inherit",
    textDecoration: resolvedTextDecoration,
  };
}

function mergeTextDecorations(
  baseDecoration: Record<string, string | number | undefined>["textDecoration"],
  decoration: string
): string {
  const tokens = new Set<string>();
  if (typeof baseDecoration === "string") {
    baseDecoration
      .split(/\s+/)
      .map((token) => token.trim())
      .filter(Boolean)
      .forEach((token) => tokens.add(token));
  }

  decoration
    .split(/\s+/)
    .map((token) => token.trim())
    .filter(Boolean)
    .forEach((token) => tokens.add(token));

  return tokens.size > 0 ? Array.from(tokens).join(" ") : "none";
}

function trackedInlineStyle(
  baseStyle: Record<string, string | number | undefined>,
  change: ParagraphTrackedInlineChange | undefined
): Record<string, string | number | undefined> {
  if (!change) {
    return baseStyle;
  }

  if (change.kind === "insertion" || change.kind === "move-to") {
    const accentColor = change.kind === "move-to" ? "#70ad47" : "#dc2626";
    return {
      ...baseStyle,
      color: accentColor,
      textDecoration: mergeTextDecorations(
        baseStyle.textDecoration,
        "underline"
      ),
    };
  }

  return baseStyle;
}

function trackedDeletedStyle(
  documentTheme: DocxDocumentTheme,
  baseRunStyle?: TextRunNode["style"] | FormFieldRunNode["style"]
): Record<string, string | number | undefined> {
  const baseStyle = runStyleToCss(baseRunStyle, documentTheme);
  return {
    ...baseStyle,
    color: documentTheme === "dark" ? "#fca5a5" : "#b91c1c",
    textDecoration: mergeTextDecorations(
      baseStyle.textDecoration,
      "line-through"
    ),
    whiteSpace: baseStyle.whiteSpace ?? "pre-wrap",
    lineHeight: "inherit",
    opacity: 0.95,
  };
}

const TABLE_OF_CONTENTS_STYLE_ID = /^toc(?:[\s_-]*\d+)?$/i;

function isTableOfContentsStyle(styleId?: string): boolean {
  if (!styleId) {
    return false;
  }
  return TABLE_OF_CONTENTS_STYLE_ID.test(styleId.trim());
}

function tableOfContentsLevel(paragraph: ParagraphNode): number | undefined {
  const candidates = [paragraph.style?.styleId, paragraph.style?.styleName];
  for (const candidate of candidates) {
    if (!candidate) {
      continue;
    }
    const match = candidate.trim().match(/^toc(?:[\s_-]*(\d+))?$/i);
    if (!match) {
      continue;
    }
    const parsedLevel = match[1] ? Number.parseInt(match[1], 10) : 1;
    if (Number.isFinite(parsedLevel) && parsedLevel > 0) {
      return Math.round(parsedLevel);
    }
    return 1;
  }
  return undefined;
}

function isTableOfContentsParagraph(paragraph: ParagraphNode): boolean {
  return (
    isTableOfContentsStyle(paragraph.style?.styleId) ||
    isTableOfContentsStyle(paragraph.style?.styleName)
  );
}

function paragraphUsesTabLeaders(paragraph: ParagraphNode): boolean {
  if (!isTableOfContentsParagraph(paragraph)) {
    return false;
  }

  const tabStops = paragraph.style?.tabStops ?? [];
  if (
    tabStops.some(
      (tabStop) => tabStop.alignment === "right" || tabStop.leader === "dot"
    )
  ) {
    return true;
  }

  return paragraph.children.some((child) => {
    if (child.type === "text") {
      return child.text.includes("\t");
    }
    if (child.type === "form-field") {
      return formFieldDisplayValue(child).includes("\t");
    }
    return false;
  });
}

function paragraphLeadingTabStop(paragraph: ParagraphNode):
  | {
      alignment?: "left" | "center" | "right" | "decimal" | "bar";
      leader?: "none" | "dot" | "hyphen" | "underscore" | "middleDot";
      positionTwips?: number;
    }
  | undefined {
  const tabStops = paragraph.style?.tabStops ?? [];
  const explicitTabStop = tabStops.find(
    (tabStop) => tabStop.alignment === "right" || tabStop.leader === "dot"
  );
  if (explicitTabStop) {
    return explicitTabStop;
  }
  if (isTableOfContentsParagraph(paragraph)) {
    return {
      alignment: "right",
      leader: "dot",
    };
  }
  return undefined;
}

function tableOfContentsLeadingLeftTabStopPx(
  paragraph: ParagraphNode
): number | undefined {
  if (!isTableOfContentsParagraph(paragraph)) {
    return undefined;
  }

  return resolveParagraphFirstLineLeftTabStopsPx(paragraph)[0];
}

function paragraphContainsTabCharacter(paragraph: ParagraphNode): boolean {
  return paragraph.children.some((child) => {
    if (child.type === "text") {
      return child.text.includes("\t");
    }
    if (child.type === "form-field") {
      return formFieldDisplayValue(child).includes("\t");
    }
    return false;
  });
}

function paragraphTabCharacterCount(paragraph: ParagraphNode): number {
  return paragraph.children.reduce((count, child) => {
    const text =
      child.type === "text"
        ? child.text
        : child.type === "form-field"
        ? formFieldDisplayValue(child)
        : "";
    if (!text) {
      return count;
    }
    const matches = text.match(/\t/g);
    return count + (matches ? matches.length : 0);
  }, 0);
}

type ParagraphAnchoredTabLayout = "none" | "center-right" | "center" | "right";

function paragraphAnchoredTabLayout(
  paragraph: ParagraphNode,
  options?: {
    withinHeaderFooter?: boolean;
  }
): ParagraphAnchoredTabLayout {
  const tabStops = paragraph.style?.tabStops ?? [];
  const hasLeft = tabStops.some((tabStop) => tabStop.alignment === "left");
  const hasCenter = tabStops.some((tabStop) => tabStop.alignment === "center");
  const hasRight = tabStops.some((tabStop) => tabStop.alignment === "right");
  const tabCount = paragraphTabCharacterCount(paragraph);
  const withinHeaderFooter = options?.withinHeaderFooter === true;

  if (hasLeft && (hasCenter || hasRight)) {
    return "none";
  }

  if (hasCenter && hasRight) {
    if (tabCount >= 2) {
      return "center-right";
    }
    if (withinHeaderFooter && tabCount >= 1) {
      return "center-right";
    }
    return "none";
  }
  if (hasCenter && tabCount === 1) {
    return "center";
  }
  if (hasRight && tabCount === 1) {
    return "right";
  }

  return "none";
}

function paragraphFirstTabStopPx(
  paragraph: ParagraphNode,
  alignment: "center" | "right"
): number | undefined {
  return (paragraph.style?.tabStops ?? [])
    .filter((tabStop) => tabStop.alignment === alignment)
    .map((tabStop) => twipsToPixels(tabStop.positionTwips))
    .filter(
      (positionPx): positionPx is number =>
        Number.isFinite(positionPx) && (positionPx as number) > 0
    )
    .sort((left, right) => left - right)[0];
}

function paragraphUsesCenterTabLayout(paragraph: ParagraphNode): boolean {
  return paragraphAnchoredTabLayout(paragraph) === "center";
}

function paragraphUsesRightTabLayout(paragraph: ParagraphNode): boolean {
  return paragraphAnchoredTabLayout(paragraph) === "right";
}

function paragraphUsesCenterRightTabLayout(paragraph: ParagraphNode): boolean {
  return paragraphAnchoredTabLayout(paragraph) === "center-right";
}

type PageFieldKind = "PAGE" | "NUMPAGES";

interface PageFieldValueToken {
  kind: PageFieldKind;
  rawText: string;
}

interface StyleRefFieldValueToken {
  target: string;
  rawText: string;
}

function decodeXmlText(text: string): string {
  if (!text) {
    return text;
  }

  const withNumericEntities = text
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => {
      const codePoint = Number.parseInt(hex, 16);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : _;
    })
    .replace(/&#([0-9]+);/g, (_, decimal: string) => {
      const codePoint = Number.parseInt(decimal, 10);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : _;
    });

  return withNumericEntities
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

interface XmlBalancedTagRange {
  start: number;
  end: number;
  tagName: string;
  openTag: string;
}

function extractBalancedTagRanges(
  xml: string,
  tagName: string
): XmlBalancedTagRange[] {
  if (!xml) {
    return [];
  }

  const escapedTagName = tagName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(`<(/?)${escapedTagName}(?=[\\s>/])[^>]*>`, "gi");
  const stack: Array<{ start: number; openTag: string }> = [];
  const ranges: XmlBalancedTagRange[] = [];

  for (const match of xml.matchAll(pattern)) {
    const fullMatch = match[0] ?? "";
    if (!fullMatch) {
      continue;
    }

    const start = match.index ?? 0;
    const isClosing = match[1] === "/";
    const isSelfClosing = !isClosing && /\/>\s*$/i.test(fullMatch);
    if (isSelfClosing) {
      ranges.push({
        start,
        end: start + fullMatch.length,
        tagName,
        openTag: fullMatch,
      });
      continue;
    }

    if (!isClosing) {
      stack.push({
        start,
        openTag: fullMatch,
      });
      continue;
    }

    const opener = stack.pop();
    if (!opener) {
      continue;
    }

    ranges.push({
      start: opener.start,
      end: start + fullMatch.length,
      tagName,
      openTag: opener.openTag,
    });
  }

  return ranges;
}

interface RevisionTagRange extends XmlBalancedTagRange {
  kind: Exclude<
    DocxTrackedChangeKind,
    "format-change" | "paragraph-format-change"
  >;
  revisionId?: string;
  author?: string;
  date?: string;
}

function trackedChangeKindFromTagName(
  tagName: string
):
  | Exclude<DocxTrackedChangeKind, "format-change" | "paragraph-format-change">
  | undefined {
  const normalized = tagName.trim().toLowerCase();
  switch (normalized) {
    case "w:ins":
      return "insertion";
    case "w:del":
      return "deletion";
    case "w:movefrom":
      return "move-from";
    case "w:moveto":
      return "move-to";
    default:
      return undefined;
  }
}

function normalizeTrackedChangeSnippet(value?: string): string | undefined {
  if (!value) {
    return undefined;
  }
  const normalized = value
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return normalized.length > 0 ? normalized : undefined;
}

function formatTrackedChangeDate(value?: string): string | undefined {
  if (!value) {
    return undefined;
  }

  const trimmed = value.trim();
  if (!trimmed) {
    return undefined;
  }

  const parsed = new Date(trimmed);
  if (Number.isNaN(parsed.getTime())) {
    return undefined;
  }

  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "2-digit",
    year: "numeric",
    timeZone: "UTC",
  }).format(parsed);
}

function stripTextBoxContentFromRunXml(runXml: string): string {
  if (!runXml.includes("w:txbxContent")) {
    return runXml;
  }
  return runXml.replace(/<w:txbxContent\b[\s\S]*?<\/w:txbxContent>/gi, "");
}

function parseTrackedRunTokens(
  runXml: string,
  includeDeletedText: boolean
): Array<{ text: string; isNote: boolean }> {
  if (!runXml) {
    return [];
  }

  const tokens: Array<{ text: string; isNote: boolean }> = [];
  const pattern =
    /<w:delText\b[^>]*>([\s\S]*?)<\/w:delText>|<(?:w|a):t\b[^>]*>([\s\S]*?)<\/(?:w|a):t>|<w:tab\b[^>]*\/?>|<w:(?:br|cr)\b[^>]*\/?>|<w:footnoteReference\b[^>]*\/?>|<w:endnoteReference\b[^>]*\/?>/gi;

  for (const match of runXml.matchAll(pattern)) {
    if (match[1] !== undefined) {
      if (!includeDeletedText) {
        continue;
      }
      const decoded = decodeXmlText(match[1] ?? "");
      tokens.push({ text: decoded, isNote: false });
      continue;
    }

    if (match[2] !== undefined) {
      const decoded = decodeXmlText(match[2] ?? "");
      tokens.push({ text: decoded, isNote: false });
      continue;
    }

    const tagXml = match[0] ?? "";
    if (/^<w:tab\b/i.test(tagXml)) {
      tokens.push({ text: "\t", isNote: false });
      continue;
    }
    if (/^<w:(?:br|cr)\b/i.test(tagXml)) {
      tokens.push({ text: "\n", isNote: false });
      continue;
    }
    if (/^<w:(?:footnoteReference|endnoteReference)\b/i.test(tagXml)) {
      tokens.push({ text: "\u2063", isNote: true });
    }
  }

  return tokens;
}

function xmlBooleanFlag(tagXml: string | undefined): boolean {
  if (!tagXml) {
    return false;
  }

  const raw = xmlAttribute(tagXml, "w:val")?.trim().toLowerCase();
  if (!raw) {
    return true;
  }

  return !(raw === "false" || raw === "0" || raw === "off" || raw === "none");
}

function xmlColorValue(tagXml: string | undefined): string | undefined {
  if (!tagXml) {
    return undefined;
  }

  const raw = xmlAttribute(tagXml, "w:val")?.trim();
  if (!raw || /^auto$/i.test(raw)) {
    return undefined;
  }

  if (/^#[0-9a-f]{6}$/i.test(raw)) {
    return raw.toLowerCase();
  }

  if (/^[0-9a-f]{6}$/i.test(raw)) {
    return `#${raw.toLowerCase()}`;
  }

  return raw;
}

function parseRunStyleFromRunXml(
  runXml: string
): TextRunNode["style"] | undefined {
  const rPrRange = extractBalancedTagRanges(runXml, "w:rPr")[0];
  const rPrXml = rPrRange ? runXml.slice(rPrRange.start, rPrRange.end) : "";
  if (!rPrXml) {
    return undefined;
  }

  const rFontsTag = rPrXml.match(/<w:rFonts\b[^>]*\/?>/i)?.[0];
  const rFontsAscii = rFontsTag
    ? xmlAttribute(rFontsTag, "w:ascii")
    : undefined;
  const rFontsHAnsi = rFontsTag
    ? xmlAttribute(rFontsTag, "w:hAnsi")
    : undefined;
  const rFontsEastAsia = rFontsTag
    ? xmlAttribute(rFontsTag, "w:eastAsia")
    : undefined;
  const rFontsCs = rFontsTag ? xmlAttribute(rFontsTag, "w:cs") : undefined;
  const fontFamily = rFontsAscii ?? rFontsHAnsi ?? rFontsEastAsia ?? rFontsCs;

  const sizeTag =
    rPrXml.match(/<w:sz\b[^>]*\/?>/i)?.[0] ??
    rPrXml.match(/<w:szCs\b[^>]*\/?>/i)?.[0];
  const sizeHalfPoints = sizeTag
    ? Number(xmlAttribute(sizeTag, "w:val"))
    : Number.NaN;
  const fontSizePt =
    Number.isFinite(sizeHalfPoints) && sizeHalfPoints > 0
      ? Number((sizeHalfPoints / 2).toFixed(2))
      : undefined;

  const bold = xmlBooleanFlag(rPrXml.match(/<w:b(?:Cs)?\b[^>]*\/?>/i)?.[0]);
  const italic = xmlBooleanFlag(rPrXml.match(/<w:i(?:Cs)?\b[^>]*\/?>/i)?.[0]);
  const underlineTag = rPrXml.match(/<w:u\b[^>]*\/?>/i)?.[0];
  const underline = xmlBooleanFlag(underlineTag);
  const strike = xmlBooleanFlag(rPrXml.match(/<w:strike\b[^>]*\/?>/i)?.[0]);
  const color = xmlColorValue(rPrXml.match(/<w:color\b[^>]*\/?>/i)?.[0]);
  const highlightTag = rPrXml.match(/<w:highlight\b[^>]*\/?>/i)?.[0];
  const highlight = highlightTag
    ? xmlAttribute(highlightTag, "w:val")
    : undefined;
  const verticalAlignTag = rPrXml.match(/<w:vertAlign\b[^>]*\/?>/i)?.[0];
  const verticalAlign = verticalAlignTag
    ? xmlAttribute(verticalAlignTag, "w:val")
    : undefined;

  const style: NonNullable<TextRunNode["style"]> = {
    fontFamily: fontFamily?.trim() || undefined,
    fontSizePt,
    bold: bold || undefined,
    italic: italic || undefined,
    underline: underline || undefined,
    strike: strike || undefined,
    color,
    highlight: highlight?.trim() || undefined,
    verticalAlign:
      verticalAlign === "superscript" || verticalAlign === "subscript"
        ? verticalAlign
        : undefined,
  };

  return Object.values(style).some((value) => value !== undefined)
    ? style
    : undefined;
}

function balancedTagXmlBlocks(xml: string, tagName: string): string[] {
  return extractBalancedTagRanges(xml, tagName).map((range) =>
    xml.slice(range.start, range.end)
  );
}

function mergeTextRunStyles(
  base?: TextRunNode["style"],
  override?: TextRunNode["style"]
): TextRunNode["style"] | undefined {
  if (!base && !override) {
    return undefined;
  }

  const merged: TextRunNode["style"] = {
    fontFamily: override?.fontFamily ?? base?.fontFamily,
    fontSizePt: override?.fontSizePt ?? base?.fontSizePt,
    bold: override?.bold ?? base?.bold,
    italic: override?.italic ?? base?.italic,
    underline: override?.underline ?? base?.underline,
    strike: override?.strike ?? base?.strike,
    color: override?.color ?? base?.color,
    highlight: override?.highlight ?? base?.highlight,
    verticalAlign: override?.verticalAlign ?? base?.verticalAlign,
  };

  return Object.values(merged).some((value) => value !== undefined)
    ? merged
    : undefined;
}

function parseParagraphAlignmentFromXml(
  paragraphPropertiesXml: string
): ParagraphAlignment | undefined {
  const jcTag = paragraphPropertiesXml.match(/<w:jc\b[^>]*\/?>/i)?.[0];
  const raw = jcTag
    ? xmlAttribute(jcTag, "w:val")?.trim().toLowerCase()
    : undefined;
  if (
    raw === "center" ||
    raw === "right" ||
    raw === "justify" ||
    raw === "left"
  ) {
    return raw;
  }
  return undefined;
}

function escapeSvgText(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function svgDataUri(svg: string): string {
  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
}

interface SyntheticTextBoxSegment {
  text: string;
  style?: TextRunNode["style"];
}

interface SyntheticTextBoxParagraph {
  align?: ParagraphAlignment;
  lineHeightPx: number;
  segments: SyntheticTextBoxSegment[];
}

interface SyntheticTextBoxFrameStyle {
  backgroundColor?: string;
  borderColor?: string;
  borderWidthPx: number;
  paddingLeftPx: number;
  paddingTopPx: number;
  paddingRightPx: number;
  paddingBottomPx: number;
}

function emuToPixels(value: number | undefined): number | undefined {
  if (!Number.isFinite(value)) {
    return undefined;
  }

  return Math.max(0, Number(((value as number) / 9525).toFixed(3)));
}

function resolveSyntheticTextBoxFieldText(
  rawText: string,
  fieldKind: PageFieldKind | undefined,
  styleRefTarget: string | undefined,
  pageNumber: number | undefined,
  totalPages: number | undefined,
  pageNumberFormat: string | undefined,
  resolveStyleRefFieldValue?: (target: string) => string | undefined
): string {
  if (fieldKind === "PAGE") {
    const resolvedFieldValue =
      Number.isFinite(pageNumber) && (pageNumber as number) > 0
        ? Math.max(1, Math.round(pageNumber as number))
        : undefined;
    if (Number.isFinite(resolvedFieldValue)) {
      const leadingWhitespace = rawText.match(/^\s*/)?.[0] ?? "";
      const trailingWhitespace = rawText.match(/\s*$/)?.[0] ?? "";
      return `${leadingWhitespace}${formatPageFieldValue(
        Math.round(resolvedFieldValue as number),
        pageNumberFormat
      )}${trailingWhitespace}`;
    }
  }

  if (fieldKind === "NUMPAGES") {
    const resolvedFieldValue =
      Number.isFinite(totalPages) && (totalPages as number) > 0
        ? Math.max(1, Math.round(totalPages as number))
        : undefined;
    if (Number.isFinite(resolvedFieldValue)) {
      const leadingWhitespace = rawText.match(/^\s*/)?.[0] ?? "";
      const trailingWhitespace = rawText.match(/\s*$/)?.[0] ?? "";
      return `${leadingWhitespace}${Math.round(
        resolvedFieldValue as number
      )}${trailingWhitespace}`;
    }
  }

  if (styleRefTarget) {
    const resolved = resolveStyleRefFieldValue?.(styleRefTarget)?.trim();
    if (resolved) {
      const leadingWhitespace = rawText.match(/^\s*/)?.[0] ?? "";
      const trailingWhitespace = rawText.match(/\s*$/)?.[0] ?? "";
      return `${leadingWhitespace}${resolved}${trailingWhitespace}`;
    }
  }

  return rawText;
}

function syntheticTextBoxParagraphsFromRunXml(
  runXml: string,
  pageNumber?: number,
  totalPages?: number,
  pageNumberFormat?: string,
  resolveStyleRefFieldValue?: (target: string) => string | undefined
): SyntheticTextBoxParagraph[] {
  const textBoxXml = balancedTagXmlBlocks(runXml, "w:txbxContent")[0];
  if (!textBoxXml) {
    return [];
  }

  const resolved: SyntheticTextBoxParagraph[] = [];
  for (const paragraphXml of balancedTagXmlBlocks(textBoxXml, "w:p")) {
    const paragraphPropertiesXml =
      balancedTagXmlBlocks(paragraphXml, "w:pPr")[0] ??
      paragraphXml.match(/<w:pPr\b[^>]*\/?>/i)?.[0] ??
      "";
    const paragraphRunPropertiesXml =
      balancedTagXmlBlocks(paragraphPropertiesXml, "w:rPr")[0] ??
      paragraphPropertiesXml.match(/<w:rPr\b[^>]*\/?>/i)?.[0] ??
      "";
    const paragraphStyle = parseRunStyleFromRunXml(
      `<w:r>${paragraphRunPropertiesXml}</w:r>`
    );
    const spacingTag =
      paragraphPropertiesXml.match(/<w:spacing\b[^>]*\/?>/i)?.[0] ?? "";
    const lineRaw = spacingTag
      ? Number(xmlAttribute(spacingTag, "w:line"))
      : Number.NaN;
    const lineRule = spacingTag
      ? xmlAttribute(spacingTag, "w:lineRule")?.trim().toLowerCase()
      : undefined;
    const lineHeightPx =
      Number.isFinite(lineRaw) &&
      (lineRaw as number) > 0 &&
      lineRule === "exact"
        ? Math.max(1, twipsToPixels(lineRaw as number) ?? 1)
        : undefined;

    const segments: SyntheticTextBoxSegment[] = [];
    const fieldStack: Array<{
      pageFieldKind?: PageFieldKind;
      styleRefTarget?: string;
      inResult: boolean;
      instructionStyle?: TextRunNode["style"];
    }> = [];

    const appendSegment = (
      text: string,
      style?: TextRunNode["style"]
    ): void => {
      if (!text) {
        return;
      }
      const previous = segments[segments.length - 1];
      if (
        previous &&
        JSON.stringify(previous.style ?? {}) === JSON.stringify(style ?? {})
      ) {
        previous.text += text;
        return;
      }
      segments.push({ text, style });
    };

    for (const runBlockXml of balancedTagXmlBlocks(paragraphXml, "w:r")) {
      const runStyle = mergeTextRunStyles(
        paragraphStyle,
        parseRunStyleFromRunXml(runBlockXml)
      );

      const beginCount =
        runBlockXml.match(/<w:fldChar\b[^>]*\bw:fldCharType="begin"[^>]*\/?>/gi)
          ?.length ?? 0;
      for (let index = 0; index < beginCount; index += 1) {
        fieldStack.push({ inResult: false });
      }

      for (const instructionMatch of runBlockXml.matchAll(
        /<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/gi
      )) {
        const pageFieldKind = instructionTextToPageFieldKind(
          instructionMatch[1] ?? ""
        );
        const styleRefTarget = instructionTextToStyleRefTarget(
          instructionMatch[1] ?? ""
        );
        if ((!pageFieldKind && !styleRefTarget) || fieldStack.length === 0) {
          continue;
        }

        for (
          let stackIndex = fieldStack.length - 1;
          stackIndex >= 0;
          stackIndex -= 1
        ) {
          if (
            !fieldStack[stackIndex].pageFieldKind &&
            !fieldStack[stackIndex].styleRefTarget
          ) {
            fieldStack[stackIndex].pageFieldKind = pageFieldKind;
            fieldStack[stackIndex].styleRefTarget = styleRefTarget;
            fieldStack[stackIndex].instructionStyle = mergeTextRunStyles(
              fieldStack[stackIndex].instructionStyle,
              runStyle
            );
            break;
          }
        }
      }

      const separateCount =
        runBlockXml.match(
          /<w:fldChar\b[^>]*\bw:fldCharType="separate"[^>]*\/?>/gi
        )?.length ?? 0;
      for (let index = 0; index < separateCount; index += 1) {
        for (
          let stackIndex = fieldStack.length - 1;
          stackIndex >= 0;
          stackIndex -= 1
        ) {
          if (!fieldStack[stackIndex].inResult) {
            fieldStack[stackIndex].inResult = true;
            break;
          }
        }
      }

      const activeField = (() => {
        for (
          let stackIndex = fieldStack.length - 1;
          stackIndex >= 0;
          stackIndex -= 1
        ) {
          const stackEntry = fieldStack[stackIndex];
          if (
            stackEntry.inResult &&
            (stackEntry.pageFieldKind || stackEntry.styleRefTarget)
          ) {
            return stackEntry;
          }
        }
        return undefined;
      })();

      for (const textMatch of runBlockXml.matchAll(
        /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi
      )) {
        const rawText = decodeXmlText(textMatch[1] ?? "");
        const resolvedText = activeField
          ? resolveSyntheticTextBoxFieldText(
              rawText,
              activeField.pageFieldKind,
              activeField.styleRefTarget,
              pageNumber,
              totalPages,
              pageNumberFormat,
              resolveStyleRefFieldValue
            )
          : rawText;
        appendSegment(
          resolvedText,
          activeField?.instructionStyle
            ? mergeTextRunStyles(activeField.instructionStyle, runStyle)
            : runStyle
        );
      }

      const endCount =
        runBlockXml.match(/<w:fldChar\b[^>]*\bw:fldCharType="end"[^>]*\/?>/gi)
          ?.length ?? 0;
      for (let index = 0; index < endCount; index += 1) {
        fieldStack.pop();
      }
    }

    if (segments.length === 0) {
      continue;
    }

    const effectiveFontSizePt =
      segments.find((segment) => segment.style?.fontSizePt)?.style
        ?.fontSizePt ?? 12;
    resolved.push({
      align: parseParagraphAlignmentFromXml(paragraphPropertiesXml),
      lineHeightPx:
        lineHeightPx ??
        Math.max(
          14,
          Math.round((((effectiveFontSizePt ?? 12) * 96) / 72) * 1.24)
        ),
      segments,
    });
  }

  return resolved;
}

function syntheticTextBoxFrameStyleFromRunXml(
  runXml: string
): SyntheticTextBoxFrameStyle {
  const bodyPrTag =
    balancedTagXmlBlocks(runXml, "wps:bodyPr")[0] ??
    runXml.match(/<wps:bodyPr\b[^>]*\/?>/i)?.[0] ??
    "";
  const lineTag =
    balancedTagXmlBlocks(runXml, "a:ln")[0] ??
    runXml.match(/<a:ln\b[\s\S]*?<\/a:ln>/i)?.[0] ??
    runXml.match(/<a:ln\b[^>]*\/?>/i)?.[0] ??
    "";
  const shapePropsXml =
    balancedTagXmlBlocks(runXml, "wps:spPr")[0] ??
    runXml.match(/<wps:spPr\b[\s\S]*?<\/wps:spPr>/i)?.[0] ??
    "";

  const hasNoFill =
    /<a:noFill\b/i.test(shapePropsXml) || /<wps:noFill\b/i.test(shapePropsXml);
  const lineHasNoFill = /<a:noFill\b/i.test(lineTag);
  const fillColor =
    shapePropsXml.match(
      /<a:solidFill>\s*<a:srgbClr\b[^>]*val="([^"]+)"/i
    )?.[1] ??
    shapePropsXml.match(
      /<a:solidFill>\s*<a:schemeClr\b[^>]*val="([^"]+)"/i
    )?.[1];
  const lineColor =
    lineTag.match(/<a:solidFill>\s*<a:srgbClr\b[^>]*val="([^"]+)"/i)?.[1] ??
    lineTag.match(/<a:solidFill>\s*<a:schemeClr\b[^>]*val="([^"]+)"/i)?.[1];
  const lineWidthEmu = Number(xmlAttribute(lineTag, "w"));
  const resolvedBorderWidthPx =
    !lineTag || lineHasNoFill
      ? 0
      : Number.isFinite(lineWidthEmu) && (lineWidthEmu as number) > 0
      ? Math.max(1, Math.round((lineWidthEmu as number) / 9525))
      : 0;

  return {
    backgroundColor: !hasNoFill && fillColor ? `#${fillColor}` : undefined,
    borderColor:
      resolvedBorderWidthPx > 0 && lineColor ? `#${lineColor}` : undefined,
    borderWidthPx: resolvedBorderWidthPx,
    paddingLeftPx: emuToPixels(Number(xmlAttribute(bodyPrTag, "lIns"))) ?? 6,
    paddingTopPx: emuToPixels(Number(xmlAttribute(bodyPrTag, "tIns"))) ?? 3,
    paddingRightPx: emuToPixels(Number(xmlAttribute(bodyPrTag, "rIns"))) ?? 6,
    paddingBottomPx: emuToPixels(Number(xmlAttribute(bodyPrTag, "bIns"))) ?? 3,
  };
}

function syntheticTextBoxTextValue(image: ImageRunNode): string | undefined {
  const explicit = image.textBoxText;
  if (typeof explicit === "string") {
    return explicit;
  }

  const paragraphs = image.sourceXml
    ? syntheticTextBoxParagraphsFromRunXml(image.sourceXml)
    : [];
  if (paragraphs.length === 0) {
    return undefined;
  }

  return paragraphs
    .map((paragraph) =>
      paragraph.segments.map((segment) => segment.text).join("")
    )
    .join("\n");
}

function resolveSyntheticTextBoxParagraphs(
  image: ImageRunNode,
  pageNumber?: number,
  totalPages?: number,
  pageNumberFormat?: string,
  resolveStyleRefFieldValue?: (target: string) => string | undefined
): SyntheticTextBoxParagraph[] {
  const baseParagraphs = image.sourceXml
    ? syntheticTextBoxParagraphsFromRunXml(
        image.sourceXml,
        pageNumber,
        totalPages,
        pageNumberFormat,
        resolveStyleRefFieldValue
      )
    : [];
  const explicitText = image.textBoxText;
  if (typeof explicitText !== "string") {
    return baseParagraphs;
  }

  const lines = explicitText.split(/\r?\n/);
  if (lines.length === 0) {
    return baseParagraphs;
  }

  return lines.map((line, index) => {
    const baseParagraph = baseParagraphs[index] ?? baseParagraphs[0];
    const baseStyle = baseParagraph?.segments[0]?.style;
    return {
      align: baseParagraph?.align,
      lineHeightPx: baseParagraph?.lineHeightPx ?? 18,
      segments: [{ text: line, style: baseStyle }],
    };
  });
}

function syntheticTextBoxForegroundZIndex(image: ImageRunNode): number {
  const normalizedZIndex = Number.isFinite(image.floating?.zIndex)
    ? Math.max(
        1,
        Math.min(
          65535,
          Math.round(
            (image.floating?.zIndex as number) / WORD_IMAGE_Z_INDEX_STEP
          )
        )
      )
    : 4;

  return Math.max(4, normalizedZIndex);
}

function syntheticTextBoxSvg(
  image: ImageRunNode,
  pageNumber?: number,
  totalPages?: number,
  pageNumberFormat?: string,
  resolveStyleRefFieldValue?: (target: string) => string | undefined
): string | undefined {
  if (!image.syntheticTextBox || !image.sourceXml) {
    return undefined;
  }

  // Grouped drawings that already import as a combined SVG image + textbox
  // should keep that original synthetic SVG. Rebuilding them here from only
  // the textbox XML drops the picture layer and regresses letterhead/logo art.
  if (syntheticTextBoxContainsPictureLayer(image)) {
    return undefined;
  }

  const paragraphs = resolveSyntheticTextBoxParagraphs(
    image,
    pageNumber,
    totalPages,
    pageNumberFormat,
    resolveStyleRefFieldValue
  );
  if (paragraphs.length === 0) {
    return undefined;
  }

  const safeWidth = Math.max(8, Math.round(image.widthPx ?? 320));
  const estimatedHeight = paragraphs.reduce(
    (sum, paragraph) => sum + paragraph.lineHeightPx,
    6
  );
  const safeHeight = Math.max(8, Math.round(image.heightPx ?? estimatedHeight));
  let cursorY = 0;
  const textBlocks: string[] = [];

  for (const paragraph of paragraphs) {
    const nextCursorY = cursorY + paragraph.lineHeightPx;
    const clampedBaselineY = Math.min(nextCursorY, Math.max(1, safeHeight - 1));
    if (nextCursorY > safeHeight + 1 && textBlocks.length > 0) {
      break;
    }
    cursorY = nextCursorY;

    const fullText = paragraph.segments.map((segment) => segment.text).join("");
    const baseStyle = paragraph.segments.find(
      (segment) => segment.style
    )?.style;
    const baseFontSizePx = Math.max(
      10,
      Math.round(((baseStyle?.fontSizePt ?? 12) * 96) / 72)
    );
    const estimatedTextWidth = Math.max(
      1,
      estimateTextAdvanceWidthPx(fullText, baseStyle)
    );
    const fitScale =
      estimatedTextWidth > safeWidth && estimatedTextWidth > 0
        ? safeWidth / estimatedTextWidth
        : 1;
    const textAlign = paragraph.align ?? "left";
    const anchor =
      textAlign === "center"
        ? "middle"
        : textAlign === "right"
        ? "end"
        : "start";
    const x =
      textAlign === "center"
        ? Math.round(safeWidth / 2)
        : textAlign === "right"
        ? safeWidth
        : 0;
    let segmentOffsetPx = 0;
    const tspans = paragraph.segments
      .map((segment, segmentIndex) => {
        const segmentStyle = mergeTextRunStyles(baseStyle, segment.style);
        const segmentFontSizePx = Math.max(
          10,
          Math.round((((segmentStyle?.fontSizePt ?? 12) * 96) / 72) * fitScale)
        );
        const textDecoration = [
          segmentStyle?.underline ? "underline" : "",
          segmentStyle?.strike ? "line-through" : "",
        ]
          .filter(Boolean)
          .join(" ");
        const attrs = [
          `font-size="${segmentFontSizePx}"`,
          segmentStyle?.fontFamily
            ? `font-family="${escapeSvgText(segmentStyle.fontFamily)}"`
            : "",
          segmentStyle?.bold ? 'font-weight="700"' : "",
          segmentStyle?.italic ? 'font-style="italic"' : "",
          segmentStyle?.color
            ? `fill="${segmentStyle.color}"`
            : 'fill="#000000"',
          textDecoration ? `text-decoration="${textDecoration}"` : "",
        ]
          .filter(Boolean)
          .join(" ");
        const xAttr =
          textAlign === "left" && segmentIndex > 0
            ? ` dx="${Math.round(segmentOffsetPx)}"`
            : segmentIndex > 0
            ? ' dx="0"'
            : "";
        segmentOffsetPx = 0;
        return `<tspan${xAttr} ${attrs}>${escapeSvgText(segment.text)}</tspan>`;
      })
      .join("");

    if (textAlign === "left") {
      let runningWidth = 0;
      const leftTspans = paragraph.segments
        .map((segment) => {
          const segmentStyle = mergeTextRunStyles(baseStyle, segment.style);
          const segmentFontSizePx = Math.max(
            10,
            Math.round(
              (((segmentStyle?.fontSizePt ?? 12) * 96) / 72) * fitScale
            )
          );
          const textDecoration = [
            segmentStyle?.underline ? "underline" : "",
            segmentStyle?.strike ? "line-through" : "",
          ]
            .filter(Boolean)
            .join(" ");
          const attrs = [
            `x="${Math.round(runningWidth)}"`,
            `font-size="${segmentFontSizePx}"`,
            segmentStyle?.fontFamily
              ? `font-family="${escapeSvgText(segmentStyle.fontFamily)}"`
              : "",
            segmentStyle?.bold ? 'font-weight="700"' : "",
            segmentStyle?.italic ? 'font-style="italic"' : "",
            segmentStyle?.color
              ? `fill="${segmentStyle.color}"`
              : 'fill="#000000"',
            textDecoration ? `text-decoration="${textDecoration}"` : "",
          ]
            .filter(Boolean)
            .join(" ");
          const block = `<tspan ${attrs}>${escapeSvgText(
            segment.text
          )}</tspan>`;
          runningWidth +=
            estimateTextAdvanceWidthPx(segment.text, segmentStyle) * fitScale;
          return block;
        })
        .join("");
      textBlocks.push(
        `<text y="${Math.round(
          clampedBaselineY
        )}" text-anchor="${anchor}">${leftTspans}</text>`
      );
      continue;
    }

    textBlocks.push(
      `<text x="${x}" y="${Math.round(
        clampedBaselineY
      )}" text-anchor="${anchor}">${tspans}</text>`
    );
  }

  return svgDataUri(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${safeWidth}" height="${safeHeight}" viewBox="0 0 ${safeWidth} ${safeHeight}">${textBlocks.join(
      ""
    )}</svg>`
  );
}

function syntheticTextBoxContainsPictureLayer(image: ImageRunNode): boolean {
  if (!image.syntheticTextBox || !image.sourceXml) {
    return false;
  }

  if (/<pic:pic\b|<a:blip\b/i.test(image.sourceXml)) {
    return true;
  }

  if (!/<wpg:wgp\b/i.test(image.sourceXml)) {
    return false;
  }

  const groupedShapes = balancedTagXmlBlocks(image.sourceXml, "wps:wsp");
  return groupedShapes.some((shapeXml) => {
    const shapePropertiesXml =
      balancedTagXmlBlocks(shapeXml, "wps:spPr")[0] ?? "";
    if (!shapePropertiesXml || /<w:txbxContent\b/i.test(shapeXml)) {
      return false;
    }

    const lineXml = balancedTagXmlBlocks(shapePropertiesXml, "a:ln")[0] ?? "";
    return (
      /<a:solidFill\b|<a:gradFill\b|<a:blipFill\b/i.test(shapePropertiesXml) ||
      (Boolean(lineXml) && !/<a:noFill\b/i.test(lineXml))
    );
  });
}

function parseDrawingImageTransformFromSourceXml(sourceXml?: string):
  | {
      rotationDegrees?: number;
      flipH?: boolean;
      flipV?: boolean;
    }
  | undefined {
  if (!sourceXml) {
    return undefined;
  }

  const transformXml =
    balancedTagXmlBlocks(sourceXml, "a:xfrm")[0] ??
    sourceXml.match(/<a:xfrm\b[^>]*\/?>/i)?.[0] ??
    "";
  if (!transformXml) {
    return undefined;
  }

  const rotationRaw = Number(xmlAttribute(transformXml, "rot"));
  const rotationDegrees = Number.isFinite(rotationRaw)
    ? Number(((rotationRaw as number) / 60000).toFixed(3))
    : undefined;
  const flipHRaw = xmlAttribute(transformXml, "flipH")?.trim().toLowerCase();
  const flipVRaw = xmlAttribute(transformXml, "flipV")?.trim().toLowerCase();
  const flipH = flipHRaw === "1" || flipHRaw === "true";
  const flipV = flipVRaw === "1" || flipVRaw === "true";

  if (!Number.isFinite(rotationDegrees) && !flipH && !flipV) {
    return undefined;
  }

  return {
    rotationDegrees,
    flipH: flipH || undefined,
    flipV: flipV || undefined,
  };
}

function joinCssTransforms(
  ...parts: Array<string | undefined>
): string | undefined {
  const resolved = parts.filter((part): part is string =>
    Boolean(part && part.trim().length > 0)
  );
  return resolved.length > 0 ? resolved.join(" ") : undefined;
}

function resolveImageRenderTransformStyle(
  image: ImageRunNode,
  options?: {
    frameWidthPx?: number;
    frameHeightPx?: number;
    fillFrame?: boolean;
    baseTransform?: string;
  }
): Record<string, string | number | undefined> {
  const transform = parseDrawingImageTransformFromSourceXml(image.sourceXml);
  const rotationDegrees = transform?.rotationDegrees;
  const frameWidthPx = Number.isFinite(options?.frameWidthPx)
    ? Math.max(1, Math.round(options?.frameWidthPx as number))
    : undefined;
  const frameHeightPx = Number.isFinite(options?.frameHeightPx)
    ? Math.max(1, Math.round(options?.frameHeightPx as number))
    : undefined;
  const fillFrame = options?.fillFrame === true;
  const normalizedQuarterTurn =
    Number.isFinite(rotationDegrees) &&
    Math.abs((Math.abs(rotationDegrees as number) % 180) - 90) < 0.5
      ? (((Math.round((rotationDegrees as number) / 90) % 4) + 4) % 4) * 90
      : undefined;
  const flipScaleTransform =
    transform?.flipH || transform?.flipV
      ? `scale(${transform?.flipH ? -1 : 1}, ${transform?.flipV ? -1 : 1})`
      : undefined;

  if (
    fillFrame &&
    Number.isFinite(frameWidthPx) &&
    Number.isFinite(frameHeightPx) &&
    Number.isFinite(normalizedQuarterTurn)
  ) {
    const safeFrameWidthPx = frameWidthPx as number;
    const safeFrameHeightPx = frameHeightPx as number;
    const rotationTransform =
      normalizedQuarterTurn === 90
        ? `translate(${safeFrameWidthPx}px, 0px) rotate(90deg)`
        : normalizedQuarterTurn === 180
        ? `translate(${safeFrameWidthPx}px, ${safeFrameHeightPx}px) rotate(180deg)`
        : normalizedQuarterTurn === 270
        ? `translate(0px, ${safeFrameHeightPx}px) rotate(-90deg)`
        : undefined;

    return {
      width: `${safeFrameHeightPx}px`,
      height: `${safeFrameWidthPx}px`,
      maxWidth: "none",
      transformOrigin: "top left",
      transform: joinCssTransforms(
        options?.baseTransform,
        rotationTransform,
        flipScaleTransform
      ),
    };
  }

  const rotationTransform =
    Number.isFinite(rotationDegrees) &&
    Math.abs(rotationDegrees as number) >= 0.01
      ? `rotate(${rotationDegrees}deg)`
      : undefined;
  const transformValue = joinCssTransforms(
    options?.baseTransform,
    rotationTransform,
    flipScaleTransform
  );
  if (!transformValue) {
    return {};
  }

  return {
    transformOrigin: "center center",
    transform: transformValue,
  };
}

function syntheticTextBoxActsAsTopAndBottomMasthead(
  image: ImageRunNode
): boolean {
  if (!image.syntheticTextBox || !image.sourceXml) {
    return false;
  }

  return (
    /<wpg:wgp\b/i.test(image.sourceXml) && /<wps:wsp\b/i.test(image.sourceXml)
  );
}

function summarizeChangeFeatures(
  prefix: string,
  features: string[],
  fallback: string
): string {
  if (features.length === 0) {
    return fallback;
  }

  const unique = Array.from(new Set(features));
  return `${prefix}: ${unique.join(", ")}`;
}

function summarizeRunFormattingChange(changeXml: string): string {
  const features: string[] = [];
  if (/<w:rFonts\b/i.test(changeXml)) {
    features.push("font");
  }
  if (/<w:sz\b/i.test(changeXml)) {
    features.push("size");
  }
  if (/<w:color\b/i.test(changeXml)) {
    features.push("color");
  }
  if (/<w:highlight\b/i.test(changeXml)) {
    features.push("highlight");
  }
  if (/<w:b(?:Cs)?\b/i.test(changeXml)) {
    features.push("bold");
  }
  if (/<w:i(?:Cs)?\b/i.test(changeXml)) {
    features.push("italic");
  }
  if (/<w:u\b/i.test(changeXml)) {
    features.push("underline");
  }
  if (/<w:strike\b/i.test(changeXml)) {
    features.push("strikethrough");
  }
  if (/<w:vertAlign\b/i.test(changeXml)) {
    features.push("baseline");
  }
  return summarizeChangeFeatures("Run formatting", features, "Run formatting");
}

function summarizeParagraphFormattingChange(changeXml: string): string {
  const features: string[] = [];
  if (/<w:ind\b/i.test(changeXml)) {
    features.push("margins/indent");
  }
  if (/<w:spacing\b/i.test(changeXml)) {
    features.push("line spacing");
  }
  if (/<w:jc\b/i.test(changeXml)) {
    features.push("alignment");
  }
  if (/<w:tabs\b/i.test(changeXml)) {
    features.push("tabs");
  }
  if (/<w:numPr\b/i.test(changeXml)) {
    features.push("numbering");
  }
  if (/<w:pBdr\b/i.test(changeXml)) {
    features.push("borders");
  }
  if (/<w:shd\b/i.test(changeXml)) {
    features.push("shading");
  }
  if (/<w:rPr\b/i.test(changeXml)) {
    features.push("text style");
  }
  return summarizeChangeFeatures(
    "Paragraph formatting",
    features,
    "Paragraph formatting"
  );
}

function summarizeTableFormattingChange(
  scope: "table" | "row" | "cell",
  changeXml: string
): string {
  const features: string[] = [];
  if (/<w:tblW\b|<w:tcW\b|<w:gridSpan\b/i.test(changeXml)) {
    features.push("width");
  }
  if (/<w:tblLayout\b/i.test(changeXml)) {
    features.push("layout");
  }
  if (/<w:tblInd\b|<w:ind\b/i.test(changeXml)) {
    features.push("indent");
  }
  if (/<w:tblCellMar\b|<w:tcMar\b/i.test(changeXml)) {
    features.push("margins");
  }
  if (/<w:(?:tblBorders|tcBorders|trBorders|pBdr)\b/i.test(changeXml)) {
    features.push("borders");
  }
  if (/<w:trHeight\b/i.test(changeXml)) {
    features.push("row height");
  }
  if (/<w:vAlign\b/i.test(changeXml)) {
    features.push("vertical align");
  }
  if (/<w:jc\b/i.test(changeXml)) {
    features.push("alignment");
  }
  if (/<w:shd\b/i.test(changeXml)) {
    features.push("shading");
  }

  const prefix =
    scope === "table"
      ? "Table formatting"
      : scope === "row"
      ? "Row formatting"
      : "Cell formatting";
  return summarizeChangeFeatures(prefix, features, prefix);
}

function resolveParagraphTrackedMarkup(
  paragraph: ParagraphNode
): ParagraphTrackedMarkup | undefined {
  const sourceXml = paragraph.sourceXml ?? "";
  if (!sourceXml) {
    return undefined;
  }

  const cached = paragraphTrackedMarkupBySourceXml.get(sourceXml);
  if (cached !== undefined) {
    return cached ?? undefined;
  }

  const revisionRanges: RevisionTagRange[] = [];
  for (const tagName of ["w:ins", "w:del", "w:moveFrom", "w:moveTo"] as const) {
    const kind = trackedChangeKindFromTagName(tagName);
    if (!kind) {
      continue;
    }

    extractBalancedTagRanges(sourceXml, tagName).forEach((range) => {
      revisionRanges.push({
        ...range,
        kind,
        revisionId: xmlAttribute(range.openTag, "w:id"),
        author: decodeXmlText(xmlAttribute(range.openTag, "w:author") ?? ""),
        date: xmlAttribute(range.openTag, "w:date"),
      });
    });
  }

  const inlineChangeByVisibleChildIndex: Array<
    ParagraphTrackedInlineChange | undefined
  > = [];
  const deletedSegmentsByVisibleChildIndex = new Map<
    number,
    ParagraphTrackedDeletionSegment[]
  >();
  const changes: ParagraphTrackedInlineChange[] = [];
  const changeByKey = new Map<string, ParagraphTrackedInlineChange>();
  let anonymousChangeCounter = 0;
  let visibleChildIndex = 0;

  const getOrCreateChange = (
    kind: DocxTrackedChangeKind,
    revisionId: string | undefined,
    author: string | undefined,
    date: string | undefined,
    text: string | undefined
  ): ParagraphTrackedInlineChange => {
    const normalizedAuthor = author?.trim() || undefined;
    const normalizedDate = date?.trim() || undefined;
    const normalizedText = normalizeTrackedChangeSnippet(text);
    const normalizedRevisionId = revisionId?.trim() || undefined;
    // Word groups tracked items primarily by revision id. Keep a single card per
    // revision id/kind to avoid over-fragmenting changes into many tiny entries.
    const key = normalizedRevisionId
      ? `${kind}:id:${normalizedRevisionId}`
      : `${kind}:anon:${normalizedAuthor ?? ""}:${normalizedDate ?? ""}:${
          normalizedText ?? ""
        }`;
    const existing = changeByKey.get(key);
    if (existing) {
      if (!existing.text && normalizedText) {
        existing.text = normalizedText;
      } else if (
        existing.text &&
        normalizedText &&
        normalizedText.length > existing.text.length
      ) {
        existing.text = normalizedText;
      }
      if (!existing.author && normalizedAuthor) {
        existing.author = normalizedAuthor;
      }
      if (!existing.date && normalizedDate) {
        existing.date = normalizedDate;
      }
      return existing;
    }

    const stableId = normalizedRevisionId
      ? `${kind}-${normalizedRevisionId}`
      : `${kind}-inline-${anonymousChangeCounter}`;
    anonymousChangeCounter += 1;
    const next: ParagraphTrackedInlineChange = {
      id: stableId,
      kind,
      author: normalizedAuthor,
      date: normalizedDate,
      text: normalizedText,
    };
    changeByKey.set(key, next);
    changes.push(next);
    return next;
  };

  const runPattern = /<w:r\b[\s\S]*?<\/w:r>/gi;
  for (const runMatch of sourceXml.matchAll(runPattern)) {
    const runXml = runMatch[0] ?? "";
    if (!runXml) {
      continue;
    }

    const runStart = runMatch.index ?? 0;
    const runEnd = runStart + runXml.length;
    const enclosingRevision = revisionRanges
      .filter(
        (revisionRange) =>
          runStart >= revisionRange.start && runEnd <= revisionRange.end
      )
      .sort(
        (left, right) => left.end - left.start - (right.end - right.start)
      )[0];

    const revisionKind = enclosingRevision?.kind;
    const revisionId = enclosingRevision?.revisionId;
    const revisionAuthor = enclosingRevision?.author;
    const revisionDate = enclosingRevision?.date;
    const isDeletionLike =
      revisionKind === "deletion" || revisionKind === "move-from";
    const contentRunXml = stripTextBoxContentFromRunXml(runXml);
    const trackedRunStyle = parseRunStyleFromRunXml(contentRunXml);
    const visibleTokens = parseTrackedRunTokens(contentRunXml, false);
    const deletedTokens = parseTrackedRunTokens(contentRunXml, true);
    const hasImage = /<w:(?:drawing|pict)\b/i.test(runXml);
    const visibleChildCount =
      visibleTokens.filter((token) => token.text.length > 0 || token.isNote)
        .length + (hasImage ? 1 : 0);
    const visibleText = normalizeTrackedChangeSnippet(
      visibleTokens.map((token) => token.text).join("")
    );
    const deletedText = normalizeTrackedChangeSnippet(
      deletedTokens.map((token) => token.text).join("")
    );

    if (isDeletionLike) {
      const deletionSnippet = deletedText ?? (hasImage ? "[image]" : undefined);
      if (!deletionSnippet) {
        continue;
      }
      const change = getOrCreateChange(
        revisionKind,
        revisionId,
        revisionAuthor,
        revisionDate,
        deletionSnippet
      );
      if (deletedText) {
        const segments =
          deletedSegmentsByVisibleChildIndex.get(visibleChildIndex) ?? [];
        segments.push({
          text: deletedText,
          change,
          style: trackedRunStyle,
        });
        deletedSegmentsByVisibleChildIndex.set(visibleChildIndex, segments);
      }
      continue;
    }

    const rPrChangeRanges = extractBalancedTagRanges(runXml, "w:rPrChange");
    rPrChangeRanges.forEach((rPrChangeRange) => {
      const rPrChangeTag = rPrChangeRange.openTag ?? "";
      const rPrChangeXml = runXml.slice(
        rPrChangeRange.start,
        rPrChangeRange.end
      );
      const formatAuthor =
        decodeXmlText(xmlAttribute(rPrChangeTag, "w:author") ?? "") ||
        revisionAuthor;
      const formatDate = xmlAttribute(rPrChangeTag, "w:date") ?? revisionDate;
      const formatId = xmlAttribute(rPrChangeTag, "w:id") ?? revisionId;
      const formatSnippet = summarizeRunFormattingChange(rPrChangeXml);
      getOrCreateChange(
        "format-change",
        formatId,
        formatAuthor,
        formatDate,
        formatSnippet
      );
    });

    if (revisionKind === "insertion" || revisionKind === "move-to") {
      const insertionSnippet =
        visibleText ?? (hasImage ? "[image]" : undefined);
      if (!insertionSnippet) {
        visibleChildIndex += visibleChildCount;
        continue;
      }
      const inlineChange = getOrCreateChange(
        revisionKind,
        revisionId,
        revisionAuthor,
        revisionDate,
        insertionSnippet
      );
      for (let index = 0; index < visibleChildCount; index += 1) {
        inlineChangeByVisibleChildIndex[visibleChildIndex + index] =
          inlineChange;
      }
    }

    visibleChildIndex += visibleChildCount;
  }

  const pPrChangeRanges = extractBalancedTagRanges(sourceXml, "w:pPrChange");
  pPrChangeRanges.forEach((pPrChangeRange) => {
    const pPrChangeTag = pPrChangeRange.openTag ?? "";
    const pPrChangeXml = sourceXml.slice(
      pPrChangeRange.start,
      pPrChangeRange.end
    );
    getOrCreateChange(
      "paragraph-format-change",
      xmlAttribute(pPrChangeTag, "w:id"),
      decodeXmlText(xmlAttribute(pPrChangeTag, "w:author") ?? ""),
      xmlAttribute(pPrChangeTag, "w:date"),
      summarizeParagraphFormattingChange(pPrChangeXml)
    );
  });

  // Some DOCX revisions wrap non-run content. Ensure every revision range is
  // represented even if the run-based pass above does not emit it.
  revisionRanges.forEach((revisionRange) => {
    const revisionXml = sourceXml.slice(revisionRange.start, revisionRange.end);
    const includeDeletedText =
      revisionRange.kind === "deletion" || revisionRange.kind === "move-from";
    const revisionTokens = parseTrackedRunTokens(
      stripTextBoxContentFromRunXml(revisionXml),
      includeDeletedText
    );
    const revisionText = normalizeTrackedChangeSnippet(
      revisionTokens.map((token) => token.text).join("")
    );
    const hasImage = /<w:(?:drawing|pict)\b/i.test(revisionXml);
    const revisionSnippet = revisionText ?? (hasImage ? "[image]" : undefined);
    if (!revisionSnippet) {
      return;
    }
    getOrCreateChange(
      revisionRange.kind,
      revisionRange.revisionId,
      revisionRange.author,
      revisionRange.date,
      revisionSnippet
    );
  });

  const hasInlineChanges = inlineChangeByVisibleChildIndex.some(Boolean);
  const hasDeletedSegments = deletedSegmentsByVisibleChildIndex.size > 0;
  if (!hasInlineChanges && !hasDeletedSegments && changes.length === 0) {
    setCacheEntry(paragraphTrackedMarkupBySourceXml, sourceXml, null);
    return undefined;
  }

  const resolved: ParagraphTrackedMarkup = {
    inlineChangeByVisibleChildIndex,
    deletedSegmentsByVisibleChildIndex,
    changes,
  };
  setCacheEntry(paragraphTrackedMarkupBySourceXml, sourceXml, resolved);
  return resolved;
}

/**
 * Maps comment ranges (`commentRangeStart`/`commentRangeEnd`) to the
 * paragraph's visible child indexes using the same run accounting as
 * `resolveParagraphTrackedMarkup`, so the run renderer can highlight
 * commented content at the matching child cursor.
 */
function resolveParagraphCommentMarkup(
  paragraph: ParagraphNode
): ParagraphCommentMarkup | undefined {
  const sourceXml = paragraph.sourceXml ?? "";
  if (!sourceXml || !/commentRange|commentReference/i.test(sourceXml)) {
    return undefined;
  }

  const cached = paragraphCommentMarkupBySourceXml.get(sourceXml);
  if (cached !== undefined) {
    return cached ?? undefined;
  }

  const rangeStartById = new Map<number, number>();
  const rangeEndById = new Map<number, number>();
  for (const match of sourceXml.matchAll(
    /<w:commentRangeStart\b[^>]*w:id="(-?\d+)"[^>]*\/?>/gi
  )) {
    const commentId = Number.parseInt(match[1] ?? "", 10);
    if (Number.isFinite(commentId) && match.index !== undefined) {
      rangeStartById.set(commentId, match.index + match[0].length);
    }
  }
  for (const match of sourceXml.matchAll(
    /<w:commentRangeEnd\b[^>]*w:id="(-?\d+)"[^>]*\/?>/gi
  )) {
    const commentId = Number.parseInt(match[1] ?? "", 10);
    if (Number.isFinite(commentId) && match.index !== undefined) {
      rangeEndById.set(commentId, match.index);
    }
  }
  // Ranges may open in an earlier paragraph (start missing) or close in a
  // later one (end missing); treat the missing side as the paragraph edge.
  const ranges: Array<{ commentId: number; start: number; end: number }> = [];
  const rangeIds = new Set<number>([
    ...rangeStartById.keys(),
    ...rangeEndById.keys(),
  ]);
  rangeIds.forEach((commentId) => {
    const start = rangeStartById.get(commentId) ?? 0;
    const end = rangeEndById.get(commentId) ?? sourceXml.length;
    if (end > start) {
      ranges.push({ commentId, start, end });
    }
  });
  if (ranges.length === 0) {
    setCacheEntry(paragraphCommentMarkupBySourceXml, sourceXml, null);
    return undefined;
  }

  const commentIdsByVisibleChildIndex: Array<number[] | undefined> = [];
  let visibleChildIndex = 0;
  const runPattern = /<w:r\b[\s\S]*?<\/w:r>/gi;
  for (const runMatch of sourceXml.matchAll(runPattern)) {
    const runXml = runMatch[0] ?? "";
    if (!runXml) {
      continue;
    }

    const runStart = runMatch.index ?? 0;
    const contentRunXml = stripTextBoxContentFromRunXml(runXml);
    const visibleTokens = parseTrackedRunTokens(contentRunXml, false);
    const hasImage = /<w:(?:drawing|pict)\b/i.test(runXml);
    const visibleChildCount =
      visibleTokens.filter((token) => token.text.length > 0 || token.isNote)
        .length + (hasImage ? 1 : 0);
    if (visibleChildCount === 0) {
      continue;
    }

    const activeCommentIds = ranges
      .filter((range) => runStart >= range.start && runStart < range.end)
      .map((range) => range.commentId);
    if (activeCommentIds.length > 0) {
      for (let index = 0; index < visibleChildCount; index += 1) {
        commentIdsByVisibleChildIndex[visibleChildIndex + index] =
          activeCommentIds;
      }
    }
    visibleChildIndex += visibleChildCount;
  }

  if (commentIdsByVisibleChildIndex.length === 0) {
    setCacheEntry(paragraphCommentMarkupBySourceXml, sourceXml, null);
    return undefined;
  }

  const resolved: ParagraphCommentMarkup = { commentIdsByVisibleChildIndex };
  setCacheEntry(paragraphCommentMarkupBySourceXml, sourceXml, resolved);
  return resolved;
}

function instructionTextToPageFieldKind(
  rawInstruction: string
): PageFieldKind | undefined {
  const normalized = decodeXmlText(rawInstruction)
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
  if (!normalized || normalized.includes("PAGEREF")) {
    return undefined;
  }

  if (/\bNUMPAGES\b/.test(normalized)) {
    return "NUMPAGES";
  }
  if (/\bPAGE\b/.test(normalized)) {
    return "PAGE";
  }

  return undefined;
}

function instructionTextToStyleRefTarget(
  rawInstruction: string
): string | undefined {
  const normalized = decodeXmlText(rawInstruction).replace(/\s+/g, " ").trim();
  if (!normalized) {
    return undefined;
  }

  const match = normalized.match(/\bSTYLEREF\b\s+(?:"([^"]+)"|([^\s\\]+))/i);
  const target = (match?.[1] ?? match?.[2] ?? "").trim();
  return target.length > 0 ? target : undefined;
}

function paragraphPageFieldSequence(paragraph: ParagraphNode): PageFieldKind[] {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return [];
  }

  const fields: PageFieldKind[] = [];
  for (const instructionMatch of xml.matchAll(
    /<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/gi
  )) {
    const kind = instructionTextToPageFieldKind(instructionMatch[1] ?? "");
    if (kind) {
      fields.push(kind);
    }
  }
  for (const simpleFieldMatch of xml.matchAll(
    /<w:fldSimple\b[^>]*\bw:instr="([^"]+)"[^>]*>/gi
  )) {
    const kind = instructionTextToPageFieldKind(simpleFieldMatch[1] ?? "");
    if (kind) {
      fields.push(kind);
    }
  }

  return fields;
}

const pageFieldValueSequenceBySourceXml = new Map<
  string,
  PageFieldValueToken[]
>();

function paragraphPageFieldValueSequence(
  paragraph: ParagraphNode
): PageFieldValueToken[] {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return [];
  }

  const cached = pageFieldValueSequenceBySourceXml.get(xml);
  if (cached) {
    return cached;
  }

  const values: PageFieldValueToken[] = [];
  const fieldStack: Array<{ kind?: PageFieldKind; inResult: boolean }> = [];
  const tokenPattern =
    /<w:fldSimple\b[^>]*\bw:instr="([^"]+)"[^>]*>[\s\S]*?<\/w:fldSimple>|<w:r\b[\s\S]*?<\/w:r>/gi;

  for (const tokenMatch of xml.matchAll(tokenPattern)) {
    const tokenXml = tokenMatch[0] ?? "";
    if (!tokenXml) {
      continue;
    }

    if (/^<w:fldSimple\b/i.test(tokenXml)) {
      const kind = instructionTextToPageFieldKind(tokenMatch[1] ?? "");
      if (!kind) {
        continue;
      }

      for (const textMatch of tokenXml.matchAll(
        /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi
      )) {
        values.push({
          kind,
          rawText: decodeXmlText(textMatch[1] ?? ""),
        });
      }
      continue;
    }

    const beginCount =
      tokenXml.match(/<w:fldChar\b[^>]*\bw:fldCharType="begin"[^>]*\/?>/gi)
        ?.length ?? 0;
    for (let index = 0; index < beginCount; index += 1) {
      fieldStack.push({ inResult: false });
    }

    for (const instructionMatch of tokenXml.matchAll(
      /<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/gi
    )) {
      const kind = instructionTextToPageFieldKind(instructionMatch[1] ?? "");
      if (!kind || fieldStack.length === 0) {
        continue;
      }

      let assigned = false;
      for (
        let stackIndex = fieldStack.length - 1;
        stackIndex >= 0;
        stackIndex -= 1
      ) {
        if (fieldStack[stackIndex].kind === undefined) {
          fieldStack[stackIndex].kind = kind;
          assigned = true;
          break;
        }
      }
      if (!assigned) {
        fieldStack[fieldStack.length - 1].kind = kind;
      }
    }

    const separateCount =
      tokenXml.match(/<w:fldChar\b[^>]*\bw:fldCharType="separate"[^>]*\/?>/gi)
        ?.length ?? 0;
    for (let index = 0; index < separateCount; index += 1) {
      for (
        let stackIndex = fieldStack.length - 1;
        stackIndex >= 0;
        stackIndex -= 1
      ) {
        if (!fieldStack[stackIndex].inResult) {
          fieldStack[stackIndex].inResult = true;
          break;
        }
      }
    }

    const activeFieldKind = (() => {
      for (
        let stackIndex = fieldStack.length - 1;
        stackIndex >= 0;
        stackIndex -= 1
      ) {
        const stackEntry = fieldStack[stackIndex];
        if (stackEntry.inResult && stackEntry.kind) {
          return stackEntry.kind;
        }
      }
      return undefined;
    })();

    if (activeFieldKind) {
      for (const textMatch of tokenXml.matchAll(
        /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi
      )) {
        values.push({
          kind: activeFieldKind,
          rawText: decodeXmlText(textMatch[1] ?? ""),
        });
      }
    }

    const endCount =
      tokenXml.match(/<w:fldChar\b[^>]*\bw:fldCharType="end"[^>]*\/?>/gi)
        ?.length ?? 0;
    for (let index = 0; index < endCount; index += 1) {
      fieldStack.pop();
    }
  }

  setCacheEntry(pageFieldValueSequenceBySourceXml, xml, values);
  return values;
}

const styleRefFieldValueSequenceBySourceXml = new Map<
  string,
  StyleRefFieldValueToken[]
>();

function paragraphStyleRefFieldValueSequence(
  paragraph: ParagraphNode
): StyleRefFieldValueToken[] {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return [];
  }

  const cached = styleRefFieldValueSequenceBySourceXml.get(xml);
  if (cached) {
    return cached;
  }

  const values: StyleRefFieldValueToken[] = [];
  const fieldStack: Array<{ target?: string; inResult: boolean }> = [];
  const tokenPattern =
    /<w:fldSimple\b[^>]*\bw:instr="([^"]+)"[^>]*>[\s\S]*?<\/w:fldSimple>|<w:r\b[\s\S]*?<\/w:r>/gi;

  for (const tokenMatch of xml.matchAll(tokenPattern)) {
    const tokenXml = tokenMatch[0] ?? "";
    if (!tokenXml) {
      continue;
    }

    if (/^<w:fldSimple\b/i.test(tokenXml)) {
      const target = instructionTextToStyleRefTarget(tokenMatch[1] ?? "");
      if (!target) {
        continue;
      }

      for (const textMatch of tokenXml.matchAll(
        /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi
      )) {
        values.push({
          target,
          rawText: decodeXmlText(textMatch[1] ?? ""),
        });
      }
      continue;
    }

    const beginCount =
      tokenXml.match(/<w:fldChar\b[^>]*\bw:fldCharType="begin"[^>]*\/?>/gi)
        ?.length ?? 0;
    for (let index = 0; index < beginCount; index += 1) {
      fieldStack.push({ inResult: false });
    }

    for (const instructionMatch of tokenXml.matchAll(
      /<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/gi
    )) {
      const target = instructionTextToStyleRefTarget(instructionMatch[1] ?? "");
      if (!target || fieldStack.length === 0) {
        continue;
      }

      let assigned = false;
      for (
        let stackIndex = fieldStack.length - 1;
        stackIndex >= 0;
        stackIndex -= 1
      ) {
        if (fieldStack[stackIndex].target === undefined) {
          fieldStack[stackIndex].target = target;
          assigned = true;
          break;
        }
      }
      if (!assigned) {
        fieldStack[fieldStack.length - 1].target = target;
      }
    }

    const separateCount =
      tokenXml.match(/<w:fldChar\b[^>]*\bw:fldCharType="separate"[^>]*\/?>/gi)
        ?.length ?? 0;
    for (let index = 0; index < separateCount; index += 1) {
      for (
        let stackIndex = fieldStack.length - 1;
        stackIndex >= 0;
        stackIndex -= 1
      ) {
        if (!fieldStack[stackIndex].inResult) {
          fieldStack[stackIndex].inResult = true;
          break;
        }
      }
    }

    const activeFieldTarget = (() => {
      for (
        let stackIndex = fieldStack.length - 1;
        stackIndex >= 0;
        stackIndex -= 1
      ) {
        const stackEntry = fieldStack[stackIndex];
        if (stackEntry.inResult && stackEntry.target) {
          return stackEntry.target;
        }
      }
      return undefined;
    })();

    if (activeFieldTarget) {
      for (const textMatch of tokenXml.matchAll(
        /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi
      )) {
        values.push({
          target: activeFieldTarget,
          rawText: decodeXmlText(textMatch[1] ?? ""),
        });
      }
    }

    const endCount =
      tokenXml.match(/<w:fldChar\b[^>]*\bw:fldCharType="end"[^>]*\/?>/gi)
        ?.length ?? 0;
    for (let index = 0; index < endCount; index += 1) {
      fieldStack.pop();
    }
  }

  setCacheEntry(styleRefFieldValueSequenceBySourceXml, xml, values);
  return values;
}

function normalizeStyleRefTarget(value: string): string {
  return value.replace(/\s+/g, " ").trim().toLowerCase();
}

function paragraphStyledRunText(
  paragraph: ParagraphNode,
  styleRefTarget: string
): string | undefined {
  const xml = paragraph.sourceXml ?? "";
  if (!xml) {
    return undefined;
  }

  const normalizedTarget = normalizeStyleRefTarget(styleRefTarget);
  if (!normalizedTarget) {
    return undefined;
  }

  const textChunks: string[] = [];
  for (const runMatch of xml.matchAll(/<w:r\b[\s\S]*?<\/w:r>/gi)) {
    const runXml = runMatch[0] ?? "";
    if (!runXml) {
      continue;
    }

    const runStyleTag = runXml.match(/<w:rStyle\b[^>]*>/i)?.[0] ?? "";
    const runStyleValue = xmlAttribute(runStyleTag, "w:val");
    if (normalizeStyleRefTarget(runStyleValue ?? "") !== normalizedTarget) {
      continue;
    }

    for (const textMatch of runXml.matchAll(
      /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi
    )) {
      const value = decodeXmlText(textMatch[1] ?? "");
      if (value.length > 0) {
        textChunks.push(value);
      }
    }

    if (/<w:tab\b[^>]*\/?>/i.test(runXml)) {
      textChunks.push("\t");
    }
  }

  const joined = textChunks.join("").trim();
  return joined.length > 0 ? joined : undefined;
}

function tabLeaderStyle(
  leader: string | undefined,
  color: string | undefined
): Record<string, string | number | undefined> {
  const normalizedLeader = leader === "middleDot" ? "dot" : leader;
  const resolvedColor = color || "currentColor";
  if (normalizedLeader === "hyphen") {
    return {
      backgroundImage: `linear-gradient(to right, transparent 0, transparent 4px, ${resolvedColor} 4px, ${resolvedColor} 5px, transparent 5px, transparent 8px)`,
      backgroundSize: "8px 1px",
      backgroundPosition: "0 80%",
      backgroundRepeat: "repeat-x",
    };
  }
  if (normalizedLeader === "underscore") {
    return {
      backgroundImage: `linear-gradient(to right, ${resolvedColor} 0, ${resolvedColor} 1px, transparent 1px, transparent 4px)`,
      backgroundSize: "4px 1px",
      backgroundPosition: "0 90%",
      backgroundRepeat: "repeat-x",
    };
  }

  return {
    backgroundImage: `radial-gradient(${resolvedColor} 1px, transparent 1px)`,
    backgroundSize: "7px 9px",
    backgroundPosition: "0 72%",
    backgroundRepeat: "repeat-x",
  };
}

function noteMarkerLabel(
  noteReference: TextRunNode["noteReference"],
  footnoteDisplayIndexById: Map<number, number>,
  endnoteDisplayIndexById: Map<number, number>
): string | undefined {
  if (!noteReference) {
    return undefined;
  }

  const index =
    noteReference.kind === "footnote"
      ? footnoteDisplayIndexById.get(noteReference.id)
      : endnoteDisplayIndexById.get(noteReference.id);

  if (noteReference.kind === "footnote") {
    const value = index ?? Math.max(1, Math.round(noteReference.id));
    if (!Number.isFinite(value) || value <= 0) {
      return undefined;
    }
    return String(Math.round(value));
  }

  const romanValue = index ?? Math.max(1, Math.round(noteReference.id));
  if (!Number.isFinite(romanValue) || romanValue <= 0) {
    return undefined;
  }

  return numberToRoman(Math.round(romanValue)).toLowerCase();
}

function normalizeDateInputValue(value?: string): string | undefined {
  if (!value) {
    return undefined;
  }

  const trimmed = value.trim();
  if (!trimmed) {
    return undefined;
  }

  const directMatch = trimmed.match(/^(\d{4}-\d{2}-\d{2})/);
  if (directMatch?.[1]) {
    return directMatch[1];
  }

  const parsed = new Date(trimmed);
  if (Number.isNaN(parsed.getTime())) {
    return undefined;
  }

  return parsed.toISOString().slice(0, 10);
}

function numberToRoman(value: number): string {
  if (!Number.isFinite(value) || value <= 0) {
    return "";
  }

  const numerals: Array<[number, string]> = [
    [1000, "M"],
    [900, "CM"],
    [500, "D"],
    [400, "CD"],
    [100, "C"],
    [90, "XC"],
    [50, "L"],
    [40, "XL"],
    [10, "X"],
    [9, "IX"],
    [5, "V"],
    [4, "IV"],
    [1, "I"],
  ];

  let remaining = Math.floor(value);
  let output = "";

  for (const [base, numeral] of numerals) {
    while (remaining >= base) {
      output += numeral;
      remaining -= base;
    }
  }

  return output;
}

function numberToLetters(value: number): string {
  if (!Number.isFinite(value) || value <= 0) {
    return "";
  }

  let remaining = Math.floor(value);
  let output = "";
  while (remaining > 0) {
    const offset = (remaining - 1) % 26;
    output = String.fromCharCode(65 + offset) + output;
    remaining = Math.floor((remaining - 1) / 26);
  }

  return output;
}

interface ParagraphNumberingLabel {
  text?: string;
  fontFamily?: string;
  color?: string;
  style?: TextRunNode["style"];
  imageSrc?: string;
  imageWidthPx?: number;
  imageHeightPx?: number;
  trailingText?: string;
}

function formatNumberingCounter(
  format: string | undefined,
  value: number
): string {
  if (!Number.isFinite(value) || value <= 0) {
    return "";
  }

  switch ((format ?? "decimal").toLowerCase()) {
    case "decimal":
      return String(value);
    case "upperroman":
      return numberToRoman(value);
    case "lowerroman":
      return numberToRoman(value).toLowerCase();
    case "upperletter":
      return numberToLetters(value);
    case "lowerletter":
      return numberToLetters(value).toLowerCase();
    case "ordinal":
    case "cardinaltext":
    case "ordinaltext":
      return String(value);
    case "bullet":
      return "\u2022";
    case "none":
      return "";
    default:
      return String(value);
  }
}

function formatPageFieldValue(
  value: number,
  format: string | undefined
): string {
  const formatted = formatNumberingCounter(format, value);
  return formatted.length > 0 ? formatted : String(value);
}

function findNumberingLevelDefinition(
  numbering: NumberingDefinitionSet,
  numId: number,
  ilvl: number
): NumberingLevelDefinition | undefined {
  const instance = numbering.instances.find((item) => item.numId === numId);
  if (!instance) {
    return undefined;
  }

  const overrideLevel = instance.levelOverrides?.find(
    (item) => item.ilvl === ilvl
  );
  if (overrideLevel) {
    return overrideLevel;
  }

  const abstract = numbering.abstracts.find(
    (item) => item.abstractNumId === instance.abstractNumId
  );
  return abstract?.levels.find((item) => item.ilvl === ilvl);
}

function numberingLevelHasVisibleMarker(
  level: NumberingLevelDefinition | undefined
): boolean {
  return Boolean(
    level?.pictureBullet?.src || (level?.text && level.text.trim().length > 0)
  );
}

function numberingLevelIsBulletLike(
  level: NumberingLevelDefinition | undefined
): boolean {
  const format = level?.format?.trim().toLowerCase();
  return (
    format === "bullet" ||
    Boolean(level?.pictureBullet?.src) ||
    isBulletLikeNumberingText(level?.text ?? "", level?.bulletFontFamily)
  );
}

function numberingAbstractLevelsForNumId(
  numberingDefinitions: NumberingDefinitionSet,
  numId: number
): NumberingLevelDefinition[] {
  const instance = numberingDefinitions.instances.find(
    (item) => item.numId === numId
  );
  if (!instance) {
    return [];
  }

  const abstract = numberingDefinitions.abstracts.find(
    (item) => item.abstractNumId === instance.abstractNumId
  );
  return [...(abstract?.levels ?? []), ...(instance.levelOverrides ?? [])];
}

function effectiveNumberingNumIdForParagraph(
  paragraph: ParagraphNode,
  numberingDefinitions: NumberingDefinitionSet | undefined
): number | undefined {
  const numbering = paragraph.style?.numbering;
  if (
    !numberingDefinitions ||
    !numbering ||
    !Number.isFinite(numbering.numId) ||
    numbering.numId <= 0
  ) {
    return numbering?.numId;
  }

  const numId = numbering.numId;
  const ilvl = Math.max(0, Math.round(numbering.ilvl ?? 0));
  const currentLevel = findNumberingLevelDefinition(
    numberingDefinitions,
    numId,
    ilvl
  );
  if (
    !numberingLevelIsBulletLike(currentLevel) ||
    numberingLevelHasVisibleMarker(currentLevel)
  ) {
    return numId;
  }

  const currentAbstractLevels = numberingAbstractLevelsForNumId(
    numberingDefinitions,
    numId
  );
  if (
    currentAbstractLevels.length === 0 ||
    currentAbstractLevels.some(
      (level) =>
        !numberingLevelIsBulletLike(level) ||
        numberingLevelHasVisibleMarker(level)
    )
  ) {
    return numId;
  }

  const paragraphTextValue = paragraphText(paragraph).trim();
  if (
    paragraphTextValue.length === 0 ||
    isBulletLikeNumberingText(paragraphTextValue) ||
    /^[\u2022\u25cf\u25cb\u25a0\u25a1\u25aa\u25ab\u25c6\u25c7\u25e6\u2043]/.test(
      paragraphTextValue
    )
  ) {
    return numId;
  }

  const candidate = numberingDefinitions.instances.find((instance) => {
    if (instance.numId === numId) {
      return false;
    }
    const candidateLevel = findNumberingLevelDefinition(
      numberingDefinitions,
      instance.numId,
      ilvl
    );
    if (!candidateLevel) {
      return false;
    }
    const format = candidateLevel.format?.trim().toLowerCase();
    if (
      !format ||
      format === "none" ||
      numberingLevelIsBulletLike(candidateLevel)
    ) {
      return false;
    }
    if (!candidateLevel.text || !/%\d+/.test(candidateLevel.text)) {
      return false;
    }
    if (ilvl > 0) {
      const parentLevel = findNumberingLevelDefinition(
        numberingDefinitions,
        instance.numId,
        ilvl - 1
      );
      if (!parentLevel) {
        return false;
      }
      const parentFormat = parentLevel.format?.trim().toLowerCase();
      if (
        !parentFormat ||
        parentFormat === "none" ||
        numberingLevelIsBulletLike(parentLevel)
      ) {
        return false;
      }
    }
    return true;
  });

  return candidate?.numId ?? numId;
}

function numberingStartValue(
  numbering: NumberingDefinitionSet,
  numId: number,
  ilvl: number
): number {
  const instance = numbering.instances.find((item) => item.numId === numId);
  const override = instance?.levelStartOverrides?.[String(ilvl)];
  if (Number.isFinite(override)) {
    return Math.max(1, Math.round(override as number));
  }

  const level = findNumberingLevelDefinition(numbering, numId, ilvl);
  if (Number.isFinite(level?.start) && (level?.start as number) > 0) {
    return Math.round(level?.start as number);
  }

  return 1;
}

function numberingSuffix(level: NumberingLevelDefinition | undefined): string {
  if (level?.suffix === "tab") {
    return "\t";
  }
  if (level?.suffix === "space") {
    return " ";
  }
  return "";
}

const LEGACY_BULLET_GLYPH_FALLBACKS: Record<string, string> = {
  "\uf0a7": "□",
  "\uf0a8": "▪",
  "\uf0b7": "•",
  "\uf0d8": "➢",
};

function normalizeLegacyBulletGlyphs(
  text: string,
  bulletFontFamily?: string
): string {
  if (!text) {
    return text;
  }

  const normalizedFont = bulletFontFamily?.trim().toLowerCase() ?? "";
  const fontUsesLegacyGlyphs =
    normalizedFont.includes("wingdings") ||
    normalizedFont.includes("symbol") ||
    normalizedFont.includes("courier");
  if (!fontUsesLegacyGlyphs && !/[\uf0a7\uf0a8\uf0b7\uf0d8]/.test(text)) {
    return text;
  }

  return Array.from(text)
    .map((glyph) => {
      const mapped = LEGACY_BULLET_GLYPH_FALLBACKS[glyph];
      if (mapped) {
        return mapped;
      }
      if (glyph === "o" && normalizedFont.includes("courier")) {
        return "◦";
      }
      return glyph;
    })
    .join("");
}

export function buildParagraphNumberingLabels(
  model: DocModel
): Map<string, ParagraphNumberingLabel> {
  const labels = new Map<string, ParagraphNumberingLabel>();
  const numbering = model.metadata.numberingDefinitions;
  if (!numbering) {
    return labels;
  }

  const numberingInstanceByNumId = new Map(
    numbering.instances.map((instance) => [instance.numId, instance])
  );
  const paragraphStyleById = new Map(
    (model.metadata.paragraphStyles ?? []).map((styleDefinition) => [
      styleDefinition.id,
      styleDefinition,
    ])
  );
  const countersByNumId = new Map<number, Array<number | undefined>>();
  const abstractCountersByAbstractNumId = new Map<
    number,
    Array<number | undefined>
  >();

  const registerParagraph = (paragraph: ParagraphNode, key: string): void => {
    const paragraphNumbering = paragraph.style?.numbering;
    if (!paragraphNumbering || paragraphNumbering.numId <= 0) {
      return;
    }

    const numId =
      effectiveNumberingNumIdForParagraph(paragraph, numbering) ??
      paragraphNumbering.numId;
    const ilvl = Math.max(0, paragraphNumbering.ilvl ?? 0);
    const level = findNumberingLevelDefinition(numbering, numId, ilvl);
    const numberingInstance = numberingInstanceByNumId.get(numId);
    const abstractNumId = numberingInstance?.abstractNumId;
    const template = level?.text ?? `%${ilvl + 1}.`;
    if (!template || template.trim().length === 0) {
      return;
    }

    const counters = countersByNumId.get(numId) ?? [];
    const sharedAbstractCounters = Number.isFinite(abstractNumId)
      ? abstractCountersByAbstractNumId.get(abstractNumId as number) ?? []
      : undefined;
    let parentCountersChanged = false;
    for (let index = 0; index < ilvl; index += 1) {
      const abstractCounterValue =
        sharedAbstractCounters && Number.isFinite(sharedAbstractCounters[index])
          ? Math.max(1, Math.round(sharedAbstractCounters[index] as number))
          : undefined;
      if (abstractCounterValue !== undefined) {
        const previousCounter = counters[index];
        if (
          !Number.isFinite(previousCounter) ||
          Math.max(1, Math.round(previousCounter as number)) !==
            abstractCounterValue
        ) {
          parentCountersChanged = true;
        }
        counters[index] = abstractCounterValue;
        continue;
      }
      if (!Number.isFinite(counters[index])) {
        counters[index] = numberingStartValue(numbering, numId, index);
      }
    }

    const currentValue = parentCountersChanged ? undefined : counters[ilvl];
    counters[ilvl] = Number.isFinite(currentValue)
      ? (currentValue as number) + 1
      : numberingStartValue(numbering, numId, ilvl);
    for (let index = ilvl + 1; index < counters.length; index += 1) {
      counters[index] = undefined;
    }
    countersByNumId.set(numId, counters);
    if (sharedAbstractCounters && Number.isFinite(abstractNumId)) {
      for (let index = 0; index <= ilvl; index += 1) {
        sharedAbstractCounters[index] = counters[index];
      }
      for (
        let index = ilvl + 1;
        index < sharedAbstractCounters.length;
        index += 1
      ) {
        sharedAbstractCounters[index] = undefined;
      }
      abstractCountersByAbstractNumId.set(
        abstractNumId as number,
        sharedAbstractCounters
      );
    }

    const label = template.replaceAll(/%(\d+)/g, (_, rawLevel) => {
      const levelIndex = Math.max(0, Number(rawLevel) - 1);
      const levelValue = counters[levelIndex];
      const safeValue = Number.isFinite(levelValue)
        ? Math.max(1, Math.round(levelValue as number))
        : numberingStartValue(numbering, numId, levelIndex);
      const levelFormat = findNumberingLevelDefinition(
        numbering,
        numId,
        levelIndex
      )?.format;
      return formatNumberingCounter(levelFormat, safeValue);
    });

    if (!label || label.trim().length === 0) {
      return;
    }

    const isBullet = (level?.format ?? "").trim().toLowerCase() === "bullet";
    const trailingText = numberingSuffix(level);
    if (isBullet && level?.pictureBullet?.src) {
      labels.set(key, {
        imageSrc: level.pictureBullet.src,
        imageWidthPx: level.pictureBullet.widthPx,
        imageHeightPx: level.pictureBullet.heightPx,
        trailingText,
      });
      return;
    }

    const firstStyledRun = paragraph.children.find(
      (child): child is TextRunNode | FormFieldRunNode =>
        (child.type === "text" || child.type === "form-field") &&
        Boolean(child.style)
    );
    const paragraphStyleRunStyle = paragraph.style?.styleId
      ? paragraphStyleById.get(paragraph.style.styleId)?.runStyle
      : undefined;
    const baseNumberingTextStyle =
      firstStyledRun?.style ?? paragraphStyleRunStyle;
    const numberingTextStyle = level?.runStyle
      ? { ...(baseNumberingTextStyle ?? {}), ...level.runStyle }
      : baseNumberingTextStyle;
    const normalizedLabel = isBullet
      ? normalizeLegacyBulletGlyphs(label, level?.bulletFontFamily)
      : label;
    labels.set(key, {
      text: `${normalizedLabel}${trailingText}`,
      fontFamily: isBullet ? level?.bulletFontFamily : undefined,
      color: isBullet ? level?.bulletColor : undefined,
      style: numberingTextStyle ? { ...numberingTextStyle } : undefined,
    });
  };

  model.nodes.forEach((node, nodeIndex) => {
    if (node.type === "paragraph") {
      registerParagraph(node, `p:${nodeIndex}`);
      return;
    }

    node.rows.forEach((row, rowIndex) => {
      row.cells.forEach((cell, cellIndex) => {
        tableCellParagraphs(cell.nodes).forEach((paragraph, paragraphIndex) => {
          registerParagraph(
            paragraph,
            `t:${nodeIndex}:${rowIndex}:${cellIndex}:${paragraphIndex}`
          );
        });
      });
    });
  });

  return labels;
}

const UNORDERED_LIST_PREFIX_PATTERN = /^\s*•\s+/;
const ORDERED_LIST_PREFIX_PATTERN = /^\s*\d+\.\s+/;
const LIST_PREFIX_PATTERN = /^\s*(?:•\s+|\d+\.\s+)/;
const ORDERED_LIST_PREFIX_CAPTURE_PATTERN = /^(\s*)(\d+)\.\s+/;
const LIST_LEVEL_STEP_TWIPS = 720;
const DEFAULT_LIST_HANGING_TWIPS = 360;
const MAX_FALLBACK_LIST_LEVEL = 8;

function isUnorderedListText(text: string): boolean {
  return UNORDERED_LIST_PREFIX_PATTERN.test(text);
}

function isOrderedListText(text: string): boolean {
  return ORDERED_LIST_PREFIX_PATTERN.test(text);
}

function isBulletLikeNumberingText(
  text: string,
  bulletFontFamily?: string
): boolean {
  const normalized = normalizeLegacyBulletGlyphs(text, bulletFontFamily).trim();
  if (!normalized || /%[\d]+/.test(normalized)) {
    return false;
  }

  return /^[\u2022\u25cf\u25cb\u25a0\u25a1\u25aa\u25ab\u25c6\u25c7\u25e6\u2043\-–—]+$/.test(
    normalized
  );
}

function paragraphHasNumbering(paragraph: ParagraphNode): boolean {
  return Boolean(
    paragraph.style?.numbering && paragraph.style.numbering.numId > 0
  );
}

function paragraphListType(
  paragraph: ParagraphNode,
  numberingDefinitions?: NumberingDefinitionSet
): DocxListType | undefined {
  const numbering = paragraph.style?.numbering;
  if (numbering && numbering.numId > 0) {
    if (numberingDefinitions) {
      const effectiveNumId =
        effectiveNumberingNumIdForParagraph(paragraph, numberingDefinitions) ??
        numbering.numId;
      const level = findNumberingLevelDefinition(
        numberingDefinitions,
        effectiveNumId,
        Math.max(0, numbering.ilvl ?? 0)
      );
      const format = level?.format?.trim().toLowerCase();
      if (
        format === "bullet" ||
        level?.pictureBullet ||
        isBulletLikeNumberingText(level?.text ?? "", level?.bulletFontFamily)
      ) {
        return "unordered";
      }
      if (format && format !== "none") {
        return "ordered";
      }
    }
    // If numbering exists but we cannot resolve a specific list format, treat it as ordered.
    return "ordered";
  }

  const text = paragraphText(paragraph);
  if (isUnorderedListText(text)) {
    return "unordered";
  }
  if (isOrderedListText(text)) {
    return "ordered";
  }
  return undefined;
}

function paragraphIsList(
  paragraph: ParagraphNode,
  textOverride?: string
): boolean {
  const text = textOverride ?? paragraphText(paragraph);
  return (
    paragraphHasNumbering(paragraph) ||
    isUnorderedListText(text) ||
    isOrderedListText(text)
  );
}

function stripListPrefix(text: string): string {
  return text.replace(/^\s*(?:•\s+|\d+\.\s+)/, "");
}

function listPrefixLength(text: string): number {
  return text.match(LIST_PREFIX_PATTERN)?.[0]?.length ?? 0;
}

function nextOrderedListItemText(
  currentText: string,
  nextItemText: string
): string {
  const match = currentText.match(ORDERED_LIST_PREFIX_CAPTURE_PATTERN);
  const leadingWhitespace = match?.[1] ?? "";
  const currentNumber = Number.parseInt(match?.[2] ?? "", 10);
  const nextNumber = Number.isFinite(currentNumber)
    ? Math.max(1, currentNumber + 1)
    : 1;
  return `${leadingWhitespace}${nextNumber}. ${stripListPrefix(nextItemText)}`;
}

function textWithListType(text: string, listType: DocxListType): string {
  const normalized = stripListPrefix(text);
  return listType === "unordered" ? `• ${normalized}` : `1. ${normalized}`;
}

function cloneParagraphBorderStyle(
  border?: ParagraphBorderStyle
): ParagraphBorderStyle | undefined {
  return border ? { ...border } : undefined;
}

function cloneParagraphBorderSet(
  borders?: ParagraphBorderSet
): ParagraphBorderSet | undefined {
  if (!borders) {
    return undefined;
  }

  return {
    top: cloneParagraphBorderStyle(borders.top),
    right: cloneParagraphBorderStyle(borders.right),
    bottom: cloneParagraphBorderStyle(borders.bottom),
    left: cloneParagraphBorderStyle(borders.left),
    between: cloneParagraphBorderStyle(borders.between),
    bar: cloneParagraphBorderStyle(borders.bar),
  };
}

function cloneParagraphStyle(
  style?: ParagraphNode["style"]
): ParagraphNode["style"] | undefined {
  if (!style) {
    return undefined;
  }

  return {
    ...style,
    numbering: style.numbering ? { ...style.numbering } : undefined,
    spacing: style.spacing ? { ...style.spacing } : undefined,
    indent: style.indent ? { ...style.indent } : undefined,
    tabStops: style.tabStops
      ? style.tabStops.map((tabStop) => ({ ...tabStop }))
      : undefined,
    borders: cloneParagraphBorderSet(style.borders),
    dropCap: style.dropCap ? { ...style.dropCap } : undefined,
  };
}

function splitParagraphStyleWithDefaultSpacing(
  style?: ParagraphNode["style"],
  sourceXml?: string
): ParagraphNode["style"] | undefined {
  const clonedStyle = cloneParagraphStyle(style);
  // Keep imported/explicit paragraph styling untouched.
  if (sourceXml || clonedStyle?.styleId) {
    return clonedStyle;
  }

  const spacing = clonedStyle?.spacing;
  const hasExplicitSpacing =
    spacing?.beforeTwips !== undefined ||
    spacing?.afterTwips !== undefined ||
    spacing?.lineTwips !== undefined ||
    spacing?.lineRule !== undefined;
  if (hasExplicitSpacing) {
    return clonedStyle;
  }

  return {
    ...(clonedStyle ?? {}),
    spacing: {
      ...(spacing ?? {}),
      lineRule: "auto",
      lineTwips: DEFAULT_SPLIT_PARAGRAPH_LINE_TWIPS,
      afterTwips: DEFAULT_SPLIT_PARAGRAPH_AFTER_TWIPS,
    },
  };
}

function cloneTableBoxSpacing(
  spacing?: TableCellStyle["marginTwips"]
): TableCellStyle["marginTwips"] | undefined {
  if (!spacing) {
    return undefined;
  }

  return {
    topTwips: spacing.topTwips,
    rightTwips: spacing.rightTwips,
    bottomTwips: spacing.bottomTwips,
    leftTwips: spacing.leftTwips,
  };
}

function cloneTableBorderStyle(
  border?: TableBorderStyle
): TableBorderStyle | undefined {
  if (!border) {
    return undefined;
  }

  return {
    type: border.type,
    color: border.color,
    sizeEighthPt: border.sizeEighthPt,
  };
}

function cloneTableBorderSet(
  borders?: TableBorderSet
): TableBorderSet | undefined {
  if (!borders) {
    return undefined;
  }

  return {
    top: cloneTableBorderStyle(borders.top),
    right: cloneTableBorderStyle(borders.right),
    bottom: cloneTableBorderStyle(borders.bottom),
    left: cloneTableBorderStyle(borders.left),
    insideH: cloneTableBorderStyle(borders.insideH),
    insideV: cloneTableBorderStyle(borders.insideV),
    tl2br: cloneTableBorderStyle(borders.tl2br),
    tr2bl: cloneTableBorderStyle(borders.tr2bl),
  };
}

function cloneTableCellStyle(
  style?: TableCellStyle
): TableCellStyle | undefined {
  if (!style) {
    return undefined;
  }

  return {
    ...style,
    marginTwips: cloneTableBoxSpacing(style.marginTwips),
    borders: cloneTableBorderSet(style.borders),
  };
}

function cloneTableRowStyle(style?: TableRowStyle): TableRowStyle | undefined {
  if (!style) {
    return undefined;
  }

  return {
    ...style,
  };
}

function createEmptyParagraphFromTemplate(
  template?: ParagraphNode
): ParagraphNode {
  const nextTextStyle = cloneTextStyle(firstRunStyle(template));

  return {
    type: "paragraph",
    style: cloneParagraphStyle(template?.style),
    children: [
      {
        type: "text",
        text: "",
        style: nextTextStyle,
      },
    ],
  };
}

function createEmptyTableCellFromTemplate(
  templateCell?: TableNode["rows"][number]["cells"][number]
): TableNode["rows"][number]["cells"][number] {
  const templateParagraph = templateCell
    ? tableCellParagraphs(templateCell.nodes)[0]
    : undefined;
  const clonedStyle = cloneTableCellStyle(templateCell?.style);
  const normalizedStyle = clonedStyle
    ? {
        ...clonedStyle,
        rowSpan: undefined,
        vMergeContinuation: undefined,
      }
    : undefined;

  return {
    type: "table-cell",
    style: normalizedStyle,
    nodes: [createEmptyParagraphFromTemplate(templateParagraph)],
  };
}

function resolveMaxNumberingLevel(
  numberingDefinitions: NumberingDefinitionSet | undefined,
  numId: number
): number {
  if (!numberingDefinitions) {
    return MAX_FALLBACK_LIST_LEVEL;
  }

  const instance = numberingDefinitions.instances.find(
    (item) => item.numId === numId
  );
  if (!instance) {
    return MAX_FALLBACK_LIST_LEVEL;
  }

  const abstract = numberingDefinitions.abstracts.find(
    (item) => item.abstractNumId === instance.abstractNumId
  );
  const levels = [
    ...(abstract?.levels ?? []),
    ...(instance.levelOverrides ?? []),
  ];
  if (levels.length === 0) {
    return MAX_FALLBACK_LIST_LEVEL;
  }

  return clampNumber(
    Math.max(
      0,
      ...levels.map((level) => Math.max(0, Math.round(level.ilvl ?? 0)))
    ),
    0,
    MAX_FALLBACK_LIST_LEVEL
  );
}

function shiftListIndent(
  indent: ParagraphIndent | undefined,
  levelDelta: number
): ParagraphIndent | undefined {
  if (!levelDelta) {
    return indent;
  }

  const nextLeftTwips = Math.max(
    0,
    (indent?.leftTwips ?? 0) + levelDelta * LIST_LEVEL_STEP_TWIPS
  );
  const nextHangingTwips =
    indent?.hangingTwips ??
    (nextLeftTwips > 0 ? DEFAULT_LIST_HANGING_TWIPS : undefined);

  return {
    ...(indent ?? {}),
    leftTwips: nextLeftTwips,
    hangingTwips: nextHangingTwips,
  };
}

function ensurePrefixListIndent(paragraph: ParagraphNode): boolean {
  const clonedStyle = cloneParagraphStyle(paragraph.style) ?? {};
  const existingIndent = clonedStyle.indent;
  const hasMeaningfulLeftIndent =
    Number.isFinite(existingIndent?.leftTwips) &&
    Math.abs(existingIndent?.leftTwips ?? 0) > 0;
  const hasMeaningfulFirstLine =
    Number.isFinite(existingIndent?.firstLineTwips) &&
    Math.abs(existingIndent?.firstLineTwips ?? 0) > 0;
  const hasMeaningfulHanging =
    Number.isFinite(existingIndent?.hangingTwips) &&
    Math.abs(existingIndent?.hangingTwips ?? 0) > 0;

  if (
    hasMeaningfulLeftIndent ||
    hasMeaningfulFirstLine ||
    hasMeaningfulHanging
  ) {
    return false;
  }

  paragraph.style = {
    ...clonedStyle,
    indent: {
      ...(existingIndent ?? {}),
      leftTwips: LIST_LEVEL_STEP_TWIPS,
      hangingTwips: DEFAULT_LIST_HANGING_TWIPS,
    },
  };
  paragraph.sourceXml = undefined;
  return true;
}

function clearAutoPrefixListIndent(paragraph: ParagraphNode): boolean {
  const clonedStyle = cloneParagraphStyle(paragraph.style);
  const existingIndent = clonedStyle?.indent;
  if (!clonedStyle || !existingIndent) {
    return false;
  }

  const leftTwips = existingIndent.leftTwips;
  const hangingTwips = existingIndent.hangingTwips;
  const firstLineTwips = existingIndent.firstLineTwips;
  const matchesAutoIndent =
    Number.isFinite(leftTwips) &&
    Math.round(leftTwips ?? 0) === LIST_LEVEL_STEP_TWIPS &&
    Number.isFinite(hangingTwips) &&
    Math.round(hangingTwips ?? 0) === DEFAULT_LIST_HANGING_TWIPS &&
    !Number.isFinite(firstLineTwips);
  if (!matchesAutoIndent) {
    return false;
  }

  const nextIndent: ParagraphIndent = {
    ...existingIndent,
    leftTwips: undefined,
    hangingTwips: undefined,
  };
  const hasRemainingIndent = [
    nextIndent.leftTwips,
    nextIndent.rightTwips,
    nextIndent.firstLineTwips,
    nextIndent.hangingTwips,
  ].some((value) => Number.isFinite(value));

  paragraph.style = {
    ...clonedStyle,
    indent: hasRemainingIndent ? nextIndent : undefined,
  };
  paragraph.sourceXml = undefined;
  return true;
}

function tableCellText(paragraphs: ParagraphNode[]): string {
  return paragraphs.map(paragraphText).join("\n");
}

function tableColumnCount(table: TableNode): number {
  return Math.max(
    1,
    ...table.rows.map((row) =>
      row.cells.reduce(
        (total, cell) =>
          total +
          (cell.style?.gridSpan && cell.style.gridSpan > 1
            ? cell.style.gridSpan
            : 1),
        0
      )
    )
  );
}

function resolveFloatingTableSide(
  table: TableNode
): "left" | "right" | undefined {
  const floating = table.style?.floating;
  if (!floating) {
    return undefined;
  }

  const horizontalAlign = floating.horizontalAlign?.toLowerCase();
  if (horizontalAlign === "right" || horizontalAlign === "outside") {
    return "right";
  }
  if (horizontalAlign === "left" || horizontalAlign === "inside") {
    return "left";
  }

  if (Number.isFinite(floating.xTwips) && (floating.xTwips as number) > 1440) {
    return "right";
  }

  return "left";
}

function tableWrapperStyle(
  table: TableNode,
  indentPx: number
): Record<string, string | number | undefined> {
  const floating = table.style?.floating;
  if (!floating) {
    return {
      marginLeft: indentPx,
      position: "relative",
    };
  }

  const side = resolveFloatingTableSide(table) ?? "left";
  const marginTop = twipsToPixels(floating.topFromTextTwips) ?? 0;
  const marginBottom = twipsToPixels(floating.bottomFromTextTwips) ?? 8;
  const marginLeftFromText = twipsToPixels(floating.leftFromTextTwips) ?? 8;
  const marginRightFromText = twipsToPixels(floating.rightFromTextTwips) ?? 8;

  return {
    float: side,
    clear: "none",
    marginTop,
    marginBottom,
    marginLeft:
      side === "left" ? marginLeftFromText + indentPx : marginLeftFromText,
    marginRight:
      side === "right" ? marginRightFromText + indentPx : marginRightFromText,
    position: "relative",
    zIndex: 1,
  };
}

interface EmbeddedTableRuntimeKeySegment {
  rowIndex: number;
  cellIndex: number;
  contentIndex: number;
}

interface EmbeddedTableRuntimeKeyLocation {
  hostTableIndex: number;
  hostRowIndex: number;
  hostCellIndex: number;
  rootContentIndex: number;
  descendants: EmbeddedTableRuntimeKeySegment[];
}

function parseEmbeddedTableRuntimeKey(
  tableRuntimeKey: string
): EmbeddedTableRuntimeKeyLocation | undefined {
  const tokens = tableRuntimeKey.split("-");
  const parseIndex = (token: string | undefined): number | undefined => {
    const value = Number(token);
    if (!Number.isInteger(value) || value < 0) {
      return undefined;
    }
    return value;
  };

  if (
    tokens.length < 8 ||
    (tokens[0] !== "body" && tokens[0] !== "active") ||
    tokens[1] !== "cell" ||
    tokens[2] !== "nested" ||
    tokens[3] !== "table"
  ) {
    return undefined;
  }

  const hostTableIndex = parseIndex(tokens[4]);
  const hostRowIndex = parseIndex(tokens[5]);
  const hostCellIndex = parseIndex(tokens[6]);
  const rootContentIndex = parseIndex(tokens[7]);
  if (
    hostTableIndex === undefined ||
    hostRowIndex === undefined ||
    hostCellIndex === undefined ||
    rootContentIndex === undefined
  ) {
    return undefined;
  }

  const descendants: EmbeddedTableRuntimeKeySegment[] = [];
  let cursor = 8;
  while (cursor < tokens.length) {
    if (tokens[cursor] !== "nested" || tokens[cursor + 1] !== "table") {
      return undefined;
    }

    const rowIndex = parseIndex(tokens[cursor + 2]);
    const cellIndex = parseIndex(tokens[cursor + 3]);
    const contentIndex = parseIndex(tokens[cursor + 4]);
    if (
      rowIndex === undefined ||
      cellIndex === undefined ||
      contentIndex === undefined
    ) {
      return undefined;
    }

    descendants.push({ rowIndex, cellIndex, contentIndex });
    cursor += 5;
  }

  return {
    hostTableIndex,
    hostRowIndex,
    hostCellIndex,
    rootContentIndex,
    descendants,
  };
}

const columnWidthsByTable = new WeakMap<
  TableNode,
  Map<number, number[] | undefined>
>();

function columnWidthsFromTableDefinition(
  table: TableNode,
  columnCount: number
): number[] | undefined {
  const cachedByCount = columnWidthsByTable.get(table);
  if (cachedByCount?.has(columnCount)) {
    return cachedByCount.get(columnCount);
  }

  const resolved = computeColumnWidthsFromTableDefinition(table, columnCount);
  const cache = cachedByCount ?? new Map<number, number[] | undefined>();
  cache.set(columnCount, resolved);
  columnWidthsByTable.set(table, cache);
  return resolved;
}

function computeColumnWidthsFromTableDefinition(
  table: TableNode,
  columnCount: number
): number[] | undefined {
  const gridWidths = table.style?.columnWidthsTwips;
  const rowDerivedWidths = deriveColumnWidthsFromTableRows(table, columnCount);

  if (gridWidths && gridWidths.length === columnCount) {
    // Some generators emit a placeholder uniform grid while the real column
    // geometry lives in per-cell tcW. Word's fixed-layout algorithm trusts
    // cell widths over the grid, so when the two disagree on most measured
    // cells, prefer the row-derived widths.
    if (
      rowDerivedWidths &&
      rowDerivedWidths.length > 0 &&
      gridConflictsWithRowWidths(table, gridWidths)
    ) {
      return rowDerivedWidths;
    }
    return normalizeColumnWidthsTwips(gridWidths, columnCount);
  }

  if (rowDerivedWidths && rowDerivedWidths.length > 0) {
    return rowDerivedWidths;
  }

  if (gridWidths && gridWidths.length > 0) {
    return normalizeColumnWidthsTwips(gridWidths, columnCount);
  }

  return undefined;
}

function gridConflictsWithRowWidths(
  table: TableNode,
  gridWidths: number[]
): boolean {
  let conflictRows = 0;
  let measuredRows = 0;

  for (const row of table.rows) {
    let columnCursor = 0;
    let measuredCells = 0;
    let conflictCells = 0;

    for (const cell of row.cells) {
      const span =
        cell.style?.gridSpan && cell.style.gridSpan > 1
          ? cell.style.gridSpan
          : 1;
      const expected = gridWidths
        .slice(columnCursor, columnCursor + span)
        .reduce((sum, value) => sum + Math.max(0, value), 0);
      columnCursor += span;

      const actual = cell.style?.widthTwips;
      if (!actual || actual <= 0 || expected <= 0) {
        continue;
      }
      measuredCells += 1;
      if (Math.abs(actual - expected) / expected > 0.2) {
        conflictCells += 1;
      }
    }

    if (measuredCells > 0) {
      measuredRows += 1;
      if (conflictCells * 2 > measuredCells) {
        conflictRows += 1;
      }
    }
  }

  return measuredRows > 0 && conflictRows * 2 > measuredRows;
}

function deriveColumnWidthsFromTableRows(
  table: TableNode,
  columnCount: number
): number[] | undefined {
  let bestCandidate: number[] | undefined;
  let bestPositiveCount = -1;
  let bestTotalWidth = -1;

  for (const row of table.rows) {
    const candidate: number[] = [];
    for (const cell of row.cells) {
      const span =
        cell.style?.gridSpan && cell.style.gridSpan > 1
          ? cell.style.gridSpan
          : 1;
      const cellWidth = cell.style?.widthTwips;

      if (cellWidth && cellWidth > 0) {
        const perColumn = cellWidth / span;
        for (let index = 0; index < span; index += 1) {
          candidate.push(perColumn);
        }
        continue;
      }

      for (let index = 0; index < span; index += 1) {
        candidate.push(0);
      }
    }

    if (candidate.length !== columnCount || candidate.length === 0) {
      continue;
    }

    const positiveCount = candidate.filter((value) => value > 0).length;
    if (positiveCount <= 0) {
      continue;
    }

    const totalWidth = candidate.reduce(
      (sum, value) => sum + (value > 0 ? value : 0),
      0
    );
    if (
      positiveCount > bestPositiveCount ||
      (positiveCount === bestPositiveCount && totalWidth > bestTotalWidth)
    ) {
      bestCandidate = candidate;
      bestPositiveCount = positiveCount;
      bestTotalWidth = totalWidth;
    }
  }

  if (!bestCandidate) {
    return undefined;
  }

  return normalizeColumnWidthsTwips(bestCandidate, columnCount);
}

function normalizeColumnWidthsTwips(
  widths: number[],
  columnCount: number
): number[] {
  const fallback = 1440 / Math.max(1, columnCount);
  const sanitized = Array.from({ length: columnCount }, (_, index) => {
    const raw = widths[index];
    if (!Number.isFinite(raw)) {
      return fallback;
    }
    return Math.max(1, raw);
  });

  if (sanitized.every((value) => value <= 0)) {
    return Array.from({ length: columnCount }, () => fallback);
  }

  return sanitized;
}

function normalizeColumnWidthsPx(
  widths: number[],
  columnCount: number,
  fallbackTableWidthPx?: number,
  minimumWidthPx = 24
): number[] {
  const fallbackWidth =
    Number.isFinite(fallbackTableWidthPx) &&
    (fallbackTableWidthPx as number) > 0
      ? (fallbackTableWidthPx as number) / Math.max(1, columnCount)
      : 140;

  return Array.from({ length: columnCount }, (_, index) => {
    const raw = widths[index];
    if (!Number.isFinite(raw) || (raw as number) <= 0) {
      return Math.max(minimumWidthPx, Math.round(fallbackWidth));
    }
    return Math.max(minimumWidthPx, Math.round(raw as number));
  });
}

function defaultColumnWidthsPx(
  columnCount: number,
  tableWidthPx?: number
): number[] {
  const fallbackWidth =
    Number.isFinite(tableWidthPx) && (tableWidthPx as number) > 0
      ? (tableWidthPx as number) / Math.max(1, columnCount)
      : 140;
  return Array.from({ length: columnCount }, () =>
    Math.max(24, Math.round(fallbackWidth))
  );
}

function clampTableWidthPx(widthPx: number, maxWidthPx?: number): number {
  if (!Number.isFinite(widthPx) || widthPx <= 0) {
    return 1;
  }

  if (!Number.isFinite(maxWidthPx) || (maxWidthPx as number) <= 0) {
    return Math.max(1, Math.round(widthPx));
  }

  return Math.max(
    1,
    Math.min(Math.round(widthPx), Math.round(maxWidthPx as number))
  );
}

function fitColumnWidthsToWidth(
  columnWidths: number[],
  targetWidthPx: number
): number[] {
  if (columnWidths.length === 0) {
    return [];
  }

  if (!Number.isFinite(targetWidthPx) || targetWidthPx <= 0) {
    return [...columnWidths];
  }

  const sanitized = columnWidths.map((value) =>
    Number.isFinite(value) && (value as number) > 0 ? (value as number) : 1
  );
  const currentTotal = sanitized.reduce((sum, value) => sum + value, 0);
  if (!Number.isFinite(currentTotal) || currentTotal <= 0) {
    const even = Math.max(1, targetWidthPx / sanitized.length);
    return Array.from({ length: sanitized.length }, () => even);
  }

  if (Math.abs(currentTotal - targetWidthPx) <= 0.5) {
    return sanitized;
  }

  if (currentTotal < targetWidthPx) {
    const scale = targetWidthPx / currentTotal;
    return sanitized.map((value) => Math.max(1, value * scale));
  }

  const minimumWidthPx = 8;
  if (sanitized.length * minimumWidthPx >= targetWidthPx) {
    const even = Math.max(1, targetWidthPx / sanitized.length);
    return Array.from({ length: sanitized.length }, () => even);
  }

  const scaled = sanitized.map((value) =>
    Math.max(minimumWidthPx, (value / currentTotal) * targetWidthPx)
  );
  let overflow = scaled.reduce((sum, value) => sum + value, 0) - targetWidthPx;
  let guard = 0;

  while (overflow > 0.25 && guard < 64) {
    const adjustableIndexes = scaled
      .map((value, index) => ({ value, index }))
      .filter((entry) => entry.value > minimumWidthPx + 0.01);
    if (adjustableIndexes.length === 0) {
      break;
    }

    const adjustableTotal = adjustableIndexes.reduce(
      (sum, entry) => sum + (entry.value - minimumWidthPx),
      0
    );
    if (adjustableTotal <= 0) {
      break;
    }

    for (const entry of adjustableIndexes) {
      const share =
        ((entry.value - minimumWidthPx) / adjustableTotal) * overflow;
      scaled[entry.index] = Math.max(minimumWidthPx, entry.value - share);
    }

    overflow = scaled.reduce((sum, value) => sum + value, 0) - targetWidthPx;
    guard += 1;
  }

  return scaled;
}

function rowGridSpanCount(
  row: TableNode["rows"][number],
  maxColumnCount: number
): number {
  const span = row.cells.reduce((total, cell) => {
    const cellSpan =
      cell.style?.gridSpan && cell.style.gridSpan > 1 ? cell.style.gridSpan : 1;
    return total + cellSpan;
  }, 0);

  return Math.max(0, Math.min(maxColumnCount, span));
}

function resolveFittedTableColumnWidths(
  table: TableNode,
  rawColumnWidthsPx: number[],
  targetWidthPx: number
): {
  columnWidthsPx: number[];
  effectiveColumnCount: number;
} {
  const columnCount = rawColumnWidthsPx.length;
  if (columnCount === 0) {
    return {
      columnWidthsPx: [],
      effectiveColumnCount: 0,
    };
  }

  const fallback = fitColumnWidthsToWidth(rawColumnWidthsPx, targetWidthPx);
  const rawTotalWidthPx = rawColumnWidthsPx.reduce(
    (sum, widthPx) => sum + widthPx,
    0
  );
  if (
    table.style?.layout !== "fixed" ||
    table.rows.length < 2 ||
    !Number.isFinite(rawTotalWidthPx) ||
    rawTotalWidthPx <= 0 ||
    !Number.isFinite(targetWidthPx) ||
    targetWidthPx <= 0 ||
    rawTotalWidthPx <= targetWidthPx + 0.5
  ) {
    return {
      columnWidthsPx: fallback,
      effectiveColumnCount: columnCount,
    };
  }

  const rowSpanCounts = table.rows.map((row) =>
    rowGridSpanCount(row, columnCount)
  );
  const spanFrequency = new Map<number, number>();
  rowSpanCounts.forEach((spanCount) => {
    if (spanCount <= 0 || spanCount >= columnCount) {
      return;
    }
    spanFrequency.set(spanCount, (spanFrequency.get(spanCount) ?? 0) + 1);
  });

  if (spanFrequency.size === 0) {
    return {
      columnWidthsPx: fallback,
      effectiveColumnCount: columnCount,
    };
  }

  let dominantSpanCount = columnCount;
  let dominantSpanFrequency = 0;
  for (const [spanCount, frequency] of spanFrequency.entries()) {
    if (
      frequency > dominantSpanFrequency ||
      (frequency === dominantSpanFrequency && spanCount > dominantSpanCount)
    ) {
      dominantSpanCount = spanCount;
      dominantSpanFrequency = frequency;
    }
  }

  const dominantCoverage =
    dominantSpanFrequency / Math.max(1, rowSpanCounts.length);
  if (
    dominantCoverage < 0.6 ||
    dominantSpanCount <= 0 ||
    dominantSpanCount >= columnCount
  ) {
    return {
      columnWidthsPx: fallback,
      effectiveColumnCount: columnCount,
    };
  }

  const trailingColumnWidthsPx = rawColumnWidthsPx.slice(dominantSpanCount);
  const trailingWidthPx = trailingColumnWidthsPx.reduce(
    (sum, widthPx) => sum + widthPx,
    0
  );
  const trailingRatio = trailingWidthPx / rawTotalWidthPx;
  if (
    !Number.isFinite(trailingWidthPx) ||
    trailingWidthPx < 24 ||
    trailingRatio < 0.12
  ) {
    return {
      columnWidthsPx: fallback,
      effectiveColumnCount: columnCount,
    };
  }

  const leadingWidthsPx = rawColumnWidthsPx.slice(0, dominantSpanCount);
  const fittedLeadingWidthsPx = fitColumnWidthsToWidth(
    leadingWidthsPx,
    targetWidthPx
  );
  return {
    columnWidthsPx: [
      ...fittedLeadingWidthsPx,
      ...Array.from({ length: columnCount - dominantSpanCount }, () => 0),
    ],
    effectiveColumnCount: dominantSpanCount,
  };
}

function resolveTableSpacingPaddingPx(spacing?: TableSpacingTwips): {
  top: number;
  right: number;
  bottom: number;
  left: number;
} {
  return {
    top:
      twipsToPixels(spacing?.topTwips) ??
      WORD_TABLE_CELL_FALLBACK_PADDING_PX.top,
    right:
      twipsToPixels(spacing?.rightTwips) ??
      WORD_TABLE_CELL_FALLBACK_PADDING_PX.right,
    bottom:
      twipsToPixels(spacing?.bottomTwips) ??
      WORD_TABLE_CELL_FALLBACK_PADDING_PX.bottom,
    left:
      twipsToPixels(spacing?.leftTwips) ??
      WORD_TABLE_CELL_FALLBACK_PADDING_PX.left,
  };
}

function tableSpacingPaddingStyle(
  spacing?: TableSpacingTwips
): Record<string, string | number | undefined> {
  const { top, right, bottom, left } = resolveTableSpacingPaddingPx(spacing);

  return {
    paddingTop: top,
    paddingRight: right,
    paddingBottom: bottom,
    paddingLeft: left,
  };
}

function mergeTableSpacing(
  baseSpacing?: TableSpacingTwips,
  overrideSpacing?: TableSpacingTwips
): TableSpacingTwips | undefined {
  if (!baseSpacing && !overrideSpacing) {
    return undefined;
  }

  return {
    topTwips: overrideSpacing?.topTwips ?? baseSpacing?.topTwips,
    rightTwips: overrideSpacing?.rightTwips ?? baseSpacing?.rightTwips,
    bottomTwips: overrideSpacing?.bottomTwips ?? baseSpacing?.bottomTwips,
    leftTwips: overrideSpacing?.leftTwips ?? baseSpacing?.leftTwips,
  };
}

type TableBorderSide = "top" | "right" | "bottom" | "left";

function normalizeBorderType(type: string | undefined): string | undefined {
  if (!type) {
    return undefined;
  }

  const normalized = type.trim().toLowerCase();
  return normalized.length > 0 ? normalized : undefined;
}

function tableBorderToCss(
  border: TableBorderStyle | undefined
): string | undefined {
  const type = normalizeBorderType(border?.type);
  if (!type) {
    return undefined;
  }

  if (type === "none" || type === "nil") {
    return "none";
  }

  const cssStyle =
    type === "double"
      ? "double"
      : type === "dashed" ||
        type === "dashsmallgap" ||
        type === "dotdash" ||
        type === "dotdotdash"
      ? "dashed"
      : type === "dotted"
      ? "dotted"
      : "solid";
  const sizeEighthPt = border?.sizeEighthPt;
  const widthPx =
    Number.isFinite(sizeEighthPt) && (sizeEighthPt as number) > 0
      ? Math.max(0.5, Number(((sizeEighthPt as number) / 6).toFixed(2)))
      : 1;
  const color = border?.color ?? "#000000";

  return `${widthPx}px ${cssStyle} ${color}`;
}

function tableBorderStrokeWidthPx(
  border: TableBorderStyle | undefined
): number {
  const type = normalizeBorderType(border?.type);
  if (!type || type === "none" || type === "nil") {
    return 0;
  }

  const sizeEighthPt = border?.sizeEighthPt;
  return Number.isFinite(sizeEighthPt) && (sizeEighthPt as number) > 0
    ? Math.max(0.5, Number(((sizeEighthPt as number) / 6).toFixed(2)))
    : 1;
}

function borderTypeVisible(type: string | undefined): boolean {
  const normalizedType = normalizeBorderType(type);
  return Boolean(
    normalizedType && normalizedType !== "none" && normalizedType !== "nil"
  );
}

function paragraphBorderVisible(
  border: ParagraphBorderStyle | undefined
): boolean {
  return borderTypeVisible(border?.type);
}

function tableBorderVisible(border: TableBorderStyle | undefined): boolean {
  return borderTypeVisible(border?.type);
}

function tableBorderSetHasVisibleEdges(
  borders: TableBorderSet | undefined
): boolean {
  return Boolean(
    tableBorderVisible(borders?.top) ||
      tableBorderVisible(borders?.right) ||
      tableBorderVisible(borders?.bottom) ||
      tableBorderVisible(borders?.left)
  );
}

function tableUsesSeparateBorderModel(table: TableNode): boolean {
  const explicitCellSpacingPx = twipsToPixels(table.style?.cellSpacingTwips);
  if (
    Number.isFinite(explicitCellSpacingPx) &&
    (explicitCellSpacingPx as number) > 0
  ) {
    return true;
  }

  const tableBorders = table.style?.borders;
  if (
    tableBorderVisible(tableBorders?.insideH) ||
    tableBorderVisible(tableBorders?.insideV)
  ) {
    return false;
  }

  return table.rows.some((row) =>
    row.cells.some((cell) => tableBorderSetHasVisibleEdges(cell.style?.borders))
  );
}

function resolveTableSeparateBorderSpacingPx(table: TableNode): number {
  const explicitCellSpacingPx = twipsToPixels(table.style?.cellSpacingTwips);
  if (
    Number.isFinite(explicitCellSpacingPx) &&
    (explicitCellSpacingPx as number) > 0
  ) {
    return Math.max(0, Math.round(explicitCellSpacingPx as number));
  }

  if (
    tableUsesSeparateBorderModel(table) &&
    tableBorderSetHasVisibleEdges(table.style?.borders)
  ) {
    return 1;
  }

  return 0;
}

function tableElementBorderStyle(
  table: TableNode,
  borderSpacingPx = resolveTableSeparateBorderSpacingPx(table)
): Record<string, string | number | undefined> {
  if (!tableUsesSeparateBorderModel(table)) {
    return {
      borderCollapse: "collapse",
    };
  }

  return {
    borderCollapse: "separate",
    borderSpacing: `${Math.max(0, Math.round(borderSpacingPx))}px`,
  };
}

function resolvePreferredParagraphBorder(
  borders: ParagraphBorderSet | undefined
): ParagraphBorderStyle | undefined {
  if (!borders) {
    return undefined;
  }

  return (
    borders.top ??
    borders.right ??
    borders.bottom ??
    borders.left ??
    borders.between ??
    borders.bar
  );
}

function resolvePreferredTableBorder(
  borders: TableBorderSet | undefined
): TableBorderStyle | undefined {
  if (!borders) {
    return undefined;
  }

  return (
    borders.top ??
    borders.right ??
    borders.bottom ??
    borders.left ??
    borders.insideH ??
    borders.insideV ??
    borders.tl2br ??
    borders.tr2bl
  );
}

function toolbarParagraphBorderStyle(
  seed: ParagraphBorderStyle | undefined
): ParagraphBorderStyle {
  return {
    type: borderTypeVisible(seed?.type) ? (seed?.type as string) : "single",
    sizeEighthPt:
      Number.isFinite(seed?.sizeEighthPt) && (seed?.sizeEighthPt as number) > 0
        ? Math.round(seed?.sizeEighthPt as number)
        : DEFAULT_TOOLBAR_BORDER_SIZE_EIGHTH_PT,
    color: seed?.color ?? DEFAULT_TOOLBAR_BORDER_COLOR,
    ...(Number.isFinite(seed?.spacePt)
      ? { spacePt: Math.max(0, Math.round(seed?.spacePt as number)) }
      : undefined),
  };
}

function toolbarTableBorderStyle(
  seed: TableBorderStyle | undefined
): TableBorderStyle {
  return {
    type: borderTypeVisible(seed?.type) ? (seed?.type as string) : "single",
    sizeEighthPt:
      Number.isFinite(seed?.sizeEighthPt) && (seed?.sizeEighthPt as number) > 0
        ? Math.round(seed?.sizeEighthPt as number)
        : DEFAULT_TOOLBAR_BORDER_SIZE_EIGHTH_PT,
    color: seed?.color ?? DEFAULT_TOOLBAR_BORDER_COLOR,
  };
}

function nilParagraphBorderStyle(): ParagraphBorderStyle {
  return { type: "nil" };
}

function nilTableBorderStyle(): TableBorderStyle {
  return { type: "nil" };
}

function paragraphBorderPresetState(
  borders: ParagraphBorderSet | undefined
): DocxBorderPresetState {
  const top = paragraphBorderVisible(borders?.top);
  const right = paragraphBorderVisible(borders?.right);
  const bottom = paragraphBorderVisible(borders?.bottom);
  const left = paragraphBorderVisible(borders?.left);
  const between = paragraphBorderVisible(borders?.between);
  const bar = paragraphBorderVisible(borders?.bar);
  const hasAny = top || right || bottom || left || between || bar;

  return {
    bottom,
    top,
    left,
    right,
    none: !hasAny,
    all: top && right && bottom && left,
    outside: top && right && bottom && left,
    inside: between || bar,
    "inside-horizontal": between,
    "inside-vertical": bar,
    "diagonal-down": false,
    "diagonal-up": false,
    "horizontal-line": bottom,
  };
}

function paragraphRangePresetActive(
  preset: DocxBorderPreset,
  bordersByParagraph: Array<ParagraphBorderSet | undefined>
): boolean {
  if (bordersByParagraph.length === 0) {
    return false;
  }

  const firstBorders = bordersByParagraph[0];
  const lastBorders = bordersByParagraph[bordersByParagraph.length - 1];
  const allLeft = bordersByParagraph.every((borders) =>
    paragraphBorderVisible(borders?.left)
  );
  const allRight = bordersByParagraph.every((borders) =>
    paragraphBorderVisible(borders?.right)
  );
  const allTop = bordersByParagraph.every((borders) =>
    paragraphBorderVisible(borders?.top)
  );
  const allBottom = bordersByParagraph.every((borders) =>
    paragraphBorderVisible(borders?.bottom)
  );
  const allBetween = bordersByParagraph.every((borders) =>
    paragraphBorderVisible(borders?.between)
  );
  const allBar = bordersByParagraph.every((borders) =>
    paragraphBorderVisible(borders?.bar)
  );

  switch (preset) {
    case "top":
      return paragraphBorderVisible(firstBorders?.top);
    case "bottom":
      return paragraphBorderVisible(lastBorders?.bottom);
    case "left":
      return allLeft;
    case "right":
      return allRight;
    case "all":
    case "outside":
      return (
        paragraphBorderVisible(firstBorders?.top) &&
        paragraphBorderVisible(lastBorders?.bottom) &&
        allLeft &&
        allRight
      );
    case "none":
      return bordersByParagraph.every(
        (borders) => paragraphBorderPresetState(borders).none
      );
    case "inside":
      return allBetween && allBar;
    case "inside-horizontal":
      return allBetween;
    case "inside-vertical":
      return allBar;
    case "horizontal-line":
      return paragraphBorderVisible(lastBorders?.bottom);
    case "diagonal-down":
    case "diagonal-up":
      return false;
    default:
      return allTop && allBottom;
  }
}

function applyParagraphBorderPresetForRangeEntry(
  borders: ParagraphBorderSet | undefined,
  preset: DocxBorderPreset,
  remove: boolean,
  index: number,
  total: number
): ParagraphBorderSet | undefined {
  if (total <= 1) {
    return applyParagraphBorderPreset(borders, preset, remove);
  }

  const nextBorders = cloneParagraphBorderSet(borders) ?? {};
  const visibleBorder = toolbarParagraphBorderStyle(
    resolvePreferredParagraphBorder(nextBorders)
  );
  const nilBorder = nilParagraphBorderStyle();
  const borderToApply = remove ? nilBorder : visibleBorder;
  const isFirst = index === 0;
  const isLast = index === total - 1;

  switch (preset) {
    case "top":
      nextBorders.top = { ...(isFirst ? borderToApply : nilBorder) };
      return nextBorders;
    case "bottom":
      nextBorders.bottom = { ...(isLast ? borderToApply : nilBorder) };
      return nextBorders;
    case "left":
      nextBorders.left = { ...borderToApply };
      return nextBorders;
    case "right":
      nextBorders.right = { ...borderToApply };
      return nextBorders;
    case "all":
    case "outside":
      if (remove) {
        nextBorders.top = { ...nilBorder };
        nextBorders.right = { ...nilBorder };
        nextBorders.bottom = { ...nilBorder };
        nextBorders.left = { ...nilBorder };
      } else {
        nextBorders.top = { ...(isFirst ? borderToApply : nilBorder) };
        nextBorders.right = { ...borderToApply };
        nextBorders.bottom = { ...(isLast ? borderToApply : nilBorder) };
        nextBorders.left = { ...borderToApply };
      }
      return nextBorders;
    case "horizontal-line":
      nextBorders.bottom = { ...(isLast ? borderToApply : nilBorder) };
      return nextBorders;
    case "none":
      nextBorders.top = { ...nilBorder };
      nextBorders.right = { ...nilBorder };
      nextBorders.bottom = { ...nilBorder };
      nextBorders.left = { ...nilBorder };
      nextBorders.between = { ...nilBorder };
      nextBorders.bar = { ...nilBorder };
      return nextBorders;
    default:
      return applyParagraphBorderPreset(borders, preset, remove);
  }
}

function tableBorderPresetState(
  borders: TableBorderSet | undefined,
  selectedCellBorders?: TableBorderSet
): DocxBorderPresetState {
  const top = tableBorderVisible(borders?.top);
  const right = tableBorderVisible(borders?.right);
  const bottom = tableBorderVisible(borders?.bottom);
  const left = tableBorderVisible(borders?.left);
  const insideH = tableBorderVisible(borders?.insideH);
  const insideV = tableBorderVisible(borders?.insideV);
  const diagonalDown = tableBorderVisible(
    selectedCellBorders?.tl2br ?? borders?.tl2br
  );
  const diagonalUp = tableBorderVisible(
    selectedCellBorders?.tr2bl ?? borders?.tr2bl
  );
  const hasAny =
    top ||
    right ||
    bottom ||
    left ||
    insideH ||
    insideV ||
    diagonalDown ||
    diagonalUp;

  return {
    bottom,
    top,
    left,
    right,
    none: !hasAny,
    all: top && right && bottom && left && insideH && insideV,
    outside: top && right && bottom && left,
    inside: insideH && insideV,
    "inside-horizontal": insideH,
    "inside-vertical": insideV,
    "diagonal-down": diagonalDown,
    "diagonal-up": diagonalUp,
    "horizontal-line": insideH,
  };
}

function applyParagraphBorderPreset(
  borders: ParagraphBorderSet | undefined,
  preset: DocxBorderPreset,
  remove = false
): ParagraphBorderSet | undefined {
  const nextBorders = cloneParagraphBorderSet(borders) ?? {};
  const visibleBorder = toolbarParagraphBorderStyle(
    resolvePreferredParagraphBorder(nextBorders)
  );
  const nilBorder = nilParagraphBorderStyle();
  const borderToApply = remove ? nilBorder : visibleBorder;

  switch (preset) {
    case "top":
      nextBorders.top = { ...borderToApply };
      break;
    case "right":
      nextBorders.right = { ...borderToApply };
      break;
    case "bottom":
      nextBorders.bottom = { ...borderToApply };
      break;
    case "left":
      nextBorders.left = { ...borderToApply };
      break;
    case "all":
    case "outside":
      nextBorders.top = { ...borderToApply };
      nextBorders.right = { ...borderToApply };
      nextBorders.bottom = { ...borderToApply };
      nextBorders.left = { ...borderToApply };
      break;
    case "inside":
      nextBorders.between = { ...borderToApply };
      nextBorders.bar = { ...borderToApply };
      break;
    case "inside-horizontal":
      nextBorders.between = { ...borderToApply };
      break;
    case "inside-vertical":
      nextBorders.bar = { ...borderToApply };
      break;
    case "horizontal-line":
      nextBorders.bottom = { ...borderToApply };
      break;
    case "none":
      nextBorders.top = { ...nilBorder };
      nextBorders.right = { ...nilBorder };
      nextBorders.bottom = { ...nilBorder };
      nextBorders.left = { ...nilBorder };
      nextBorders.between = { ...nilBorder };
      nextBorders.bar = { ...nilBorder };
      break;
    case "diagonal-down":
    case "diagonal-up":
      return undefined;
    default:
      return nextBorders;
  }

  return nextBorders;
}

function applyTableBorderPreset(
  borders: TableBorderSet | undefined,
  preset: DocxBorderPreset,
  remove = false
): TableBorderSet | undefined {
  const nextBorders = cloneTableBorderSet(borders) ?? {};
  const visibleBorder = toolbarTableBorderStyle(
    resolvePreferredTableBorder(nextBorders)
  );
  const nilBorder = nilTableBorderStyle();
  const borderToApply = remove ? nilBorder : visibleBorder;

  switch (preset) {
    case "top":
      nextBorders.top = { ...borderToApply };
      break;
    case "right":
      nextBorders.right = { ...borderToApply };
      break;
    case "bottom":
      nextBorders.bottom = { ...borderToApply };
      break;
    case "left":
      nextBorders.left = { ...borderToApply };
      break;
    case "all":
      nextBorders.top = { ...borderToApply };
      nextBorders.right = { ...borderToApply };
      nextBorders.bottom = { ...borderToApply };
      nextBorders.left = { ...borderToApply };
      nextBorders.insideH = { ...borderToApply };
      nextBorders.insideV = { ...borderToApply };
      break;
    case "outside":
      nextBorders.top = { ...borderToApply };
      nextBorders.right = { ...borderToApply };
      nextBorders.bottom = { ...borderToApply };
      nextBorders.left = { ...borderToApply };
      if (!remove) {
        nextBorders.insideH = { ...nilBorder };
        nextBorders.insideV = { ...nilBorder };
      }
      break;
    case "inside":
      nextBorders.insideH = { ...borderToApply };
      nextBorders.insideV = { ...borderToApply };
      break;
    case "inside-horizontal":
      nextBorders.insideH = { ...borderToApply };
      break;
    case "inside-vertical":
      nextBorders.insideV = { ...borderToApply };
      break;
    case "horizontal-line":
      nextBorders.insideH = { ...borderToApply };
      break;
    case "none":
      nextBorders.top = { ...nilBorder };
      nextBorders.right = { ...nilBorder };
      nextBorders.bottom = { ...nilBorder };
      nextBorders.left = { ...nilBorder };
      nextBorders.insideH = { ...nilBorder };
      nextBorders.insideV = { ...nilBorder };
      nextBorders.tl2br = { ...nilBorder };
      nextBorders.tr2bl = { ...nilBorder };
      break;
    case "diagonal-down":
      nextBorders.tl2br = { ...borderToApply };
      break;
    case "diagonal-up":
      nextBorders.tr2bl = { ...borderToApply };
      break;
    default:
      return nextBorders;
  }

  return nextBorders;
}

function resolveTableBorder(
  tableBorders: TableBorderSet | undefined,
  cellBorders: TableBorderSet | undefined,
  side: TableBorderSide,
  rowIndex: number,
  rowCount: number,
  startColumnIndex: number,
  endColumnIndex: number,
  columnCount: number
): TableBorderStyle | undefined {
  const directCellBorder = cellBorders?.[side];
  if (directCellBorder) {
    return directCellBorder;
  }

  const isTopRow = rowIndex === 0;
  const isBottomRow = rowIndex >= rowCount - 1;
  const isFirstColumn = startColumnIndex === 0;
  const isLastColumn = endColumnIndex >= columnCount - 1;

  if (side === "top") {
    return isTopRow ? tableBorders?.top : tableBorders?.insideH;
  }
  if (side === "bottom") {
    return isBottomRow ? tableBorders?.bottom : tableBorders?.insideH;
  }
  if (side === "left") {
    return isFirstColumn ? tableBorders?.left : tableBorders?.insideV;
  }

  return isLastColumn ? tableBorders?.right : tableBorders?.insideV;
}

function resolveTableCellBorderCss(
  tableBorders: TableBorderSet | undefined,
  cellBorders: TableBorderSet | undefined,
  rowIndex: number,
  rowCount: number,
  startColumnIndex: number,
  endColumnIndex: number,
  columnCount: number
): Record<string, string | number | undefined> {
  const top = tableBorderToCss(
    resolveTableBorder(
      tableBorders,
      cellBorders,
      "top",
      rowIndex,
      rowCount,
      startColumnIndex,
      endColumnIndex,
      columnCount
    )
  );
  const right = tableBorderToCss(
    resolveTableBorder(
      tableBorders,
      cellBorders,
      "right",
      rowIndex,
      rowCount,
      startColumnIndex,
      endColumnIndex,
      columnCount
    )
  );
  const bottom = tableBorderToCss(
    resolveTableBorder(
      tableBorders,
      cellBorders,
      "bottom",
      rowIndex,
      rowCount,
      startColumnIndex,
      endColumnIndex,
      columnCount
    )
  );
  const left = tableBorderToCss(
    resolveTableBorder(
      tableBorders,
      cellBorders,
      "left",
      rowIndex,
      rowCount,
      startColumnIndex,
      endColumnIndex,
      columnCount
    )
  );

  return {
    ...(top !== undefined ? { borderTop: top } : undefined),
    ...(right !== undefined ? { borderRight: right } : undefined),
    ...(bottom !== undefined ? { borderBottom: bottom } : undefined),
    ...(left !== undefined ? { borderLeft: left } : undefined),
  };
}

function resolveCollapsedTableHorizontalOuterBleedPx(
  table: TableNode,
  columnCount = tableColumnCount(table)
): number {
  if (columnCount <= 0 || tableUsesSeparateBorderModel(table)) {
    return 0;
  }

  const rowCount = table.rows.length;
  const tableBorders = table.style?.borders;
  let maxLeftBorderWidthPx = tableBorderStrokeWidthPx(tableBorders?.left);
  let maxRightBorderWidthPx = tableBorderStrokeWidthPx(tableBorders?.right);

  table.rows.forEach((row, rowIndex) => {
    let columnCursor = 0;

    row.cells.forEach((cell) => {
      const columnSpan =
        cell.style?.gridSpan && cell.style.gridSpan > 1
          ? cell.style.gridSpan
          : 1;
      const startColumnIndex = columnCursor;
      const endColumnIndex = Math.min(
        columnCount - 1,
        startColumnIndex + columnSpan - 1
      );
      columnCursor += columnSpan;

      if (cell.style?.vMergeContinuation) {
        return;
      }

      if (startColumnIndex === 0) {
        maxLeftBorderWidthPx = Math.max(
          maxLeftBorderWidthPx,
          tableBorderStrokeWidthPx(
            resolveTableBorder(
              tableBorders,
              cell.style?.borders,
              "left",
              rowIndex,
              rowCount,
              startColumnIndex,
              endColumnIndex,
              columnCount
            )
          )
        );
      }

      if (endColumnIndex >= columnCount - 1) {
        maxRightBorderWidthPx = Math.max(
          maxRightBorderWidthPx,
          tableBorderStrokeWidthPx(
            resolveTableBorder(
              tableBorders,
              cell.style?.borders,
              "right",
              rowIndex,
              rowCount,
              startColumnIndex,
              endColumnIndex,
              columnCount
            )
          )
        );
      }
    });
  });

  return Math.max(
    0,
    Math.ceil((maxLeftBorderWidthPx + maxRightBorderWidthPx) / 2)
  );
}

function resolveTableCellDiagonalOverlayCss(
  tableBorders: TableBorderSet | undefined,
  cellBorders: TableBorderSet | undefined
): Record<string, string | number | undefined> {
  const diagonalDownBorder = cellBorders?.tl2br ?? tableBorders?.tl2br;
  const diagonalUpBorder = cellBorders?.tr2bl ?? tableBorders?.tr2bl;
  const layers: string[] = [];

  const addLayer = (
    border: TableBorderStyle | undefined,
    direction: "to bottom right" | "to bottom left"
  ): void => {
    if (!tableBorderVisible(border)) {
      return;
    }

    const color = border?.color ?? "#000000";
    const widthPx =
      Number.isFinite(border?.sizeEighthPt) &&
      (border?.sizeEighthPt as number) > 0
        ? Math.max(
            0.75,
            Number(((border?.sizeEighthPt as number) / 6).toFixed(2))
          )
        : 1;
    const halfWidthPx = Number((widthPx / 2).toFixed(2));
    layers.push(
      `linear-gradient(${direction}, transparent calc(50% - ${halfWidthPx}px), ${color} calc(50% - ${halfWidthPx}px), ${color} calc(50% + ${halfWidthPx}px), transparent calc(50% + ${halfWidthPx}px))`
    );
  };

  addLayer(diagonalDownBorder, "to bottom right");
  addLayer(diagonalUpBorder, "to bottom left");

  if (layers.length === 0) {
    return {};
  }

  return {
    backgroundImage: layers.join(", "),
    backgroundRepeat: "no-repeat",
    backgroundSize: "100% 100%",
  };
}

function cloneTextStyle(
  style?: TextRunNode["style"]
): TextRunNode["style"] | undefined {
  return style ? { ...style } : undefined;
}

function cloneFormFieldWidget(
  widget?: FormFieldRunNode["widget"]
): FormFieldRunNode["widget"] | undefined {
  if (!widget) {
    return undefined;
  }

  return {
    name: widget.name,
    enabled: widget.enabled,
    calcOnExit: widget.calcOnExit,
    text: widget.text
      ? {
          inputType: widget.text.inputType,
          defaultText: widget.text.defaultText,
          maxLength: widget.text.maxLength,
          textFormat: widget.text.textFormat,
        }
      : undefined,
    checkbox: widget.checkbox
      ? {
          defaultChecked: widget.checkbox.defaultChecked,
          sizeMode: widget.checkbox.sizeMode,
          sizePt: widget.checkbox.sizePt,
        }
      : undefined,
    dropdown: widget.dropdown
      ? {
          defaultValue: widget.dropdown.defaultValue,
        }
      : undefined,
  };
}

function mergeFormFieldWidgetPatch(
  current: FormFieldRunNode["widget"] | undefined,
  patch: Partial<NonNullable<FormFieldRunNode["widget"]>>
): FormFieldRunNode["widget"] | undefined {
  const hasPatch =
    patch.name !== undefined ||
    patch.enabled !== undefined ||
    patch.calcOnExit !== undefined ||
    patch.text !== undefined ||
    patch.checkbox !== undefined ||
    patch.dropdown !== undefined;
  if (!hasPatch) {
    return cloneFormFieldWidget(current);
  }

  const mergedText =
    patch.text === undefined
      ? cloneFormFieldWidget(current)?.text
      : {
          ...(current?.text ?? {}),
          ...patch.text,
        };
  const mergedCheckbox =
    patch.checkbox === undefined
      ? cloneFormFieldWidget(current)?.checkbox
      : {
          ...(current?.checkbox ?? {}),
          ...patch.checkbox,
        };
  const mergedDropdown =
    patch.dropdown === undefined
      ? cloneFormFieldWidget(current)?.dropdown
      : {
          ...(current?.dropdown ?? {}),
          ...patch.dropdown,
        };

  return {
    ...(current ?? {}),
    ...(patch.name !== undefined ? { name: patch.name } : undefined),
    ...(patch.enabled !== undefined ? { enabled: patch.enabled } : undefined),
    ...(patch.calcOnExit !== undefined
      ? { calcOnExit: patch.calcOnExit }
      : undefined),
    ...(mergedText ? { text: mergedText } : undefined),
    ...(mergedCheckbox ? { checkbox: mergedCheckbox } : undefined),
    ...(mergedDropdown ? { dropdown: mergedDropdown } : undefined),
  };
}

function cloneFormFieldRun(field: FormFieldRunNode): FormFieldRunNode {
  return {
    type: "form-field",
    fieldType: field.fieldType,
    sourceKind: field.sourceKind,
    id: field.id,
    tag: field.tag,
    title: field.title,
    placeholder: field.placeholder,
    checked: field.checked,
    value: field.value,
    options: field.options?.map((option) => ({
      displayText: option.displayText,
      value: option.value,
    })),
    widget: cloneFormFieldWidget(field.widget),
    checkedSymbol: field.checkedSymbol,
    uncheckedSymbol: field.uncheckedSymbol,
    style: cloneTextStyle(field.style),
    link: field.link,
    sourceXml: field.sourceXml,
  };
}

function textStylesEqual(
  a?: TextRunNode["style"],
  b?: TextRunNode["style"]
): boolean {
  return JSON.stringify(a ?? {}) === JSON.stringify(b ?? {});
}

function mergeAdjacentTextRuns(
  children: ParagraphNode["children"]
): ParagraphNode["children"] {
  const merged: ParagraphNode["children"] = [];

  for (const child of children) {
    const previous = merged[merged.length - 1];
    if (
      previous &&
      previous.type === "text" &&
      child.type === "text" &&
      textStylesEqual(previous.style, child.style) &&
      previous.link === child.link
    ) {
      previous.text += child.text;
      continue;
    }

    merged.push(
      child.type === "text"
        ? {
            type: "text",
            text: child.text,
            style: cloneTextStyle(child.style),
            link: child.link,
          }
        : child.type === "form-field"
        ? cloneFormFieldRun(child)
        : {
            type: "image",
            src: child.src,
            alt: child.alt,
            widthPx: child.widthPx,
            heightPx: child.heightPx,
            partName: child.partName,
            contentType: child.contentType,
            data: child.data ? new Uint8Array(child.data) : undefined,
            floating: child.floating ? { ...child.floating } : undefined,
            syntheticTextBox: child.syntheticTextBox,
            textBoxText: child.textBoxText,
          }
    );
  }

  return merged;
}

function paragraphHasOnlyTextRuns(paragraph: ParagraphNode): boolean {
  return paragraph.children.every(
    (child): child is TextRunNode => child.type === "text"
  );
}

function cloneTextRunWithMetadata(run: TextRunNode): TextRunNode {
  return {
    type: "text",
    text: run.text,
    style: cloneTextStyle(run.style),
    link: run.link,
    noteReference: run.noteReference ? { ...run.noteReference } : undefined,
  };
}

function splitTextRunsAtOffset(
  runs: TextRunNode[],
  offset: number
): {
  left: TextRunNode[];
  right: TextRunNode[];
} {
  const safeOffset = Math.max(0, Math.round(offset));
  const left: TextRunNode[] = [];
  const right: TextRunNode[] = [];
  let cursor = 0;

  for (const run of runs) {
    const runLength = run.text.length;
    const runStart = cursor;
    const runEnd = runStart + runLength;
    cursor = runEnd;

    if (runEnd <= safeOffset) {
      left.push(cloneTextRunWithMetadata(run));
      continue;
    }

    if (runStart >= safeOffset) {
      right.push(cloneTextRunWithMetadata(run));
      continue;
    }

    const localSplit = Math.max(0, Math.min(runLength, safeOffset - runStart));
    const before = run.text.slice(0, localSplit);
    const after = run.text.slice(localSplit);
    if (before.length > 0) {
      left.push({
        ...cloneTextRunWithMetadata(run),
        text: before,
      });
    }
    if (after.length > 0) {
      right.push({
        ...cloneTextRunWithMetadata(run),
        text: after,
      });
    }
  }

  return { left, right };
}

function firstTextStyleAtOffset(
  paragraph: ParagraphNode,
  offset: number,
  preferPreviousAtBoundary: boolean
): TextRunNode["style"] | undefined {
  const textChildren = paragraph.children.filter(
    (child): child is TextRunNode => child.type === "text"
  );
  if (textChildren.length === 0) {
    return undefined;
  }

  const safeOffset = Math.max(0, offset);
  let cursor = 0;

  for (let index = 0; index < textChildren.length; index += 1) {
    const run = textChildren[index];
    const runLength = run.text.length;
    const runStart = cursor;
    const runEnd = runStart + runLength;
    cursor = runEnd;

    if (safeOffset < runEnd) {
      return cloneTextStyle(run.style);
    }

    if (safeOffset === runEnd) {
      if (preferPreviousAtBoundary || index === textChildren.length - 1) {
        return cloneTextStyle(run.style);
      }

      return cloneTextStyle(textChildren[index + 1]?.style);
    }
  }

  return cloneTextStyle(textChildren[textChildren.length - 1]?.style);
}

function linkAtOffset(
  paragraph: ParagraphNode,
  offset: number,
  preferPreviousAtBoundary: boolean
): string | undefined {
  const textChildren = paragraph.children.filter(
    (child): child is TextRunNode => child.type === "text"
  );
  if (textChildren.length === 0) {
    return undefined;
  }

  const safeOffset = Math.max(0, offset);
  let cursor = 0;

  for (let index = 0; index < textChildren.length; index += 1) {
    const run = textChildren[index];
    const runLength = run.text.length;
    const runStart = cursor;
    const runEnd = runStart + runLength;
    cursor = runEnd;

    if (safeOffset < runEnd) {
      return run.link;
    }

    if (safeOffset === runEnd) {
      if (preferPreviousAtBoundary || index === textChildren.length - 1) {
        return run.link;
      }

      return textChildren[index + 1]?.link;
    }
  }

  return textChildren[textChildren.length - 1]?.link;
}

function uniformLinkInRange(
  paragraph: ParagraphNode,
  startOffset: number,
  endOffset: number
): string | undefined {
  const safeStart = Math.max(0, Math.min(startOffset, endOffset));
  const safeEnd = Math.max(safeStart, Math.max(startOffset, endOffset));
  if (safeStart === safeEnd) {
    return linkAtOffset(paragraph, safeStart, true);
  }

  let cursor = 0;
  let candidateLink: string | undefined;

  for (const child of paragraph.children) {
    if (child.type !== "text") {
      continue;
    }

    const runStart = cursor;
    const runEnd = runStart + child.text.length;
    cursor = runEnd;

    if (runEnd <= safeStart || runStart >= safeEnd) {
      continue;
    }

    if (!candidateLink) {
      candidateLink = child.link;
      continue;
    }

    if (candidateLink !== child.link) {
      return undefined;
    }
  }

  return candidateLink;
}

function linkRangeAtOffset(
  paragraph: ParagraphNode,
  offset: number
):
  | {
      start: number;
      end: number;
      link: string;
    }
  | undefined {
  const textChildren = paragraph.children.filter(
    (child): child is TextRunNode => child.type === "text"
  );
  if (textChildren.length === 0) {
    return undefined;
  }

  const safeOffset = Math.max(0, offset);
  let cursor = 0;
  let runIndex = -1;

  for (let index = 0; index < textChildren.length; index += 1) {
    const run = textChildren[index];
    const runStart = cursor;
    const runEnd = runStart + run.text.length;
    cursor = runEnd;

    if (
      safeOffset < runEnd ||
      safeOffset === runEnd ||
      index === textChildren.length - 1
    ) {
      runIndex = index;
      break;
    }
  }

  if (runIndex < 0) {
    return undefined;
  }

  const currentRun = textChildren[runIndex];
  const currentLink = currentRun.link;
  if (!currentLink) {
    return undefined;
  }

  const runStarts: number[] = [];
  cursor = 0;
  for (const run of textChildren) {
    runStarts.push(cursor);
    cursor += run.text.length;
  }

  let startIndex = runIndex;
  while (startIndex > 0 && textChildren[startIndex - 1]?.link === currentLink) {
    startIndex -= 1;
  }

  let endIndex = runIndex;
  while (
    endIndex + 1 < textChildren.length &&
    textChildren[endIndex + 1]?.link === currentLink
  ) {
    endIndex += 1;
  }

  const start = runStarts[startIndex] ?? 0;
  const end =
    (runStarts[endIndex] ?? 0) + (textChildren[endIndex]?.text.length ?? 0);
  if (end <= start) {
    return undefined;
  }

  return {
    start,
    end,
    link: currentLink,
  };
}

function mutateParagraphTextStyleInRange(
  paragraph: ParagraphNode,
  startOffset: number,
  endOffset: number,
  mutator: (
    currentStyle: TextRunNode["style"] | undefined
  ) => TextRunNode["style"] | undefined
): boolean {
  const safeStart = Math.max(0, Math.min(startOffset, endOffset));
  const safeEnd = Math.max(safeStart, Math.max(startOffset, endOffset));
  if (safeStart === safeEnd) {
    return false;
  }

  const nextChildren: ParagraphNode["children"] = [];
  let cursor = 0;
  let touched = false;

  for (const child of paragraph.children) {
    if (child.type !== "text") {
      if (child.type === "form-field") {
        nextChildren.push(cloneFormFieldRun(child));
      } else {
        nextChildren.push({
          type: "image",
          src: child.src,
          alt: child.alt,
          widthPx: child.widthPx,
          heightPx: child.heightPx,
          partName: child.partName,
          contentType: child.contentType,
          data: child.data ? new Uint8Array(child.data) : undefined,
          floating: child.floating ? { ...child.floating } : undefined,
          syntheticTextBox: child.syntheticTextBox,
          textBoxText: child.textBoxText,
        });
      }
      continue;
    }

    const text = child.text;
    const runStart = cursor;
    const runEnd = runStart + text.length;
    cursor = runEnd;

    if (runEnd <= safeStart || runStart >= safeEnd || text.length === 0) {
      nextChildren.push({
        type: "text",
        text,
        style: cloneTextStyle(child.style),
        link: child.link,
      });
      continue;
    }

    touched = true;
    const localStart = Math.max(0, safeStart - runStart);
    const localEnd = Math.min(text.length, safeEnd - runStart);

    if (localStart > 0) {
      nextChildren.push({
        type: "text",
        text: text.slice(0, localStart),
        style: cloneTextStyle(child.style),
        link: child.link,
      });
    }

    const selectedText = text.slice(localStart, localEnd);
    if (selectedText.length > 0) {
      nextChildren.push({
        type: "text",
        text: selectedText,
        style: mutator(cloneTextStyle(child.style)),
        link: child.link,
      });
    }

    if (localEnd < text.length) {
      nextChildren.push({
        type: "text",
        text: text.slice(localEnd),
        style: cloneTextStyle(child.style),
        link: child.link,
      });
    }
  }

  if (!touched) {
    return false;
  }

  paragraph.children = mergeAdjacentTextRuns(nextChildren);
  return true;
}

function mutateParagraphLinkInRange(
  paragraph: ParagraphNode,
  startOffset: number,
  endOffset: number,
  link?: string
): boolean {
  const safeStart = Math.max(0, Math.min(startOffset, endOffset));
  const safeEnd = Math.max(safeStart, Math.max(startOffset, endOffset));
  if (safeStart === safeEnd) {
    return false;
  }

  const nextChildren: ParagraphNode["children"] = [];
  let cursor = 0;
  let touched = false;

  for (const child of paragraph.children) {
    if (child.type !== "text") {
      if (child.type === "form-field") {
        nextChildren.push(cloneFormFieldRun(child));
      } else {
        nextChildren.push({
          type: "image",
          src: child.src,
          alt: child.alt,
          widthPx: child.widthPx,
          heightPx: child.heightPx,
          partName: child.partName,
          contentType: child.contentType,
          data: child.data ? new Uint8Array(child.data) : undefined,
          floating: child.floating ? { ...child.floating } : undefined,
          syntheticTextBox: child.syntheticTextBox,
          textBoxText: child.textBoxText,
        });
      }
      continue;
    }

    const text = child.text;
    const runStart = cursor;
    const runEnd = runStart + text.length;
    cursor = runEnd;

    if (runEnd <= safeStart || runStart >= safeEnd || text.length === 0) {
      nextChildren.push({
        type: "text",
        text,
        style: cloneTextStyle(child.style),
        link: child.link,
      });
      continue;
    }

    touched = true;
    const localStart = Math.max(0, safeStart - runStart);
    const localEnd = Math.min(text.length, safeEnd - runStart);

    if (localStart > 0) {
      nextChildren.push({
        type: "text",
        text: text.slice(0, localStart),
        style: cloneTextStyle(child.style),
        link: child.link,
      });
    }

    const selectedText = text.slice(localStart, localEnd);
    if (selectedText.length > 0) {
      nextChildren.push({
        type: "text",
        text: selectedText,
        style: cloneTextStyle(child.style),
        link,
      });
    }

    if (localEnd < text.length) {
      nextChildren.push({
        type: "text",
        text: text.slice(localEnd),
        style: cloneTextStyle(child.style),
        link: child.link,
      });
    }
  }

  if (!touched) {
    return false;
  }

  paragraph.children = mergeAdjacentTextRuns(nextChildren);
  return true;
}

function toBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(bytes).toString("base64");
  }

  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.byteLength; index += chunkSize) {
    const chunk = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode(...chunk);
  }

  if (typeof btoa === "function") {
    return btoa(binary);
  }

  return "";
}

function isParagraphSelected(
  selection: DocxEditorSelection,
  nodeIndex: number
): boolean {
  return selection.kind === "paragraph" && selection.nodeIndex === nodeIndex;
}

function isCellSelected(
  selection: DocxEditorSelection,
  tableIndex: number,
  rowIndex: number,
  cellIndex: number
): boolean {
  return (
    selection.kind === "table-cell" &&
    selection.tableIndex === tableIndex &&
    selection.rowIndex === rowIndex &&
    selection.cellIndex === cellIndex
  );
}

function paragraphLocationKey(location: ParagraphLocation): string {
  if (location.kind === "paragraph") {
    return `p:${location.nodeIndex}`;
  }

  return `t:${location.tableIndex}:${location.rowIndex}:${location.cellIndex}:${location.paragraphIndex}`;
}

function imageLocationKey(location: DocxImageLocation): string {
  return `${paragraphLocationKey(location)}:${location.childIndex}`;
}

function dropTargetKey(target: DocxImageDropTarget): string {
  return `${paragraphLocationKey(target)}:${target.childIndex}`;
}

function parseImageDropTargetFromDataset(
  dataset: DOMStringMap
): DocxImageDropTarget | undefined {
  const kind = dataset.docxTargetKind;
  const childIndex = Number.parseInt(dataset.docxChildIndex ?? "", 10);
  if (!Number.isFinite(childIndex) || childIndex < 0) {
    return undefined;
  }

  if (kind === "paragraph") {
    const nodeIndex = Number.parseInt(dataset.docxNodeIndex ?? "", 10);
    if (!Number.isFinite(nodeIndex) || nodeIndex < 0) {
      return undefined;
    }

    return {
      kind: "paragraph",
      nodeIndex,
      childIndex,
    };
  }

  if (kind === "table-cell") {
    const tableIndex = Number.parseInt(dataset.docxTableIndex ?? "", 10);
    const rowIndex = Number.parseInt(dataset.docxRowIndex ?? "", 10);
    const cellIndex = Number.parseInt(dataset.docxCellIndex ?? "", 10);
    const paragraphIndex = Number.parseInt(
      dataset.docxParagraphIndex ?? "",
      10
    );
    if (
      !Number.isFinite(tableIndex) ||
      !Number.isFinite(rowIndex) ||
      !Number.isFinite(cellIndex) ||
      !Number.isFinite(paragraphIndex) ||
      tableIndex < 0 ||
      rowIndex < 0 ||
      cellIndex < 0 ||
      paragraphIndex < 0
    ) {
      return undefined;
    }

    return {
      kind: "table-cell",
      tableIndex,
      rowIndex,
      cellIndex,
      paragraphIndex,
      childIndex,
    };
  }

  return undefined;
}

function firstTableCellAnchorLocation(
  table: TableNode,
  tableIndex: number
): Extract<DocxTextRangeLocation, { kind: "table-cell" }> | undefined {
  for (let rowIndex = 0; rowIndex < table.rows.length; rowIndex += 1) {
    const row = table.rows[rowIndex];
    if (!row) {
      continue;
    }

    for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex += 1) {
      const cell = row.cells[cellIndex];
      if (!cell) {
        continue;
      }

      return {
        kind: "table-cell",
        tableIndex,
        rowIndex,
        cellIndex,
        paragraphIndex: 0,
      };
    }
  }

  return undefined;
}

function collectTablePropertyTrackedChanges(
  table: TableNode,
  tableIndex: number
): Array<{
  stableId: string;
  kind: DocxTrackedChangeKind;
  author?: string;
  date?: string;
  text?: string;
  location: DocxTextRangeLocation;
}> {
  const sourceXml = table.sourceXml ?? "";
  if (!sourceXml) {
    return [];
  }

  const anchorLocation = firstTableCellAnchorLocation(table, tableIndex);
  if (!anchorLocation) {
    return [];
  }

  const entries: Array<{
    stableId: string;
    kind: DocxTrackedChangeKind;
    author?: string;
    date?: string;
    text?: string;
    location: DocxTextRangeLocation;
  }> = [];
  const entryByKey = new Map<string, (typeof entries)[number]>();

  const append = (
    scope: "table" | "row" | "cell",
    changeTag: XmlBalancedTagRange
  ): void => {
    const kind: DocxTrackedChangeKind = "paragraph-format-change";
    const author =
      decodeXmlText(xmlAttribute(changeTag.openTag, "w:author") ?? "") ||
      undefined;
    const date = xmlAttribute(changeTag.openTag, "w:date")?.trim() || undefined;
    const revisionId =
      xmlAttribute(changeTag.openTag, "w:id")?.trim() || undefined;
    const changeXml = sourceXml.slice(changeTag.start, changeTag.end);
    const text = summarizeTableFormattingChange(scope, changeXml);
    const key = revisionId
      ? `${scope}:${kind}:id:${revisionId}`
      : `${scope}:${kind}:${author ?? ""}:${date ?? ""}:${text}`;
    const existing = entryByKey.get(key);
    if (existing) {
      if (!existing.text && text) {
        existing.text = text;
      }
      if (!existing.author && author) {
        existing.author = author;
      }
      if (!existing.date && date) {
        existing.date = date;
      }
      return;
    }

    const stableId = revisionId
      ? `${scope}-${kind}-${revisionId}`
      : `${scope}-${kind}-${entryByKey.size}`;
    const next = {
      stableId,
      kind,
      author,
      date,
      text,
      location: {
        kind: "table-cell" as const,
        tableIndex,
        rowIndex: anchorLocation.rowIndex,
        cellIndex: anchorLocation.cellIndex,
        paragraphIndex: anchorLocation.paragraphIndex,
      },
    };
    entries.push(next);
    entryByKey.set(key, next);
  };

  extractBalancedTagRanges(sourceXml, "w:tblPrChange").forEach((range) =>
    append("table", range)
  );
  extractBalancedTagRanges(sourceXml, "w:trPrChange").forEach((range) =>
    append("row", range)
  );
  extractBalancedTagRanges(sourceXml, "w:tcPrChange").forEach((range) =>
    append("cell", range)
  );

  return entries;
}

function collectTrackedChangesFromModel(model: DocModel): DocxTrackedChange[] {
  const trackedChanges: DocxTrackedChange[] = [];

  const appendParagraphChanges = (
    paragraph: ParagraphNode,
    nodeIndex: number,
    location: ParagraphLocation
  ): void => {
    const trackedMarkup = resolveParagraphTrackedMarkup(paragraph);
    if (!trackedMarkup) {
      return;
    }

    trackedMarkup.changes.forEach((change, changeIndex) => {
      trackedChanges.push({
        id: `${paragraphLocationKey(location)}:${change.id}:${changeIndex}`,
        inlineAnchorId: change.id,
        kind: change.kind,
        author: change.author,
        date: change.date,
        text: change.text,
        nodeIndex,
        location:
          location.kind === "paragraph"
            ? { kind: "paragraph", nodeIndex: location.nodeIndex }
            : {
                kind: "table-cell",
                tableIndex: location.tableIndex,
                rowIndex: location.rowIndex,
                cellIndex: location.cellIndex,
                paragraphIndex: location.paragraphIndex,
              },
      });
    });
  };

  model.nodes.forEach((node, nodeIndex) => {
    if (node.type === "paragraph") {
      appendParagraphChanges(node, nodeIndex, {
        kind: "paragraph",
        nodeIndex,
      });
      return;
    }

    node.rows.forEach((row, rowIndex) => {
      row.cells.forEach((cell, cellIndex) => {
        const directParagraphs = tableCellParagraphs(cell.nodes);
        directParagraphs.forEach((paragraph, paragraphIndex) => {
          appendParagraphChanges(paragraph, nodeIndex, {
            kind: "table-cell",
            tableIndex: nodeIndex,
            rowIndex,
            cellIndex,
            paragraphIndex,
          });
        });

        // Nested tables inside a cell are rendered without per-paragraph location
        // attributes, so anchor these changes to the owning cell via a negative
        // paragraph index to avoid colliding with direct paragraph indexes.
        const nestedParagraphs = tableCellParagraphsRecursively(
          cell.nodes
        ).filter((paragraph) => !directParagraphs.includes(paragraph));
        nestedParagraphs.forEach((paragraph, nestedParagraphIndex) => {
          appendParagraphChanges(paragraph, nodeIndex, {
            kind: "table-cell",
            tableIndex: nodeIndex,
            rowIndex,
            cellIndex,
            paragraphIndex: -(nestedParagraphIndex + 1),
          });
        });
      });
    });

    collectTablePropertyTrackedChanges(node, nodeIndex).forEach(
      (change, changeIndex) => {
        trackedChanges.push({
          id: `${paragraphLocationKey(change.location)}:${
            change.stableId
          }:${changeIndex}`,
          kind: change.kind,
          author: change.author,
          date: change.date,
          text: change.text,
          nodeIndex,
          location: change.location,
        });
      }
    );
  });

  return trackedChanges;
}

function decodeCommentRangeText(rangeXml: string): string | undefined {
  const texts: string[] = [];
  for (const match of rangeXml.matchAll(
    /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi
  )) {
    texts.push(decodeXmlText(match[1] ?? ""));
  }
  const combined = texts.join("").replace(/\s+/g, " ").trim();
  if (!combined) {
    return undefined;
  }
  return combined.length > 120 ? `${combined.slice(0, 119)}…` : combined;
}

function resolveCommentAnchorText(
  sourceXml: string,
  commentId: number
): string | undefined {
  const startMatch = sourceXml.match(
    new RegExp(`<w:commentRangeStart\\b[^>]*w:id="${commentId}"[^>]*/?>`, "i")
  );
  const endMatch = sourceXml.match(
    new RegExp(`<w:commentRangeEnd\\b[^>]*w:id="${commentId}"[^>]*/?>`, "i")
  );
  const startIndex =
    startMatch?.index !== undefined
      ? startMatch.index + startMatch[0].length
      : // Range opened in an earlier paragraph: take from the paragraph start.
        endMatch?.index !== undefined
      ? 0
      : undefined;
  if (startIndex === undefined) {
    return undefined;
  }
  const endIndex =
    endMatch?.index !== undefined ? endMatch.index : sourceXml.length;
  if (endIndex <= startIndex) {
    return undefined;
  }
  return decodeCommentRangeText(sourceXml.slice(startIndex, endIndex));
}

function collectCommentsFromModel(model: DocModel): DocxComment[] {
  const definitions = model.metadata.comments ?? [];
  if (definitions.length === 0) {
    return [];
  }
  const definitionById = new Map(
    definitions.map((definition) => [definition.id, definition])
  );

  const comments: DocxComment[] = [];
  const appendParagraphComments = (
    paragraph: ParagraphNode,
    nodeIndex: number,
    location: ParagraphLocation
  ): void => {
    const sourceXml = paragraph.sourceXml ?? "";
    if (!sourceXml || !/commentReference/i.test(sourceXml)) {
      return;
    }

    for (const match of sourceXml.matchAll(
      /<w:commentReference\b[^>]*w:id="(-?\d+)"/gi
    )) {
      const commentId = Number.parseInt(match[1] ?? "", 10);
      const definition = Number.isFinite(commentId)
        ? definitionById.get(commentId)
        : undefined;
      if (!definition) {
        continue;
      }

      comments.push({
        id: `${paragraphLocationKey(location)}:comment:${commentId}`,
        commentId,
        author: definition.author,
        initials: definition.initials,
        date: definition.date,
        text: definition.text,
        parentId: definition.parentId,
        resolved: definition.resolved,
        anchorText: resolveCommentAnchorText(sourceXml, commentId),
        nodeIndex,
        location:
          location.kind === "paragraph"
            ? { kind: "paragraph", nodeIndex: location.nodeIndex }
            : {
                kind: "table-cell",
                tableIndex: location.tableIndex,
                rowIndex: location.rowIndex,
                cellIndex: location.cellIndex,
                paragraphIndex: location.paragraphIndex,
              },
      });
    }
  };

  model.nodes.forEach((node, nodeIndex) => {
    if (node.type === "paragraph") {
      appendParagraphComments(node, nodeIndex, {
        kind: "paragraph",
        nodeIndex,
      });
      return;
    }

    node.rows.forEach((row, rowIndex) => {
      row.cells.forEach((cell, cellIndex) => {
        const directParagraphs = tableCellParagraphs(cell.nodes);
        directParagraphs.forEach((paragraph, paragraphIndex) => {
          appendParagraphComments(paragraph, nodeIndex, {
            kind: "table-cell",
            tableIndex: nodeIndex,
            rowIndex,
            cellIndex,
            paragraphIndex,
          });
        });

        const nestedParagraphs = tableCellParagraphsRecursively(
          cell.nodes
        ).filter((paragraph) => !directParagraphs.includes(paragraph));
        nestedParagraphs.forEach((paragraph, nestedParagraphIndex) => {
          appendParagraphComments(paragraph, nodeIndex, {
            kind: "table-cell",
            tableIndex: nodeIndex,
            rowIndex,
            cellIndex,
            paragraphIndex: -(nestedParagraphIndex + 1),
          });
        });
      });
    });
  });

  return comments;
}

function hexColorWithAlpha(color: string, alpha: number): string {
  const normalized = color.trim().replace(/^#/, "");
  if (!/^[\da-f]{6}$/i.test(normalized)) {
    return color;
  }

  const red = Number.parseInt(normalized.slice(0, 2), 16);
  const green = Number.parseInt(normalized.slice(2, 4), 16);
  const blue = Number.parseInt(normalized.slice(4, 6), 16);
  return `rgba(${red}, ${green}, ${blue}, ${alpha})`;
}

function commentAccentColor(
  documentTheme: DocxDocumentTheme,
  commentId?: number
): string {
  const lightPalette = ["#5b9bd5", "#d65f5f", "#8f6ac8", "#70ad47", "#d9872b"];
  const darkPalette = ["#7dd3fc", "#fca5a5", "#c4b5fd", "#86efac", "#fdba74"];
  if (Number.isFinite(commentId)) {
    const palette = documentTheme === "dark" ? darkPalette : lightPalette;
    return palette[Math.abs(Math.round(commentId as number)) % palette.length];
  }

  return documentTheme === "dark" ? darkPalette[0] : lightPalette[0];
}

function commentHighlightStyle(
  documentTheme: DocxDocumentTheme,
  commentId?: number
): Record<string, string | number | undefined> {
  const accent = commentAccentColor(documentTheme, commentId);
  return {
    backgroundColor: hexColorWithAlpha(
      accent,
      documentTheme === "dark" ? 0.3 : 0.22
    ),
    boxShadow: `inset 0 -1px 0 ${hexColorWithAlpha(
      accent,
      documentTheme === "dark" ? 0.7 : 0.48
    )}`,
  };
}

function estimateCommentCardHeight(comment: DocxComment): number {
  const snippet = comment.text || "Comment";
  const lines = Math.min(2, Math.max(1, Math.ceil(snippet.length / 42)));
  return Math.max(TRACKED_CHANGE_GUTTER_CARD_MIN_HEIGHT_PX, 24 + lines * 11);
}

function trackedChangeKindLabel(kind: DocxTrackedChangeKind): string {
  switch (kind) {
    case "insertion":
      return "Inserted";
    case "deletion":
      return "Deleted";
    case "move-from":
      return "Moved from";
    case "move-to":
      return "Moved to";
    case "format-change":
      return "Formatted";
    case "paragraph-format-change":
      return "Paragraph formatting";
    default:
      return kind;
  }
}

function trackedChangeAccentColor(
  kind: DocxTrackedChangeKind,
  documentTheme: DocxDocumentTheme
): string {
  const palette =
    documentTheme === "dark"
      ? {
          insertion: "#f87171",
          deletion: "#f87171",
          moveFrom: "#86efac",
          moveTo: "#86efac",
          format: "#c084fc",
        }
      : {
          insertion: "#dc2626",
          deletion: "#dc2626",
          moveFrom: "#70ad47",
          moveTo: "#70ad47",
          format: "#7c3aed",
        };

  switch (kind) {
    case "insertion":
      return palette.insertion;
    case "deletion":
      return palette.deletion;
    case "move-from":
      return palette.moveFrom;
    case "move-to":
      return palette.moveTo;
    case "format-change":
    case "paragraph-format-change":
      return palette.format;
    default:
      return palette.format;
  }
}

function trackedChangeUsesGutterBalloon(
  change: DocxTrackedChange
): boolean {
  return change.kind !== "insertion" && change.kind !== "move-to";
}

function gutterAnnotationSortTuple(
  location: DocxTextRangeLocation
): [number, number, number, number] {
  if (location.kind === "paragraph") {
    return [location.nodeIndex, 0, 0, 0];
  }

  return [
    location.tableIndex,
    location.rowIndex,
    location.cellIndex,
    location.paragraphIndex,
  ];
}

function trackedChangeBelongsToPageSegments(
  location: DocxTextRangeLocation,
  pageSegments: DocumentPageNodeSegment[]
): boolean {
  if (location.kind === "paragraph") {
    return pageSegments.some(
      (segment) => segment.nodeIndex === location.nodeIndex
    );
  }

  return pageSegments.some((segment) => {
    if (segment.nodeIndex !== location.tableIndex) {
      return false;
    }

    if (!segment.tableRowRange) {
      return true;
    }

    return (
      location.rowIndex >= segment.tableRowRange.startRowIndex &&
      location.rowIndex < segment.tableRowRange.endRowIndex
    );
  });
}

function resolveGutterAnnotationPageIndex(
  location: DocxTextRangeLocation,
  pageNodeSegmentsByPage: DocumentPageNodeSegment[][]
): number {
  for (
    let pageIndex = 0;
    pageIndex < pageNodeSegmentsByPage.length;
    pageIndex += 1
  ) {
    if (
      trackedChangeBelongsToPageSegments(
        location,
        pageNodeSegmentsByPage[pageIndex] ?? []
      )
    ) {
      return pageIndex;
    }
  }

  return -1;
}

function findTrackedChangeAnchorElementInPage(
  pageElement: HTMLElement,
  location: DocxTextRangeLocation
): HTMLElement | undefined {
  if (location.kind === "paragraph") {
    const paragraphElement = pageElement.querySelector(
      `[data-docx-paragraph-kind="paragraph"][data-docx-paragraph-node-index="${location.nodeIndex}"]`
    ) as HTMLElement | null;
    return paragraphElement ?? undefined;
  }

  const paragraphElement = pageElement.querySelector(
    `[data-docx-paragraph-kind="table-cell"][data-docx-table-index="${location.tableIndex}"][data-docx-row-index="${location.rowIndex}"][data-docx-cell-index="${location.cellIndex}"][data-docx-paragraph-index="${location.paragraphIndex}"]`
  ) as HTMLElement | null;
  if (paragraphElement) {
    return paragraphElement;
  }

  const cellElement = pageElement.querySelector(
    `[data-docx-table-cell="true"][data-docx-table-index="${location.tableIndex}"][data-docx-row-index="${location.rowIndex}"][data-docx-cell-index="${location.cellIndex}"]`
  ) as HTMLElement | null;
  return cellElement ?? undefined;
}

function elementRectWithinContainer(
  element: HTMLElement,
  container: HTMLElement
):
  | {
      left: number;
      width: number;
      right: number;
      top: number;
      height: number;
    }
  | undefined {
  const containerRect = container.getBoundingClientRect();
  const elementRect = element.getBoundingClientRect();
  if (containerRect.height <= 0 || containerRect.width <= 0) {
    return undefined;
  }

  const scaleY = containerRect.height / Math.max(1, container.offsetHeight);
  if (!Number.isFinite(scaleY) || scaleY <= 0) {
    return undefined;
  }

  const scaleX = containerRect.width / Math.max(1, container.offsetWidth);
  if (!Number.isFinite(scaleX) || scaleX <= 0) {
    return undefined;
  }

  const left = (elementRect.left - containerRect.left) / scaleX;
  const width = elementRect.width / scaleX;

  return {
    left,
    width,
    right: left + width,
    top: (elementRect.top - containerRect.top) / scaleY,
    height: elementRect.height / scaleY,
  };
}

interface TrackedChangeAnchorPoint {
  x: number;
  y: number;
}

/**
 * One entry in the page gutter: either a tracked change or a comment. Both
 * share the anchor/stacking pipeline so they interleave in document order.
 */
interface DocxGutterAnnotation {
  id: string;
  location: DocxTextRangeLocation;
  trackedChange?: DocxTrackedChange;
  comment?: DocxComment;
}

function findFirstElementWithSpaceSeparatedDataValue(
  rootElement: HTMLElement,
  attributeName: string,
  value: string
): HTMLElement | undefined {
  const candidateElements: HTMLElement[] = [];
  if (rootElement.hasAttribute(attributeName)) {
    candidateElements.push(rootElement);
  }
  candidateElements.push(
    ...rootElement.querySelectorAll<HTMLElement>(
      `[${attributeName}]`
    )
  );

  for (const candidate of candidateElements) {
    const rawValue = candidate.getAttribute(attributeName);
    if (!rawValue) {
      continue;
    }

    if (rawValue.split(/\s+/).includes(value)) {
      return candidate;
    }
  }

  return undefined;
}

function findGutterAnnotationScopeElementInPage(
  pageElement: HTMLElement,
  annotation: DocxGutterAnnotation
): HTMLElement | undefined {
  return findTrackedChangeAnchorElementInPage(pageElement, annotation.location);
}

function findGutterAnnotationDataAnchorInPage(
  pageElement: HTMLElement,
  annotation: DocxGutterAnnotation,
  attributeName: string,
  value: string
): HTMLElement | undefined {
  const scopedRoot = findGutterAnnotationScopeElementInPage(
    pageElement,
    annotation
  );
  if (scopedRoot) {
    const scopedAnchor = findFirstElementWithSpaceSeparatedDataValue(
      scopedRoot,
      attributeName,
      value
    );
    if (scopedAnchor) {
      return scopedAnchor;
    }
  }

  return findFirstElementWithSpaceSeparatedDataValue(
    pageElement,
    attributeName,
    value
  );
}

function findGutterAnnotationAnchorElementInPage(
  pageElement: HTMLElement,
  annotation: DocxGutterAnnotation
): HTMLElement | undefined {
  if (annotation.trackedChange) {
    const trackedAnchor = findGutterAnnotationDataAnchorInPage(
      pageElement,
      annotation,
      "data-docx-tracked-change-id",
      annotation.trackedChange.id
    );
    if (trackedAnchor) {
      return trackedAnchor;
    }

    const inlineAnchorId = annotation.trackedChange.inlineAnchorId;
    if (inlineAnchorId && inlineAnchorId !== annotation.trackedChange.id) {
      const inlineTrackedAnchor = findGutterAnnotationDataAnchorInPage(
        pageElement,
        annotation,
        "data-docx-tracked-change-id",
        inlineAnchorId
      );
      if (inlineTrackedAnchor) {
        return inlineTrackedAnchor;
      }
    }
  }

  if (annotation.comment) {
    const commentAnchor = findGutterAnnotationDataAnchorInPage(
      pageElement,
      annotation,
      "data-docx-comment-ids",
      String(annotation.comment.commentId)
    );
    if (commentAnchor) {
      return commentAnchor;
    }
  }

  return findTrackedChangeAnchorElementInPage(pageElement, annotation.location);
}

interface PositionedGutterAnnotation {
  annotation: DocxGutterAnnotation;
  anchorX: number;
  anchorY: number;
  top: number;
  heightPx: number;
  connectorLane: number;
}

function estimateTrackedChangeCardHeight(change: DocxTrackedChange): number {
  const snippet =
    normalizeTrackedChangeSnippet(change.text) ??
    trackedChangeKindLabel(change.kind);
  const lines = Math.min(2, Math.max(1, Math.ceil(snippet.length / 42)));
  return Math.max(TRACKED_CHANGE_GUTTER_CARD_MIN_HEIGHT_PX, 22 + lines * 11);
}

function assignGutterConnectorLanes(
  entries: PositionedGutterAnnotation[]
): void {
  if (entries.length <= 1) {
    entries.forEach((entry) => {
      entry.connectorLane = 0;
    });
    return;
  }

  const intervalPaddingPx = 3;
  const laneEndY = Array.from(
    { length: TRACKED_CHANGE_GUTTER_CONNECTOR_LANE_COUNT },
    () => Number.NEGATIVE_INFINITY
  );
  const routedEntries = entries
    .map((entry, index) => {
      const cardCenterY = entry.top + entry.heightPx / 2;
      return {
        entry,
        index,
        startY: Math.min(entry.anchorY, cardCenterY) - intervalPaddingPx,
        endY: Math.max(entry.anchorY, cardCenterY) + intervalPaddingPx,
      };
    })
    .sort((left, right) => {
      const startDelta = left.startY - right.startY;
      if (startDelta !== 0) {
        return startDelta;
      }

      const endDelta = left.endY - right.endY;
      if (endDelta !== 0) {
        return endDelta;
      }

      return left.index - right.index;
    });

  routedEntries.forEach((route) => {
    let laneIndex = laneEndY.findIndex(
      (laneEnd) => route.startY > laneEnd + intervalPaddingPx
    );

    if (laneIndex < 0) {
      laneIndex = laneEndY.reduce(
        (bestLane, laneEnd, index) =>
          laneEnd < laneEndY[bestLane] ? index : bestLane,
        0
      );
    }

    route.entry.connectorLane = laneIndex;
    laneEndY[laneIndex] = Math.max(laneEndY[laneIndex], route.endY);
  });
}

function layoutTrackedChangesForPage(
  annotations: DocxGutterAnnotation[],
  anchorByChangeId: Map<string, TrackedChangeAnchorPoint>,
  cardHeightsByChangeId: Map<string, number> | undefined,
  pageWidthPx: number,
  pageHeightPx: number
): PositionedGutterAnnotation[] {
  if (annotations.length === 0) {
    return [];
  }

  const fallbackStride = Math.max(
    18,
    pageHeightPx / Math.max(1, annotations.length + 1)
  );
  const withAnchors = annotations.map((annotation, index) => {
    const defaultAnchor = Math.min(
      Math.max(10, Math.round((index + 1) * fallbackStride)),
      Math.max(10, pageHeightPx - 10)
    );
    const anchorPoint = anchorByChangeId.get(annotation.id);
    const anchorY = anchorPoint?.y ?? defaultAnchor;
    const anchorX = anchorPoint?.x ?? Math.max(10, pageWidthPx - 10);
    const measuredHeightPx = cardHeightsByChangeId?.get(annotation.id);
    const estimatedHeightPx = annotation.trackedChange
      ? estimateTrackedChangeCardHeight(annotation.trackedChange)
      : annotation.comment
      ? estimateCommentCardHeight(annotation.comment)
      : TRACKED_CHANGE_GUTTER_CARD_MIN_HEIGHT_PX;
    const heightPx =
      Number.isFinite(measuredHeightPx) && (measuredHeightPx as number) > 0
        ? Math.max(
            TRACKED_CHANGE_GUTTER_CARD_MIN_HEIGHT_PX,
            Math.round(measuredHeightPx as number)
          )
        : estimatedHeightPx;
    return {
      annotation,
      anchorX: clampNumber(
        Math.round(anchorX),
        10,
        Math.max(10, pageWidthPx - 10)
      ),
      anchorY: clampNumber(
        Math.round(anchorY),
        10,
        Math.max(10, pageHeightPx - 10)
      ),
      top: 0,
      heightPx,
      connectorLane: 0,
    };
  });

  withAnchors.sort((left, right) => {
    const yDelta = left.anchorY - right.anchorY;
    if (Math.abs(yDelta) > 2) {
      return yDelta;
    }

    const xDelta = left.anchorX - right.anchorX;
    if (xDelta !== 0) {
      return xDelta;
    }

    return left.annotation.id.localeCompare(right.annotation.id);
  });

  const minTopPx = 8;
  let cursorTopPx = minTopPx;
  withAnchors.forEach((entry) => {
    const desiredTop = entry.anchorY - 8;
    entry.top = Math.max(desiredTop, cursorTopPx);
    cursorTopPx =
      entry.top + entry.heightPx + TRACKED_CHANGE_GUTTER_CARD_GAP_PX;
  });

  const maxBottom = Math.max(0, pageHeightPx - minTopPx);
  const bottomAfterInitialPlacement =
    cursorTopPx - TRACKED_CHANGE_GUTTER_CARD_GAP_PX;
  const overflowPx = bottomAfterInitialPlacement - maxBottom;
  if (overflowPx > 0) {
    let shiftedCursorTopPx = minTopPx;
    withAnchors.forEach((entry) => {
      const desiredTop = entry.top - overflowPx;
      entry.top = Math.max(desiredTop, shiftedCursorTopPx);
      shiftedCursorTopPx =
        entry.top + entry.heightPx + TRACKED_CHANGE_GUTTER_CARD_GAP_PX;
    });
  }

  assignGutterConnectorLanes(withAnchors);

  return withAnchors;
}

function sameParagraphLocation(
  a: ParagraphLocation,
  b: ParagraphLocation
): boolean {
  if (a.kind === "paragraph") {
    return b.kind === "paragraph" && a.nodeIndex === b.nodeIndex;
  }

  if (b.kind === "paragraph") {
    return false;
  }

  return (
    a.tableIndex === b.tableIndex &&
    a.rowIndex === b.rowIndex &&
    a.cellIndex === b.cellIndex &&
    a.paragraphIndex === b.paragraphIndex
  );
}

function firstParagraphLocationInTable(
  model: DocModel,
  tableIndex: number
): ParagraphLocation | undefined {
  const tableNode = model.nodes[tableIndex];
  if (!tableNode || tableNode.type !== "table") {
    return undefined;
  }

  for (let rowIndex = 0; rowIndex < tableNode.rows.length; rowIndex += 1) {
    const row = tableNode.rows[rowIndex];
    if (!row) {
      continue;
    }

    for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex += 1) {
      const cell = row.cells[cellIndex];
      if (!cell || cell.style?.vMergeContinuation) {
        continue;
      }

      const cellParagraphs = tableCellParagraphs(cell.nodes);
      if (cellParagraphs.length === 0) {
        continue;
      }

      return {
        kind: "table-cell",
        tableIndex,
        rowIndex,
        cellIndex,
        paragraphIndex: 0,
      };
    }
  }

  return undefined;
}

function lastParagraphLocationInTable(
  model: DocModel,
  tableIndex: number
): ParagraphLocation | undefined {
  const tableNode = model.nodes[tableIndex];
  if (!tableNode || tableNode.type !== "table") {
    return undefined;
  }

  for (let rowIndex = tableNode.rows.length - 1; rowIndex >= 0; rowIndex -= 1) {
    const row = tableNode.rows[rowIndex];
    if (!row) {
      continue;
    }

    for (let cellIndex = row.cells.length - 1; cellIndex >= 0; cellIndex -= 1) {
      const cell = row.cells[cellIndex];
      if (!cell || cell.style?.vMergeContinuation) {
        continue;
      }

      const cellParagraphs = tableCellParagraphs(cell.nodes);
      if (cellParagraphs.length === 0) {
        continue;
      }

      return {
        kind: "table-cell",
        tableIndex,
        rowIndex,
        cellIndex,
        paragraphIndex: cellParagraphs.length - 1,
      };
    }
  }

  return undefined;
}

function nodeIndexFromParagraphLocation(location: ParagraphLocation): number {
  return location.kind === "paragraph"
    ? location.nodeIndex
    : location.tableIndex;
}

function adjustLocationAfterRemovedNodeIndexes(
  location: DocxTextRangeLocation,
  removedNodeIndexes: number[]
): DocxTextRangeLocation | undefined {
  if (removedNodeIndexes.length === 0) {
    return cloneTextRangeLocation(location);
  }

  const normalizedRemoved = [...removedNodeIndexes]
    .filter((value) => Number.isFinite(value) && value >= 0)
    .map((value) => Math.round(value))
    .sort((left, right) => left - right);

  const sourceNodeIndex =
    location.kind === "paragraph" ? location.nodeIndex : location.tableIndex;
  let adjustedNodeIndex = sourceNodeIndex;
  for (const removedIndex of normalizedRemoved) {
    if (removedIndex === adjustedNodeIndex) {
      return undefined;
    }
    if (removedIndex < adjustedNodeIndex) {
      adjustedNodeIndex -= 1;
    }
  }

  if (location.kind === "paragraph") {
    return {
      kind: "paragraph",
      nodeIndex: Math.max(0, adjustedNodeIndex),
    };
  }

  return {
    kind: "table-cell",
    tableIndex: Math.max(0, adjustedNodeIndex),
    rowIndex: location.rowIndex,
    cellIndex: location.cellIndex,
    paragraphIndex: location.paragraphIndex,
  };
}

function tableCoverageBoundaries(
  model: DocModel,
  tableIndex: number
):
  | {
      start: DocxTextRangeBoundary;
      end: DocxTextRangeBoundary;
    }
  | undefined {
  const firstLocation = firstParagraphLocationInTable(model, tableIndex);
  const lastLocation = lastParagraphLocationInTable(model, tableIndex);
  if (!firstLocation || !lastLocation) {
    return undefined;
  }

  const firstParagraph = getParagraphAtLocation(model, firstLocation).paragraph;
  const lastParagraph = getParagraphAtLocation(model, lastLocation).paragraph;
  if (!firstParagraph || !lastParagraph) {
    return undefined;
  }

  return {
    start: {
      location: cloneTextRangeLocation(firstLocation),
      offset: 0,
    },
    end: {
      location: cloneTextRangeLocation(lastLocation),
      offset: paragraphText(lastParagraph).length,
    },
  };
}

function fullyCoveredTableNodeIndexesForRange(
  model: DocModel,
  normalizedRange: DocxTextRange
): number[] {
  const startNodeIndex = nodeIndexFromParagraphLocation(
    normalizedRange.start.location
  );
  const endNodeIndex = nodeIndexFromParagraphLocation(
    normalizedRange.end.location
  );
  const firstIndex = Math.min(startNodeIndex, endNodeIndex);
  const lastIndex = Math.max(startNodeIndex, endNodeIndex);
  const coveredTableIndexes: number[] = [];

  for (let nodeIndex = firstIndex; nodeIndex <= lastIndex; nodeIndex += 1) {
    const node = model.nodes[nodeIndex];
    if (!node || node.type !== "table") {
      continue;
    }

    const boundaries = tableCoverageBoundaries(model, nodeIndex);
    if (!boundaries) {
      continue;
    }

    const coversFromStart =
      compareTextRangeBoundaries(normalizedRange.start, boundaries.start) <= 0;
    const coversToEnd =
      compareTextRangeBoundaries(normalizedRange.end, boundaries.end) >= 0;
    if (!coversFromStart || !coversToEnd) {
      continue;
    }

    coveredTableIndexes.push(nodeIndex);
  }

  return coveredTableIndexes;
}

function firstParagraphLocationFromNode(
  model: DocModel,
  nodeIndex: number
): ParagraphLocation | undefined {
  const node = model.nodes[nodeIndex];
  if (!node) {
    return undefined;
  }

  if (node.type === "paragraph") {
    return {
      kind: "paragraph",
      nodeIndex,
    };
  }

  return firstParagraphLocationInTable(model, nodeIndex);
}

function lastParagraphLocationInNode(
  model: DocModel,
  nodeIndex: number
): ParagraphLocation | undefined {
  const node = model.nodes[nodeIndex];
  if (!node) {
    return undefined;
  }

  if (node.type === "paragraph") {
    return {
      kind: "paragraph",
      nodeIndex,
    };
  }

  return lastParagraphLocationInTable(model, nodeIndex);
}

function nextParagraphLocation(
  model: DocModel,
  location: ParagraphLocation
): ParagraphLocation | undefined {
  if (location.kind === "paragraph") {
    for (
      let nodeIndex = location.nodeIndex + 1;
      nodeIndex < model.nodes.length;
      nodeIndex += 1
    ) {
      const next = firstParagraphLocationFromNode(model, nodeIndex);
      if (next) {
        return next;
      }
    }
    return undefined;
  }

  const tableNode = model.nodes[location.tableIndex];
  if (!tableNode || tableNode.type !== "table") {
    return undefined;
  }

  const currentCell =
    tableNode.rows[location.rowIndex]?.cells[location.cellIndex];
  if (!currentCell) {
    return undefined;
  }

  const cellParagraphs = tableCellParagraphs(currentCell.nodes);
  if (location.paragraphIndex < cellParagraphs.length - 1) {
    return {
      kind: "table-cell",
      tableIndex: location.tableIndex,
      rowIndex: location.rowIndex,
      cellIndex: location.cellIndex,
      paragraphIndex: location.paragraphIndex + 1,
    };
  }

  for (
    let rowIndex = location.rowIndex;
    rowIndex < tableNode.rows.length;
    rowIndex += 1
  ) {
    const row = tableNode.rows[rowIndex];
    if (!row) {
      continue;
    }

    for (
      let cellIndex =
        rowIndex === location.rowIndex ? location.cellIndex + 1 : 0;
      cellIndex < row.cells.length;
      cellIndex += 1
    ) {
      const cell = row.cells[cellIndex];
      if (!cell || cell.style?.vMergeContinuation) {
        continue;
      }

      if (tableCellParagraphs(cell.nodes).length === 0) {
        continue;
      }

      return {
        kind: "table-cell",
        tableIndex: location.tableIndex,
        rowIndex,
        cellIndex,
        paragraphIndex: 0,
      };
    }
  }

  for (
    let nodeIndex = location.tableIndex + 1;
    nodeIndex < model.nodes.length;
    nodeIndex += 1
  ) {
    const next = firstParagraphLocationFromNode(model, nodeIndex);
    if (next) {
      return next;
    }
  }

  return undefined;
}

function firstParagraphLocationInDocument(
  model: DocModel
): ParagraphLocation | undefined {
  for (let nodeIndex = 0; nodeIndex < model.nodes.length; nodeIndex += 1) {
    const first = firstParagraphLocationFromNode(model, nodeIndex);
    if (first) {
      return first;
    }
  }

  return undefined;
}

function lastParagraphLocationInDocument(
  model: DocModel
): ParagraphLocation | undefined {
  for (let nodeIndex = model.nodes.length - 1; nodeIndex >= 0; nodeIndex -= 1) {
    const node = model.nodes[nodeIndex];
    if (!node) {
      continue;
    }

    if (node.type === "paragraph") {
      return {
        kind: "paragraph",
        nodeIndex,
      };
    }

    for (let rowIndex = node.rows.length - 1; rowIndex >= 0; rowIndex -= 1) {
      const row = node.rows[rowIndex];
      if (!row) {
        continue;
      }

      for (
        let cellIndex = row.cells.length - 1;
        cellIndex >= 0;
        cellIndex -= 1
      ) {
        const cell = row.cells[cellIndex];
        if (!cell || cell.style?.vMergeContinuation) {
          continue;
        }

        const paragraphs = tableCellParagraphs(cell.nodes);
        if (paragraphs.length === 0) {
          continue;
        }

        return {
          kind: "table-cell",
          tableIndex: nodeIndex,
          rowIndex,
          cellIndex,
          paragraphIndex: paragraphs.length - 1,
        };
      }
    }
  }

  return undefined;
}

function paragraphRangeForMutate(
  model: DocModel,
  start: ParagraphLocation,
  end: ParagraphLocation
): {
  location: ParagraphLocation;
}[] {
  const items: { location: ParagraphLocation }[] = [];

  const ordered =
    compareParagraphLocations(start, end) <= 0 ? [start, end] : [end, start];
  let current: ParagraphLocation | undefined = ordered[0];
  const limit = ordered[1];
  while (current) {
    items.push({ location: current });
    if (compareParagraphLocations(current, limit) >= 0) {
      break;
    }
    current = nextParagraphLocation(model, current);
  }

  return items;
}

function normalizeRangeBoundaryParagraphOffset(
  paragraph: ParagraphNode,
  offset: number
): number {
  const length = paragraphText(paragraph).length;
  return Math.max(0, Math.min(Math.max(0, length), Math.round(offset)));
}

function rangeCoversEntireDocument(
  model: DocModel,
  range: DocxTextRange
): boolean {
  const normalizedRange = normalizeTextRange(range);
  const firstLocation = firstParagraphLocationInDocument(model);
  const lastLocation = lastParagraphLocationInDocument(model);
  if (!firstLocation || !lastLocation) {
    return false;
  }

  if (!sameParagraphLocation(normalizedRange.start.location, firstLocation)) {
    return false;
  }

  if (!sameParagraphLocation(normalizedRange.end.location, lastLocation)) {
    return false;
  }

  const firstParagraph = getParagraphAtLocation(model, firstLocation).paragraph;
  const lastParagraph = getParagraphAtLocation(model, lastLocation).paragraph;
  if (!firstParagraph || !lastParagraph) {
    return false;
  }

  const safeStart = normalizeRangeBoundaryParagraphOffset(
    firstParagraph,
    normalizedRange.start.offset
  );
  const safeEnd = normalizeRangeBoundaryParagraphOffset(
    lastParagraph,
    normalizedRange.end.offset
  );
  return safeStart <= 0 && safeEnd >= paragraphText(lastParagraph).length;
}

function resolveRangeBoundaryOffsetsForParagraph(
  currentLocation: ParagraphLocation,
  rangeStart: DocxTextRangeBoundary,
  rangeEnd: DocxTextRangeBoundary,
  paragraph: ParagraphNode
): [number, number] {
  const ordered =
    compareTextRangeBoundaries(rangeStart, rangeEnd) <= 0
      ? [rangeStart, rangeEnd]
      : [rangeEnd, rangeStart];
  const startBoundary = ordered[0];
  const endBoundary = ordered[1];
  const isStartBoundaryHere = sameParagraphLocation(
    currentLocation,
    startBoundary.location
  );
  const isEndBoundaryHere = sameParagraphLocation(
    currentLocation,
    endBoundary.location
  );

  const safeStart = isStartBoundaryHere
    ? normalizeRangeBoundaryParagraphOffset(paragraph, startBoundary.offset)
    : 0;
  const safeEnd = isEndBoundaryHere
    ? normalizeRangeBoundaryParagraphOffset(paragraph, endBoundary.offset)
    : paragraphText(paragraph).length;

  return [safeStart, safeEnd];
}

function cloneTextRangeLocation(
  location: DocxTextRangeLocation
): DocxTextRangeLocation {
  if (location.kind === "paragraph") {
    return {
      kind: "paragraph",
      nodeIndex: location.nodeIndex,
    };
  }

  return {
    kind: "table-cell",
    tableIndex: location.tableIndex,
    rowIndex: location.rowIndex,
    cellIndex: location.cellIndex,
    paragraphIndex: location.paragraphIndex,
  };
}

function cloneTextRange(range?: DocxTextRange): DocxTextRange | undefined {
  if (!range) {
    return undefined;
  }

  return {
    start: {
      location: cloneTextRangeLocation(range.start.location),
      offset: range.start.offset,
    },
    end: {
      location: cloneTextRangeLocation(range.end.location),
      offset: range.end.offset,
    },
  };
}

function cloneEditorSelection(
  selection: DocxEditorSelection
): DocxEditorSelection {
  if (selection.kind === "paragraph") {
    return {
      kind: "paragraph",
      nodeIndex: selection.nodeIndex,
    };
  }

  return {
    kind: "table-cell",
    tableIndex: selection.tableIndex,
    rowIndex: selection.rowIndex,
    cellIndex: selection.cellIndex,
  };
}

function sameEditorSelection(
  a: DocxEditorSelection,
  b: DocxEditorSelection
): boolean {
  if (a.kind !== b.kind) {
    return false;
  }

  if (a.kind === "paragraph") {
    return (
      a.nodeIndex ===
      (b as Extract<DocxEditorSelection, { kind: "paragraph" }>).nodeIndex
    );
  }

  const tableSelection = b as Extract<
    DocxEditorSelection,
    { kind: "table-cell" }
  >;
  return (
    a.tableIndex === tableSelection.tableIndex &&
    a.rowIndex === tableSelection.rowIndex &&
    a.cellIndex === tableSelection.cellIndex
  );
}

function sameTextRangeBoundary(
  a: DocxTextRangeBoundary,
  b: DocxTextRangeBoundary
): boolean {
  return compareTextRangeBoundaries(a, b) === 0;
}

function sameTextRange(a?: DocxTextRange, b?: DocxTextRange): boolean {
  if (!a && !b) {
    return true;
  }
  if (!a || !b) {
    return false;
  }

  const normalizedA = normalizeTextRange(a);
  const normalizedB = normalizeTextRange(b);
  return (
    sameTextRangeBoundary(normalizedA.start, normalizedB.start) &&
    sameTextRangeBoundary(normalizedA.end, normalizedB.end)
  );
}

export function shouldReissueDomSelectionRestore(options: {
  modelChanged: boolean;
  selectionChanged: boolean;
  rangeChanged: boolean;
  activeTextRange?: DocxTextRange;
  suppressNext: boolean;
  selectionSessionKind: DocxSelectionSessionKind;
}): boolean {
  if (options.suppressNext) {
    return false;
  }

  if (
    options.selectionSessionKind === "pointer" ||
    options.selectionSessionKind === "keyboard" ||
    options.selectionSessionKind === "composition"
  ) {
    return false;
  }

  if (!options.activeTextRange) {
    return false;
  }

  if (options.selectionChanged || options.rangeChanged) {
    return false;
  }

  return options.modelChanged;
}

function shouldSyncActiveRangeOnKeyUp(
  event: { key: string; shiftKey: boolean; ctrlKey: boolean; metaKey: boolean; preventDefault: () => void; nativeEvent: KeyboardEvent }
): boolean {
  const nativeKeyboardEvent = event.nativeEvent as KeyboardEvent | undefined;
  if (nativeKeyboardEvent?.isComposing) {
    return false;
  }

  if (MODIFIER_ONLY_KEYS.has(event.key)) {
    return false;
  }

  if (ACTIVE_RANGE_SYNC_KEYS.has(event.key)) {
    return true;
  }

  if (event.ctrlKey || event.metaKey) {
    const normalizedKey = event.key.toLowerCase();
    return (
      normalizedKey === "a" ||
      normalizedKey === "x" ||
      normalizedKey === "z" ||
      normalizedKey === "y"
    );
  }

  return false;
}

function isCollapsedSelectionAtElementStart(
  element: HTMLElement,
  range: Range,
  pointX?: number
): boolean {
  if (!range.collapsed) {
    return false;
  }

  if (range.startContainer === element && range.startOffset === 0) {
    return true;
  }

  if (!element.contains(range.startContainer)) {
    return false;
  }

  try {
    const offsetProbe = document.createRange();
    offsetProbe.selectNodeContents(element);
    offsetProbe.setEnd(range.startContainer, range.startOffset);
    const textOffset = offsetProbe.toString().length;
    if (textOffset !== 0) {
      return false;
    }

    if (Number.isFinite(pointX)) {
      const elementRect = element.getBoundingClientRect();
      if ((pointX as number) <= elementRect.left + 12) {
        return false;
      }
    }

    return true;
  } catch {
    return false;
  }
}

function isSuspiciousCollapsedSelectionAtElementStart(
  element: HTMLElement,
  range: Range,
  pointX?: number
): boolean {
  if (!isCollapsedSelectionAtElementStart(element, range, pointX)) {
    return false;
  }

  return (
    range.startContainer === element || !(range.startContainer instanceof Text)
  );
}

function imageLocationToParagraphLocation(
  location: DocxImageLocation
): ParagraphLocation {
  if (location.kind === "paragraph") {
    return {
      kind: "paragraph",
      nodeIndex: location.nodeIndex,
    };
  }

  return {
    kind: "table-cell",
    tableIndex: location.tableIndex,
    rowIndex: location.rowIndex,
    cellIndex: location.cellIndex,
    paragraphIndex: location.paragraphIndex,
  };
}

function getParagraphAtLocation(
  model: DocModel,
  location: ParagraphLocation
): {
  paragraph?: ParagraphNode;
  tableNode?: TableNode;
} {
  if (location.kind === "paragraph") {
    const node = model.nodes[location.nodeIndex];
    if (!node || node.type !== "paragraph") {
      return {};
    }

    return { paragraph: node };
  }

  const tableNode = model.nodes[location.tableIndex];
  if (!tableNode || tableNode.type !== "table") {
    return {};
  }

  const cell = tableNode.rows[location.rowIndex]?.cells[location.cellIndex];
  if (!cell) {
    return {};
  }

  const paragraph = tableCellParagraphs(cell.nodes)[location.paragraphIndex];
  if (!paragraph) {
    return {};
  }

  return { paragraph, tableNode };
}

function collectFormFieldsFromModel(model: DocModel): DocxSelectedFormField[] {
  const collected: DocxSelectedFormField[] = [];

  model.nodes.forEach((node, nodeIndex) => {
    if (node.type === "paragraph") {
      node.children.forEach((child, childIndex) => {
        if (child.type !== "form-field") {
          return;
        }

        collected.push({
          location: {
            kind: "paragraph",
            nodeIndex,
            childIndex,
          },
          field: child,
        });
      });
      return;
    }

    node.rows.forEach((row, rowIndex) => {
      row.cells.forEach((cell, cellIndex) => {
        tableCellParagraphs(cell.nodes).forEach((paragraph, paragraphIndex) => {
          paragraph.children.forEach((child, childIndex) => {
            if (child.type !== "form-field") {
              return;
            }

            collected.push({
              location: {
                kind: "table-cell",
                tableIndex: nodeIndex,
                rowIndex,
                cellIndex,
                paragraphIndex,
                childIndex,
              },
              field: child,
            });
          });
        });
      });
    });
  });

  return collected;
}

type ParagraphTextUpdateOptions = Parameters<typeof updateParagraphText>[3];

function selectionFallbackParagraphLocation(
  selection: DocxEditorSelection
): ParagraphLocation {
  if (selection.kind === "paragraph") {
    return {
      kind: "paragraph",
      nodeIndex: selection.nodeIndex,
    };
  }

  return {
    kind: "table-cell",
    tableIndex: selection.tableIndex,
    rowIndex: selection.rowIndex,
    cellIndex: selection.cellIndex,
    paragraphIndex: 0,
  };
}

function selectionFromTextRangeLocation(
  location: DocxTextRangeLocation
): DocxEditorSelection {
  if (location.kind === "paragraph") {
    return {
      kind: "paragraph",
      nodeIndex: location.nodeIndex,
    };
  }

  return {
    kind: "table-cell",
    tableIndex: location.tableIndex,
    rowIndex: location.rowIndex,
    cellIndex: location.cellIndex,
  };
}

function paragraphLocationFromTextRangeLocation(
  location: DocxTextRangeLocation
): ParagraphLocation {
  if (location.kind === "paragraph") {
    return {
      kind: "paragraph",
      nodeIndex: location.nodeIndex,
    };
  }

  return {
    kind: "table-cell",
    tableIndex: location.tableIndex,
    rowIndex: location.rowIndex,
    cellIndex: location.cellIndex,
    paragraphIndex: Math.max(0, Math.round(location.paragraphIndex)),
  };
}

function resolveSelectedParagraphLocation(
  selection: DocxEditorSelection,
  activeTextRange?: DocxTextRange
): ParagraphLocation {
  const activeLocation = activeTextRange?.start.location;
  if (activeLocation) {
    return paragraphLocationFromTextRangeLocation(activeLocation);
  }

  return selectionFallbackParagraphLocation(selection);
}

function normalizeParagraphLocationForModel(
  model: DocModel,
  location: ParagraphLocation
): ParagraphLocation | undefined {
  const exact = getParagraphAtLocation(model, location).paragraph;
  if (exact) {
    return cloneTextRangeLocation(location);
  }

  if (location.kind === "table-cell") {
    const tableNode = model.nodes[location.tableIndex];
    if (tableNode && tableNode.type === "table") {
      const safeRowIndex = clampNumber(
        location.rowIndex,
        0,
        Math.max(0, tableNode.rows.length - 1)
      );
      const row = tableNode.rows[safeRowIndex];
      if (row) {
        const safeCellIndex = clampNumber(
          location.cellIndex,
          0,
          Math.max(0, row.cells.length - 1)
        );
        const candidateCell = row.cells[safeCellIndex];
        if (candidateCell && !candidateCell.style?.vMergeContinuation) {
          const paragraphs = tableCellParagraphs(candidateCell.nodes);
          if (paragraphs.length > 0) {
            return {
              kind: "table-cell",
              tableIndex: location.tableIndex,
              rowIndex: safeRowIndex,
              cellIndex: safeCellIndex,
              paragraphIndex: clampNumber(
                location.paragraphIndex,
                0,
                Math.max(0, paragraphs.length - 1)
              ),
            };
          }
        }
      }

      const firstInTable = firstParagraphLocationInTable(
        model,
        location.tableIndex
      );
      if (firstInTable) {
        return firstInTable;
      }
    }
  }

  if (model.nodes.length === 0) {
    return undefined;
  }

  const anchorNodeIndex =
    location.kind === "paragraph" ? location.nodeIndex : location.tableIndex;
  const clampedAnchorNodeIndex = clampNumber(
    anchorNodeIndex,
    0,
    Math.max(0, model.nodes.length - 1)
  );

  const sameNodeFallback = firstParagraphLocationFromNode(
    model,
    clampedAnchorNodeIndex
  );
  if (sameNodeFallback) {
    return sameNodeFallback;
  }

  for (
    let nodeIndex = clampedAnchorNodeIndex + 1;
    nodeIndex < model.nodes.length;
    nodeIndex += 1
  ) {
    const forward = firstParagraphLocationFromNode(model, nodeIndex);
    if (forward) {
      return forward;
    }
  }

  for (
    let nodeIndex = clampedAnchorNodeIndex - 1;
    nodeIndex >= 0;
    nodeIndex -= 1
  ) {
    const backward = firstParagraphLocationFromNode(model, nodeIndex);
    if (backward) {
      return backward;
    }
  }

  return firstParagraphLocationInDocument(model);
}

function normalizeTextRangeForModel(
  model: DocModel,
  range?: DocxTextRange
): DocxTextRange | undefined {
  if (!range) {
    return undefined;
  }

  const normalized = normalizeTextRange(range);
  const startLocation = normalizeParagraphLocationForModel(
    model,
    paragraphLocationFromTextRangeLocation(normalized.start.location)
  );
  const endLocation = normalizeParagraphLocationForModel(
    model,
    paragraphLocationFromTextRangeLocation(normalized.end.location)
  );
  if (!startLocation || !endLocation) {
    return undefined;
  }

  const startParagraph = getParagraphAtLocation(model, startLocation).paragraph;
  const endParagraph = getParagraphAtLocation(model, endLocation).paragraph;
  if (!startParagraph || !endParagraph) {
    return undefined;
  }

  return normalizeTextRange({
    start: {
      location: cloneTextRangeLocation(startLocation),
      offset: normalizeRangeBoundaryParagraphOffset(
        startParagraph,
        normalized.start.offset
      ),
    },
    end: {
      location: cloneTextRangeLocation(endLocation),
      offset: normalizeRangeBoundaryParagraphOffset(
        endParagraph,
        normalized.end.offset
      ),
    },
  });
}

function normalizeSelectionForModel(
  model: DocModel,
  selection: DocxEditorSelection
): DocxEditorSelection {
  const normalizedParagraphLocation = normalizeParagraphLocationForModel(
    model,
    selectionFallbackParagraphLocation(selection)
  );
  if (!normalizedParagraphLocation) {
    return {
      kind: "paragraph",
      nodeIndex: 0,
    };
  }

  return selectionFromTextRangeLocation(normalizedParagraphLocation);
}

function normalizeEditorCursorStateForModel(
  model: DocModel,
  selection: DocxEditorSelection,
  activeTextRange?: DocxTextRange
): {
  selection: DocxEditorSelection;
  activeTextRange?: DocxTextRange;
} {
  const normalizedRange = normalizeTextRangeForModel(model, activeTextRange);
  if (normalizedRange) {
    return {
      selection: normalizeSelectionForModel(
        model,
        selectionFromTextRangeLocation(normalizedRange.start.location)
      ),
      activeTextRange: normalizedRange,
    };
  }

  return {
    selection: normalizeSelectionForModel(model, selection),
    activeTextRange: undefined,
  };
}

function updateParagraphTextAtLocation(
  model: DocModel,
  location: ParagraphLocation,
  text: string,
  options?: ParagraphTextUpdateOptions
): DocModel {
  if (location.kind === "paragraph") {
    return updateParagraphText(model, location.nodeIndex, text, options);
  }

  return updateTableCellParagraphText(
    model,
    location.tableIndex,
    location.rowIndex,
    location.cellIndex,
    location.paragraphIndex,
    text,
    options
  );
}

function sectionParagraphLocationKey(
  location: DocxSectionParagraphLocation
): string {
  return JSON.stringify({
    region: location.region,
    partName: location.partName,
    nodeIndex: location.nodeIndex,
    rowIndex: location.rowIndex ?? -1,
    cellIndex: location.cellIndex ?? -1,
    paragraphIndex: location.paragraphIndex ?? -1,
  });
}

function sectionImageLocationKey(location: DocxSectionImageLocation): string {
  return JSON.stringify({
    region: location.region,
    partName: location.partName,
    nodeIndex: location.nodeIndex,
    rowIndex: location.rowIndex ?? -1,
    cellIndex: location.cellIndex ?? -1,
    paragraphIndex: location.paragraphIndex ?? -1,
    childIndex: location.childIndex,
  });
}

function tableCellParagraphDraftKey(
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  paragraphIndex: number
): string {
  return `${tableIndex}:${rowIndex}:${cellIndex}:${paragraphIndex}`;
}

function parseSectionParagraphLocationKey(
  raw: string
): DocxSectionParagraphLocation | undefined {
  try {
    const parsed = JSON.parse(raw) as {
      region?: string;
      partName?: string;
      nodeIndex?: number;
      rowIndex?: number;
      cellIndex?: number;
      paragraphIndex?: number;
    };

    if (
      !parsed ||
      (parsed.region !== "header" && parsed.region !== "footer") ||
      typeof parsed.partName !== "string" ||
      !Number.isFinite(parsed.nodeIndex)
    ) {
      return undefined;
    }

    const rowIndex =
      Number.isFinite(parsed.rowIndex) && (parsed.rowIndex as number) >= 0
        ? Math.round(parsed.rowIndex as number)
        : undefined;
    const cellIndex =
      Number.isFinite(parsed.cellIndex) && (parsed.cellIndex as number) >= 0
        ? Math.round(parsed.cellIndex as number)
        : undefined;
    const paragraphIndex =
      Number.isFinite(parsed.paragraphIndex) &&
      (parsed.paragraphIndex as number) >= 0
        ? Math.round(parsed.paragraphIndex as number)
        : undefined;

    return {
      region: parsed.region,
      partName: parsed.partName,
      nodeIndex: Math.max(0, Math.round(parsed.nodeIndex as number)),
      rowIndex,
      cellIndex,
      paragraphIndex,
    };
  } catch {
    return undefined;
  }
}

function paragraphTextFromSectionLocation(
  sectionNodes: DocModel["nodes"],
  location: Omit<DocxSectionParagraphLocation, "region" | "partName">
): string | undefined {
  const rootNode = sectionNodes[location.nodeIndex];
  if (!rootNode) {
    return undefined;
  }

  if (location.rowIndex === undefined || location.cellIndex === undefined) {
    if (rootNode.type !== "paragraph") {
      return undefined;
    }
    return paragraphText(rootNode);
  }

  if (rootNode.type !== "table") {
    return undefined;
  }

  const paragraphIndex = Math.max(0, Math.round(location.paragraphIndex ?? 0));
  const cell = rootNode.rows[location.rowIndex]?.cells[location.cellIndex];
  if (!cell) {
    return undefined;
  }

  const paragraph = tableCellParagraphs(cell.nodes)[paragraphIndex];
  return paragraph ? paragraphText(paragraph) : undefined;
}

function sectionParagraphFromLocation(
  sectionNodes: DocModel["nodes"],
  location: Omit<DocxSectionParagraphLocation, "region" | "partName">
): {
  paragraph?: ParagraphNode;
  tableNode?: TableNode;
} {
  const rootNode = sectionNodes[location.nodeIndex];
  if (!rootNode) {
    return {};
  }

  if (location.rowIndex === undefined || location.cellIndex === undefined) {
    if (rootNode.type !== "paragraph") {
      return {};
    }
    return { paragraph: rootNode };
  }

  if (rootNode.type !== "table") {
    return {};
  }

  const paragraphIndex = Math.max(0, Math.round(location.paragraphIndex ?? 0));
  const cell = rootNode.rows[location.rowIndex]?.cells[location.cellIndex];
  if (!cell) {
    return {};
  }

  const paragraph = tableCellParagraphs(cell.nodes)[paragraphIndex];
  if (!paragraph) {
    return {};
  }

  return {
    paragraph,
    tableNode: rootNode,
  };
}

function updateSectionParagraphTextAtLocation(
  model: DocModel,
  location: DocxSectionParagraphLocation,
  text: string,
  options?: ParagraphTextUpdateOptions
): DocModel {
  const next = cloneDocModel(model);
  const candidateSectionLists =
    location.region === "header"
      ? [
          next.metadata.headerSections ?? [],
          ...(next.metadata.sections?.map(
            (section) => section.headerSections ?? []
          ) ?? []),
        ]
      : [
          next.metadata.footerSections ?? [],
          ...(next.metadata.sections?.map(
            (section) => section.footerSections ?? []
          ) ?? []),
        ];

  const normalizedLocation = {
    nodeIndex: Math.max(0, Math.round(location.nodeIndex)),
    rowIndex:
      Number.isFinite(location.rowIndex) && (location.rowIndex as number) >= 0
        ? Math.round(location.rowIndex as number)
        : undefined,
    cellIndex:
      Number.isFinite(location.cellIndex) && (location.cellIndex as number) >= 0
        ? Math.round(location.cellIndex as number)
        : undefined,
    paragraphIndex:
      Number.isFinite(location.paragraphIndex) &&
      (location.paragraphIndex as number) >= 0
        ? Math.round(location.paragraphIndex as number)
        : undefined,
  };

  let changed = false;
  candidateSectionLists.forEach((sections) => {
    sections.forEach((section) => {
      if (section.partName !== location.partName) {
        return;
      }

      const currentText = paragraphTextFromSectionLocation(
        section.nodes,
        normalizedLocation
      );
      if (currentText === undefined || currentText === text) {
        return;
      }

      const scopedModel: DocModel = {
        ...next,
        nodes: section.nodes,
      };
      const updatedScopedModel =
        normalizedLocation.rowIndex === undefined ||
        normalizedLocation.cellIndex === undefined
          ? updateParagraphText(
              scopedModel,
              normalizedLocation.nodeIndex,
              text,
              options
            )
          : updateTableCellParagraphText(
              scopedModel,
              normalizedLocation.nodeIndex,
              normalizedLocation.rowIndex,
              normalizedLocation.cellIndex,
              normalizedLocation.paragraphIndex ?? 0,
              text,
              options
            );

      section.nodes = updatedScopedModel.nodes;
      changed = true;
    });
  });

  return changed ? next : model;
}

function updateSectionImageFloatingAtLocation(
  model: DocModel,
  location: DocxSectionImageLocation,
  patch: Partial<NonNullable<ImageRunNode["floating"]>>
): DocModel {
  const next = cloneDocModel(model);
  const candidateSectionLists =
    location.region === "header"
      ? [
          next.metadata.headerSections ?? [],
          ...(next.metadata.sections?.map(
            (section) => section.headerSections ?? []
          ) ?? []),
        ]
      : [
          next.metadata.footerSections ?? [],
          ...(next.metadata.sections?.map(
            (section) => section.footerSections ?? []
          ) ?? []),
        ];

  const normalizedLocation = {
    nodeIndex: Math.max(0, Math.round(location.nodeIndex)),
    rowIndex:
      Number.isFinite(location.rowIndex) && (location.rowIndex as number) >= 0
        ? Math.round(location.rowIndex as number)
        : undefined,
    cellIndex:
      Number.isFinite(location.cellIndex) && (location.cellIndex as number) >= 0
        ? Math.round(location.cellIndex as number)
        : undefined,
    paragraphIndex:
      Number.isFinite(location.paragraphIndex) &&
      (location.paragraphIndex as number) >= 0
        ? Math.round(location.paragraphIndex as number)
        : undefined,
    childIndex: Math.max(0, Math.round(location.childIndex)),
  };

  let changed = false;
  candidateSectionLists.forEach((sections) => {
    sections.forEach((section) => {
      if (section.partName !== location.partName) {
        return;
      }

      const lookup = sectionParagraphFromLocation(
        section.nodes,
        normalizedLocation
      );
      const paragraph = lookup.paragraph;
      if (!paragraph) {
        return;
      }

      const child = paragraph.children[normalizedLocation.childIndex];
      if (!child || child.type !== "image") {
        return;
      }

      child.floating = {
        ...(child.floating ?? {}),
        ...patch,
      };
      paragraph.sourceXml = undefined;
      if (lookup.tableNode) {
        lookup.tableNode.sourceXml = undefined;
      }
      changed = true;
    });
  });

  return changed ? next : model;
}


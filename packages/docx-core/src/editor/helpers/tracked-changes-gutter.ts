// Tracked-changes gutter: card positioning, coloring and layout pure functions.
// Upstream editor.tsx: lines 23005-23536.
//
// These functions turn the model-derived `DocxTrackedChange` / `DocxComment`
// entries (produced by tracked-changes.ts) into positioned gutter cards for a
// single rendered page: accent colors, height estimates, anchor lookup in the
// live DOM, connector-lane assignment and the vertical stacking layout.
//
// All functions are framework-agnostic and read-only with respect to the model.
// `React.CSSProperties` is replaced with `Record<string, string | number | undefined>`.

import type { DocxTextRangeLocation } from "./editor-types";
import type {
  DocxComment,
  DocxDocumentTheme,
  DocxTrackedChange,
  DocxTrackedChangeKind
} from "./editor-types";
import type { DocumentPageNodeSegment } from "../../layout/page-segmentation-core";
import { commentThreadIsResolved } from "./tracked-changes";
import {
  TRACKED_CHANGE_GUTTER_CARD_GAP_PX,
  TRACKED_CHANGE_GUTTER_CARD_MIN_HEIGHT_PX,
  TRACKED_CHANGE_GUTTER_CONNECTOR_LANE_COUNT
} from "./constants";
import { clampNumber } from "./zoom-utils";
import { normalizeTrackedChangeSnippet } from "./xml-parsing";

// ---------------------------------------------------------------------------
// Accent colors & highlight styles (upstream 23005-23046)
// ---------------------------------------------------------------------------

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

export function commentAccentColor(
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

export function commentHighlightStyle(
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

// ---------------------------------------------------------------------------
// Card height estimates (upstream 23048-23052, 23377-23383)
// ---------------------------------------------------------------------------

export function estimateCommentCardHeight(comment: DocxComment): number {
  const snippet = comment.text || "Comment";
  const lines = Math.min(2, Math.max(1, Math.ceil(snippet.length / 42)));
  return Math.max(TRACKED_CHANGE_GUTTER_CARD_MIN_HEIGHT_PX, 24 + lines * 11);
}

export function trackedChangeKindLabel(kind: DocxTrackedChangeKind): string {
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
    case "table-format-change":
      return "Table formatting";
    case "row-format-change":
      return "Row formatting";
    case "cell-format-change":
      return "Cell formatting";
    default:
      return kind;
  }
}

export function trackedChangeAccentColor(
  kind: DocxTrackedChangeKind,
  documentTheme: DocxDocumentTheme
): string {
  const palette =
    documentTheme === "dark"
      ? {
          insertion: "#4ade80",
          deletion: "#fca5a5",
          moveFrom: "#86efac",
          moveTo: "#86efac",
          format: "#c084fc",
        }
      : {
          insertion: "#15803d",
          deletion: "#b91c1c",
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
    case "table-format-change":
    case "row-format-change":
    case "cell-format-change":
      return palette.format;
    default:
      return palette.format;
  }
}

export function trackedChangeUsesGutterBalloon(
  change: DocxTrackedChange
): boolean {
  return change.kind !== "insertion" && change.kind !== "move-to";
}

// ---------------------------------------------------------------------------
// Annotation sort / page-membership helpers (upstream 23117-23178)
// ---------------------------------------------------------------------------

export function gutterAnnotationSortTuple(
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

export function trackedChangeBelongsToPageSegments(
  location: DocxTextRangeLocation,
  pageSegments: readonly DocumentPageNodeSegment[]
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

export function resolveGutterAnnotationPageIndex(
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

// ---------------------------------------------------------------------------
// DOM anchor lookup (upstream 23180-23322)
// ---------------------------------------------------------------------------

export function findTrackedChangeAnchorElementInPage(
  pageElement: HTMLElement,
  location: DocxTextRangeLocation
): HTMLElement | undefined {
  if (location.kind === "paragraph") {
    // The paragraph host rendered by `DocxParagraphHost` / `DocxPageBody`.
    const paragraphElement = pageElement.querySelector(
      `[data-docx-paragraph-host="true"][data-docx-paragraph-node-index="${location.nodeIndex}"]`
    ) as HTMLElement | null;
    return paragraphElement ?? undefined;
  }

  const paragraphElement = pageElement.querySelector(
    `[data-docx-table-cell-paragraph-host="true"][data-docx-table-index="${location.tableIndex}"][data-docx-table-row-index="${location.rowIndex}"][data-docx-table-cell-index="${location.cellIndex}"][data-docx-paragraph-index="${location.paragraphIndex}"]`
  ) as HTMLElement | null;
  if (paragraphElement) {
    return paragraphElement;
  }

  const cellElement = pageElement.querySelector(
    `[data-docx-table-cell="true"][data-docx-table-index="${location.tableIndex}"][data-docx-table-row-index="${location.rowIndex}"][data-docx-table-cell-index="${location.cellIndex}"]`
  ) as HTMLElement | null;
  return cellElement ?? undefined;
}

export function elementRectWithinContainer(
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

// ---------------------------------------------------------------------------
// Gutter annotation model (upstream 23244-23375)
// ---------------------------------------------------------------------------

export interface TrackedChangeAnchorPoint {
  x: number;
  y: number;
}

/**
 * One entry in the page gutter: either a tracked change or a comment. Both
 * share the anchor/stacking pipeline so they interleave in document order.
 */
export interface DocxGutterAnnotation {
  /** Unique within one page; `comment:`/`change:` prefixed. */
  id: string;
  location: DocxTextRangeLocation;
  trackedChange?: DocxTrackedChange;
  comment?: DocxComment;
  /** 1-based document-order position on the page (card / marker index). */
  ordinal?: number;
}

// ---------------------------------------------------------------------------
// Page annotation collection (comments + tracked changes)
// ---------------------------------------------------------------------------

/**
 * One page's gutter annotations, in stable document order. `ordinal` is the
 * 1-based position of the annotation inside its page, used as the comment
 * index shown both on the anchor marker and on the gutter card.
 */
export interface PageGutterAnnotationsOptions {
  pageSegments: readonly DocumentPageNodeSegment[];
  comments?: readonly DocxComment[];
  trackedChanges?: readonly DocxTrackedChange[];
  /** Defaults to true. */
  showComments?: boolean;
  /** Defaults to true. */
  showTrackedChanges?: boolean;
  /** Header part rendered on this page, used to place header comments. */
  headerPartName?: string;
  /** Footer part rendered on this page, used to place footer comments. */
  footerPartName?: string;
  /** Resolved threads are hidden unless this is true. Defaults to false. */
  showResolvedComments?: boolean;
}

function gutterAnnotationRegionRank(annotation: DocxGutterAnnotation): number {
  switch (annotation.comment?.region) {
    case "header":
      return 1;
    case "footer":
      return 2;
    default:
      return 0;
  }
}

function gutterPartNamesMatch(left?: string, right?: string): boolean {
  if (left === right) {
    return true;
  }
  if (!left || !right) {
    return false;
  }
  return (left.split("/").pop() ?? left) === (right.split("/").pop() ?? right);
}

function commentBelongsToPage(
  comment: DocxComment,
  options: PageGutterAnnotationsOptions
): boolean {
  if (comment.region === "header") {
    return gutterPartNamesMatch(comment.partName, options.headerPartName);
  }
  if (comment.region === "footer") {
    return gutterPartNamesMatch(comment.partName, options.footerPartName);
  }
  return trackedChangeBelongsToPageSegments(
    resolveGutterAnnotationLocation(comment, -1),
    options.pageSegments
  );
}

function commentsVisibleInGutter(
  options: PageGutterAnnotationsOptions
): readonly DocxComment[] {
  const comments = options.comments ?? [];
  if (options.showResolvedComments) {
    return comments;
  }
  const commentsById = new Map(
    comments.map((comment) => [comment.commentId, comment])
  );
  return comments.filter(
    (comment) => !commentThreadIsResolved(comment, commentsById)
  );
}

function gutterAnnotationKindRank(annotation: DocxGutterAnnotation): number {
  // Comments are numbered, so they sort before unnumbered tracked changes when
  // both anchor to the same paragraph.
  return annotation.comment ? 0 : 1;
}

function gutterAnnotationDetailKey(annotation: DocxGutterAnnotation): string {
  if (annotation.comment) {
    const commentId = annotation.comment.commentId;
    return Number.isFinite(commentId)
      ? `c${String(commentId).padStart(12, "0")}`
      : `c${annotation.comment.id}`;
  }

  const change = annotation.trackedChange;
  if (!change) {
    return "";
  }

  return `t${change.kind}:${change.inlineAnchorId ?? change.id}`;
}

/**
 * Hosts may hand the gutter a comment/tracked-change record without an
 * explicit location (only a node index); treat it as a body paragraph so page
 * membership and anchor lookup still work.
 */
function resolveGutterAnnotationLocation(
  entry: { location?: DocxTextRangeLocation; nodeIndex?: number },
  fallbackNodeIndex: number
): DocxTextRangeLocation {
  if (entry.location) {
    return entry.location;
  }

  const nodeIndex = Number.isFinite(entry.nodeIndex)
    ? (entry.nodeIndex as number)
    : fallbackNodeIndex;
  return { kind: "paragraph", nodeIndex };
}

function compareGutterAnnotations(
  left: DocxGutterAnnotation,
  right: DocxGutterAnnotation
): number {
  const regionDelta =
    gutterAnnotationRegionRank(left) - gutterAnnotationRegionRank(right);
  if (regionDelta !== 0) {
    return regionDelta;
  }

  const leftTuple = gutterAnnotationSortTuple(left.location);
  const rightTuple = gutterAnnotationSortTuple(right.location);
  for (let index = 0; index < leftTuple.length; index += 1) {
    const delta = leftTuple[index] - rightTuple[index];
    if (delta !== 0) {
      return delta;
    }
  }

  const rankDelta =
    gutterAnnotationKindRank(left) - gutterAnnotationKindRank(right);
  if (rankDelta !== 0) {
    return rankDelta;
  }

  const detailDelta = gutterAnnotationDetailKey(left).localeCompare(
    gutterAnnotationDetailKey(right)
  );
  if (detailDelta !== 0) {
    return detailDelta;
  }

  return left.id.localeCompare(right.id);
}

/**
 * Builds the gutter annotation list for a single page: filters the document's
 * comments and tracked changes down to the nodes rendered on that page, then
 * orders them by document position (never by kind) and numbers them. The
 * result is deterministic for a given input, so card order stays stable across
 * re-renders and while the user pans or zooms.
 */
export function buildPageGutterAnnotations(
  options: PageGutterAnnotationsOptions
): DocxGutterAnnotation[] {
  const pageSegments = options.pageSegments;
  const annotations: DocxGutterAnnotation[] = [];

  if (options.showComments ?? true) {
    for (const comment of commentsVisibleInGutter(options)) {
      if (!commentBelongsToPage(comment, options)) {
        continue;
      }
      const location = resolveGutterAnnotationLocation(comment, -1);

      annotations.push({
        id: `comment:${comment.id}`,
        location,
        comment,
      });
    }
  }

  if (options.showTrackedChanges ?? true) {
    for (const change of options.trackedChanges ?? []) {
      const location = resolveGutterAnnotationLocation(change, -1);
      if (!trackedChangeBelongsToPageSegments(location, pageSegments)) {
        continue;
      }

      annotations.push({
        id: `change:${change.id}`,
        location: change.location,
        trackedChange: change,
      });
    }
  }

  annotations.sort(compareGutterAnnotations);
  annotations.forEach((annotation, index) => {
    annotation.ordinal = index + 1;
  });
  return annotations;
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

export function findGutterAnnotationScopeElementInPage(
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

function findCommentRegionAnchorElementInPage(
  pageElement: HTMLElement,
  annotation: DocxGutterAnnotation
): HTMLElement | undefined {
  const comment = annotation.comment;
  if (!comment?.region || comment.region === "body") {
    return undefined;
  }

  const regionRoot = pageElement.querySelector(
    `[data-docx-page-${comment.region}="true"]`
  ) as HTMLElement | null;
  if (!regionRoot) {
    return undefined;
  }

  const inlineCommentAnchor = findFirstElementWithSpaceSeparatedDataValue(
    regionRoot,
    "data-docx-comment-ids",
    String(comment.commentId)
  );
  if (inlineCommentAnchor) {
    return inlineCommentAnchor;
  }

  const nodeIndex =
    annotation.location.kind === "paragraph"
      ? annotation.location.nodeIndex
      : annotation.location.tableIndex;
  const nodeSelector =
    annotation.location.kind === "paragraph"
      ? `[data-docx-paragraph-host="true"][data-docx-${comment.region}-node-index="${nodeIndex}"]`
      : `[data-docx-table-host="true"][data-docx-${comment.region}-node-index="${nodeIndex}"]`;
  return (
    (regionRoot.querySelector(nodeSelector) as HTMLElement | null) ??
    regionRoot
  );
}

export function findGutterAnnotationAnchorElementInPage(
  pageElement: HTMLElement,
  annotation: DocxGutterAnnotation
): HTMLElement | undefined {
  const regionAnchor = findCommentRegionAnchorElementInPage(
    pageElement,
    annotation
  );
  if (regionAnchor) {
    return regionAnchor;
  }

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

// ---------------------------------------------------------------------------
// Page-space anchor resolution
// ---------------------------------------------------------------------------

/**
 * Resolves each annotation's anchor point in *page space* (CSS px measured
 * against the page surface's untransformed box, origin at the page's top-left).
 *
 * Page space means the result is independent of the viewer's zoom: the anchor
 * of a paragraph does not move when the page is scaled, so cards can be laid
 * out once per page geometry instead of re-deriving positions from client
 * coordinates at every zoom level.
 *
 * Prefers the inline anchor element (the highlighted run / revision markup)
 * and falls back to the annotation's paragraph or table cell, so a card still
 * has a position when the paragraph has been edited and the inline markup is
 * gone.
 */
export function resolveGutterAnnotationAnchorsInPage(
  pageElement: HTMLElement,
  annotations: readonly DocxGutterAnnotation[]
): Map<string, TrackedChangeAnchorPoint> {
  const anchors = new Map<string, TrackedChangeAnchorPoint>();
  for (const annotation of annotations) {
    const anchorElement = findGutterAnnotationAnchorElementInPage(
      pageElement,
      annotation
    );
    if (!anchorElement) {
      continue;
    }

    const rect = elementRectWithinContainer(anchorElement, pageElement);
    if (!rect) {
      continue;
    }

    // The left edge (not the right edge) identifies where the anchor starts on
    // its line, so same-line anchors sort in document order.
    const x = Math.round(rect.left);
    const y = Math.round(rect.top);
    if (!Number.isFinite(x) || !Number.isFinite(y)) {
      continue;
    }

    anchors.set(annotation.id, { x, y });
  }

  return anchors;
}

// ---------------------------------------------------------------------------
// Positioned annotation & connector-lane assignment (upstream 23368-23440)
// ---------------------------------------------------------------------------

export interface PositionedGutterAnnotation {
  annotation: DocxGutterAnnotation;
  anchorX: number;
  anchorY: number;
  top: number;
  heightPx: number;
  connectorLane: number;
}

/**
 * Vertical bucket used only for ordering cards. Anchors on the same text line
 * can differ by a pixel or two (run borders, sub/superscript); bucketing keeps
 * their relative order fixed instead of letting 1px jitter swap two cards.
 */
const GUTTER_ANNOTATION_SORT_BUCKET_PX = 4;

export function estimateTrackedChangeCardHeight(
  change: DocxTrackedChange
): number {
  const snippet =
    normalizeTrackedChangeSnippet(change.text) ??
    trackedChangeKindLabel(change.kind);
  const lines = Math.min(2, Math.max(1, Math.ceil(snippet.length / 42)));
  return Math.max(TRACKED_CHANGE_GUTTER_CARD_MIN_HEIGHT_PX, 22 + lines * 11);
}

export function assignGutterConnectorLanes(
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

// ---------------------------------------------------------------------------
// layoutTrackedChangesForPage (upstream 23442-23536)
// ---------------------------------------------------------------------------

/**
 * Lays out a page's gutter annotations: assigns anchor coordinates, estimates
 * card heights, stacks cards vertically with collision avoidance against the
 * page bottom, and assigns connector lanes. Pure with respect to the DOM —
 * callers supply pre-resolved anchor points and optional measured heights.
 *
 * All inputs and outputs are in page space (CSS px against the page's
 * untransformed box), never client/zoomed coordinates: the caller converts the
 * resulting `top`/`heightPx` for display when the page is rendered under a
 * scale transform.
 */
export function layoutTrackedChangesForPage(
  annotations: readonly DocxGutterAnnotation[],
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
  const resolvedAnchors = annotations.map((annotation) =>
    anchorByChangeId.get(annotation.id)
  );
  const hasResolvedAnchor = resolvedAnchors.some(Boolean);
  // When nothing measured (SSR, a hidden page, or a test renderer) the cards
  // are spread evenly down the page in document order. When only some anchors
  // measured, the unresolved ones inherit the nearest measured y instead of
  // jumping to a synthetic page-wide slot above or below their neighbours.
  const firstResolvedY = resolvedAnchors.find(Boolean)?.y;
  let lastResolvedY: number | undefined;
  const withAnchors = annotations.map((annotation, index) => {
    const defaultAnchor = Math.min(
      Math.max(10, Math.round((index + 1) * fallbackStride)),
      Math.max(10, pageHeightPx - 10)
    );
    const anchorPoint = resolvedAnchors[index];
    let anchorY: number;
    if (anchorPoint && Number.isFinite(anchorPoint.y)) {
      anchorY = anchorPoint.y;
      lastResolvedY = anchorPoint.y;
    } else if (hasResolvedAnchor && Number.isFinite(firstResolvedY)) {
      anchorY = lastResolvedY ?? (firstResolvedY as number);
    } else {
      anchorY = defaultAnchor;
    }
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

  // Total order: bucketed y, then x, then input (document) order. A comparator
  // with a tolerance band is not transitive, so it cannot produce a stable sort
  // and cards would swap places between renders.
  const orderById = new Map(
    annotations.map((annotation, index) => [annotation.id, index])
  );
  withAnchors.sort((left, right) => {
    const yDelta =
      Math.round(left.anchorY / GUTTER_ANNOTATION_SORT_BUCKET_PX) -
      Math.round(right.anchorY / GUTTER_ANNOTATION_SORT_BUCKET_PX);
    if (yDelta !== 0) {
      return yDelta;
    }

    const xDelta = left.anchorX - right.anchorX;
    if (xDelta !== 0) {
      return xDelta;
    }

    return (
      (orderById.get(left.annotation.id) ?? 0) -
      (orderById.get(right.annotation.id) ?? 0)
    );
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

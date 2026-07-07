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

export function findGutterAnnotationAnchorElementInPage(
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
 */
export function layoutTrackedChangesForPage(
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

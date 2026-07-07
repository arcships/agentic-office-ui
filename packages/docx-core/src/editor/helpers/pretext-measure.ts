// Pretext measurement integration: dual-wrapped floating-image geometry,
// page-flow wrap-obstacle collection, foreign-exclusion resolution, explicit
// column-break splitting, and alignment-aware pretext exclusion layout.
//
// Upstream editor.tsx:
//   5884-5903  (resolveDualWrapParagraphRenderBlockHeightPx)
//   5905-5930  (resolveForeignWrapExclusionsForFlowRange)
//   5932-5963  (resolveWrappedImageGeometryForPageFlow)
//   5965-6095  (collectPageFlowWrapObstaclesForParagraph)
//   6097-6129  (resolveParagraphForeignOnlyWrappedTextLayout)
//   6764-6844  (splitParagraphAtExplicitColumnBreaks)
//   7232-7386  (resolveDualWrappedFloatingImageGeometry)
//   7388-7543  (resolveParagraphDualWrappedTextLayout)
//   7545-7623  (resolveParagraphPretextExclusionLayout)
//
// Functions that depend on not-yet-ported modules
// (precomputePageSegmentForeignWrapExclusions,
// applyWrappedFloatingInteractionPreviewToParagraph,
// estimateParagraphContentHeightPx,
// projectParagraphConsumedHeightWithExplicitColumnBreaks) are deferred to the
// table-height / pagination-plan / paragraph-render modules. See
// docs/docx-editor-helpers-split-plan.md.

import type {
  ImageRunNode,
  ParagraphNode
} from "../../engine/types";
import type {
  PretextExclusionRect,
  PretextVariableWidthLayout
} from "../../viewer/pretext-layout";
import { resolveCaretRectAtOffset } from "../../viewer/pretext-selection";
import type {
  ParagraphLocation
} from "./editor-types";
import {
  MIN_DUAL_WRAPPED_INTERIOR_BAND_PX
} from "./pretext-build";
import {
  MIN_PARAGRAPH_LINE_HEIGHT_PX
} from "./constants";
import { clampNumber } from "./zoom-utils";
import {
  pretextLayoutContentBottomPx,
  wrappedPretextParagraphBlockHeightPx
} from "./line-height";
import {
  buildParagraphPretextLayoutSource,
  layoutParagraphPretextSource,
  paragraphChildAnchorOffset,
  type DualWrappedFloatingImageGeometry,
  type PageFlowFloatingWrapObstacle,
  type ParagraphDualWrappedTextLayout,
  type ParagraphPretextLayoutSource
} from "./pretext-build";
import {
  shouldRenderWrappedFloatingImage
} from "./paragraph-geometry";
import {
  syntheticTextBoxActsAsTopAndBottomMasthead
} from "./synthetic-textbox";
import {
  cloneFormFieldRun,
  cloneTextRunWithMetadata,
  imageLocationKey
} from "./text-mutation";
import { cloneParagraphStyle } from "./numbering";
import { paragraphHasExplicitColumnBreak } from "./paragraph-tracked";

// --- Dual-wrap block height (upstream 5884-5903) ---

export function resolveDualWrapParagraphRenderBlockHeightPx(
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

// --- Foreign wrap exclusions (upstream 5905-5930) ---

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

// --- Page-flow wrapped image geometry (upstream 5932-5963) ---

export function resolveWrappedImageGeometryForPageFlow(
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

// --- Collect page-flow wrap obstacles (upstream 5965-6095) ---

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

// --- Foreign-only wrapped text layout (upstream 6097-6129) ---

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

// --- Explicit column-break splitting (upstream 6764-6844) ---

export function splitParagraphAtExplicitColumnBreaks(
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

// --- Dual-wrapped floating image geometry (upstream 7232-7386) ---

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

// --- Dual-wrapped text layout (upstream 7388-7543) ---

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

// --- Pretext exclusion layout with alignment (upstream 7545-7623) ---

export function resolveParagraphPretextExclusionLayout(
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

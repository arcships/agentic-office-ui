// Cross-paragraph wrap clearance for floating images.
//
// A CSS float inside a paragraph does not grow that paragraph's box: the line
// boxes beside the float are shortened, but when the float is taller than the
// text the anchor paragraph stays short and the float overflows into the next
// flex item (every paragraph is an isolated flex item in the page body), which
// then paints its text on top of the image.
//
// Word instead keeps the following text clear of the float. With the flex page
// model we reproduce that vertically: a paragraph that would start inside a
// float which outlives its anchor paragraph is pushed down by the overlap.
//
// Known limitation: the float does not narrow the following paragraphs
// horizontally, so text flows below the float rather than beside its free band.
// See docs/docx-floating-images.md.

import type {
  NumberingDefinitionSet,
  ParagraphNode,
} from "../../engine/types";
import { MIN_PARAGRAPH_LINE_HEIGHT_PX } from "./constants";
import {
  shouldRenderWrappedFloatingImage,
} from "./paragraph-geometry";
import { estimateParagraphHeightPx } from "./line-height-table";

export interface DocxPageFlowWrapClearance {
  /** Space to insert above the paragraph so it clears an overflowing float. */
  marginTopPx: number;
}

/**
 * Vertical extent a paragraph's wrapped floating images reach below the
 * paragraph's own top, in paragraph-local pixels.
 *
 * Only floats that are painted outside the text flow and are not
 * behind-the-text backgrounds are considered: `topAndBottom` floats are already
 * an in-flow block (the anchor paragraph grows with them) and behind-text
 * images are meant to sit under the text.
 */
export function paragraphWrappedImageFlowBottomPx(
  paragraph: ParagraphNode
): number {
  let bottomPx = 0;
  for (const child of paragraph.children) {
    if (child.type !== "image") {
      continue;
    }
    if (!shouldRenderWrappedFloatingImage(child)) {
      continue;
    }
    const floating = child.floating;
    if (!floating) {
      continue;
    }
    const wrapType = floating.wrapType?.trim().toLowerCase();
    if (wrapType === "topandbottom" || floating.behindDocument === true) {
      continue;
    }

    const heightPx = Math.max(
      1,
      Math.round(child.heightPx ?? child.widthPx ?? MIN_PARAGRAPH_LINE_HEIGHT_PX)
    );
    const topPx =
      Math.max(0, Math.round(floating.yPx ?? 0)) +
      Math.max(0, Math.round(floating.distTPx ?? 0));
    const bottomSpacingPx = Math.max(0, Math.round(floating.distBPx ?? 0));
    bottomPx = Math.max(bottomPx, topPx + heightPx + bottomSpacingPx);
  }
  return bottomPx;
}

/**
 * Resolve the top clearance each paragraph in a page flow needs to stay clear
 * of floating images anchored in earlier paragraphs of the same flow.
 *
 * The estimates use the same paragraph height model as pagination, so the
 * rendered clearance and the page-break estimate agree.
 */
export function resolveWrappedImageFlowClearancePx(
  paragraphs: ParagraphNode[],
  availableWidthPx: number,
  numberingDefinitions?: NumberingDefinitionSet
): DocxPageFlowWrapClearance[] {
  const clearances: DocxPageFlowWrapClearance[] = [];
  let flowTopPx = 0;
  let obstacleBottomPx = Number.NEGATIVE_INFINITY;

  for (const paragraph of paragraphs) {
    const marginTopPx = Math.max(0, Math.round(obstacleBottomPx - flowTopPx));
    clearances.push({ marginTopPx });

    const paragraphTopPx = flowTopPx + marginTopPx;
    const paragraphHeightPx = Math.max(
      MIN_PARAGRAPH_LINE_HEIGHT_PX,
      estimateParagraphHeightPx(
        paragraph,
        availableWidthPx,
        numberingDefinitions
      )
    );
    flowTopPx = paragraphTopPx + paragraphHeightPx;

    const floatBottomPx = paragraphWrappedImageFlowBottomPx(paragraph);
    if (floatBottomPx > 0) {
      obstacleBottomPx = Math.max(
        obstacleBottomPx,
        paragraphTopPx + floatBottomPx
      );
    }
  }

  return clearances;
}

// Paragraph geometry: floating-image render-mode predicates, style helpers,
// image crop layout, and paragraph available-text-width calculation.
//
// Upstream editor.tsx:
//   6368-6377        (floatingImageMovesWithText)
//   14008-14052      (shouldRenderWrappedFloatingImage, isFixedPositionWrappedFloatingImage)
//   14054-14098      (shouldRenderTopAnchoredMarginFloatAsAbsolute)
//   14720-14785      (shouldRenderAbsoluteFloatingImage, isPageOrMarginAnchored*)
//   2045-2054        (shouldReserveHeaderFooterFloatingImageSpace)
//   14100-14886      (wrapped/absolute floating image styles)
//   9936-9980        (paragraphAvailableTextWidthPx)
//
// Split from paragraph-geometry.ts to stay within the ≤1000-line constraint.

import type {
  ImageRunNode,
  NumberingDefinitionSet,
  ParagraphNode,
} from "../../engine/types";
import { syntheticTextBoxContainsPictureLayer } from "./synthetic-textbox";
import { resolveListParagraphIndent } from "./xml-parsing-extra";
import { twipsToSignedPixels } from "./ooxml-helpers";
import { paragraphBorderInsetPx } from "./style-block-css";

export function floatingImageMovesWithText(
  floating?: ImageRunNode["floating"]
): boolean {
  if (!floating) {
    return true;
  }

  const verticalRelativeTo = floating.verticalRelativeTo?.trim().toLowerCase();
  return verticalRelativeTo !== "page" && verticalRelativeTo !== "margin";
}
export function shouldRenderWrappedFloatingImage(image: ImageRunNode): boolean {
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

export function isFixedPositionWrappedFloatingImage(image: ImageRunNode): boolean {
  return (
    shouldRenderWrappedFloatingImage(image) &&
    !floatingImageMovesWithText(image.floating)
  );
}
export function shouldRenderTopAnchoredMarginFloatAsAbsolute(
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
export function shouldRenderAbsoluteFloatingImage(image: ImageRunNode): boolean {
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

export function isPageOrMarginAnchoredAbsoluteFloatingImage(
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

export function isPageOrMarginAnchoredWrappedFloatingImage(
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
export function shouldReserveHeaderFooterFloatingImageSpace(
  image: ImageRunNode
): boolean {
  return (
    shouldRenderAbsoluteFloatingImage(image) &&
    Boolean(image.floating) &&
    image.floating?.behindDocument !== true
  );
}


/** Whether the paragraph contains a page/margin-anchored absolute floating image.
 *  Upstream line 14808; lives here because it depends on
 *  isPageOrMarginAnchoredAbsoluteFloatingImage (defined above). */
export function paragraphHasPageAnchoredAbsoluteFloatingImage(
  paragraph: ParagraphNode
): boolean {
  return paragraph.children.some(
    (child) =>
      child.type === "image" &&
      isPageOrMarginAnchoredAbsoluteFloatingImage(child)
  );
}

// ---- Render-focused floating-image style helpers ----
// Upstream editor.tsx: 14100-14148, 14150-14886

export function resolveWrappedFloatingSide(
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
  const intrinsicBlockWidthStyle: Record<string, string | number> =
    imageWidthPx ? { width: imageWidthPx } : { width: "fit-content" };
  const rightOffsetPx =
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

export function resolvePageSpanningAbsoluteFloatingDimensions(
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

export function imageCropLayout(
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

// --- Paragraph available text width (upstream 9936-9980) ---

export function paragraphAvailableTextWidthPx(
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

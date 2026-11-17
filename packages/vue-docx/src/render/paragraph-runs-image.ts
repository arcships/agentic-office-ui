// Image run rendering for renderParagraphRuns.
// Upstream editor.tsx: 18639-18987 (image branch of renderRun).
//
// Translates JSX to Vue h() calls.

import { h, type VNode } from "vue"
import { officeMessagesEn } from "@arcships/office-interaction"
import type { ImageRunNode, ParagraphNode } from "@arcships/docx-core"
import type { DocxSectionImageLocation } from "@arcships/docx-core"
import {
  shouldRenderWrappedFloatingImage,
  shouldRenderAbsoluteFloatingImage,
  shouldRenderTopAnchoredMarginFloatAsAbsolute,
  isPageOrMarginAnchoredWrappedFloatingImage,
  imageBehavesAsDecorativeBehindTextBackground,
  wrappedFloatingImageStyle,
  absoluteFloatingImageStyle,
  resolvePageSpanningAbsoluteFloatingDimensions,
  resolveImageFrameSizePx,
  imageCropLayout,
  resolveRenderableImageSource,
  imageUsesPlaceholderFallback,
  unsupportedImageFallbackLabel,
  appendCssFilters,
  resolveImageRenderTransformStyle,
  syntheticTextBoxSvg,
  sectionImageLocationKey,
} from "@arcships/docx-core"

export interface ImageRunRenderExtra {
  /** Optional overlay representing a section-level image move preview. */
  movePreview?: {
    imageKey: string
    deltaX: number
    deltaY: number
  } | undefined
  /** CSS filter suffix injected by the host when images are exported. */
  imageFilterSuffix?: string
  /** Reports a browser decode failure without exposing document bytes. */
  onImageDecodeError?: (error: DocxImageDecodeError) => void

  // Section image interaction callbacks
  sectionImageInteraction?: ImageInteraction | undefined
}

export interface DocxImageDecodeError {
  code: "IMAGE_DECODE_FAILED"
  phase: "image-decode"
  imageKey: string
  message: string
}

export interface ImageInteraction {
  isReadOnly?: boolean
  imageLocationForChild?: (childIndex: number) => DocxSectionImageLocation | undefined
  floatingMovePreview?: {
    imageKey: string
    deltaX: number
    deltaY: number
  }
  onImagePointerDown?: (
    event: PointerEvent,
    location: DocxSectionImageLocation,
    image: ImageRunNode,
    isWrapped: boolean,
    isAbsolute: boolean
  ) => void
  onImageClick?: (location: DocxSectionImageLocation) => void
}

export interface ImageRunRenderContext {
  paragraph: ParagraphNode
  child: ImageRunNode
  key: string
  childIndex: number
  pageNumber?: number
  totalPages?: number
  pageNumberFormat?: string
  resolveStyleRefFieldValue?: (target: string) => string | undefined
  floatingPageOriginPx?: {
    left: number
    top: number
    marginLeft?: number
    marginTop?: number
    columnLeft?: number
    columnTop?: number
    pageWidth?: number
  }
  floatingAnchorOriginCorrectionXPx: number
  withinHeaderFooter?: boolean
  headerFooterRegion?: "header" | "footer"
  /** Bounds of the table cell that contains this paragraph, when applicable. */
  tableCellLayout?: {
    widthPx?: number
    heightPx?: number
  }
  extra?: ImageRunRenderExtra
}

export interface ImageRunRenderResult {
  vnode: VNode
  isWrappedFloating: boolean
  isAbsoluteFloating: boolean
  inlineAdvancePx?: number
}

/**
 * Render a single image run child as a Vue VNode.
 */
export function renderImageRun(ctx: ImageRunRenderContext): ImageRunRenderResult {
  const { child, paragraph, key, childIndex, extra } = ctx

  const behavesAsDecorativeBehindTextBackground =
    imageBehavesAsDecorativeBehindTextBackground(child, paragraph)

  const sectionImageInteraction = extra?.sectionImageInteraction
  const sectionImageLocation =
    sectionImageInteraction?.imageLocationForChild?.(childIndex)
  const sectionImageKey = sectionImageLocation
    ? sectionImageLocationKey(sectionImageLocation)
    : undefined
  const movePreview =
    sectionImageKey &&
    extra?.sectionImageInteraction?.floatingMovePreview?.imageKey ===
      sectionImageKey
      ? extra.sectionImageInteraction.floatingMovePreview
      : undefined

  const forceAbsoluteFixedPositionSectionFloat =
    ctx.withinHeaderFooter === true &&
    isPageOrMarginAnchoredWrappedFloatingImage(child)

  const forceWrappedTopAnchoredSectionFloat =
    ctx.withinHeaderFooter === true &&
    Boolean(child.floating?.wrapType && child.floating.wrapType !== "none") &&
    shouldRenderTopAnchoredMarginFloatAsAbsolute(child)

  const isWrappedFloatingImage =
    !forceAbsoluteFixedPositionSectionFloat &&
    (forceWrappedTopAnchoredSectionFloat
      ? true
      : shouldRenderWrappedFloatingImage(child))

  const isAbsoluteFloatingImage = forceAbsoluteFixedPositionSectionFloat
    ? true
    : forceWrappedTopAnchoredSectionFloat
    ? false
    : shouldRenderAbsoluteFloatingImage(child)

  const imageLayoutAttributes: Record<string, string> = {
    "data-docx-image-child-index": String(childIndex),
    "data-docx-image-layout": isWrappedFloatingImage
      ? "wrapped"
      : isAbsoluteFloatingImage
      ? "absolute"
      : "inline",
  }

  const tableCellLayout = ctx.tableCellLayout
  const layoutInCellActive =
    isAbsoluteFloatingImage &&
    child.floating?.layoutInCell !== false &&
    Number.isFinite(tableCellLayout?.widthPx) &&
    (tableCellLayout?.widthPx as number) > 0
  if (layoutInCellActive) {
    imageLayoutAttributes["data-docx-layout-in-cell"] = "true"
  }

  const horizontalRelativeTo =
    child.floating?.horizontalRelativeTo?.toLowerCase()
  let widthPx = child.widthPx
  let heightPx = child.heightPx
  ;({ widthPx, heightPx } = resolvePageSpanningAbsoluteFloatingDimensions(
    child,
    widthPx,
    heightPx,
    ctx.floatingPageOriginPx
  ))

  // A 90°/270° drawing rotation turns the visual footprint a quarter turn. Lay
  // the frame out swapped and rotate the image content into it, otherwise the
  // rotated image keeps its unrotated box, overflows it, and paints over the
  // neighbouring runs.
  const contentWidthPx = widthPx
  const contentHeightPx = heightPx
  const imageFrameSize = resolveImageFrameSizePx(
    child,
    contentWidthPx,
    contentHeightPx
  )
  const frameWidthPx = imageFrameSize.widthPx
  const frameHeightPx = imageFrameSize.heightPx
  const rotationSwapsFrame = imageFrameSize.rotationSwapped

  const usesExternalHorizontalAnchorOrigin =
    horizontalRelativeTo === "page" ||
    horizontalRelativeTo === "margin" ||
    horizontalRelativeTo === "column"
  const horizontalAnchorCorrectionPx = usesExternalHorizontalAnchorOrigin
    ? ctx.floatingAnchorOriginCorrectionXPx
    : 0

  const absoluteColumnOriginLeftPx =
    ctx.withinHeaderFooter === true &&
    ctx.headerFooterRegion === "header" &&
    child.floating?.behindDocument === true &&
    horizontalRelativeTo === "column"
      ? ctx.floatingPageOriginPx?.marginLeft
      : ctx.floatingPageOriginPx?.columnLeft

  const floatingStyle = isWrappedFloatingImage
    ? wrappedFloatingImageStyle(child, {
        containerWidthPx: ctx.floatingPageOriginPx?.pageWidth,
        deltaX: (movePreview?.deltaX ?? 0) + horizontalAnchorCorrectionPx,
        deltaY: movePreview?.deltaY ?? 0,
        allowNegativeOffsets: true,
      })
    : isAbsoluteFloatingImage
    ? absoluteFloatingImageStyle(child, {
        pageOriginLeft: ctx.floatingPageOriginPx?.left,
        pageOriginTop: ctx.floatingPageOriginPx?.top,
        marginOriginLeft: ctx.floatingPageOriginPx?.marginLeft,
        marginOriginTop: ctx.floatingPageOriginPx?.marginTop,
        columnOriginLeft: absoluteColumnOriginLeftPx,
        columnOriginTop: ctx.floatingPageOriginPx?.columnTop,
        deltaX: (movePreview?.deltaX ?? 0) + horizontalAnchorCorrectionPx,
        deltaY: movePreview?.deltaY ?? 0,
      })
    : movePreview
    ? {
        transform: `translate(${movePreview.deltaX}px, ${movePreview.deltaY}px)`,
        position: "relative" as const,
        zIndex: 4,
      }
    : {}

  let floatingStyleWithMovePreview = isWrappedFloatingImage
    ? { ...floatingStyle }
    : floatingStyle

  if (layoutInCellActive) {
    floatingStyleWithMovePreview = constrainFloatingImageToTableCell(
      floatingStyleWithMovePreview,
      child,
      tableCellLayout,
      widthPx,
      heightPx
    )
  }

  const onSectionImagePointerDown =
    sectionImageLocation &&
    sectionImageInteraction?.onImagePointerDown &&
    !sectionImageInteraction.isReadOnly &&
    !behavesAsDecorativeBehindTextBackground
      ? (event: PointerEvent) => {
          sectionImageInteraction.onImagePointerDown?.(
            event,
            sectionImageLocation,
            child,
            isWrappedFloatingImage,
            isAbsoluteFloatingImage
          )
        }
      : undefined

  const onSectionImageClick =
    sectionImageLocation &&
    sectionImageInteraction?.onImageClick &&
    !sectionImageInteraction.isReadOnly &&
    !behavesAsDecorativeBehindTextBackground
      ? (event: MouseEvent) => {
          event.stopPropagation()
          sectionImageInteraction.onImageClick?.(sectionImageLocation)
        }
      : undefined

  const sectionImageCursor =
    sectionImageLocation &&
    !sectionImageInteraction?.isReadOnly &&
    !behavesAsDecorativeBehindTextBackground &&
    (isWrappedFloatingImage || isAbsoluteFloatingImage)
      ? "move"
      : undefined

  const isCenteredStandaloneInlineImage =
    !isWrappedFloatingImage &&
    !isAbsoluteFloatingImage &&
    paragraph.style?.align === "center" &&
    paragraph.children.length === 1 &&
    paragraph.children[0]?.type === "image"

  const decorativeStyle = behavesAsDecorativeBehindTextBackground
    ? { pointerEvents: "none" as const, userSelect: "none" as const }
    : undefined

  // Resolve image source
  const renderableImageSrc =
    syntheticTextBoxSvg(
      child,
      ctx.pageNumber,
      ctx.totalPages,
      ctx.pageNumberFormat,
      ctx.resolveStyleRefFieldValue
    ) ?? resolveRenderableImageSource(child)

  // Missing image
  if (!child.src) {
    return {
      vnode: h(
        "span",
        {
          key,
          ...imageLayoutAttributes,
          style: buildMissingImageStyle(
            isWrappedFloatingImage,
            isAbsoluteFloatingImage,
            floatingStyleWithMovePreview,
            sectionImageCursor,
            decorativeStyle
          ),
          onPointerdown: onSectionImagePointerDown,
          onClick: onSectionImageClick,
        },
        "Missing image"
      ),
      isWrappedFloating: isWrappedFloatingImage,
      isAbsoluteFloating: isAbsoluteFloatingImage,
    }
  }

  // Placeholder fallback
  if (
    imageUsesPlaceholderFallback(child) ||
    (child.src && !renderableImageSrc)
  ) {
    const isSmallIcon = (widthPx ?? 0) <= 56 && (heightPx ?? 0) <= 56
    const fallbackText = unsupportedImageFallbackLabel(child, widthPx, heightPx)
    return {
      vnode: h(
        "span",
        {
          key,
          ...imageLayoutAttributes,
          role: "img",
          "aria-label": child.alt ?? "DOCX image",
          style: buildPlaceholderStyle(
            frameWidthPx,
            frameHeightPx,
            isSmallIcon,
            isWrappedFloatingImage,
            isAbsoluteFloatingImage,
            floatingStyleWithMovePreview,
            sectionImageCursor,
            fallbackText,
            decorativeStyle
          ),
          onPointerdown: onSectionImagePointerDown,
          onClick: onSectionImageClick,
        },
        fallbackText
      ),
      isWrappedFloating: isWrappedFloatingImage,
      isAbsoluteFloating: isAbsoluteFloatingImage,
    }
  }

  // Rendered image
  const cropLayout = imageCropLayout(child, contentWidthPx, contentHeightPx)
  const imageVisualStyle: Record<string, string | number | undefined> = {
    filter: appendCssFilters(child.cssFilter, extra?.imageFilterSuffix),
    opacity: child.cssOpacity,
  }
  // The rotated image content keeps the unrotated (content) box; the frame that
  // holds it is swapped. `fillFrame: true` returns exactly that content box plus
  // the quarter-turn transform that lands it inside the swapped frame.
  const rotatedFrameTransformStyle = rotationSwapsFrame
    ? resolveImageRenderTransformStyle(child, {
        frameWidthPx: cropLayout
          ? cropLayout.frameHeightPx
          : frameWidthPx,
        frameHeightPx: cropLayout
          ? cropLayout.frameWidthPx
          : frameHeightPx,
        fillFrame: true,
        baseTransform: cropLayout
          ? `translate(${-cropLayout.offsetXPx}px, ${-cropLayout.offsetYPx}px)`
          : undefined,
      })
    : undefined
  const imageTransformStyle = rotationSwapsFrame
    ? undefined
    : resolveImageRenderTransformStyle(child, {
        frameWidthPx: cropLayout?.frameWidthPx ?? widthPx,
        frameHeightPx: cropLayout?.frameHeightPx ?? heightPx,
        baseTransform: cropLayout
          ? `translate(${-cropLayout.offsetXPx}px, ${-cropLayout.offsetYPx}px)`
          : undefined,
      })
  const onImageLoad = (event: Event) => {
    const image = event.currentTarget as HTMLImageElement | null
    if (image) image.dataset.imageState = "ready"
  }
  const onImageError = (event: Event) => {
    const image = event.currentTarget as HTMLImageElement | null
    if (image) {
      image.dataset.imageState = "error"
      image.setAttribute("aria-invalid", "true")
    }
    extra?.onImageDecodeError?.({
      code: "IMAGE_DECODE_FAILED",
      phase: "image-decode",
      imageKey: key,
      message: officeMessagesEn.docx.diagnostics.imageDecodeFailed,
    })
  }

  if (cropLayout) {
    const cropImageVNode = h("img", {
      src: renderableImageSrc,
      alt: child.alt ?? "DOCX image",
      draggable: false,
      loading: "lazy",
      decoding: "async",
      "data-image-state": "loading",
      onLoad: onImageLoad,
      onError: onImageError,
      style: {
        width: `${cropLayout.imageWidthPx}px`,
        height: `${cropLayout.imageHeightPx}px`,
        maxWidth: "none",
        display: "block",
        ...imageVisualStyle,
        ...imageTransformStyle,
      },
    })

    // The crop window itself is rotated into the swapped frame, so the floating
    // box (and therefore the drag hit box) matches the rotated footprint.
    const cropFrameStyle = buildCropFrameStyle(
      cropLayout,
      rotationSwapsFrame ? false : isCenteredStandaloneInlineImage,
      isWrappedFloatingImage,
      isAbsoluteFloatingImage,
      rotationSwapsFrame ? {} : floatingStyleWithMovePreview,
      rotationSwapsFrame ? undefined : sectionImageCursor,
      rotationSwapsFrame ? undefined : decorativeStyle
    )
    const cropContents = rotationSwapsFrame
      ? [
          h(
            "span",
            { style: { ...cropFrameStyle, ...rotatedFrameTransformStyle } },
            [cropImageVNode]
          ),
        ]
      : [cropImageVNode]

    if (!rotationSwapsFrame) {
      return {
        vnode: h(
          "span",
          {
            key,
            ...imageLayoutAttributes,
            style: cropFrameStyle,
            onPointerdown: onSectionImagePointerDown,
            onClick: onSectionImageClick,
          },
          cropContents
        ),
        isWrappedFloating: isWrappedFloatingImage,
        isAbsoluteFloating: isAbsoluteFloatingImage,
      }
    }

    return {
      vnode: h(
        "span",
        {
          key,
          ...imageLayoutAttributes,
          style: buildRotationFrameStyle(
            frameWidthPx ?? cropLayout.frameHeightPx,
            frameHeightPx ?? cropLayout.frameWidthPx,
            isCenteredStandaloneInlineImage,
            isWrappedFloatingImage,
            isAbsoluteFloatingImage,
            floatingStyleWithMovePreview,
            sectionImageCursor,
            decorativeStyle
          ),
          onPointerdown: onSectionImagePointerDown,
          onClick: onSectionImageClick,
        },
        cropContents
      ),
      isWrappedFloating: isWrappedFloatingImage,
      isAbsoluteFloating: isAbsoluteFloatingImage,
    }
  }

  if (rotationSwapsFrame) {
    return {
      vnode: h(
        "span",
        {
          key,
          ...imageLayoutAttributes,
          style: buildRotationFrameStyle(
            frameWidthPx ?? 0,
            frameHeightPx ?? 0,
            isCenteredStandaloneInlineImage,
            isWrappedFloatingImage,
            isAbsoluteFloatingImage,
            floatingStyleWithMovePreview,
            sectionImageCursor,
            decorativeStyle
          ),
          onPointerdown: onSectionImagePointerDown,
          onClick: onSectionImageClick,
        },
        [
          h("img", {
            src: renderableImageSrc,
            alt: child.alt ?? "DOCX image",
            draggable: false,
            loading: "lazy",
            decoding: "async",
            "data-image-state": "loading",
            onLoad: onImageLoad,
            onError: onImageError,
            style: {
              display: "block",
              maxWidth: "none",
              ...imageVisualStyle,
              ...rotatedFrameTransformStyle,
            },
          }),
        ]
      ),
      isWrappedFloating: isWrappedFloatingImage,
      isAbsoluteFloating: isAbsoluteFloatingImage,
      inlineAdvancePx:
        !isWrappedFloatingImage && !isAbsoluteFloatingImage
          ? (frameWidthPx ?? 0)
          : undefined,
    }
  }

  return {
    vnode: h("img", {
      key,
      ...imageLayoutAttributes,
      src: renderableImageSrc,
      alt: child.alt ?? "DOCX image",
      draggable: false,
      loading: "lazy",
      decoding: "async",
      "data-image-state": "loading",
      onLoad: onImageLoad,
      onError: onImageError,
      style: buildDirectImageStyle(
        widthPx,
        heightPx,
        isCenteredStandaloneInlineImage,
        isWrappedFloatingImage,
        isAbsoluteFloatingImage,
        floatingStyleWithMovePreview,
        sectionImageCursor,
        decorativeStyle,
        imageVisualStyle,
        resolveImageRenderTransformStyle(child, {
          frameWidthPx: widthPx,
          frameHeightPx: heightPx,
        })
      ),
      onPointerdown: onSectionImagePointerDown,
      onClick: onSectionImageClick,
    }),
    isWrappedFloating: isWrappedFloatingImage,
    isAbsoluteFloating: isAbsoluteFloatingImage,
    inlineAdvancePx:
      !isWrappedFloatingImage && !isAbsoluteFloatingImage
        ? (widthPx ?? 0)
        : undefined,
  }
}

// ---- Style builders ----

function positivePixels(value: number | undefined): number | undefined {
  return Number.isFinite(value) && (value as number) > 0
    ? Math.max(1, Math.round(value as number))
    : undefined
}

/**
 * Keep a floating image inside its table cell when `wp:anchor/@layoutInCell`
 * is enabled (the OOXML default). Without an explicit cell width the image is
 * still clipped by the cell's positioned container in the host component.
 */
function constrainFloatingImageToTableCell(
  style: Record<string, string | number | undefined>,
  image: ImageRunNode,
  tableCellLayout: { widthPx?: number; heightPx?: number } | undefined,
  widthPx: number | undefined,
  heightPx: number | undefined
): Record<string, string | number | undefined> {
  const cellWidthPx = positivePixels(tableCellLayout?.widthPx)
  if (!cellWidthPx) {
    return style
  }

  const next: Record<string, string | number | undefined> = {
    ...style,
    maxWidth: `${cellWidthPx}px`,
  }
  const floating = image.floating
  const horizontalRelativeTo = floating?.simplePos
    ? "page"
    : floating?.horizontalRelativeTo?.toLowerCase()
  const externalHorizontalAnchor =
    horizontalRelativeTo === "page" || horizontalRelativeTo === "margin"
  const imageWidthPx = Math.max(1, Math.round(widthPx ?? image.widthPx ?? cellWidthPx))
  if (!externalHorizontalAnchor && typeof next.left === "number") {
    next.left = Math.max(
      0,
      Math.min(next.left, Math.max(0, cellWidthPx - imageWidthPx))
    )
  }

  const cellHeightPx = positivePixels(tableCellLayout?.heightPx)
  if (cellHeightPx) {
    next.maxHeight = `${cellHeightPx}px`
    const verticalRelativeTo = floating?.simplePos
      ? "page"
      : floating?.verticalRelativeTo?.toLowerCase()
    const externalVerticalAnchor =
      verticalRelativeTo === "page" || verticalRelativeTo === "margin"
    const imageHeightPx = Math.max(
      1,
      Math.round(heightPx ?? image.heightPx ?? cellHeightPx)
    )
    if (!externalVerticalAnchor && typeof next.top === "number") {
      next.top = Math.max(
        0,
        Math.min(next.top, Math.max(0, cellHeightPx - imageHeightPx))
      )
    }
  }

  return next
}

function buildMissingImageStyle(
  isWrapped: boolean,
  isAbsolute: boolean,
  floatingStyle: Record<string, string | number | undefined>,
  cursor: string | undefined,
  decorative?: Record<string, string>
): Record<string, string | number | undefined> {
  return {
    display: "inline-flex",
    minWidth: 112,
    minHeight: 80,
    alignItems: "center",
    justifyContent: "center",
    border: "1px dashed #d1d5db",
    borderRadius: 4,
    color: "#6b7280",
    fontSize: 12,
    marginInline: isWrapped || isAbsolute ? undefined : 0,
    paddingInline: 8,
    ...floatingStyle,
    cursor,
    ...decorative,
  }
}

function buildPlaceholderStyle(
  widthPx: number | undefined,
  heightPx: number | undefined,
  isSmallIcon: boolean,
  isWrapped: boolean,
  isAbsolute: boolean,
  floatingStyle: Record<string, string | number | undefined>,
  cursor: string | undefined,
  _fallbackText: string,
  decorative?: Record<string, string>
): Record<string, string | number | undefined> {
  return {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    width: widthPx ? `${widthPx}px` : "1.8em",
    height: heightPx ? `${heightPx}px` : "1.8em",
    minWidth: 16,
    minHeight: 16,
    border: "1px solid #d1d5db",
    borderRadius: 3,
    backgroundColor: "#ffffff",
    color: "#0f172a",
    fontSize: isSmallIcon ? 12 : 10,
    fontWeight: 700,
    textTransform: "lowercase",
    fontFamily: "Arial, sans-serif",
    lineHeight: 1,
    verticalAlign: "middle",
    marginInline: isWrapped || isAbsolute ? undefined : 0,
    ...floatingStyle,
    cursor,
    ...decorative,
  }
}

function buildCropFrameStyle(
  cropLayout: {
    frameWidthPx: number
    frameHeightPx: number
  },
  isCentered: boolean,
  isWrapped: boolean,
  isAbsolute: boolean,
  floatingStyle: Record<string, string | number | undefined>,
  cursor: string | undefined,
  decorative?: Record<string, string>
): Record<string, string | number | undefined> {
  return {
    width: `${cropLayout.frameWidthPx}px`,
    height: `${cropLayout.frameHeightPx}px`,
    overflow: "clip",
    display: isCentered ? "block" : "inline-block",
    marginLeft: isCentered ? "auto" : undefined,
    marginRight: isCentered ? "auto" : undefined,
    verticalAlign: "middle",
    marginInline: isWrapped || isAbsolute ? undefined : 0,
    ...floatingStyle,
    cursor,
    ...decorative,
  }
}

/**
 * Outer frame of a quarter-turned image. It carries the layout box
 * (width/height swapped for the rotation) plus the floating/drag styles, while
 * the rotated content inside is sized to the unrotated content box.
 */
function buildRotationFrameStyle(
  frameWidthPx: number,
  frameHeightPx: number,
  isCentered: boolean,
  isWrapped: boolean,
  isAbsolute: boolean,
  floatingStyle: Record<string, string | number | undefined>,
  cursor: string | undefined,
  decorative?: Record<string, string>
): Record<string, string | number | undefined> {
  return {
    display: isCentered ? "block" : "inline-block",
    marginLeft: isCentered ? "auto" : undefined,
    marginRight: isCentered ? "auto" : undefined,
    verticalAlign: "middle",
    lineHeight: 0,
    marginInline: isWrapped || isAbsolute ? undefined : 0,
    ...floatingStyle,
    // The swapped frame is the layout box: it must win over the float's
    // intrinsic (unrotated) width.
    width: `${Math.max(1, Math.round(frameWidthPx))}px`,
    height: `${Math.max(1, Math.round(frameHeightPx))}px`,
    cursor,
    ...decorative,
  }
}

function buildDirectImageStyle(
  widthPx: number | undefined,
  heightPx: number | undefined,
  isCentered: boolean,
  isWrapped: boolean,
  isAbsolute: boolean,
  floatingStyle: Record<string, string | number | undefined>,
  cursor: string | undefined,
  decorative: Record<string, string> | undefined,
  imageVisualStyle: Record<string, string | number | undefined>,
  transformStyle: Record<string, string | number | undefined>
): Record<string, string | number | undefined> {
  return {
    width: widthPx ? `${widthPx}px` : undefined,
    height: heightPx ? `${heightPx}px` : undefined,
    maxWidth: isWrapped || isAbsolute ? undefined : "100%",
    display: isCentered ? "block" : undefined,
    marginLeft: isCentered ? "auto" : undefined,
    marginRight: isCentered ? "auto" : undefined,
    verticalAlign: "middle",
    marginInline: isWrapped || isAbsolute ? undefined : 0,
    ...imageVisualStyle,
    ...transformStyle,
    ...floatingStyle,
    cursor,
    ...decorative,
  }
}

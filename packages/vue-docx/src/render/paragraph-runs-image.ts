// Image run rendering for renderParagraphRuns.
// Upstream editor.tsx: 18639-18987 (image branch of renderRun).
//
// Translates JSX to Vue h() calls.

import { h, type VNode } from "vue"
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

  const imageLayoutAttributes = {
    "data-docx-image-child-index": String(childIndex),
    "data-docx-image-layout": isWrappedFloatingImage
      ? "wrapped"
      : isAbsoluteFloatingImage
      ? "absolute"
      : "inline",
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

  const floatingStyleWithMovePreview = isWrappedFloatingImage
    ? { ...floatingStyle }
    : floatingStyle

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
            widthPx,
            heightPx,
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
  const cropLayout = imageCropLayout(child, widthPx, heightPx)
  const imageVisualStyle: Record<string, string | number | undefined> = {
    filter: appendCssFilters(child.cssFilter, extra?.imageFilterSuffix),
    opacity: child.cssOpacity,
  }
  const imageTransformStyle = resolveImageRenderTransformStyle(child, {
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
      message: "DOCX 图片无法解码。",
    })
  }

  if (cropLayout) {
    return {
      vnode: h(
        "span",
        {
          key,
          ...imageLayoutAttributes,
          style: buildCropFrameStyle(
            cropLayout,
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
              width: `${cropLayout.imageWidthPx}px`,
              height: `${cropLayout.imageHeightPx}px`,
              maxWidth: "none",
              display: "block",
              ...imageVisualStyle,
              ...imageTransformStyle,
            },
          }),
        ]
      ),
      isWrappedFloating: isWrappedFloatingImage,
      isAbsoluteFloating: isAbsoluteFloatingImage,
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

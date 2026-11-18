// renderParagraphRuns main module.
// Upstream editor.tsx: 18109-19581.
//
// Rewrites JSX to Vue h() calls. Orchestrates text/image/form-field
// run rendering via type-specific sub-modules.

import { h, type VNode } from "vue"
import { sanitizeOfficeLink } from "@arcships/office-runtime/link-policy"
import { officeMessagesEn } from "@arcships/office-interaction"
import type {
  ParagraphNode,
  TextRunNode,
  FormFieldRunNode,
  NumberingDefinitionSet,
  ParagraphNumberingLabel,
} from "@arcships/docx-core"
import type { DocxDocumentTheme } from "@arcships/docx-core"
import type { ParagraphTrackedInlineChange } from "@arcships/docx-core"
import type { ParagraphCommentTextRange } from "@arcships/docx-core"
import {
  paragraphUsesTabLeaders,
  paragraphAnchoredTabLayout,
  paragraphPageFieldSequence,
  paragraphPageFieldValueSequence,
  paragraphStyleRefFieldValueSequence,
  paragraphLeadingTabStop,
  paragraphFirstTabStopPx,
  DEFAULT_TAB_STOP_PX,
  resolveParagraphFirstLineLeftTabStopsPx,
  paragraphLooksLikeCheckboxChoiceRow,
  checkboxChoiceRowTabWidthPx,
  updateEstimatedLineWidthPxForText,
  resolveTabSpacerWidthPx,
  resolveParagraphTrackedMarkup,
  resolveParagraphCommentMarkup,
  trackedDeletedStyle,
  runStyleToCss,
  linkStyleToCss,
  trackedInlineStyle,
  commentHighlightStyle,
  tableOfContentsLevel,
  isTableOfContentsParagraph,
  themedRunColor,
  numberingMarkerStyle,
  tabLeaderStyle,
  formFieldDisplayValue,
  attachTextToPreviousCheckbox,
  shouldRenderWrappedFloatingImage,
  isJustifiedAlignment,
} from "@arcships/docx-core"
import type { ImageRunRenderExtra } from "./paragraph-runs-image"
import { renderImageRun } from "./paragraph-runs-image"
import {
  type FieldResolutionContext,
  type FieldResolutionState,
  resolveFieldText,
} from "./paragraph-runs-field-resolve"
import { renderParagraphMarkChange } from "./paragraph-runs-annotation"

// ---- Public types ----

export interface ParagraphRunRenderOptions {
  showTrackedChanges?: boolean
  showCommentHighlights?: boolean
  numberingDefinitions?: NumberingDefinitionSet
  tocLinkColorByLevel?: Partial<Record<number, string | undefined>>
  trackedMarkupMode?: "inline" | "gutter"
  withinHeaderFooter?: boolean
  headerFooterRegion?: "header" | "footer"
  pageNumberFormat?: string
  sectionPages?: number
  resolveStyleRefFieldValue?: (target: string) => string | undefined
  resolveBookmarkFieldValue?: (target: string) => string | undefined
  resolveBookmarkPageNumber?: (target: string) => number | undefined
  floatingAnchorOriginCorrectionXPx?: number
  sectionImageInteraction?: ImageRunRenderExtra["sectionImageInteraction"]
  paragraphOriginLeftPx?: number
  paragraphOriginTopPx?: number
  imageFilterSuffix?: string
  onImageDecodeError?: ImageRunRenderExtra["onImageDecodeError"]
  /** Bounds of the table cell that hosts this paragraph. */
  tableCellLayout?: {
    widthPx?: number
    heightPx?: number
  }
  /** Floating body images are rendered by the editor interaction layer. */
  skipFloatingImages?: boolean
  /** Keep wrapped images in document flow while the interaction layer supplies controls. */
  renderWrappedFloatingImages?: boolean
  /**
   * Pre-formatted footnote/endnote marker labels keyed by the raw w:id, as
   * resolved by `buildDocumentPageNotes`. Without them the marker falls back to
   * the raw id, which is not the displayed note number.
   */
  noteMarkerLabels?: NoteMarkerLabels
  /** Child indexes rendered by another surface layer (for example a page cover). */
  skipImageChildIndexes?: readonly number[]
  /** Localized alt text and missing-source fallback for images. */
  imageAltText?: string
  imageMissingText?: string
  /**
   * Document `w:defaultTabStop` in CSS pixels; Word's 48px grid when omitted.
   * Tabs that run past the last explicit stop keep advancing on this grid.
   */
  defaultTabStopPx?: number
}

export interface NoteMarkerLabels {
  footnote?: Map<number, string>
  endnote?: Map<number, string>
}

// ---- Main function ----

export function renderParagraphRuns(
  paragraph: ParagraphNode,
  keyPrefix: string,
  documentTheme: DocxDocumentTheme = "light",
  numberingLabel?: ParagraphNumberingLabel,
  onInternalLinkClick?: (bookmarkName: string) => void,
  floatingPageOriginPx?: {
    left: number
    top: number
    marginLeft?: number
    marginTop?: number
    columnLeft?: number
    columnTop?: number
    pageWidth?: number
  },
  noteMarkerIndexes?: {
    footnote: Map<number, number>
    endnote: Map<number, number>
  },
  pageNumber?: number,
  totalPages?: number,
  options?: ParagraphRunRenderOptions
): VNode | VNode[] {
  const runs: VNode[] = []
  const runsLeft: VNode[] = []
  const runsRight: VNode[] = []

  const safeNoteMarkerIndexes = {
    footnote: noteMarkerIndexes?.footnote ?? new Map<number, number>(),
    endnote: noteMarkerIndexes?.endnote ?? new Map<number, number>(),
  }

  // Tab layout detection
  const useTabLeaderLayout = paragraphUsesTabLeaders(paragraph)
  const anchoredTabLayout = paragraphAnchoredTabLayout(paragraph, {
    withinHeaderFooter: options?.withinHeaderFooter,
  })
  const useCenterRightTabLayout =
    !useTabLeaderLayout && anchoredTabLayout === "center-right"
  const useCenterTabLayout =
    !useTabLeaderLayout && anchoredTabLayout === "center"
  const useRightTabLayout =
    !useTabLeaderLayout && anchoredTabLayout === "right"
  const useAnchoredTabLayout =
    useCenterRightTabLayout || useCenterTabLayout || useRightTabLayout

  // Field resolution context (extracted to paragraph-runs-field-resolve.ts)
  const fieldCtx: FieldResolutionContext = {
    hasPageField: false,
    hasStyleRefField: false,
    pageFieldSequence: [],
    pageFieldValueSequence: [],
    styleRefFieldValueSequence: [],
    pageNumber,
    totalPages,
    sectionPages: options?.sectionPages,
    pageNumberFormat: options?.pageNumberFormat,
    resolveStyleRefFieldValue: options?.resolveStyleRefFieldValue,
    resolveBookmarkFieldValue: options?.resolveBookmarkFieldValue,
    resolveBookmarkPageNumber: options?.resolveBookmarkPageNumber,
  }
  fieldCtx.pageFieldSequence = paragraphPageFieldSequence(paragraph)
  fieldCtx.pageFieldValueSequence = paragraphPageFieldValueSequence(paragraph)
  fieldCtx.styleRefFieldValueSequence = paragraphStyleRefFieldValueSequence(paragraph)
  fieldCtx.hasPageField =
    fieldCtx.pageFieldSequence.length > 0 || fieldCtx.pageFieldValueSequence.length > 0
  fieldCtx.hasStyleRefField = fieldCtx.styleRefFieldValueSequence.length > 0

  const fieldState: FieldResolutionState = {
    consumedPageFieldValues: 0,
    consumedStyleRefFieldValues: 0,
  }

  // Tab stops
  const tabStop = paragraphLeadingTabStop(paragraph)
  const tabStopPositionsPx: number[] = (paragraph.style?.tabStops ?? [])
    .map((entry) => entry.positionTwips !== undefined ? entry.positionTwips / 20 * (96 / 72) : undefined)
    .filter((v): v is number => v !== undefined && Number.isFinite(v) && v > 0)
    .sort((a, b) => a - b)

  let hasTabSplit = false
  let tabLeaderColor: string | undefined

  // Tracked changes & comments
  const showTrackedChanges = options?.showTrackedChanges === true
  const showTrackedInlineMarkup =
    showTrackedChanges && options?.trackedMarkupMode !== "gutter"
  const trackedMarkup = showTrackedInlineMarkup
    ? resolveParagraphTrackedMarkup(paragraph)
    : undefined
  const commentMarkup =
    options?.showCommentHighlights === true
      ? resolveParagraphCommentMarkup(paragraph)
      : undefined

  // TOC
  const tocParagraphLevel = tableOfContentsLevel(paragraph)
  const tocLinkColor = tocParagraphLevel
    ? options?.tocLinkColorByLevel?.[tocParagraphLevel]
    : undefined

  const floatingAnchorOriginCorrectionXPx = Number.isFinite(
    options?.floatingAnchorOriginCorrectionXPx
  )
    ? Math.round(options?.floatingAnchorOriginCorrectionXPx as number)
    : 0

  const checkboxChoiceRow = paragraphLooksLikeCheckboxChoiceRow(paragraph)
  const documentTabStopPx =
    Number.isFinite(options?.defaultTabStopPx) &&
    (options?.defaultTabStopPx as number) > 0
      ? Math.max(1, Math.round(options?.defaultTabStopPx as number))
      : DEFAULT_TAB_STOP_PX
  const fallbackTabWidthPx = checkboxChoiceRow
    ? checkboxChoiceRowTabWidthPx(paragraph)
    : documentTabStopPx

  const tabLeaderLeftTabStopPositionsPx = useTabLeaderLayout
    ? resolveParagraphFirstLineLeftTabStopsPx(paragraph)
    : []

  const shouldTrackTabLineWidth = !useTabLeaderLayout && !useAnchoredTabLayout
  let approximateLineWidthPx = 0
  let trackedVisibleChildCursor = 0

  // ---- Inner closures ----

  const appendTrackedDeletionSegments = (
    target: VNode[],
    keySeed: string,
    fallbackStyle?: TextRunNode["style"] | FormFieldRunNode["style"]
  ): void => {
    if (!trackedMarkup) return
    const segments = trackedMarkup.deletedSegmentsByVisibleChildIndex.get(
      trackedVisibleChildCursor
    )
    if (!segments || segments.length === 0) return

    segments.forEach((segment, segmentIndex) => {
      target.push(
        h(
          "span",
          {
            key: `${keySeed}-tracked-del-${trackedVisibleChildCursor}-${segmentIndex}`,
            "data-docx-tracked-change": "deletion",
            "data-docx-tracked-change-id": segment.change.id,
            title: segment.change.author
              ? `${segment.change.author} deleted`
              : "Deleted",
            style: trackedDeletedStyle(documentTheme, segment.style ?? fallbackStyle),
          },
          segment.text
        )
      )
    })
  }

  const currentTrackedInlineChange = ():
    | ParagraphTrackedInlineChange
    | undefined =>
    trackedMarkup?.inlineChangeByVisibleChildIndex[trackedVisibleChildCursor]

  const currentCommentHighlightStyle = ():
    | Record<string, string | number | undefined>
    | undefined => {
    const commentIds =
      commentMarkup?.commentIdsByVisibleChildIndex[trackedVisibleChildCursor]
    if (!commentIds || commentIds.length === 0) return undefined
    return commentHighlightStyle(documentTheme, commentIds[0])
  }

  // Exact character windows for the child being rendered. When the paragraph
  // carries model comment offsets the renderer splits the run's text at these
  // boundaries instead of highlighting the whole run.
  const currentCommentTextRanges = ():
    | ParagraphCommentTextRange[]
    | undefined => {
    const ranges =
      commentMarkup?.rangesByVisibleChildIndex?.[trackedVisibleChildCursor]
    return ranges && ranges.length > 0 ? ranges : undefined
  }

  const commentStyleForIds = (
    commentIds: readonly number[]
  ): Record<string, string | number | undefined> | undefined =>
    commentIds.length > 0
      ? commentHighlightStyle(documentTheme, commentIds[0])
      : undefined

  const commentAttributesForIds = (
    commentIds: readonly number[]
  ): Record<string, string> => {
    if (commentIds.length === 0) return {}
    const attributes: Record<string, string> = {
      "data-docx-comment-ids": commentIds.join(" "),
    }
    if (commentIds.length === 1) {
      attributes["data-docx-comment-id"] = String(commentIds[0])
    }
    return attributes
  }

  /**
   * Splits `text` into the segments defined by `ranges`, so each character is
   * highlighted by exactly the comments that cover it.
   */
  const commentTextSegments = (
    text: string,
    ranges: readonly ParagraphCommentTextRange[]
  ): Array<{ start: number; end: number; commentIds: number[] }> => {
    const boundaries = new Set<number>([0, text.length])
    for (const range of ranges) {
      if (range.start > 0 && range.start < text.length) boundaries.add(range.start)
      if (range.end > 0 && range.end < text.length) boundaries.add(range.end)
    }
    const positions = [...boundaries].sort((left, right) => left - right)
    const segments: Array<{ start: number; end: number; commentIds: number[] }> = []
    for (let index = 0; index < positions.length - 1; index += 1) {
      const start = positions[index]
      const end = positions[index + 1]
      if (end <= start) continue
      const commentIds: number[] = []
      for (const range of ranges) {
        if (range.start <= start && range.end >= end && !commentIds.includes(range.commentId)) {
          commentIds.push(range.commentId)
        }
      }
      segments.push({ start, end, commentIds: commentIds.sort((a, b) => a - b) })
    }
    return segments
  }

  const currentAnnotationAttributes = (
    trackedInlineChange: ParagraphTrackedInlineChange | undefined
  ): Record<string, string> => {
    const attrs: Record<string, string> = {}
    if (trackedInlineChange) {
      attrs["data-docx-tracked-change"] = trackedInlineChange.kind
      attrs["data-docx-tracked-change-id"] = trackedInlineChange.id
    }
    const commentIds =
      commentMarkup?.commentIdsByVisibleChildIndex[trackedVisibleChildCursor]
    if (commentIds && commentIds.length > 0) {
      attrs["data-docx-comment-ids"] = commentIds.join(" ")
      if (commentIds.length === 1) {
        attrs["data-docx-comment-id"] = String(commentIds[0])
      }
    }
    return attrs
  }

  const appendParagraphMarkChange = (target: VNode[]): void => {
    const marker = renderParagraphMarkChange({
      keyPrefix,
      change: trackedMarkup?.paragraphMarkChange,
      documentTheme,
    })
    if (marker) target.push(marker)
  }

  const consumeTrackedVisibleChild = (
    child: ParagraphNode["children"][number]
  ): void => {
    if (!trackedMarkup && !commentMarkup) return
    if (child.type === "form-field") return
    trackedVisibleChildCursor += 1
  }

  const trackTextAdvance = (
    text: string,
    style?: TextRunNode["style"] | FormFieldRunNode["style"]
  ): void => {
    if (!shouldTrackTabLineWidth) return
    approximateLineWidthPx = updateEstimatedLineWidthPxForText(
      approximateLineWidthPx,
      text,
      style
    )
  }

  const trackInlineAdvance = (widthPx: number): void => {
    if (!shouldTrackTabLineWidth) return
    approximateLineWidthPx += Math.max(0, Math.round(widthPx))
  }

  const resolveNextTabWidthPx = (): number =>
    resolveTabSpacerWidthPx(
      tabStopPositionsPx,
      approximateLineWidthPx,
      fallbackTabWidthPx,
      checkboxChoiceRow
    )

  const tabTextStyle = (
    style: TextRunNode["style"] | FormFieldRunNode["style"],
    textStyle: Record<string, string | number | undefined>
  ): Record<string, string | number | undefined> => {
    const tabWidthPx = resolveNextTabWidthPx()
    trackInlineAdvance(tabWidthPx)
    const hasUnderline = Boolean(style?.underline)
    return {
      ...textStyle,
      display: "inline-block",
      width: tabWidthPx,
      minWidth: tabWidthPx,
      whiteSpace: "pre",
      textDecoration: hasUnderline ? "none" : textStyle.textDecoration,
      borderBottom: hasUnderline ? "1px solid currentColor" : undefined,
      lineHeight: "1em",
    }
  }

  const renderTabLeaderLeftSpacer = (
    target: VNode[],
    key: string,
    style: TextRunNode["style"] | undefined,
    trackedInlineChange: ParagraphTrackedInlineChange | undefined,
    widthPx: number
  ): void => {
    const textStyle = trackedInlineStyle(
      runStyleToCss(style, documentTheme),
      trackedInlineChange,
      documentTheme
    )
    const hasUnderline = Boolean(style?.underline)
    target.push(
      h(
        "span",
        {
          key,
          style: {
            ...textStyle,
            display: "inline-block",
            width: widthPx,
            minWidth: widthPx,
            whiteSpace: "pre",
            textDecoration: hasUnderline ? "none" : textStyle.textDecoration,
            borderBottom: hasUnderline ? "1px solid currentColor" : undefined,
            lineHeight: "1em",
          },
        },
        "\u00a0"
      )
    )
  }

  // Tracked link style
  const trackedLinkStyle = (
    style: TextRunNode["style"] | FormFieldRunNode["style"] | undefined,
    trackedInlineChange: ParagraphTrackedInlineChange | undefined
  ): Record<string, string | number | undefined> => {
    if (!isTableOfContentsParagraph(paragraph)) {
      return {
        ...trackedInlineStyle(
          linkStyleToCss(style, documentTheme),
          trackedInlineChange,
          documentTheme
        ),
        ...currentCommentHighlightStyle(),
      }
    }
    const base = runStyleToCss(style, documentTheme)
    return {
      ...trackedInlineStyle(
        {
          ...base,
          color: tocLinkColor
            ? themedRunColor(tocLinkColor, documentTheme)
            : "inherit",
          textDecoration: "none",
        },
        trackedInlineChange,
        documentTheme
      ),
      ...currentCommentHighlightStyle(),
    }
  }

  // ---- renderRun dispatcher ----
  const renderRun = (
    target: VNode[],
    child: ParagraphNode["children"][number],
    key: string,
    textOverride?: string,
    trackedInlineChange?: ParagraphTrackedInlineChange,
    childIndex = -1
  ): void => {
    const annotationAttributes = currentAnnotationAttributes(trackedInlineChange)

    // form-field
    if (child.type === "form-field") {
      renderFormFieldInline(
        target, child, key, textOverride, annotationAttributes,
        trackedInlineChange, childIndex
      )
      return
    }

    // image
    if (child.type === "image") {
      renderImageInline(target, child, key, childIndex)
      return
    }

    // text
    renderTextInline(
      target, child, key, textOverride, annotationAttributes,
      trackedInlineChange
    )
  }

  // Inline render helpers (delegates to renderRun internals)
  const renderFormFieldInline = (
    target: VNode[],
    child: Extract<ParagraphNode["children"][number], { type: "form-field" }>,
    key: string,
    textOverride: string | undefined,
    annotationAttributes: Record<string, string>,
    trackedInlineChange: ParagraphTrackedInlineChange | undefined,
    _childIndex: number
  ): void => {
    const textValue = textOverride ??
      (formFieldDisplayValue(child) || "\u00a0").replace(/\t/g, " ")
    const formFieldAttributes = {
      ...annotationAttributes,
      "data-docx-form-field": "true",
    }

    const linkedHref = child.link ? sanitizeOfficeLink(child.link) : undefined
    if (linkedHref) {
      const isInternalLink = linkedHref.startsWith("#")
      target.push(
        h(
          "a",
          {
            key,
            ...formFieldAttributes,
            href: linkedHref,
            target: isInternalLink ? undefined : "_blank",
            rel: isInternalLink ? undefined : "noreferrer noopener",
            onMouseDown(event: MouseEvent) {
              if (!isInternalLink || !onInternalLinkClick) return
              event.preventDefault(); event.stopPropagation()
            },
            onClick(event: MouseEvent) {
              if (!isInternalLink || !onInternalLinkClick) return
              event.preventDefault(); event.stopPropagation()
              onInternalLinkClick(linkedHref.slice(1))
            },
            style: trackedLinkStyle(child.style, trackedInlineChange),
          },
          textValue
        )
      )
      trackTextAdvance(textValue, child.style)
      return
    }

    const trackedStyle = {
      ...trackedInlineStyle(
        runStyleToCss(child.style, documentTheme),
        trackedInlineChange,
        documentTheme
      ),
      ...currentCommentHighlightStyle(),
    }
    if (textValue === "\t" && !useTabLeaderLayout && !useAnchoredTabLayout) {
      target.push(h("span", {
        key, ...formFieldAttributes,
        style: tabTextStyle(child.style, trackedStyle),
      }, "\u00a0"))
      return
    }
    target.push(h("span", {
      key, ...formFieldAttributes, style: trackedStyle,
    }, textValue))
    trackTextAdvance(textValue, child.style)
  }

  const renderImageInline = (
    target: VNode[],
    child: Extract<ParagraphNode["children"][number], { type: "image" }>,
    key: string,
    childIndex: number
  ): void => {
    if (options?.skipImageChildIndexes?.includes(childIndex)) return
    if (
      options?.skipFloatingImages &&
      child.floating &&
      !(
        options.renderWrappedFloatingImages &&
        shouldRenderWrappedFloatingImage(child)
      )
    ) return
    const result = renderImageRun({
      paragraph,
      child,
      key,
      childIndex,
      pageNumber,
      totalPages,
      pageNumberFormat: options?.pageNumberFormat,
      resolveStyleRefFieldValue: options?.resolveStyleRefFieldValue,
      floatingPageOriginPx,
      floatingAnchorOriginCorrectionXPx,
      withinHeaderFooter: options?.withinHeaderFooter,
      headerFooterRegion: options?.headerFooterRegion,
      tableCellLayout: options?.tableCellLayout,
      extra: {
        imageFilterSuffix: options?.imageFilterSuffix,
        imageAltText: options?.imageAltText,
        imageMissingText: options?.imageMissingText,
        onImageDecodeError: options?.onImageDecodeError,
        sectionImageInteraction: options?.sectionImageInteraction as ImageRunRenderExtra["sectionImageInteraction"],
      },
    })
    target.push(result.vnode)
    if (
      !result.isWrappedFloating &&
      !result.isAbsoluteFloating &&
      result.inlineAdvancePx !== undefined
    ) {
      trackInlineAdvance(result.inlineAdvancePx)
    }
  }

  const renderTextInline = (
    target: VNode[],
    child: Extract<ParagraphNode["children"][number], { type: "text" }>,
    key: string,
    textOverride: string | undefined,
    annotationAttributes: Record<string, string>,
    trackedInlineChange: ParagraphTrackedInlineChange | undefined
  ): void => {
    const baseTextStyle = trackedInlineStyle(
      runStyleToCss(child.style, documentTheme),
      trackedInlineChange,
      documentTheme
    )
    const effectiveText = textOverride ?? child.text
    // Character-level ranges only line up when the rendered text is the child's
    // own text (a checkbox leading space becomes an nbsp of the same length).
    // Tab-leader parts are substrings, so they keep the run-level highlight.
    const commentTextRanges =
      effectiveText.length === child.text.length ? currentCommentTextRanges() : undefined
    const textStyle = commentTextRanges
      ? baseTextStyle
      : { ...baseTextStyle, ...currentCommentHighlightStyle() }
    // Single-unit runs (note markers, tabs) keep the run-level highlight: a
    // character window inside them would not change what is painted.
    const singleUnitTextStyle = {
      ...textStyle,
      ...currentCommentHighlightStyle(),
    }

    // Note marker
    const noteLabel = resolveNoteLabel(child)
    if (noteLabel) {
      target.push(h("span", {
        key, ...annotationAttributes,
        "data-docx-note-marker": "true",
        style: { ...singleUnitTextStyle, verticalAlign: "super", fontSize: "0.75em" },
      }, noteLabel))
      trackTextAdvance(noteLabel, child.style)
      return
    }

    // Tab
    if (effectiveText === "\t" && !useTabLeaderLayout && !useAnchoredTabLayout) {
      target.push(h("span", {
        key, ...annotationAttributes,
        style: tabTextStyle(child.style, singleUnitTextStyle),
      }, "\u00a0"))
      return
    }

    // Hyperlink
    const linkHref = child.link ? sanitizeOfficeLink(child.link) : undefined
    if (linkHref) {
      const isInternalLink = linkHref.startsWith("#")
      target.push(
        h("a", {
          key, ...annotationAttributes,
          href: linkHref,
          target: isInternalLink ? undefined : "_blank",
          rel: isInternalLink ? undefined : "noreferrer noopener",
          onMouseDown(event: MouseEvent) {
            if (!isInternalLink || !onInternalLinkClick) return
            event.preventDefault(); event.stopPropagation()
          },
          onClick(event: MouseEvent) {
            if (!isInternalLink || !onInternalLinkClick) return
            event.preventDefault(); event.stopPropagation()
            onInternalLinkClick(linkHref.slice(1))
          },
          style: trackedLinkStyle(child.style, trackedInlineChange),
        },
        commentTextRanges
          ? commentTextSegments(effectiveText, commentTextRanges).map((segment) =>
              h(
                "span",
                {
                  key: `${key}-comment-${segment.start}`,
                  ...commentAttributesForIds(segment.commentIds),
                  style: commentStyleForIds(segment.commentIds),
                },
                effectiveText.slice(segment.start, segment.end)
              )
            )
          : effectiveText
        )
      )
      trackTextAdvance(effectiveText, child.style)
      return
    }

    // Plain text with soft-break control
    appendPlainTextWithSoftBreakControl(
      target, key, effectiveText, textStyle,
      child.style, annotationAttributes, commentTextRanges
    )
  }

  // ---- appendPlainTextWithSoftBreakControl ----
  const appendPlainTextWithSoftBreakControl = (
    target: VNode[],
    keySeed: string,
    text: string,
    style: Record<string, string | number | undefined>,
    measureStyle?: TextRunNode["style"] | FormFieldRunNode["style"],
    spanAttributes?: Record<string, string>,
    commentTextRanges?: readonly ParagraphCommentTextRange[]
  ): void => {
    // Exact comment windows: emit one span per run of identically-commented
    // characters, so a range covering half a run highlights only that half.
    if (commentTextRanges && commentTextRanges.length > 0) {
      for (const segment of commentTextSegments(text, commentTextRanges)) {
        target.push(
          h("span", {
            key: `${keySeed}-comment-${segment.start}`,
            ...spanAttributes,
            ...commentAttributesForIds(segment.commentIds),
            style: { ...style, ...commentStyleForIds(segment.commentIds) },
          }, text.slice(segment.start, segment.end))
        )
      }
      trackTextAdvance(text, measureStyle)
      return
    }

    const shouldControlSoftBreakStretch =
      isJustifiedAlignment(paragraph.style?.align) && text.includes("\n")
    if (!shouldControlSoftBreakStretch) {
      target.push(h("span", { key: keySeed, ...spanAttributes, style }, text))
      trackTextAdvance(text, measureStyle)
      return
    }

    const segments = text.split("\n")
    segments.forEach((segment, segmentIndex) => {
      const isLastSegment = segmentIndex === segments.length - 1
      if (segmentIndex > 0) {
        target.push(h("br", { key: `${keySeed}-soft-break-${segmentIndex}` }))
      }
      if (segment.length > 0) {
        target.push(
          h("span", {
            key: `${keySeed}-segment-${segmentIndex}`,
            ...spanAttributes,
            style: isLastSegment
              ? { ...style, display: "inline-block", maxWidth: "100%", whiteSpace: "pre-wrap", verticalAlign: "baseline" }
              : style,
          }, segment)
        )
      }
      trackTextAdvance(segment, measureStyle)
      if (!isLastSegment) approximateLineWidthPx = 0
    })
  }

  // ---- resolveNoteLabel ----
  const resolveNoteLabel = (child: TextRunNode): string | undefined => {
    if (!child.noteReference) return undefined
    const noteRef = child.noteReference
    const supplied =
      noteRef.kind === "footnote"
        ? options?.noteMarkerLabels?.footnote
        : options?.noteMarkerLabels?.endnote
    const label = supplied?.get(noteRef.id)
    if (label) return label
    const index =
      noteRef.kind === "footnote"
        ? safeNoteMarkerIndexes.footnote.get(noteRef.id)
        : safeNoteMarkerIndexes.endnote.get(noteRef.id)
    if (!Number.isFinite(index)) return undefined
    return String(index as number)
  }

  // ---- renderNumberingIntoTarget ----
  const renderNumberingIntoTarget = (
    target: VNode[],
    numberingKey: string
  ): void => {
    if (!numberingLabel) return
    const numberingTextStyle = numberingLabel.style
      ? runStyleToCss(numberingLabel.style, documentTheme)
      : undefined
    target.push(
      h(
        "span",
        {
          key: numberingKey,
          "data-docx-numbering-label": "true",
          style: numberingMarkerStyle(
            paragraph,
            options?.numberingDefinitions,
            numberingLabel,
            numberingTextStyle,
            documentTheme
          ),
        },
        numberingLabel.imageSrc
          ? [
              h("img", {
                src: numberingLabel.imageSrc,
                alt: "",
                "aria-hidden": true,
                loading: "lazy",
                decoding: "async",
                "data-image-state": "loading",
                onLoad: (event: Event) => {
                  const image = event.currentTarget as HTMLImageElement | null
                  if (image) image.dataset.imageState = "ready"
                },
                onError: (event: Event) => {
                  const image = event.currentTarget as HTMLImageElement | null
                  if (image) image.dataset.imageState = "error"
                  options?.onImageDecodeError?.({
                    code: "IMAGE_DECODE_FAILED",
                    phase: "image-decode",
                    imageKey: numberingKey,
                    message: officeMessagesEn.docx.diagnostics.numberingImageDecodeFailed,
                  })
                },
                style: {
                  display: "inline-block",
                  verticalAlign: "text-bottom",
                  width: numberingLabel.imageWidthPx ?? 12,
                  height: numberingLabel.imageHeightPx ?? 12,
                  marginRight: 2,
                },
              }),
              numberingLabel.trailingText ?? "",
            ]
          : numberingLabel.text ?? ""
      )
    )
  }

  // ---- buildAnchoredTabZones ----
  const buildAnchoredTabZones = (zoneCount: 2 | 3): VNode[][] => {
    const zones = Array.from({ length: zoneCount }, () => [] as VNode[])
    let activeZone = 0

    renderNumberingIntoTarget(zones[0], `${keyPrefix}-numbering-anchored`)

    paragraph.children.forEach((child, childIndex) => {
      const zoneIndex = Math.max(0, Math.min(zoneCount - 1, activeZone))
      const zoneTarget = zones[zoneIndex]
      const key = `${keyPrefix}-anchored-run-${childIndex}`
      appendTrackedDeletionSegments(
        zoneTarget,
        `${key}-before`,
        child.type === "text" || child.type === "form-field" ? child.style : undefined
      )
      const trackedInlineChange = currentTrackedInlineChange()

      const rawText =
        child.type === "text"
          ? child.text
          : child.type === "form-field"
          ? formFieldDisplayValue(child)
          : undefined
      if (typeof rawText !== "string" || !rawText.includes("\t")) {
        if (child.type === "text") {
          const resolvedText = resolveFieldText(rawText ?? "", fieldCtx, fieldState, zoneIndex)
          renderRun(
            zoneTarget, child, key,
            attachTextToPreviousCheckbox(paragraph, childIndex, resolvedText),
            trackedInlineChange, childIndex
          )
        } else {
          renderRun(zoneTarget, child, key, undefined, trackedInlineChange, childIndex)
        }
        consumeTrackedVisibleChild(child)
        return
      }

      const parts = rawText.split("\t")
      parts.forEach((part, partIndex) => {
        const currentZone = Math.max(0, Math.min(zoneCount - 1, activeZone))
        if (part.length > 0) {
          renderRun(
            zones[currentZone], child,
            `${key}-part-${partIndex}`,
            resolveFieldText(part, fieldCtx, fieldState, currentZone),
            trackedInlineChange, childIndex
          )
        }
        if (partIndex < parts.length - 1 && activeZone < zoneCount - 1) {
          activeZone += 1
        }
      })
      consumeTrackedVisibleChild(child)
    })

    appendTrackedDeletionSegments(
      zones[Math.max(0, Math.min(zoneCount - 1, activeZone))],
      `${keyPrefix}-tail`
    )
    appendParagraphMarkChange(
      zones[Math.max(0, Math.min(zoneCount - 1, activeZone))]
    )
    return zones
  }

  // ---- anchoredTabZoneStyle ----
  const anchoredTabZoneStyle: Record<string, string> = {
    display: "inline-flex", minWidth: "0", whiteSpace: "pre-wrap",
    alignItems: "baseline", wordBreak: "normal", overflowWrap: "normal",
    flexWrap: "nowrap",
  }

  // ---- Main rendering logic ----

  if (numberingLabel && !useAnchoredTabLayout) {
    renderNumberingIntoTarget(
      useTabLeaderLayout ? runsLeft : runs,
      `${keyPrefix}-numbering`
    )
    if (numberingLabel.imageSrc) {
      trackInlineAdvance(numberingLabel.imageWidthPx ?? 12)
      trackTextAdvance(numberingLabel.trailingText ?? "", numberingLabel.style)
    } else {
      trackTextAdvance(numberingLabel.text ?? "", numberingLabel.style)
    }
  }

  if (!useAnchoredTabLayout) {
    paragraph.children.forEach((child, childIndex) => {
      const key = `${keyPrefix}-run-${childIndex}`
      const target = useTabLeaderLayout
        ? hasTabSplit ? runsRight : runsLeft
        : runs
      appendTrackedDeletionSegments(
        target, `${key}-before`,
        child.type === "text" || child.type === "form-field" ? child.style : undefined
      )
      const trackedInlineChange = currentTrackedInlineChange()

      if (
        useTabLeaderLayout &&
        child.type === "text" &&
        child.text.includes("\t") &&
        !hasTabSplit
      ) {
        handleTabLeaderSplit(
          child, childIndex, trackedInlineChange
        )
        return
      }

      if (child.type === "text") {
        const resolvedText = resolveFieldText(child.text, fieldCtx, fieldState)
        renderRun(
          target, child, key,
          attachTextToPreviousCheckbox(paragraph, childIndex, resolvedText),
          trackedInlineChange, childIndex
        )
      } else {
        renderRun(target, child, key, undefined, trackedInlineChange, childIndex)
      }
      consumeTrackedVisibleChild(child)
    })

    const trailingTarget = useTabLeaderLayout
      ? hasTabSplit ? runsRight : runsLeft
      : runs
    appendTrackedDeletionSegments(trailingTarget, `${keyPrefix}-tail`)
  }

  // Tab leader split handler
  function handleTabLeaderSplit(
    child: Extract<ParagraphNode["children"][number], { type: "text" }>,
    childIndex: number,
    trackedInlineChange: ParagraphTrackedInlineChange | undefined
  ): void {
    const text = child.text
    const tabIndex = text.lastIndexOf("\t")
    const leftText = text.slice(0, tabIndex)
    const rightText = text.slice(tabIndex + 1)
    if (!tabLeaderColor && child.style?.color) {
      tabLeaderColor = child.style.color
    }

    const hasLaterTab = paragraph.children.slice(childIndex + 1).some(
      (nextChild) =>
        (nextChild.type === "text" && nextChild.text.includes("\t")) ||
        (nextChild.type === "form-field" &&
          formFieldDisplayValue(nextChild).includes("\t"))
    )

    if (leftText.length === 0 && rightText.length === 0 && hasLaterTab) {
      renderRun(
        runsLeft, child,
        `${keyPrefix}-run-${childIndex}-spacer`,
        " ", trackedInlineChange, childIndex
      )
      consumeTrackedVisibleChild(child)
      return
    }

    if (leftText.length > 0) {
      let tabLeaderLeftLineWidthPx = 0
      const leftTextParts = leftText.split("\t")
      leftTextParts.forEach((part, partIndex) => {
        if (part.length > 0) {
          const resolvedPart = resolveFieldText(part, fieldCtx, fieldState, 0)
          renderRun(
            runsLeft, child,
            `${keyPrefix}-run-${childIndex}-left-${partIndex}`,
            resolvedPart, trackedInlineChange, childIndex
          )
          tabLeaderLeftLineWidthPx = updateEstimatedLineWidthPxForText(
            tabLeaderLeftLineWidthPx, resolvedPart, child.style
          )
        }
        if (partIndex >= leftTextParts.length - 1) return
        const tabWidthPx = resolveTabSpacerWidthPx(
          tabLeaderLeftTabStopPositionsPx,
          tabLeaderLeftLineWidthPx, fallbackTabWidthPx
        )
        renderTabLeaderLeftSpacer(
          runsLeft,
          `${keyPrefix}-run-${childIndex}-left-tab-${partIndex}`,
          child.style, trackedInlineChange, tabWidthPx
        )
        tabLeaderLeftLineWidthPx += tabWidthPx
      })
    }
    renderRun(
      runsRight, child,
      `${keyPrefix}-run-${childIndex}-right`,
      resolveFieldText(rightText, fieldCtx, fieldState, 1),
      trackedInlineChange, childIndex
    )
    hasTabSplit = true
    consumeTrackedVisibleChild(child)
  }

  // ---- Return based on tab layout ----

  if (useCenterRightTabLayout) {
    const zones = buildAnchoredTabZones(3)
    const centerStopPx = Math.max(0, Math.round(paragraphFirstTabStopPx(paragraph, "center") ?? 0))
    const rightStopPx = Math.max(centerStopPx, Math.round(paragraphFirstTabStopPx(paragraph, "right") ?? centerStopPx))
    return h("div", {
      key: `${keyPrefix}-center-right-tabs`,
      "data-docx-tab-layout": "center-right",
      style: {
        display: "grid",
        gridTemplateColumns: `${centerStopPx}px 0px ${Math.max(0, rightStopPx - centerStopPx)}px 0px minmax(0, 1fr)`,
        alignItems: "start", width: "100%",
      },
    }, [
      h("span", { "data-docx-tab-zone": "0", style: { ...anchoredTabZoneStyle, gridColumn: "1 / 2", gridRow: "1 / 2" } }, zones[0]),
      h("span", { "data-docx-tab-zone": "1", style: { ...anchoredTabZoneStyle, gridColumn: "1 / -1", gridRow: "1 / 2", justifySelf: "start", marginLeft: centerStopPx, textAlign: "center", transform: "translateX(-50%)", width: "max-content", maxWidth: "100%" } }, zones[1]),
      h("span", { "data-docx-tab-zone": "2", style: { ...anchoredTabZoneStyle, gridColumn: "3 / 4", gridRow: "1 / 2", justifySelf: "end", justifyContent: "flex-end", textAlign: "right" } }, zones[2]),
    ])
  }

  if (useCenterTabLayout) {
    const zones = buildAnchoredTabZones(2)
    const centerStopPx = Math.max(0, Math.round(paragraphFirstTabStopPx(paragraph, "center") ?? 0))
    return h("div", {
      key: `${keyPrefix}-center-tabs`,
      "data-docx-tab-layout": "center",
      style: {
        display: "grid",
        gridTemplateColumns: `${centerStopPx}px 0px minmax(0, 1fr)`,
        alignItems: "start", width: "100%",
      },
    }, [
      zones[0].length > 0
        ? h("span", { "data-docx-tab-zone": "left", style: { ...anchoredTabZoneStyle, gridColumn: "1 / 2", gridRow: "1 / 2" } }, zones[0])
        : null,
      h("span", { "data-docx-tab-zone": "center", style: { ...anchoredTabZoneStyle, gridColumn: "1 / -1", gridRow: "1 / 2", justifySelf: "start", marginLeft: centerStopPx, textAlign: "center", transform: "translateX(-50%)", width: "max-content", maxWidth: "100%" } }, zones[1]),
    ])
  }

  if (useRightTabLayout) {
    const zones = buildAnchoredTabZones(2)
    const rightStopPx = Math.max(0, Math.round(paragraphFirstTabStopPx(paragraph, "right") ?? 0))
    return h("div", {
      key: `${keyPrefix}-right-tabs`,
      "data-docx-tab-layout": "right",
      style: {
        display: "grid",
        gridTemplateColumns: `${rightStopPx}px 0px minmax(0, 1fr)`,
        alignItems: "start", width: "100%",
      },
    }, [
      h("span", { "data-docx-tab-zone": "left", style: { ...anchoredTabZoneStyle, gridColumn: "1 / 2", gridRow: "1 / 2" } }, zones[0]),
      h("span", { "data-docx-tab-zone": "right", style: { ...anchoredTabZoneStyle, gridColumn: "1 / 2", gridRow: "1 / 2", justifySelf: "end", justifyContent: "flex-end", textAlign: "right" } }, zones[1]),
    ])
  }

  if (useTabLeaderLayout && hasTabSplit) {
    appendParagraphMarkChange(runsRight)
    return h("div", {
      key: `${keyPrefix}-toc`,
      "data-docx-tab-layout": "leader",
      style: { display: "flex", alignItems: "baseline", width: "100%" },
    }, [
      h("span", { "data-docx-tab-zone": "left", style: { display: "block", flex: "0 1 auto", minWidth: 0, whiteSpace: "pre-wrap" } }, runsLeft),
      h("span", { "aria-hidden": true, style: { ...tabLeaderStyle(tabStop?.leader, themedRunColor(tabLeaderColor, documentTheme)), flex: "1 1 auto", minWidth: 8, marginLeft: 6, marginRight: 6, height: "1em", alignSelf: "center" } }),
      h("span", { "data-docx-tab-zone": "right", style: { display: "block", flex: "0 0 max-content", minWidth: "max-content", width: "max-content", textAlign: "right", whiteSpace: "nowrap", textIndent: 0 } }, runsRight),
    ])
  }

  const finalRuns = runsLeft.length > 0 ? runsLeft : runs
  appendParagraphMarkChange(finalRuns)
  return finalRuns
}

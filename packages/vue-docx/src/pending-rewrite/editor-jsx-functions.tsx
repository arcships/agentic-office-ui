// Pending rewrite for vue-docx.
// Extracted from upstream editor.tsx during docx-002 (editor-helpers cleanup).
// These functions use React JSX and renderToStaticMarkup and must be rewritten
// as Vue render functions (h()/v-html) before being used.

import { renderToStaticMarkup } from "react-dom/server";
import type {
  NumberingDefinitionSet,
  ParagraphNode
} from "@extend-ai/docx-core";

// HeaderFooterImageInteraction is defined in the editor.tsx hooks region
// (upstream line 31736) and will be provided by the vue-docx composables.
interface HeaderFooterImageInteraction {
  [key: string]: unknown;
}

function renderStaticHtml(node: unknown): string {
  return renderToStaticMarkup(<>{node}</>);
}

interface ParagraphRunRenderOptions {
  showTrackedChanges?: boolean;
  showCommentHighlights?: boolean;
  numberingDefinitions?: NumberingDefinitionSet;
  tocLinkColorByLevel?: Partial<Record<number, string | undefined>>;
  trackedMarkupMode?: "inline" | "gutter";
  withinHeaderFooter?: boolean;
  headerFooterRegion?: "header" | "footer";
  pageNumberFormat?: string;
  resolveStyleRefFieldValue?: (target: string) => string | undefined;
  floatingAnchorOriginCorrectionXPx?: number;
  sectionImageInteraction?: HeaderFooterImageInteraction;
  paragraphOriginLeftPx?: number;
  paragraphOriginTopPx?: number;
  imageFilterSuffix?: string;
}

function renderParagraphRuns(
  paragraph: ParagraphNode,
  keyPrefix: string,
  documentTheme: DocxDocumentTheme = "light",
  numberingLabel?: ParagraphNumberingLabel,
  onInternalLinkClick?: (bookmarkName: string) => void,
  floatingPageOriginPx?: {
    left: number;
    top: number;
    marginLeft?: number;
    marginTop?: number;
    columnLeft?: number;
    columnTop?: number;
    pageWidth?: number;
  },
  noteMarkerIndexes?: {
    footnote: Map<number, number>;
    endnote: Map<number, number>;
  },
  pageNumber?: number,
  totalPages?: number,
  options?: ParagraphRunRenderOptions
): unknown {
  const runs: unknown[] = [];
  const runsLeft: unknown[] = [];
  const runsRight: unknown[] = [];
  const safeNoteMarkerIndexes = {
    footnote: noteMarkerIndexes?.footnote ?? new Map<number, number>(),
    endnote: noteMarkerIndexes?.endnote ?? new Map<number, number>(),
  };
  const useTabLeaderLayout = paragraphUsesTabLeaders(paragraph);
  const anchoredTabLayout = paragraphAnchoredTabLayout(paragraph, {
    withinHeaderFooter: options?.withinHeaderFooter,
  });
  const useCenterRightTabLayout =
    !useTabLeaderLayout && anchoredTabLayout === "center-right";
  const useCenterTabLayout =
    !useTabLeaderLayout && anchoredTabLayout === "center";
  const useRightTabLayout =
    !useTabLeaderLayout && anchoredTabLayout === "right";
  const useAnchoredTabLayout =
    useCenterRightTabLayout || useCenterTabLayout || useRightTabLayout;
  const pageFieldSequence = paragraphPageFieldSequence(paragraph);
  const pageFieldValueSequence = paragraphPageFieldValueSequence(paragraph);
  const styleRefFieldValueSequence =
    paragraphStyleRefFieldValueSequence(paragraph);
  const hasPageField =
    pageFieldSequence.length > 0 || pageFieldValueSequence.length > 0;
  const hasStyleRefField = styleRefFieldValueSequence.length > 0;
  let consumedPageFieldValues = 0;
  let consumedStyleRefFieldValues = 0;
  const tabStop = paragraphLeadingTabStop(paragraph);
  const tabStopPositionsPx = (paragraph.style?.tabStops ?? [])
    .map((tabStopEntry) => twipsToPixels(tabStopEntry.positionTwips))
    .filter(
      (value): value is number =>
        Number.isFinite(value) && (value as number) > 0
    )
    .sort((left, right) => left - right);
  let hasTabSplit = false;
  let tabLeaderColor: string | undefined;
  const showTrackedChanges = options?.showTrackedChanges === true;
  const showTrackedInlineMarkup =
    showTrackedChanges && options?.trackedMarkupMode !== "gutter";
  const trackedMarkup = showTrackedInlineMarkup
    ? resolveParagraphTrackedMarkup(paragraph)
    : undefined;
  const commentMarkup =
    options?.showCommentHighlights === true
      ? resolveParagraphCommentMarkup(paragraph)
      : undefined;
  const tocParagraphLevel = tableOfContentsLevel(paragraph);
  const tocLinkColor = tocParagraphLevel
    ? options?.tocLinkColorByLevel?.[tocParagraphLevel]
    : undefined;
  const floatingAnchorOriginCorrectionXPx = Number.isFinite(
    options?.floatingAnchorOriginCorrectionXPx
  )
    ? Math.round(options?.floatingAnchorOriginCorrectionXPx as number)
    : 0;
  const checkboxChoiceRow = paragraphLooksLikeCheckboxChoiceRow(paragraph);
  const fallbackTabWidthPx = checkboxChoiceRow
    ? checkboxChoiceRowTabWidthPx(paragraph)
    : DEFAULT_TAB_STOP_PX;
  const tabLeaderLeftTabStopPositionsPx = useTabLeaderLayout
    ? resolveParagraphFirstLineLeftTabStopsPx(paragraph)
    : [];
  const shouldTrackTabLineWidth = !useTabLeaderLayout && !useAnchoredTabLayout;
  let approximateLineWidthPx = 0;
  let trackedVisibleChildCursor = 0;
  const appendTrackedDeletionSegments = (
    target: unknown[],
    keySeed: string,
    fallbackStyle?: TextRunNode["style"] | FormFieldRunNode["style"]
  ): void => {
    if (!trackedMarkup) {
      return;
    }

    const segments = trackedMarkup.deletedSegmentsByVisibleChildIndex.get(
      trackedVisibleChildCursor
    );
    if (!segments || segments.length === 0) {
      return;
    }

    segments.forEach((segment, segmentIndex) => {
      target.push(
        <span
          key={`${keySeed}-tracked-del-${trackedVisibleChildCursor}-${segmentIndex}`}
          data-docx-tracked-change="deletion"
          data-docx-tracked-change-id={segment.change.id}
          title={
            segment.change.author
              ? `${segment.change.author} deleted`
              : "Deleted"
          }
          style={trackedDeletedStyle(
            documentTheme,
            segment.style ?? fallbackStyle
          )}
        >
          {segment.text}
        </span>
      );
    });
  };
  const currentTrackedInlineChange = ():
    | ParagraphTrackedInlineChange
    | undefined =>
    trackedMarkup?.inlineChangeByVisibleChildIndex[trackedVisibleChildCursor];
  const currentCommentHighlightStyle = (): Record<string, string | number | undefined> | undefined => {
    const commentIds =
      commentMarkup?.commentIdsByVisibleChildIndex[trackedVisibleChildCursor];
    if (!commentIds || commentIds.length === 0) {
      return undefined;
    }
    return commentHighlightStyle(documentTheme, commentIds[0]);
  };
  const currentAnnotationAttributes = (
    trackedInlineChange: ParagraphTrackedInlineChange | undefined
  ): Record<string, string> => {
    const attributes: Record<string, string> = {};
    if (trackedInlineChange) {
      attributes["data-docx-tracked-change"] = trackedInlineChange.kind;
      attributes["data-docx-tracked-change-id"] = trackedInlineChange.id;
    }

    const commentIds =
      commentMarkup?.commentIdsByVisibleChildIndex[trackedVisibleChildCursor];
    if (commentIds && commentIds.length > 0) {
      attributes["data-docx-comment-ids"] = commentIds.join(" ");
      if (commentIds.length === 1) {
        attributes["data-docx-comment-id"] = String(commentIds[0]);
      }
    }

    return attributes;
  };
  const consumeTrackedVisibleChild = (
    child: ParagraphNode["children"][number]
  ): void => {
    if (!trackedMarkup && !commentMarkup) {
      return;
    }
    if (child.type === "form-field") {
      return;
    }
    trackedVisibleChildCursor += 1;
  };
  const normalizeFieldComparableText = (input: string): string =>
    input
      .replace(/\u00a0/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  const trackTextAdvance = (
    text: string,
    style?: TextRunNode["style"] | FormFieldRunNode["style"]
  ): void => {
    if (!shouldTrackTabLineWidth) {
      return;
    }
    approximateLineWidthPx = updateEstimatedLineWidthPxForText(
      approximateLineWidthPx,
      text,
      style
    );
  };
  const trackInlineAdvance = (widthPx: number): void => {
    if (!shouldTrackTabLineWidth) {
      return;
    }
    approximateLineWidthPx += Math.max(0, Math.round(widthPx));
  };
  const resolveNextTabWidthPx = (): number =>
    resolveTabSpacerWidthPx(
      tabStopPositionsPx,
      approximateLineWidthPx,
      fallbackTabWidthPx,
      checkboxChoiceRow
    );
  const appendPlainTextWithSoftBreakControl = (
    target: unknown[],
    keySeed: string,
    text: string,
    style: Record<string, string | number | undefined>,
    measureStyle?: TextRunNode["style"] | FormFieldRunNode["style"],
    spanAttributes?: Record<string, string>
  ): void => {
    const shouldControlSoftBreakStretch =
      paragraph.style?.align === "justify" && text.includes("\n");
    if (!shouldControlSoftBreakStretch) {
      target.push(
        <span key={keySeed} {...spanAttributes} style={style}>
          {text}
        </span>
      );
      trackTextAdvance(text, measureStyle);
      return;
    }

    const segments = text.split("\n");
    segments.forEach((segment, segmentIndex) => {
      const isLastSegment = segmentIndex === segments.length - 1;
      if (segmentIndex > 0) {
        target.push(<br key={`${keySeed}-soft-break-${segmentIndex}`} />);
      }

      if (segment.length > 0) {
        target.push(
          <span
            key={`${keySeed}-segment-${segmentIndex}`}
            {...spanAttributes}
            style={
              isLastSegment
                ? {
                    ...style,
                    display: "inline-block",
                    maxWidth: "100%",
                    whiteSpace: "pre-wrap",
                    verticalAlign: "baseline",
                  }
                : style
            }
          >
            {segment}
          </span>
        );
      }

      trackTextAdvance(segment, measureStyle);
      if (!isLastSegment) {
        approximateLineWidthPx = 0;
      }
    });
  };
  const tabTextStyle = (
    style: TextRunNode["style"] | FormFieldRunNode["style"],
    textStyle: Record<string, string | number | undefined>
  ): Record<string, string | number | undefined> => {
    const tabWidthPx = resolveNextTabWidthPx();
    trackInlineAdvance(tabWidthPx);
    const hasUnderline = Boolean(style?.underline);
    return {
      ...textStyle,
      display: "inline-block",
      width: tabWidthPx,
      minWidth: tabWidthPx,
      whiteSpace: "pre",
      textDecoration: hasUnderline ? "none" : textStyle.textDecoration,
      borderBottom: hasUnderline ? "1px solid currentColor" : undefined,
      lineHeight: "1em",
    };
  };
  const renderTabLeaderLeftSpacer = (
    target: unknown[],
    key: string,
    style: TextRunNode["style"] | undefined,
    trackedInlineChange: ParagraphTrackedInlineChange | undefined,
    widthPx: number
  ): void => {
    const textStyle = trackedInlineStyle(
      runStyleToCss(style, documentTheme),
      trackedInlineChange
    );
    const hasUnderline = Boolean(style?.underline);
    target.push(
      <span
        key={key}
        style={{
          ...textStyle,
          display: "inline-block",
          width: widthPx,
          minWidth: widthPx,
          whiteSpace: "pre",
          textDecoration: hasUnderline ? "none" : textStyle.textDecoration,
          borderBottom: hasUnderline ? "1px solid currentColor" : undefined,
          lineHeight: "1em",
        }}
      >
        {"\u00a0"}
      </span>
    );
  };
  const resolvePageFieldText = (
    value: string,
    preferredZone?: number
  ): string => {
    if (!hasPageField || value.trim().length === 0) {
      return value;
    }

    let fieldKind = pageFieldSequence[consumedPageFieldValues];
    const valueToken = pageFieldValueSequence[consumedPageFieldValues];
    if (valueToken) {
      if (
        normalizeFieldComparableText(value) !==
        normalizeFieldComparableText(valueToken.rawText)
      ) {
        return value;
      }
      fieldKind = valueToken.kind;
    } else {
      if (!fieldKind) {
        return value;
      }

      const normalized = value.trim();
      const likelyFieldResult =
        /^\d+$/.test(normalized) || /^[ivxlcdm]+$/i.test(normalized);
      if (!likelyFieldResult && preferredZone !== 1) {
        return value;
      }
    }

    const resolvedFieldValue =
      fieldKind === "NUMPAGES"
        ? Number.isFinite(totalPages) && (totalPages as number) > 0
          ? Math.max(1, Math.round(totalPages as number))
          : undefined
        : Number.isFinite(pageNumber) && (pageNumber as number) > 0
        ? Math.max(1, Math.round(pageNumber as number))
        : undefined;
    if (!Number.isFinite(resolvedFieldValue)) {
      return value;
    }

    consumedPageFieldValues += 1;
    const leadingWhitespace = value.match(/^\s*/)?.[0] ?? "";
    const trailingWhitespace = value.match(/\s*$/)?.[0] ?? "";
    const normalizedValue = Math.max(
      1,
      Math.round(resolvedFieldValue as number)
    );
    const formattedFieldValue =
      fieldKind === "PAGE"
        ? formatPageFieldValue(normalizedValue, options?.pageNumberFormat)
        : String(normalizedValue);
    return `${leadingWhitespace}${formattedFieldValue}${trailingWhitespace}`;
  };
  const resolveStyleRefFieldText = (value: string): string => {
    if (!hasStyleRefField || value.trim().length === 0) {
      return value;
    }

    const valueToken = styleRefFieldValueSequence[consumedStyleRefFieldValues];
    if (!valueToken) {
      return value;
    }

    if (
      normalizeFieldComparableText(value) !==
      normalizeFieldComparableText(valueToken.rawText)
    ) {
      return value;
    }

    const resolvedValue = options
      ?.resolveStyleRefFieldValue?.(valueToken.target)
      ?.trim();
    if (!resolvedValue) {
      return value;
    }

    consumedStyleRefFieldValues += 1;
    const leadingWhitespace = value.match(/^\s*/)?.[0] ?? "";
    const trailingWhitespace = value.match(/\s*$/)?.[0] ?? "";
    return `${leadingWhitespace}${resolvedValue}${trailingWhitespace}`;
  };
  const resolveFieldText = (value: string, preferredZone?: number): string =>
    resolvePageFieldText(resolveStyleRefFieldText(value), preferredZone);

  const trackedLinkStyle = (
    style: TextRunNode["style"] | FormFieldRunNode["style"] | undefined,
    trackedInlineChange: ParagraphTrackedInlineChange | undefined
  ): Record<string, string | number | undefined> => {
    if (!isTableOfContentsParagraph(paragraph)) {
      return {
        ...trackedInlineStyle(
          linkStyleToCss(style, documentTheme),
          trackedInlineChange
        ),
        ...currentCommentHighlightStyle(),
      };
    }

    const base = runStyleToCss(style, documentTheme);
    return {
      ...trackedInlineStyle(
        {
          ...base,
          color: tocLinkColor
            ? themedRunColor(tocLinkColor, documentTheme)
            : "inherit",
          textDecoration: "none",
        },
        trackedInlineChange
      ),
      ...currentCommentHighlightStyle(),
    };
  };
  const usesExternalHorizontalAnchorOrigin = (image: ImageRunNode): boolean => {
    const horizontalRelativeTo = image.floating?.horizontalRelativeTo
      ?.trim()
      .toLowerCase();
    return (
      horizontalRelativeTo === "page" ||
      horizontalRelativeTo === "margin" ||
      horizontalRelativeTo === "column"
    );
  };

  const renderRun = (
    target: unknown[],
    child: ParagraphNode["children"][number],
    key: string,
    textOverride?: string,
    trackedInlineChange?: ParagraphTrackedInlineChange,
    childIndex = -1
  ): void => {
    if (child.type === "form-field") {
      const textValue = (formFieldDisplayValue(child) || "\u00a0").replace(
        /\t/g,
        " "
      );
      const text = textOverride ?? textValue;
      const annotationAttributes =
        currentAnnotationAttributes(trackedInlineChange);
      if (child.link) {
        const linkHref = child.link;
        const isInternalLink = linkHref.startsWith("#");
        target.push(
          <a
            key={key}
            {...annotationAttributes}
            href={linkHref}
            target={isInternalLink ? undefined : "_blank"}
            rel={isInternalLink ? undefined : "noreferrer noopener"}
            onMouseDown={(event) => {
              if (!isInternalLink || !onInternalLinkClick) {
                return;
              }

              event.preventDefault();
              event.stopPropagation();
            }}
            onClick={(event) => {
              if (!isInternalLink || !onInternalLinkClick) {
                return;
              }

              event.preventDefault();
              event.stopPropagation();
              onInternalLinkClick(linkHref.slice(1));
            }}
            style={trackedLinkStyle(child.style, trackedInlineChange)}
          >
            {text}
          </a>
        );
        trackTextAdvance(text, child.style);
        return;
      }

      const trackedStyle = {
        ...trackedInlineStyle(
          runStyleToCss(child.style, documentTheme),
          trackedInlineChange
        ),
        ...currentCommentHighlightStyle(),
      };
      if (text === "\t" && !useTabLeaderLayout && !useAnchoredTabLayout) {
        target.push(
          <span
            key={key}
            {...annotationAttributes}
            style={tabTextStyle(child.style, trackedStyle)}
          >
            {"\u00a0"}
          </span>
        );
        return;
      }

      target.push(
        <span key={key} {...annotationAttributes} style={trackedStyle}>
          {text}
        </span>
      );
      trackTextAdvance(text, child.style);
      return;
    }

    if (child.type === "image") {
      const behavesAsDecorativeBehindTextBackground =
        imageBehavesAsDecorativeBehindTextBackground(child, paragraph);
      const sectionImageInteraction = options?.sectionImageInteraction;
      const sectionImageLocation =
        sectionImageInteraction?.imageLocationForChild?.(childIndex);
      const sectionImageKey = sectionImageLocation
        ? sectionImageLocationKey(sectionImageLocation)
        : undefined;
      const movePreview =
        sectionImageKey &&
        sectionImageInteraction?.floatingMovePreview?.imageKey ===
          sectionImageKey
          ? sectionImageInteraction.floatingMovePreview
          : undefined;
      const forceAbsoluteFixedPositionSectionFloat =
        options?.withinHeaderFooter === true &&
        isPageOrMarginAnchoredWrappedFloatingImage(child);
      const forceWrappedTopAnchoredSectionFloat =
        options?.withinHeaderFooter === true &&
        Boolean(
          child.floating?.wrapType && child.floating.wrapType !== "none"
        ) &&
        shouldRenderTopAnchoredMarginFloatAsAbsolute(child);
      const isWrappedFloatingImage =
        !forceAbsoluteFixedPositionSectionFloat &&
        (forceWrappedTopAnchoredSectionFloat
          ? true
          : shouldRenderWrappedFloatingImage(child));
      const isAbsoluteFloatingImage = forceAbsoluteFixedPositionSectionFloat
        ? true
        : forceWrappedTopAnchoredSectionFloat
        ? false
        : shouldRenderAbsoluteFloatingImage(child);
      const horizontalRelativeTo =
        child.floating?.horizontalRelativeTo?.toLowerCase();
      const verticalRelativeTo =
        child.floating?.verticalRelativeTo?.toLowerCase();
      let widthPx = child.widthPx;
      let heightPx = child.heightPx;
      ({ widthPx, heightPx } = resolvePageSpanningAbsoluteFloatingDimensions(
        child,
        widthPx,
        heightPx,
        floatingPageOriginPx
      ));
      const horizontalAnchorCorrectionPx = usesExternalHorizontalAnchorOrigin(
        child
      )
        ? floatingAnchorOriginCorrectionXPx
        : 0;
      const absoluteColumnOriginLeftPx =
        options?.withinHeaderFooter === true &&
        options?.headerFooterRegion === "header" &&
        child.floating?.behindDocument === true &&
        horizontalRelativeTo === "column"
          ? floatingPageOriginPx?.marginLeft
          : floatingPageOriginPx?.columnLeft;
      const floatingStyle: Record<string, string | number | undefined> = isWrappedFloatingImage
        ? wrappedFloatingImageStyle(child, {
            containerWidthPx: floatingPageOriginPx?.pageWidth,
            deltaX: (movePreview?.deltaX ?? 0) + horizontalAnchorCorrectionPx,
            deltaY: movePreview?.deltaY ?? 0,
            allowNegativeOffsets: true,
          })
        : isAbsoluteFloatingImage
        ? absoluteFloatingImageStyle(child, {
            pageOriginLeft: floatingPageOriginPx?.left,
            pageOriginTop: floatingPageOriginPx?.top,
            marginOriginLeft: floatingPageOriginPx?.marginLeft,
            marginOriginTop: floatingPageOriginPx?.marginTop,
            columnOriginLeft: absoluteColumnOriginLeftPx,
            columnOriginTop: floatingPageOriginPx?.columnTop,
            deltaX: (movePreview?.deltaX ?? 0) + horizontalAnchorCorrectionPx,
            deltaY: movePreview?.deltaY ?? 0,
          })
        : movePreview
        ? {
            transform: `translate(${movePreview.deltaX}px, ${movePreview.deltaY}px)`,
            position: "relative",
            zIndex: 4,
          }
        : {};
      const floatingStyleWithMovePreview = isWrappedFloatingImage
        ? {
            ...floatingStyle,
          }
        : floatingStyle;
      const onSectionImagePointerDown =
        sectionImageLocation &&
        sectionImageInteraction?.onImagePointerDown &&
        !sectionImageInteraction.isReadOnly &&
        !behavesAsDecorativeBehindTextBackground
          ? (event: Record<string, unknown>) => {
              sectionImageInteraction.onImagePointerDown?.(
                event,
                sectionImageLocation,
                child,
                isWrappedFloatingImage,
                isAbsoluteFloatingImage
              );
            }
          : undefined;
      const onSectionImageClick =
        sectionImageLocation &&
        sectionImageInteraction?.onImageClick &&
        !sectionImageInteraction.isReadOnly &&
        !behavesAsDecorativeBehindTextBackground
          ? (event: Record<string, unknown>) => {
              event.stopPropagation();
              sectionImageInteraction.onImageClick?.(sectionImageLocation);
            }
          : undefined;
      const sectionImageCursor =
        sectionImageLocation &&
        !sectionImageInteraction?.isReadOnly &&
        !behavesAsDecorativeBehindTextBackground &&
        (isWrappedFloatingImage || isAbsoluteFloatingImage)
          ? "move"
          : undefined;
      const isCenteredStandaloneInlineImage =
        !isWrappedFloatingImage &&
        !isAbsoluteFloatingImage &&
        paragraph.style?.align === "center" &&
        paragraph.children.length === 1 &&
        paragraph.children[0]?.type === "image";
      const trackedImageStyle =
        trackedInlineChange?.kind === "insertion" ||
        trackedInlineChange?.kind === "move-to"
          ? {
              outline:
                trackedInlineChange.kind === "move-to"
                  ? "2px solid rgba(112, 173, 71, 0.35)"
                  : "2px solid rgba(220, 38, 38, 0.3)",
              outlineOffset: 1,
            }
          : undefined;

      const renderableImageSrc =
        syntheticTextBoxSvg(
          child,
          pageNumber,
          totalPages,
          options?.pageNumberFormat,
          options?.resolveStyleRefFieldValue
        ) ?? resolveRenderableImageSource(child);

      if (!child.src) {
        target.push(
          <span
            key={key}
            style={{
              display: "inline-flex",
              minWidth: 112,
              minHeight: 80,
              alignItems: "center",
              justifyContent: "center",
              border: "1px dashed #d1d5db",
              borderRadius: 4,
              color: "#6b7280",
              fontSize: 12,
              marginInline:
                isWrappedFloatingImage || isAbsoluteFloatingImage
                  ? undefined
                  : 0,
              paddingInline: 8,
              ...trackedImageStyle,
              ...floatingStyleWithMovePreview,
              cursor: sectionImageCursor,
              ...(behavesAsDecorativeBehindTextBackground
                ? { pointerEvents: "none", userSelect: "none" }
                : undefined),
            }}
            onPointerDown={onSectionImagePointerDown}
            onClick={onSectionImageClick}
          >
            Missing image
          </span>
        );
        return;
      }

      if (
        imageUsesPlaceholderFallback(child) ||
        (child.src && !renderableImageSrc)
      ) {
        const isSmallIcon = (widthPx ?? 0) <= 56 && (heightPx ?? 0) <= 56;
        const fallbackText = unsupportedImageFallbackLabel(
          child,
          widthPx,
          heightPx
        );
        target.push(
          <span
            key={key}
            role="img"
            aria-label={child.alt ?? "DOCX image"}
            style={{
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
              marginInline:
                isWrappedFloatingImage || isAbsoluteFloatingImage
                  ? undefined
                  : 0,
              ...trackedImageStyle,
              ...floatingStyleWithMovePreview,
              cursor: sectionImageCursor,
              ...(behavesAsDecorativeBehindTextBackground
                ? { pointerEvents: "none", userSelect: "none" }
                : undefined),
            }}
            onPointerDown={onSectionImagePointerDown}
            onClick={onSectionImageClick}
          >
            {fallbackText}
          </span>
        );
        return;
      }

      target.push(
        (() => {
          const cropLayout = imageCropLayout(child, widthPx, heightPx);
          const imageVisualStyle: Record<string, string | number | undefined> = {
            filter: appendCssFilters(
              child.cssFilter,
              options?.imageFilterSuffix
            ),
            opacity: child.cssOpacity,
          };
          const imageTransformStyle = resolveImageRenderTransformStyle(child, {
            frameWidthPx: cropLayout?.frameWidthPx ?? widthPx,
            frameHeightPx: cropLayout?.frameHeightPx ?? heightPx,
            baseTransform: cropLayout
              ? `translate(${-cropLayout.offsetXPx}px, ${-cropLayout.offsetYPx}px)`
              : undefined,
          });

          if (cropLayout) {
            return (
              <span
                key={key}
                style={{
                  width: `${cropLayout.frameWidthPx}px`,
                  height: `${cropLayout.frameHeightPx}px`,
                  overflow: "clip",
                  display: isCenteredStandaloneInlineImage
                    ? "block"
                    : "inline-block",
                  marginLeft: isCenteredStandaloneInlineImage
                    ? "auto"
                    : undefined,
                  marginRight: isCenteredStandaloneInlineImage
                    ? "auto"
                    : undefined,
                  verticalAlign: "middle",
                  marginInline:
                    isWrappedFloatingImage || isAbsoluteFloatingImage
                      ? undefined
                      : 0,
                  ...trackedImageStyle,
                  ...floatingStyleWithMovePreview,
                  cursor: sectionImageCursor,
                  ...(behavesAsDecorativeBehindTextBackground
                    ? { pointerEvents: "none", userSelect: "none" }
                    : undefined),
                }}
                onPointerDown={onSectionImagePointerDown}
                onClick={onSectionImageClick}
              >
                <img
                  src={renderableImageSrc}
                  alt={child.alt ?? "DOCX image"}
                  draggable={false}
                  style={{
                    width: `${cropLayout.imageWidthPx}px`,
                    height: `${cropLayout.imageHeightPx}px`,
                    maxWidth: "none",
                    display: "block",
                    ...imageVisualStyle,
                    ...imageTransformStyle,
                  }}
                />
              </span>
            );
          }

          return (
            <img
              key={key}
              src={renderableImageSrc}
              alt={child.alt ?? "DOCX image"}
              draggable={false}
              onPointerDown={onSectionImagePointerDown}
              onClick={onSectionImageClick}
              style={{
                width: widthPx ? `${widthPx}px` : undefined,
                height: heightPx ? `${heightPx}px` : undefined,
                maxWidth:
                  isWrappedFloatingImage || isAbsoluteFloatingImage
                    ? undefined
                    : "100%",
                display: isCenteredStandaloneInlineImage ? "block" : undefined,
                marginLeft: isCenteredStandaloneInlineImage
                  ? "auto"
                  : undefined,
                marginRight: isCenteredStandaloneInlineImage
                  ? "auto"
                  : undefined,
                verticalAlign: "middle",
                marginInline:
                  isWrappedFloatingImage || isAbsoluteFloatingImage
                    ? undefined
                    : 0,
                ...imageVisualStyle,
                ...resolveImageRenderTransformStyle(child, {
                  frameWidthPx: widthPx,
                  frameHeightPx: heightPx,
                }),
                ...trackedImageStyle,
                ...floatingStyleWithMovePreview,
                cursor: sectionImageCursor,
                ...(behavesAsDecorativeBehindTextBackground
                  ? { pointerEvents: "none", userSelect: "none" }
                  : undefined),
              }}
            />
          );
        })()
      );
      if (!isWrappedFloatingImage && !isAbsoluteFloatingImage) {
        trackInlineAdvance(widthPx ?? 0);
      }
      return;
    }

    const textStyle = {
      ...trackedInlineStyle(
        runStyleToCss(child.style, documentTheme),
        trackedInlineChange
      ),
      ...currentCommentHighlightStyle(),
    };
    const annotationAttributes =
      currentAnnotationAttributes(trackedInlineChange);
    const noteLabel = noteMarkerLabel(
      child.noteReference,
      safeNoteMarkerIndexes.footnote,
      safeNoteMarkerIndexes.endnote
    );
    if (noteLabel) {
      target.push(
        <span
          key={key}
          {...annotationAttributes}
          style={{
            ...textStyle,
            verticalAlign: "super",
            fontSize: "0.75em",
          }}
        >
          {noteLabel}
        </span>
      );
      trackTextAdvance(noteLabel, child.style);
      return;
    }

    if (
      (textOverride ?? child.text) === "\t" &&
      !useTabLeaderLayout &&
      !useAnchoredTabLayout
    ) {
      target.push(
        <span
          key={key}
          {...annotationAttributes}
          style={tabTextStyle(child.style, textStyle)}
        >
          {"\u00a0"}
        </span>
      );
      return;
    }

    if (child.link) {
      const linkHref = child.link;
      const isInternalLink = linkHref.startsWith("#");
      target.push(
        <a
          key={key}
          {...annotationAttributes}
          href={linkHref}
          target={isInternalLink ? undefined : "_blank"}
          rel={isInternalLink ? undefined : "noreferrer noopener"}
          onMouseDown={(event) => {
            if (!isInternalLink || !onInternalLinkClick) {
              return;
            }

            event.preventDefault();
            event.stopPropagation();
          }}
          onClick={(event) => {
            if (!isInternalLink || !onInternalLinkClick) {
              return;
            }

            event.preventDefault();
            event.stopPropagation();
            onInternalLinkClick(linkHref.slice(1));
          }}
          style={trackedLinkStyle(child.style, trackedInlineChange)}
        >
          {textOverride ?? child.text}
        </a>
      );
      trackTextAdvance(textOverride ?? child.text, child.style);
      return;
    }

    appendPlainTextWithSoftBreakControl(
      target,
      key,
      textOverride ?? child.text,
      textStyle,
      child.style,
      annotationAttributes
    );
  };

  const anchoredTabZoneStyle: Record<string, string | number | undefined> = {
    display: "inline-flex",
    minWidth: 0,
    whiteSpace: "pre-wrap",
    alignItems: "baseline",
    wordBreak: "normal",
    overflowWrap: "normal",
    flexWrap: "nowrap",
  };

  const renderNumberingIntoTarget = (
    target: unknown[],
    numberingKey: string
  ): void => {
    if (!numberingLabel) {
      return;
    }

    const numberingTextStyle = numberingLabel.style
      ? runStyleToCss(numberingLabel.style, documentTheme)
      : undefined;
    target.push(
      <span
        key={numberingKey}
        style={numberingMarkerStyle(
          paragraph,
          options?.numberingDefinitions,
          numberingLabel,
          numberingTextStyle,
          documentTheme
        )}
      >
        {numberingLabel.imageSrc ? (
          <>
            <img
              src={numberingLabel.imageSrc}
              alt=""
              aria-hidden="true"
              style={{
                display: "inline-block",
                verticalAlign: "text-bottom",
                width: numberingLabel.imageWidthPx ?? 12,
                height: numberingLabel.imageHeightPx ?? 12,
                marginRight: 2,
              }}
            />
            {numberingLabel.trailingText ?? ""}
          </>
        ) : (
          numberingLabel.text ?? ""
        )}
      </span>
    );
  };

  const buildAnchoredTabZones = (zoneCount: 2 | 3): unknown[][] => {
    const zones = Array.from(
      { length: zoneCount },
      () => [] as unknown[]
    );
    let activeZone = 0;

    renderNumberingIntoTarget(zones[0], `${keyPrefix}-numbering-anchored`);

    paragraph.children.forEach((child, childIndex) => {
      const zoneIndex = Math.max(0, Math.min(zoneCount - 1, activeZone));
      const zoneTarget = zones[zoneIndex];
      const key = `${keyPrefix}-anchored-run-${childIndex}`;
      appendTrackedDeletionSegments(
        zoneTarget,
        `${key}-before`,
        child.type === "text" || child.type === "form-field"
          ? child.style
          : undefined
      );
      const trackedInlineChange = currentTrackedInlineChange();

      const rawText =
        child.type === "text"
          ? child.text
          : child.type === "form-field"
          ? formFieldDisplayValue(child)
          : undefined;
      if (typeof rawText !== "string" || !rawText.includes("\t")) {
        if (child.type === "text") {
          const resolvedText = resolveFieldText(rawText ?? "", zoneIndex);
          renderRun(
            zoneTarget,
            child,
            key,
            attachTextToPreviousCheckbox(paragraph, childIndex, resolvedText),
            trackedInlineChange,
            childIndex
          );
        } else {
          renderRun(
            zoneTarget,
            child,
            key,
            undefined,
            trackedInlineChange,
            childIndex
          );
        }
        consumeTrackedVisibleChild(child);
        return;
      }

      const parts = rawText.split("\t");
      parts.forEach((part, partIndex) => {
        const currentZone = Math.max(0, Math.min(zoneCount - 1, activeZone));
        if (part.length > 0) {
          renderRun(
            zones[currentZone],
            child,
            `${key}-part-${partIndex}`,
            resolveFieldText(part, currentZone),
            trackedInlineChange,
            childIndex
          );
        }
        if (partIndex < parts.length - 1 && activeZone < zoneCount - 1) {
          activeZone += 1;
        }
      });
      consumeTrackedVisibleChild(child);
    });

    appendTrackedDeletionSegments(
      zones[Math.max(0, Math.min(zoneCount - 1, activeZone))],
      `${keyPrefix}-tail`
    );
    return zones;
  };

  if (numberingLabel && !useAnchoredTabLayout) {
    renderNumberingIntoTarget(
      useTabLeaderLayout ? runsLeft : runs,
      `${keyPrefix}-numbering`
    );
    if (numberingLabel.imageSrc) {
      trackInlineAdvance(numberingLabel.imageWidthPx ?? 12);
      trackTextAdvance(numberingLabel.trailingText ?? "", numberingLabel.style);
    } else {
      trackTextAdvance(numberingLabel.text ?? "", numberingLabel.style);
    }
  }

  if (!useAnchoredTabLayout) {
    paragraph.children.forEach((child, childIndex) => {
      const key = `${keyPrefix}-run-${childIndex}`;
      const target = useTabLeaderLayout
        ? hasTabSplit
          ? runsRight
          : runsLeft
        : runs;
      appendTrackedDeletionSegments(
        target,
        `${key}-before`,
        child.type === "text" || child.type === "form-field"
          ? child.style
          : undefined
      );
      const trackedInlineChange = currentTrackedInlineChange();
      if (
        useTabLeaderLayout &&
        child.type === "text" &&
        child.text.includes("\t") &&
        !hasTabSplit
      ) {
        const text = child.text;
        const tabIndex = text.lastIndexOf("\t");
        const leftText = text.slice(0, tabIndex);
        const rightText = text.slice(tabIndex + 1);
        if (!tabLeaderColor && child.style?.color) {
          tabLeaderColor = child.style.color;
        }

        const hasLaterTab = paragraph.children
          .slice(childIndex + 1)
          .some(
            (nextChild) =>
              (nextChild.type === "text" && nextChild.text.includes("\t")) ||
              (nextChild.type === "form-field" &&
                formFieldDisplayValue(nextChild).includes("\t"))
          );
        if (leftText.length === 0 && rightText.length === 0 && hasLaterTab) {
          renderRun(
            runsLeft,
            child,
            `${keyPrefix}-run-${childIndex}-spacer`,
            " ",
            trackedInlineChange,
            childIndex
          );
          consumeTrackedVisibleChild(child);
          return;
        }

        if (leftText.length > 0) {
          let tabLeaderLeftLineWidthPx = 0;
          const leftTextParts = leftText.split("\t");
          leftTextParts.forEach((part, partIndex) => {
            if (part.length > 0) {
              const resolvedPart = resolveFieldText(part, 0);
              renderRun(
                runsLeft,
                child,
                `${keyPrefix}-run-${childIndex}-left-${partIndex}`,
                resolvedPart,
                trackedInlineChange,
                childIndex
              );
              tabLeaderLeftLineWidthPx = updateEstimatedLineWidthPxForText(
                tabLeaderLeftLineWidthPx,
                resolvedPart,
                child.style
              );
            }

            if (partIndex >= leftTextParts.length - 1) {
              return;
            }

            const tabWidthPx = resolveTabSpacerWidthPx(
              tabLeaderLeftTabStopPositionsPx,
              tabLeaderLeftLineWidthPx,
              fallbackTabWidthPx
            );
            renderTabLeaderLeftSpacer(
              runsLeft,
              `${keyPrefix}-run-${childIndex}-left-tab-${partIndex}`,
              child.style,
              trackedInlineChange,
              tabWidthPx
            );
            tabLeaderLeftLineWidthPx += tabWidthPx;
          });
        }
        renderRun(
          runsRight,
          child,
          `${keyPrefix}-run-${childIndex}-right`,
          resolveFieldText(rightText, 1),
          trackedInlineChange,
          childIndex
        );
        hasTabSplit = true;
        consumeTrackedVisibleChild(child);
        return;
      }

      if (child.type === "text") {
        const resolvedText = resolveFieldText(child.text);
        renderRun(
          target,
          child,
          key,
          attachTextToPreviousCheckbox(paragraph, childIndex, resolvedText),
          trackedInlineChange,
          childIndex
        );
        consumeTrackedVisibleChild(child);
        return;
      }
      renderRun(target, child, key, undefined, trackedInlineChange, childIndex);
      consumeTrackedVisibleChild(child);
    });

    const trailingTarget = useTabLeaderLayout
      ? hasTabSplit
        ? runsRight
        : runsLeft
      : runs;
    appendTrackedDeletionSegments(trailingTarget, `${keyPrefix}-tail`);
  }

  if (useCenterRightTabLayout) {
    const zones = buildAnchoredTabZones(3);
    const centerStopPx = Math.max(
      0,
      Math.round(paragraphFirstTabStopPx(paragraph, "center") ?? 0)
    );
    const rightStopPx = Math.max(
      centerStopPx,
      Math.round(paragraphFirstTabStopPx(paragraph, "right") ?? centerStopPx)
    );

    return (
      <div
        key={`${keyPrefix}-center-right-tabs`}
        data-docx-tab-layout="center-right"
        style={{
          display: "grid",
          gridTemplateColumns: `${centerStopPx}px 0px ${Math.max(
            0,
            rightStopPx - centerStopPx
          )}px 0px minmax(0, 1fr)`,
          alignItems: "start",
          width: "100%",
        }}
      >
        <span
          data-docx-tab-zone="0"
          style={{
            ...anchoredTabZoneStyle,
            gridColumn: "1 / 2",
            gridRow: "1 / 2",
          }}
        >
          {zones[0]}
        </span>
        <span
          data-docx-tab-zone="1"
          style={{
            ...anchoredTabZoneStyle,
            gridColumn: "1 / -1",
            gridRow: "1 / 2",
            justifySelf: "start",
            marginLeft: centerStopPx,
            textAlign: "center",
            transform: "translateX(-50%)",
            width: "max-content",
            maxWidth: "100%",
          }}
        >
          {zones[1]}
        </span>
        <span
          data-docx-tab-zone="2"
          style={{
            ...anchoredTabZoneStyle,
            gridColumn: "3 / 4",
            gridRow: "1 / 2",
            justifySelf: "end",
            justifyContent: "flex-end",
            textAlign: "right",
          }}
        >
          {zones[2]}
        </span>
      </div>
    );
  }

  if (useCenterTabLayout) {
    const zones = buildAnchoredTabZones(2);
    const centerStopPx = Math.max(
      0,
      Math.round(paragraphFirstTabStopPx(paragraph, "center") ?? 0)
    );
    return (
      <div
        key={`${keyPrefix}-center-tabs`}
        data-docx-tab-layout="center"
        style={{
          display: "grid",
          gridTemplateColumns: `${centerStopPx}px 0px minmax(0, 1fr)`,
          alignItems: "start",
          width: "100%",
        }}
      >
        {zones[0].length > 0 ? (
          <span
            data-docx-tab-zone="left"
            style={{
              ...anchoredTabZoneStyle,
              gridColumn: "1 / 2",
              gridRow: "1 / 2",
            }}
          >
            {zones[0]}
          </span>
        ) : null}
        <span
          data-docx-tab-zone="center"
          style={{
            ...anchoredTabZoneStyle,
            gridColumn: "1 / -1",
            gridRow: "1 / 2",
            justifySelf: "start",
            marginLeft: centerStopPx,
            textAlign: "center",
            transform: "translateX(-50%)",
            width: "max-content",
            maxWidth: "100%",
          }}
        >
          {zones[1]}
        </span>
      </div>
    );
  }

  if (useRightTabLayout) {
    const zones = buildAnchoredTabZones(2);
    const rightStopPx = Math.max(
      0,
      Math.round(paragraphFirstTabStopPx(paragraph, "right") ?? 0)
    );
    return (
      <div
        key={`${keyPrefix}-right-tabs`}
        data-docx-tab-layout="right"
        style={{
          display: "grid",
          gridTemplateColumns: `${rightStopPx}px 0px minmax(0, 1fr)`,
          alignItems: "start",
          width: "100%",
        }}
      >
        <span
          data-docx-tab-zone="left"
          style={{
            ...anchoredTabZoneStyle,
            gridColumn: "1 / 2",
            gridRow: "1 / 2",
          }}
        >
          {zones[0]}
        </span>
        <span
          data-docx-tab-zone="right"
          style={{
            ...anchoredTabZoneStyle,
            gridColumn: "1 / 2",
            gridRow: "1 / 2",
            justifySelf: "end",
            justifyContent: "flex-end",
            textAlign: "right",
          }}
        >
          {zones[1]}
        </span>
      </div>
    );
  }

  if (!useTabLeaderLayout || !hasTabSplit) {
    return runsLeft.length > 0 ? runsLeft : runs;
  }

  return (
    <div
      key={`${keyPrefix}-toc`}
      data-docx-tab-layout="leader"
      style={{
        display: "flex",
        alignItems: "baseline",
        width: "100%",
      }}
    >
      <span
        data-docx-tab-zone="left"
        style={{
          display: "block",
          flex: "0 1 auto",
          minWidth: 0,
          whiteSpace: "pre-wrap",
        }}
      >
        {runsLeft}
      </span>
      <span
        aria-hidden="true"
        style={{
          ...tabLeaderStyle(
            tabStop?.leader,
            themedRunColor(tabLeaderColor, documentTheme)
          ),
          flex: "1 1 auto",
          minWidth: 8,
          marginLeft: 6,
          marginRight: 6,
          height: "1em",
          alignSelf: "center",
        }}
      />
      <span
        data-docx-tab-zone="right"
        style={{
          display: "block",
          flex: "0 0 max-content",
          minWidth: "max-content",
          width: "max-content",
          textAlign: "right",
          whiteSpace: "nowrap",
          textIndent: 0,
        }}
      >
        {runsRight}
      </span>
    </div>
  );
}

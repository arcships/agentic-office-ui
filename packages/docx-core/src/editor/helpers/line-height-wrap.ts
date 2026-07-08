// Paragraph line wrapping estimation and line-count calculation.
// Upstream editor.tsx: lines 9342-10054.
//
// Ports `estimateWrappedLineCountForParagraph` (9658-9934) and
// `paragraphLineCountWithinWidth` (9982-10054). These live in a
// separate file to keep line-height.ts within the ≤1000-line limit.
//
// Note: The upstream tab-leader fast path (estimateTabLeaderWrappedLineCountForParagraph)
// is not yet ported; tab-leader paragraphs fall through to the general
// character-by-character wrapping path.

import type {
  NumberingDefinitionSet,
  ParagraphNode,
  TextRunNode,
  FormFieldRunNode,
} from "../../engine/types";
import {
  buildParagraphPretextLayoutSource,
  layoutParagraphPretextSource,
  resolveUniformPretextSourceFont,
  pretextWordBreakModeForText,
} from "./pretext-build";
import {
  measurePretextPlainTextLineCount,
} from "../../viewer/pretext-layout";
import {
  estimateParagraphLineHeightPx,
  measureTextWidthPx,
  paragraphBaseFontSizePx,
  paragraphContainsExplicitLineBreakText,
  resolveNextTabStopPx,
  resolveParagraphTabStopsPx,
} from "./line-height";
import {
  paragraphAnchoredTabLayout,
  paragraphUsesTabLeaders,
} from "./field-helpers";
import { formFieldDisplayValue } from "./paragraph-inspect";
import {
  shouldRenderAbsoluteFloatingImage,
  shouldRenderWrappedFloatingImage,
  paragraphAvailableTextWidthPx,
} from "./paragraph-geometry";
import { resolveParagraphDualWrappedTextLayout } from "./pretext-measure";
import { resolveListParagraphIndent } from "./xml-parsing-extra";
import { twipsToSignedPixels } from "./ooxml-helpers";
import { resolveNumberingMarkerBoxWidthPx } from "./style-to-css";
import {
  PAGE_OVERFLOW_TOLERANCE_PX,
} from "./constants";

// ── Wrapped line count cache ──

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

// ── Line wrapping estimator ──

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

  // Fast path: pretext-powered line count for plain paragraphs
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

      // Uniform-font paragraphs use arithmetic wrapping over cached
      // segment widths (no per-line string allocations).
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

  // General character-by-character wrapping path
  let lineCount = 1;
  let currentLineWidthPx = 0;
  let hasVisibleContent = false;

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
      lineCount += 1;
      currentLineWidthPx = 0;
      return;
    }

    if (token === "\t") {
      if (useTabLeaderLayout) {
        // Tab-leader layout renders right-side zones independently;
        // fall through to treat as regular tab for now.
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
      // Use the child's style for text runs; form-field style comes from
      // its own style field (same TextStyle shape).
      const style: TextRunNode["style"] | FormFieldRunNode["style"] | undefined =
        child.type === "text" ? child.style : child.style;
      commitToken(token, style);
    }
  }

  return rememberLineCount(hasVisibleContent ? lineCount : 1);
}

// ── Paragraph line count within available width ──

export function paragraphLineCountWithinWidth(
  paragraph: ParagraphNode,
  availableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet
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
          numberingDefinitions
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
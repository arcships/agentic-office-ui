// Pretext layout source construction: build per-paragraph pretext layout
// sources/runs, tab-spacer text, image-wrap-mode helpers, word-break mode,
// layout items, uniform-font resolution, and the low-level
// layoutParagraphPretextSource entry point.
//
// Upstream editor.tsx:
//   5813-5883  (Pretext layout run/source/geometry types + KEEP_ALL_SCRIPT_RE)
//   6331-6531  (pretext caches + image wrap-mode helpers)
//   6533-6762  (buildParagraphPretextTabSpacerText + buildParagraphPretextLayoutSource)
//   6945-7125  (synthetic source, word-break mode, sanitize, layout items,
//               uniform font, measure-segments items, layoutParagraphPretextSource,
//               wrappedParagraphSessionText)
//   7127-7230  (inline-image placeholder, child anchor offset, expandOffsetToWord)
//   7228-7230  (pixelsToTwips)
//   7936-7944  (paragraphIsEffectivelyEmpty)
//
// Functions that depend on not-yet-ported modules (estimateParagraphHeightPx,
// paragraphLineCountWithinWidth, paragraphAvailableTextWidthPx,
// estimateRenderedPageSegmentHeightPx, resolveWrappedFloatingImageDropPatch)
// are deferred to the line-height-table / table-height / pagination-plan /
// paragraph-render modules. See docs/docx-editor-helpers-split-plan.md.

import type {
  FormFieldRunNode,
  ImageRunNode,
  ParagraphNode,
  TextRunNode
} from "../../engine/types";
import type {
  PretextExclusionRect,
  PretextLayoutItem,
  PretextVariableWidthLayout
} from "../../viewer/pretext-layout";
import { layoutItemsWithPretextAroundExclusions } from "../../viewer/pretext-items-layout";
import { layoutTextWithPretextAroundExclusions } from "../../viewer/pretext-layout";
import { TWIPS_PER_PIXEL } from "../../viewer/section-layout";
import type {
  DocxContextMenuActionId,
  DocxImageWrapMode,
  DocxImageWrapState
} from "./editor-types-extra";
import {
  DEFAULT_TAB_STOP_PX,
  MIN_PARAGRAPH_LINE_HEIGHT_PX,
  TEXT_MEASURE_CACHE_MAX_ENTRIES
} from "./constants";
import {
  pretextWordBreakModeByText,
  setCacheEntry
} from "./cache-utils";
import {
  checkboxChoiceRowTabWidthPx,
  resolveTabSpacerWidthPx,
  updateEstimatedLineWidthPxForText
} from "./drop-cap";
import {
  measureTextWidthPx,
  paragraphBaseFontSizePx,
  paragraphContainsExplicitLineBreakText,
  resolveMeasureFont,
  resolveParagraphTabStopsPx
} from "./line-height";
import {
  firstRunStyle,
  paragraphText
} from "./paragraph-inspect";
import {
  floatingImageMovesWithText,
  paragraphHasFormField,
  paragraphLooksLikeCheckboxChoiceRow
} from "./paragraph-geometry";

// --- Types (upstream 5813-5883) ---

export interface ParagraphPretextLayoutRun {
  kind: "text" | "image" | "tab";
  key: string;
  text: string;
  startOffset: number;
  endOffset: number;
  style?: TextRunNode["style"];
  link?: string;
  image?: ImageRunNode;
  tabWidthPx?: number;
}

export interface ParagraphPretextLayoutSource {
  text: string;
  runs: ParagraphPretextLayoutRun[];
}

export const KEEP_ALL_SCRIPT_RE =
  /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u;

export interface DualWrappedFloatingImageGeometry {
  image: ImageRunNode;
  imageIndex: number;
  containerWidthPx: number;
  imageLeftPx: number;
  imageTopPx: number;
  imageWidthPx: number;
  imageHeightPx: number;
  exclusion: PretextExclusionRect;
}

export const MIN_DUAL_WRAPPED_INTERIOR_BAND_PX = 72;

export interface ParagraphDualWrappedTextLayout {
  source: ParagraphPretextLayoutSource;
  geometries: DualWrappedFloatingImageGeometry[];
  lineHeightPx: number;
  layout: PretextVariableWidthLayout;
}

export interface PageFlowFloatingWrapObstacle extends PretextExclusionRect {
  sourceNodeIndex: number;
}

/**
 * Minimal measure-segment shape used by buildMeasureSegmentsPretextLayoutItems.
 * The full ParagraphMeasureSegment interface (upstream 9342) lands in the
 * line-height-table module; only the { text, style } fields are needed here.
 */
export interface ParagraphMeasureSegment {
  text: string;
  style?: TextRunNode["style"] | FormFieldRunNode["style"];
}

// --- Pretext caches (upstream 6331-6342) ---
// Declared here (not in cache-utils.ts) because they key on
// ParagraphPretextLayoutSource, which is defined in this module — putting them
// in cache-utils would create a circular import.

const paragraphPretextLayoutSourceCache = new WeakMap<
  ParagraphNode,
  Map<number, ParagraphPretextLayoutSource | null>
>();
const paragraphPretextLayoutItemsBySource = new WeakMap<
  ParagraphPretextLayoutSource,
  PretextLayoutItem[]
>();
const paragraphPretextUniformFontBySource = new WeakMap<
  ParagraphPretextLayoutSource,
  string | null
>();

// --- Image wrap-mode helpers (upstream 6344-6531) ---

export function imageWrapModeFromFloating(
  floating?: ImageRunNode["floating"]
): DocxImageWrapMode {
  if (!floating) {
    return "inline";
  }

  const wrapType = floating.wrapType;
  if (!wrapType || wrapType === "none") {
    return floating.behindDocument ? "behindText" : "inFrontOfText";
  }

  if (
    wrapType === "square" ||
    wrapType === "tight" ||
    wrapType === "through" ||
    wrapType === "topAndBottom"
  ) {
    return wrapType;
  }

  return "square";
}

export function resolveDocxImageWrapState(
  floating?: ImageRunNode["floating"]
): DocxImageWrapState {
  const moveWithText = floatingImageMovesWithText(floating);
  return {
    mode: imageWrapModeFromFloating(floating),
    moveWithText,
    fixedPositionOnPage: !moveWithText,
  };
}

export function imageWrapModeActionId(
  mode: DocxImageWrapMode
): DocxContextMenuActionId {
  switch (mode) {
    case "inline":
      return "image-wrap-inline";
    case "square":
      return "image-wrap-square";
    case "tight":
      return "image-wrap-tight";
    case "through":
      return "image-wrap-through";
    case "topAndBottom":
      return "image-wrap-top-and-bottom";
    case "behindText":
      return "image-wrap-behind-text";
    case "inFrontOfText":
      return "image-wrap-in-front-of-text";
    default:
      return "image-wrap-square";
  }
}

export function imageWrapModeFromActionId(
  actionId: DocxContextMenuActionId | (string & {})
): DocxImageWrapMode | undefined {
  switch (actionId) {
    case "image-wrap-inline":
      return "inline";
    case "image-wrap-square":
      return "square";
    case "image-wrap-tight":
      return "tight";
    case "image-wrap-through":
      return "through";
    case "image-wrap-top-and-bottom":
      return "topAndBottom";
    case "image-wrap-behind-text":
    case "image-behind-text":
      return "behindText";
    case "image-wrap-in-front-of-text":
    case "image-in-front-of-text":
      return "inFrontOfText";
    default:
      return undefined;
  }
}

// --- Tab-spacer text (upstream 6533-6584) ---

function buildParagraphPretextTabSpacerText(
  widthPx: number,
  style: TextRunNode["style"] | undefined,
  paragraphBaseFontPx: number
): string {
  const safeWidthPx = Math.max(8, Math.round(widthPx));
  const spacerCharacter = "\u00a0";
  const spacerAdvancePx = Math.max(
    1,
    measureTextWidthPx(spacerCharacter, style, paragraphBaseFontPx)
  );
  let spacerCount = Math.max(1, Math.round(safeWidthPx / spacerAdvancePx));
  let spacerText = spacerCharacter.repeat(spacerCount);
  let measuredWidthPx = measureTextWidthPx(
    spacerText,
    style,
    paragraphBaseFontPx
  );

  while (
    measuredWidthPx + spacerAdvancePx * 0.5 < safeWidthPx &&
    spacerCount < 64
  ) {
    spacerCount += 1;
    spacerText = spacerCharacter.repeat(spacerCount);
    measuredWidthPx = measureTextWidthPx(
      spacerText,
      style,
      paragraphBaseFontPx
    );
  }

  while (spacerCount > 1) {
    const nextText = spacerCharacter.repeat(spacerCount - 1);
    const nextWidthPx = measureTextWidthPx(
      nextText,
      style,
      paragraphBaseFontPx
    );
    if (
      Math.abs(nextWidthPx - safeWidthPx) >=
      Math.abs(measuredWidthPx - safeWidthPx)
    ) {
      break;
    }
    spacerCount -= 1;
    spacerText = nextText;
    measuredWidthPx = nextWidthPx;
  }

  return spacerText;
}

// --- Layout source construction (upstream 6586-6762) ---

export function buildParagraphPretextLayoutSource(
  paragraph: ParagraphNode,
  options?: {
    allowExplicitLineBreakText?: boolean;
    expandTabsForLayout?: boolean;
  }
): ParagraphPretextLayoutSource | undefined {
  const cacheKey =
    (options?.allowExplicitLineBreakText ? 1 : 0) |
    (options?.expandTabsForLayout ? 2 : 0);
  const cachedVariants = paragraphPretextLayoutSourceCache.get(paragraph);
  const cachedSource = cachedVariants?.get(cacheKey);
  if (cachedSource !== undefined) {
    return cachedSource ?? undefined;
  }
  const storeCachedSource = (
    source: ParagraphPretextLayoutSource | undefined
  ): ParagraphPretextLayoutSource | undefined => {
    if (!cachedVariants) {
      paragraphPretextLayoutSourceCache.set(
        paragraph,
        new Map([[cacheKey, source ?? null]])
      );
    } else {
      cachedVariants.set(cacheKey, source ?? null);
    }
    return source;
  };
  if (paragraphHasFormField(paragraph)) {
    return storeCachedSource(undefined);
  }
  if (
    !options?.allowExplicitLineBreakText &&
    paragraphContainsExplicitLineBreakText(paragraph)
  ) {
    return storeCachedSource(undefined);
  }

  const runs: ParagraphPretextLayoutRun[] = [];
  let combinedText = "";
  const paragraphBaseFontPx = paragraphBaseFontSizePx(paragraph);
  const tabStopPositionsPx = options?.expandTabsForLayout
    ? resolveParagraphTabStopsPx(paragraph)
    : [];
  const usesCheckboxRowTabFallback =
    paragraphLooksLikeCheckboxChoiceRow(paragraph);
  const fallbackTabWidthPx = usesCheckboxRowTabFallback
    ? checkboxChoiceRowTabWidthPx(paragraph)
    : DEFAULT_TAB_STOP_PX;
  let approximateLineWidthPx = 0;
  let lastTextStyle: TextRunNode["style"] | undefined =
    firstRunStyle(paragraph);

  for (
    let childIndex = 0;
    childIndex < paragraph.children.length;
    childIndex += 1
  ) {
    const child = paragraph.children[childIndex];
    if (!child) {
      continue;
    }
    if (child.type === "image") {
      if (child.floating) {
        continue;
      }

      const placeholderText = buildInlineImagePlaceholderText(
        child,
        lastTextStyle,
        paragraphBaseFontPx
      );
      const startOffset = combinedText.length;
      combinedText += placeholderText;
      runs.push({
        kind: "image",
        key: `run-${childIndex}`,
        text: placeholderText,
        startOffset,
        endOffset: combinedText.length,
        style: lastTextStyle,
        image: child,
      });
      continue;
    }
    if (child.type !== "text") {
      return storeCachedSource(undefined);
    }
    if (child.noteReference) {
      return storeCachedSource(undefined);
    }
    if (child.text.includes("\t")) {
      if (!options?.expandTabsForLayout) {
        return storeCachedSource(undefined);
      }

      const tabSegments = child.text.split("\t");
      tabSegments.forEach((segmentText, segmentIndex) => {
        if (segmentText.length > 0) {
          const startOffset = combinedText.length;
          combinedText += segmentText;
          runs.push({
            kind: "text",
            key: `run-${childIndex}-${segmentIndex}`,
            text: segmentText,
            startOffset,
            endOffset: combinedText.length,
            style: child.style,
            link: child.link,
          });
          approximateLineWidthPx = updateEstimatedLineWidthPxForText(
            approximateLineWidthPx,
            segmentText,
            child.style
          );
        }

        if (segmentIndex >= tabSegments.length - 1) {
          return;
        }

        const tabWidthPx = resolveTabSpacerWidthPx(
          tabStopPositionsPx,
          approximateLineWidthPx,
          fallbackTabWidthPx,
          usesCheckboxRowTabFallback
        );
        const spacerText = buildParagraphPretextTabSpacerText(
          tabWidthPx,
          child.style,
          paragraphBaseFontPx
        );
        const startOffset = combinedText.length;
        combinedText += spacerText;
        runs.push({
          kind: "tab",
          key: `run-${childIndex}-${segmentIndex}-tab`,
          text: spacerText,
          startOffset,
          endOffset: combinedText.length,
          style: child.style,
          tabWidthPx,
        });
        approximateLineWidthPx += tabWidthPx;
      });
      lastTextStyle = child.style ?? lastTextStyle;
      continue;
    }

    const startOffset = combinedText.length;
    combinedText += child.text;
    runs.push({
      kind: "text",
      key: `run-${childIndex}`,
      text: child.text,
      startOffset,
      endOffset: combinedText.length,
      style: child.style,
      link: child.link,
    });
    approximateLineWidthPx = updateEstimatedLineWidthPxForText(
      approximateLineWidthPx,
      child.text,
      child.style
    );
    lastTextStyle = child.style ?? lastTextStyle;
  }

  if (combinedText.length === 0) {
    return storeCachedSource(undefined);
  }

  return storeCachedSource({
    text: combinedText,
    runs,
  });
}

// --- Synthetic source / word-break / sanitize (upstream 6945-6986) ---

export function buildSyntheticPretextLayoutSource(
  text: string,
  style?: TextRunNode["style"]
): ParagraphPretextLayoutSource {
  return {
    text,
    runs: [
      {
        kind: "text",
        key: "synthetic-0",
        text,
        startOffset: 0,
        endOffset: text.length,
        style,
      },
    ],
  };
}

export function pretextWordBreakModeForText(text: string): "normal" | "keep-all" {
  const cached = pretextWordBreakModeByText.get(text);
  if (cached) {
    return cached;
  }

  const mode = KEEP_ALL_SCRIPT_RE.test(text) ? "keep-all" : "normal";
  setCacheEntry(pretextWordBreakModeByText, text, mode);
  while (pretextWordBreakModeByText.size > TEXT_MEASURE_CACHE_MAX_ENTRIES) {
    const firstKey = pretextWordBreakModeByText.keys().next().value as
      | string
      | undefined;
    if (!firstKey) {
      break;
    }
    pretextWordBreakModeByText.delete(firstKey);
  }
  return mode;
}

export function sanitizeRenderedPretextFragmentText(text: string): string {
  return text.replace(/\r\n?|\n/g, "");
}

// --- Layout items (upstream 6988-7078) ---

export function buildParagraphPretextLayoutItems(
  paragraph: ParagraphNode,
  source: ParagraphPretextLayoutSource
): PretextLayoutItem[] {
  const cachedItems = paragraphPretextLayoutItemsBySource.get(source);
  if (cachedItems) {
    return cachedItems;
  }

  const paragraphBaseFontPx = paragraphBaseFontSizePx(paragraph);
  const wordBreak = pretextWordBreakModeForText(source.text);

  const items: PretextLayoutItem[] = source.runs
    .filter((run) => run.endOffset > run.startOffset)
    .map((run) => ({
      text: run.text,
      font: resolveMeasureFont(run.style, paragraphBaseFontPx),
      startOffset: run.startOffset,
      endOffset: run.endOffset,
      break: run.kind === "image" || run.kind === "tab" ? "never" : "normal",
      wordBreak,
    }));
  paragraphPretextLayoutItemsBySource.set(source, items);
  return items;
}

/**
 * Returns the single measurement font shared by every non-empty run of this
 * pretext source, or `undefined` if there are atomic items (images, tabs)
 * or the runs vary in font. Callers can use this to take a fast
 * `measureLineStats`-based line count path when it returns a value.
 */
export function resolveUniformPretextSourceFont(
  paragraph: ParagraphNode,
  source: ParagraphPretextLayoutSource
): string | undefined {
  const cachedUniformFont = paragraphPretextUniformFontBySource.get(source);
  if (cachedUniformFont !== undefined) {
    return cachedUniformFont ?? undefined;
  }

  const paragraphBaseFontPx = paragraphBaseFontSizePx(paragraph);
  let uniformFont: string | undefined;

  for (const run of source.runs) {
    if (run.endOffset <= run.startOffset) {
      continue;
    }
    if (run.kind !== "text") {
      paragraphPretextUniformFontBySource.set(source, null);
      return undefined;
    }

    const runFont = resolveMeasureFont(run.style, paragraphBaseFontPx);
    if (uniformFont === undefined) {
      uniformFont = runFont;
      continue;
    }
    if (runFont !== uniformFont) {
      paragraphPretextUniformFontBySource.set(source, null);
      return undefined;
    }
  }

  paragraphPretextUniformFontBySource.set(source, uniformFont ?? null);
  return uniformFont;
}

export function buildMeasureSegmentsPretextLayoutItems(
  segments: ParagraphMeasureSegment[],
  paragraphBaseFontPx: number,
  text: string
): PretextLayoutItem[] {
  const wordBreak = pretextWordBreakModeForText(text);
  let nextOffset = 0;

  return segments
    .map((segment) => {
      const startOffset = nextOffset;
      nextOffset += segment.text.length;
      return {
        text: segment.text,
        font: resolveMeasureFont(segment.style, paragraphBaseFontPx),
        startOffset,
        endOffset: nextOffset,
        break: "normal" as const,
        wordBreak,
      };
    })
    .filter((item) => item.endOffset > item.startOffset);
}

// --- Layout entry point (upstream 7080-7118) ---

export function layoutParagraphPretextSource(
  paragraph: ParagraphNode,
  source: ParagraphPretextLayoutSource,
  containerWidthPx: number,
  lineHeightPx: number,
  exclusions?: PretextExclusionRect[]
): PretextVariableWidthLayout | undefined {
  const fallbackFont = resolveMeasureFont(
    firstRunStyle(paragraph),
    paragraphBaseFontSizePx(paragraph)
  );
  const wordBreak = pretextWordBreakModeForText(source.text);
  const items = buildParagraphPretextLayoutItems(paragraph, source);
  const richLayout =
    items.length > 0
      ? layoutItemsWithPretextAroundExclusions(
          source.text,
          items,
          containerWidthPx,
          lineHeightPx,
          exclusions,
          fallbackFont
        )
      : undefined;
  if (richLayout) {
    return richLayout;
  }

  return layoutTextWithPretextAroundExclusions(
    source.text,
    fallbackFont,
    containerWidthPx,
    lineHeightPx,
    exclusions,
    {
      wordBreak,
    }
  );
}

// --- Session text / inline-image placeholder / anchor offset (upstream 7120-7180) ---

export function wrappedParagraphSessionText(paragraph: ParagraphNode): string {
  return (
    buildParagraphPretextLayoutSource(paragraph)?.text ??
    paragraphText(paragraph)
  );
}

export function buildInlineImagePlaceholderText(
  image: ImageRunNode,
  style: TextRunNode["style"] | undefined,
  paragraphBaseFontPx: number
): string {
  const targetWidthPx = Math.max(
    1,
    Math.round(image.widthPx ?? image.heightPx ?? MIN_PARAGRAPH_LINE_HEIGHT_PX)
  );
  let placeholderText = "\u00a0";
  while (
    measureTextWidthPx(placeholderText, style, paragraphBaseFontPx) <
      targetWidthPx &&
    placeholderText.length < 64
  ) {
    placeholderText += "\u00a0";
  }
  return placeholderText;
}

export function paragraphChildAnchorOffset(
  paragraph: ParagraphNode,
  childIndex: number
): number {
  const paragraphBaseFontPx = paragraphBaseFontSizePx(paragraph);
  let lastTextStyle: TextRunNode["style"] | undefined =
    firstRunStyle(paragraph);
  let offset = 0;

  for (let currentIndex = 0; currentIndex < childIndex; currentIndex += 1) {
    const child = paragraph.children[currentIndex];
    if (!child) {
      continue;
    }
    if (child.type === "image") {
      if (child.floating) {
        continue;
      }
      offset += buildInlineImagePlaceholderText(
        child,
        lastTextStyle,
        paragraphBaseFontPx
      ).length;
      continue;
    }
    if (child.type !== "text") {
      continue;
    }
    offset += child.text.length;
    lastTextStyle = child.style ?? lastTextStyle;
  }

  return offset;
}

// --- Word expansion (upstream 7182-7226) ---

export function expandOffsetToWord(
  text: string,
  offset: number
): {
  start: number;
  end: number;
} {
  const safeOffset = Math.max(0, Math.min(Math.round(offset), text.length));
  if (!text) {
    return { start: 0, end: 0 };
  }

  const characterAt = (index: number): string => text.slice(index, index + 1);
  const isWordCharacter = (value: string): boolean =>
    /[0-9A-Za-z_]/.test(value);
  let probeIndex = safeOffset;
  if (probeIndex > 0 && probeIndex === text.length) {
    probeIndex -= 1;
  }
  if (
    !isWordCharacter(characterAt(probeIndex)) &&
    probeIndex > 0 &&
    isWordCharacter(characterAt(probeIndex - 1))
  ) {
    probeIndex -= 1;
  }

  if (!isWordCharacter(characterAt(probeIndex))) {
    return {
      start: safeOffset,
      end: safeOffset,
    };
  }

  let start = probeIndex;
  let end = probeIndex + 1;
  while (start > 0 && isWordCharacter(characterAt(start - 1))) {
    start -= 1;
  }
  while (end < text.length && isWordCharacter(characterAt(end))) {
    end += 1;
  }

  return { start, end };
}

// --- Pixels-to-twips conversion (upstream 7228-7230) ---

export function pixelsToTwips(valuePx: number): number {
  return Math.round(valuePx * TWIPS_PER_PIXEL);
}

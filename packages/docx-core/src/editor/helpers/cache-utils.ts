// Global cache Map/WeakMap declarations and cache utility functions.
// Upstream editor.tsx: lines 352-381, 419-468.
//
// These module-level caches are shared across editor helpers. They store XML
// parse results, height estimates, text-width measurements and font metrics.
// The ParagraphTrackedMarkup / ParagraphCommentMarkup caches (upstream
// 382-418) live in tracked-changes.ts to keep that domain self-contained.

import type {
  ParagraphIndent,
  ParagraphNode,
  TableNode
} from "../../engine/types";
import {
  documentLayoutCacheKey,
  type DocumentLayoutDefaults,
} from "../../engine/layout-defaults";
import { textMeasurementGeneration } from "../../viewer/text-measurement-generation";
import { XML_CACHE_MAX_ENTRIES } from "./constants";

export const paragraphBreakFlagsBySourceXml = new Map<
  string,
  {
    explicitPageBreak: boolean;
    explicitColumnBreak: boolean;
    lastRenderedPageBreak: boolean;
    pageBreakBefore: boolean;
    sectionBreakStartsNewPage: boolean;
  }
>();
export const paragraphEstimatedHeightBySourceXml = new Map<
  string,
  Map<number | string, number>
>();
export const tableEstimatedHeightBySourceXml = new Map<
  string,
  Map<number | string, number>
>();
export const tableEstimatedRowHeightsByNode = new WeakMap<
  TableNode,
  Map<number | string, number[]>
>();
export const paragraphExplicitIndentBySourceXml = new Map<
  string,
  ParagraphIndent | null
>();
export const paragraphDropCapBySourceXml = new Map<
  string,
  {
    type: "drop" | "margin";
    lines?: number;
  } | null
>();

export let paragraphMeasureCanvasContext: CanvasRenderingContext2D | undefined;
export const textWidthByFontAndValue = new Map<string, number>();
export const estimatedTextAdvanceWidthByFontAndValue = new Map<string, number>();
export const pretextWordBreakModeByText = new Map<string, "normal" | "keep-all">();
export const paragraphBaseFontSizePxByParagraph = new WeakMap<ParagraphNode, number>();
export const paragraphDominantFontFamilyByParagraph = new WeakMap<
  ParagraphNode,
  string | null
>();

export function setParagraphMeasureCanvasContext(
  context: CanvasRenderingContext2D | undefined
): void {
  paragraphMeasureCanvasContext = context;
}

export interface TableSpacingTwips {
  topTwips?: number;
  rightTwips?: number;
  bottomTwips?: number;
  leftTwips?: number;
}

export function setCacheEntry<K, V>(cache: Map<K, V>, key: K, value: V): void {
  if (!cache.has(key) && cache.size >= XML_CACHE_MAX_ENTRIES) {
    const oldestKey = cache.keys().next().value as K | undefined;
    if (oldestKey !== undefined) {
      cache.delete(oldestKey);
    }
  }
  cache.set(key, value);
}

export function widthCacheKeyPx(widthPx?: number): number {
  if (!Number.isFinite(widthPx) || (widthPx as number) <= 0) {
    return -1;
  }

  return Math.max(1, Math.round(widthPx as number));
}

export function heightEstimateCacheKeyPx(
  widthPx?: number,
  docGridLinePitchPx?: number,
  disableDocGridSnap = false,
  layoutDefaults?: DocumentLayoutDefaults
): number | string {
  const widthKey = widthCacheKeyPx(widthPx);
  const docGridKey =
    Number.isFinite(docGridLinePitchPx) && (docGridLinePitchPx as number) > 0
      ? Math.max(0, Math.round(docGridLinePitchPx as number))
      : 0;
  // The generation prefix retires entries measured before a web font loaded,
  // so a stale (typically narrower) estimate can never be reused.
  const measurementGeneration = textMeasurementGeneration() + 1;

  const baseKey =
    measurementGeneration * 1_000_000_000 +
    (widthKey + 2) * 10_000 +
    docGridKey * 2 +
    (disableDocGridSnap ? 1 : 0);
  // Document defaults (compatibility mode, default tab stop) change the same
  // estimate, so they extend the key instead of producing a second cache.
  const documentKey = documentLayoutCacheKey(layoutDefaults);
  return documentKey ? `${baseKey}|${documentKey}` : baseKey;
}

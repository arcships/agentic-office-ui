// Real line-index provider for Word's cached `w:lastRenderedPageBreak` marker.
//
// The layout layer can split a paragraph at a mid-paragraph marker, but it
// cannot wrap text itself. This module runs the same pretext layout used by the
// renderer and maps the marker's text offset to the wrap line that starts the
// next page.

import type { NumberingDefinitionSet, ParagraphNode } from "../../engine/types";
import { paragraphLastRenderedPageBreakTextOffset } from "../../layout/pagination";
import {
  buildParagraphPretextLayoutSource,
  layoutParagraphPretextSource,
} from "./pretext-build";
import { paragraphAvailableTextWidthPx } from "./paragraph-geometry";
import { estimateParagraphLineHeightPx } from "./line-height";

/**
 * Wrap line index at which a mid-paragraph `lastRenderedPageBreak` splits the
 * paragraph: all lines before the returned index stay on the current page.
 *
 * Returns `undefined` when the marker is absent, sits at the paragraph start
 * (node-level break) or at its end, or when pretext cannot lay the paragraph
 * out. The caller falls back to its proportional estimate in that case.
 */
export function paragraphLastRenderedPageBreakLineIndex(
  paragraph: ParagraphNode,
  lineCount: number,
  availableWidthPx?: number,
  numberingDefinitions?: NumberingDefinitionSet
): number | undefined {
  if (!Number.isFinite(lineCount) || lineCount <= 1) {
    return undefined;
  }

  const textOffset = paragraphLastRenderedPageBreakTextOffset(paragraph);
  if (textOffset === undefined || textOffset <= 0) {
    return undefined;
  }

  const source = buildParagraphPretextLayoutSource(paragraph, {
    allowExplicitLineBreakText: true,
    expandTabsForLayout: true,
  });
  if (!source || source.text.length === 0) {
    return undefined;
  }

  const lineHeightPx = Math.max(1, estimateParagraphLineHeightPx(paragraph));
  const textWidthPx = paragraphAvailableTextWidthPx(
    paragraph,
    Math.max(24, Math.round(availableWidthPx ?? 0)),
    numberingDefinitions
  );
  const layout = layoutParagraphPretextSource(
    paragraph,
    source,
    textWidthPx,
    lineHeightPx,
    []
  );
  if (!layout || layout.lines.length === 0) {
    return undefined;
  }

  const clampedOffset = Math.max(0, Math.min(textOffset, source.text.length));
  let lineIndex = -1;
  let lineStartOffset = 0;
  for (let index = 0; index < layout.lines.length; index += 1) {
    const fragments = layout.lines[index]?.fragments ?? [];
    if (fragments.length === 0) {
      continue;
    }
    const startOffset = Math.min(...fragments.map((fragment) => fragment.startOffset));
    const endOffset = Math.max(...fragments.map((fragment) => fragment.endOffset));
    if (clampedOffset >= startOffset && clampedOffset < endOffset) {
      lineIndex = index;
      lineStartOffset = startOffset;
      break;
    }
  }

  if (lineIndex < 0) {
    // Markers can sit between the last fragment of a line and the next line's
    // first character (e.g. right after a line break). Use the following line
    // in that case; if the marker is past the last line, there is no split.
    lineIndex = layout.lines.findIndex((line) => {
      const fragments = line.fragments ?? [];
      if (fragments.length === 0) {
        return false;
      }
      return (
        Math.min(...fragments.map((fragment) => fragment.startOffset)) >=
        clampedOffset
      );
    });
    if (lineIndex < 0) {
      return undefined;
    }
    return lineIndex > 0 && lineIndex < lineCount ? lineIndex : undefined;
  }

  // A marker exactly at a line start means that line already belongs to the
  // next page; a marker inside a line moves the following line.
  const nextPageLineIndex =
    clampedOffset <= lineStartOffset ? lineIndex : lineIndex + 1;
  if (nextPageLineIndex <= 0 || nextPageLineIndex >= lineCount) {
    return undefined;
  }
  return nextPageLineIndex;
}

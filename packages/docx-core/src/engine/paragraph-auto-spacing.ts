// Word's automatic (HTML) paragraph spacing.
//
// `w:spacing/@w:beforeAutospacing` and `@w:afterAutospacing` name no value: the
// consumer picks the amount. Word applies them only in compatibility mode 12
// (Word 2007) and 15 (Word 2013+), skips them between adjacent items of the
// same list, and uses a fixed 14 pt gap — the HTML paragraph rhythm that Word
// and LibreOffice both render (measured at 14.0 pt for 10 pt, 10.5 pt, 12 pt and
// 24 pt text alike).
//
// Upstream: `paragraph-auto-spacing.ts:3-19`.

import type { DocNode, ParagraphNode, ParagraphSpacing } from "./types";
import {
  documentHonoursAutoSpacing,
  twipsToPixelsExact,
  twipsToWholePixels,
  type DocumentLayoutDefaults,
} from "./layout-defaults";

/** The gap Word inserts for one `*Autospacing` flag, in twips (14 pt). */
export const HTML_AUTO_SPACING_TWIPS = 280;

/** The same gap in CSS pixels (18.67 at 96 dpi). */
export const HTML_AUTO_SPACING_PX = twipsToPixelsExact(
  HTML_AUTO_SPACING_TWIPS
);

export interface ParagraphAutoSpacingContext extends DocumentLayoutDefaults {
  /** The paragraph before this one in the same story, when known. */
  previousParagraph?: ParagraphNode;
  /** The paragraph after this one in the same story, when known. */
  nextParagraph?: ParagraphNode;
}

export interface ParagraphAutoSpacingPx {
  beforePx: number;
  afterPx: number;
}

const NO_AUTO_SPACING: ParagraphAutoSpacingPx = Object.freeze({
  beforePx: 0,
  afterPx: 0,
});

/**
 * True when both paragraphs are items of the same numbering instance. Word
 * collapses the automatic spacing across such a boundary so a list keeps its
 * own rhythm instead of growing an HTML paragraph gap per item.
 */
export function paragraphsShareListNumbering(
  previousParagraph?: ParagraphNode,
  currentParagraph?: ParagraphNode
): boolean {
  const previousNumId = previousParagraph?.style?.numbering?.numId;
  const currentNumId = currentParagraph?.style?.numbering?.numId;
  return (
    previousNumId !== undefined &&
    currentNumId !== undefined &&
    previousNumId === currentNumId
  );
}

/**
 * The automatic before/after spacing a paragraph contributes, in CSS pixels.
 * Returns zeros when the paragraph names no flags, when its compatibility mode
 * ignores them, or when the neighbour continues the same list.
 */
export function resolveParagraphAutoSpacingPx(
  paragraph: ParagraphNode,
  context?: ParagraphAutoSpacingContext
): ParagraphAutoSpacingPx {
  const spacing: ParagraphSpacing | undefined = paragraph.style?.spacing;
  if (!spacing || !documentHonoursAutoSpacing(context?.compatibilityMode)) {
    return NO_AUTO_SPACING;
  }

  const beforeApplies =
    spacing.beforeAutoSpacing === true &&
    !paragraphsShareListNumbering(context?.previousParagraph, paragraph);
  const afterApplies =
    spacing.afterAutoSpacing === true &&
    !paragraphsShareListNumbering(paragraph, context?.nextParagraph);
  if (!beforeApplies && !afterApplies) {
    return NO_AUTO_SPACING;
  }
  return {
    beforePx: beforeApplies ? HTML_AUTO_SPACING_PX : 0,
    afterPx: afterApplies ? HTML_AUTO_SPACING_PX : 0,
  };
}

/**
 * Automatic spacing of one paragraph inside a story (body, header, footer or a
 * note body): the neighbours decide whether a list boundary collapses it.
 */
export function resolveParagraphAutoSpacingInStory(
  nodes: readonly DocNode[],
  nodeIndex: number,
  paragraph: ParagraphNode,
  defaults?: DocumentLayoutDefaults
): ParagraphAutoSpacingPx {
  const previousNode = nodes[nodeIndex - 1];
  const nextNode = nodes[nodeIndex + 1];
  return resolveParagraphAutoSpacingPx(paragraph, {
    compatibilityMode: defaults?.compatibilityMode,
    previousParagraph:
      previousNode?.type === "paragraph" ? previousNode : undefined,
    nextParagraph: nextNode?.type === "paragraph" ? nextNode : undefined,
  });
}

/**
 * True when the automatic spacing between the two paragraphs is collapsed
 * because they are adjacent items of the same list.
 */
export function paragraphsSuppressAutoSpacing(
  previousParagraph: ParagraphNode | undefined,
  currentParagraph: ParagraphNode,
  defaults?: DocumentLayoutDefaults
): boolean {
  if (!documentHonoursAutoSpacing(defaults?.compatibilityMode)) {
    return false;
  }
  return paragraphsShareListNumbering(previousParagraph, currentParagraph);
}

/**
 * Explicit plus automatic spacing of one paragraph in CSS pixels. Suppression
 * is derived from the neighbours supplied in `context`, so the renderer and the
 * paginator resolve the same numbers from the same model.
 */
export function resolveParagraphSpacingPx(
  paragraph: ParagraphNode,
  context?: ParagraphAutoSpacingContext
): ParagraphAutoSpacingPx {
  const spacing = paragraph.style?.spacing;
  const auto = resolveParagraphAutoSpacingPx(paragraph, context);
  return {
    // Explicit twips follow the model/renderer conversion: whole pixels.
    beforePx: twipsToWholePixels(spacing?.beforeTwips) + auto.beforePx,
    afterPx: twipsToWholePixels(spacing?.afterTwips) + auto.afterPx,
  };
}

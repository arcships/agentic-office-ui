// Paragraph geometry: floating-image / cover-image detection and the
// floating-image render-mode predicates (shouldRenderWrappedFloatingImage etc.).
//
// Upstream editor.tsx:
//   lines 5012-5811  (floating/cover detection)
//   6368-6377        (floatingImageMovesWithText)
//   14008-14052      (shouldRenderWrappedFloatingImage, isFixedPositionWrappedFloatingImage)
//   14054-14098      (shouldRenderTopAnchoredMarginFloatAsAbsolute)
//   14720-14785      (shouldRenderAbsoluteFloatingImage, isPageOrMarginAnchored*)
//   2045-2054        (shouldReserveHeaderFooterFloatingImageSpace)
//
// The render-mode predicates live here (rather than in pretext-integration) to
// break the paragraph-inspect <-> pretext-integration cycle.

import type {
  DocModel,
  FormFieldRunNode,
  ImageRunNode,
  NumberingDefinitionSet,
  ParagraphNode,
  TableNode,
  TextRunNode
} from "../../engine/types";
import type { DocumentLayoutMetrics } from "../../viewer/section-layout";
import {
  headingLevelFromStyleLabel,
  paragraphHasImage,
  paragraphText,
  tableCellParagraphsRecursively,
  formFieldDisplayValue,
  resolveParagraphStyleHeadingLevel
} from "./paragraph-inspect";
import {
  syntheticTextBoxContainsPictureLayer,
  syntheticTextBoxParagraphsFromRunXml
} from "./synthetic-textbox";
import { WORD_IMAGE_Z_INDEX_STEP } from "./constants";
import type { TableCellContentNode } from "../../engine/types";
import { resolveListParagraphIndent } from "./xml-parsing-extra";
import { twipsToSignedPixels } from "./ooxml-helpers";
import { paragraphBorderInsetPx } from "./style-block-css";

export function paragraphHasInFlowImage(paragraph: ParagraphNode): boolean {
  return paragraph.children.some((child) => {
    if (child.type !== "image") {
      return false;
    }

    return (
      !shouldRenderAbsoluteFloatingImage(child) &&
      !shouldRenderWrappedFloatingImage(child)
    );
  });
}

export function paragraphHasTextBearingAbsoluteFloatingTextBox(
  paragraph: ParagraphNode
): boolean {
  return paragraph.children.some(
    (child) =>
      child.type === "image" &&
      shouldRenderAbsoluteFloatingImage(child) &&
      child.syntheticTextBox === true &&
      !syntheticTextBoxContainsPictureLayer(child) &&
      Boolean(floatingTextBoxVisibleTextFromImage(child))
  );
}

export function paragraphIsFloatingImageAnchorOnly(paragraph: ParagraphNode): boolean {
  if (
    paragraphHasVisibleText(paragraph) ||
    paragraphHasFormField(paragraph) ||
    paragraphHasTextBearingAbsoluteFloatingTextBox(paragraph)
  ) {
    return false;
  }

  let hasFloatingImage = false;
  for (const child of paragraph.children) {
    if (child.type === "text") {
      if (child.text.trim().length > 0) {
        return false;
      }
      continue;
    }

    if (child.type !== "image") {
      return false;
    }

    if (
      !shouldRenderWrappedFloatingImage(child) &&
      !shouldRenderAbsoluteFloatingImage(child)
    ) {
      return false;
    }

    hasFloatingImage = true;
  }

  return hasFloatingImage;
}

export function paragraphIsAbsoluteFloatingImageAnchorOnly(
  paragraph: ParagraphNode
): boolean {
  const allowSectionBreakTextBearingAnchor =
    paragraphContainsSectionBreakProperties(paragraph) &&
    !paragraphHasVisibleText(paragraph) &&
    !paragraphHasFormField(paragraph);

  if (
    paragraphHasVisibleText(paragraph) ||
    paragraphHasFormField(paragraph) ||
    (paragraphHasTextBearingAbsoluteFloatingTextBox(paragraph) &&
      !allowSectionBreakTextBearingAnchor)
  ) {
    return false;
  }

  let hasAbsoluteFloatingImage = false;
  for (const child of paragraph.children) {
    if (child.type === "text") {
      if (child.text.trim().length > 0) {
        return false;
      }
      continue;
    }

    if (child.type !== "image" || !shouldRenderAbsoluteFloatingImage(child)) {
      return false;
    }

    hasAbsoluteFloatingImage = true;
  }

  return hasAbsoluteFloatingImage;
}

export function paragraphContainsOnlyAbsoluteFloatingContent(
  paragraph: ParagraphNode
): boolean {
  if (paragraphHasFormField(paragraph)) {
    return false;
  }

  let hasAbsoluteFloatingImage = false;
  for (const child of paragraph.children) {
    if (child.type === "text") {
      continue;
    }

    if (child.type !== "image" || !shouldRenderAbsoluteFloatingImage(child)) {
      return false;
    }

    hasAbsoluteFloatingImage = true;
  }

  return hasAbsoluteFloatingImage;
}

export function paragraphIsBehindTextAbsoluteFloatingImageAnchorOnly(
  paragraph: ParagraphNode
): boolean {
  if (!paragraphIsAbsoluteFloatingImageAnchorOnly(paragraph)) {
    return false;
  }

  return paragraph.children.every(
    (child) =>
      child.type === "text" ||
      (child.type === "image" && child.floating?.behindDocument === true)
  );
}

export function imageBehavesAsDecorativeBehindTextBackground(
  image: ImageRunNode,
  paragraph: ParagraphNode
): boolean {
  return (
    image.floating?.behindDocument === true &&
    shouldRenderAbsoluteFloatingImage(image) &&
    paragraphIsBehindTextAbsoluteFloatingImageAnchorOnly(paragraph)
  );
}

export function paragraphActsAsDecorativeBehindTextBackgroundOverlay(
  paragraph: ParagraphNode
): boolean {
  return (
    paragraphIsBehindTextAbsoluteFloatingImageAnchorOnly(paragraph) &&
    !paragraphHasVisibleText(paragraph) &&
    !paragraphHasFormField(paragraph) &&
    paragraph.children.every(
      (child) => child.type !== "image" || child.syntheticTextBox !== true
    )
  );
}

export function paragraphNeedsPageWidthAnchorHost(paragraph: ParagraphNode): boolean {
  if (!paragraphIsFloatingImageAnchorOnly(paragraph)) {
    return false;
  }

  return paragraph.children.some((child) => {
    if (
      child.type !== "image" ||
      !shouldRenderAbsoluteFloatingImage(child) ||
      !child.floating
    ) {
      return false;
    }

    const horizontalRelativeTo = child.floating.horizontalRelativeTo
      ?.trim()
      .toLowerCase();
    return horizontalRelativeTo === "page" || horizontalRelativeTo === "margin";
  });
}

export function paragraphHasAbsoluteFloatingImage(paragraph: ParagraphNode): boolean {
  return paragraph.children.some(
    (child) =>
      child.type === "image" && shouldRenderAbsoluteFloatingImage(child)
  );
}

export function collectHeadingTextColorByLevel(
  model: DocModel
): Partial<Record<number, string>> {
  const colorsByLevel: Partial<Record<number, string>> = {};
  const paragraphStyleById = new Map(
    (model.metadata.paragraphStyles ?? []).map((styleDefinition) => [
      styleDefinition.id,
      styleDefinition,
    ])
  );

  const resolveParagraphHeadingLevel = (
    paragraph: ParagraphNode
  ): number | undefined => {
    if (
      Number.isFinite(paragraph.style?.headingLevel) &&
      (paragraph.style?.headingLevel as number) > 0
    ) {
      return Math.round(paragraph.style?.headingLevel as number);
    }

    const styleDefinition = paragraph.style?.styleId
      ? paragraphStyleById.get(paragraph.style.styleId)
      : undefined;
    const fromStyleDefinition = styleDefinition
      ? resolveParagraphStyleHeadingLevel(styleDefinition)
      : undefined;
    if (fromStyleDefinition) {
      return fromStyleDefinition;
    }

    return (
      headingLevelFromStyleLabel(paragraph.style?.styleId) ??
      headingLevelFromStyleLabel(paragraph.style?.styleName)
    );
  };

  const resolveParagraphTextColor = (
    paragraph: ParagraphNode
  ): string | undefined => {
    for (const child of paragraph.children) {
      if (
        (child.type === "text" || child.type === "form-field") &&
        child.style?.color
      ) {
        return child.style.color;
      }
    }

    const styleDefinition = paragraph.style?.styleId
      ? paragraphStyleById.get(paragraph.style.styleId)
      : undefined;
    return styleDefinition?.runStyle?.color;
  };

  const registerParagraph = (paragraph: ParagraphNode): void => {
    const headingLevel = resolveParagraphHeadingLevel(paragraph);
    if (!headingLevel || colorsByLevel[headingLevel]) {
      return;
    }

    const color = resolveParagraphTextColor(paragraph);
    if (!color) {
      return;
    }

    colorsByLevel[headingLevel] = color;
  };

  const walkCellNodes = (nodes: TableCellContentNode[]): void => {
    nodes.forEach((node) => {
      if (node.type === "paragraph") {
        registerParagraph(node);
        return;
      }

      node.rows.forEach((row) => {
        row.cells.forEach((cell) => {
          walkCellNodes(cell.nodes);
        });
      });
    });
  };

  model.nodes.forEach((node) => {
    if (node.type === "paragraph") {
      registerParagraph(node);
      return;
    }

    node.rows.forEach((row) => {
      row.cells.forEach((cell) => {
        walkCellNodes(cell.nodes);
      });
    });
  });

  return colorsByLevel;
}

export function sectionNodesNeedPageWideLayout(
  nodes: DocModel["nodes"],
  pageWidthPx: number,
  contentWidthPx: number
): boolean {
  const safePageWidthPx = Math.max(1, Math.round(pageWidthPx));
  const safeContentWidthPx = Math.max(1, Math.round(contentWidthPx));
  const requiresWideImageLayout = (image: ImageRunNode): boolean => {
    if (!image.floating) {
      return false;
    }

    const horizontalRelativeTo =
      image.floating.horizontalRelativeTo?.toLowerCase();
    const usesPageWidthAnchorOrigin =
      horizontalRelativeTo === "page" || horizontalRelativeTo === "margin";
    if (image.floating.behindDocument && usesPageWidthAnchorOrigin) {
      return true;
    }

    if (
      usesPageWidthAnchorOrigin &&
      (shouldRenderAbsoluteFloatingImage(image) ||
        shouldRenderWrappedFloatingImage(image))
    ) {
      return true;
    }

    if (
      Number.isFinite(image.widthPx) &&
      Number.isFinite(image.floating.xPx) &&
      (image.widthPx as number) >= safeContentWidthPx - 12 &&
      (image.floating.xPx as number) <= 12
    ) {
      return true;
    }

    if (
      Number.isFinite(image.widthPx) &&
      Number.isFinite(image.floating.xPx) &&
      (image.widthPx as number) >= safePageWidthPx - 12 &&
      (image.floating.xPx as number) <= 12
    ) {
      return true;
    }

    return false;
  };

  const paragraphNeedsWideLayout = (paragraph: ParagraphNode): boolean =>
    paragraph.children.some(
      (child) => child.type === "image" && requiresWideImageLayout(child)
    );
  const tableNeedsWideLayout = (table: TableNode): boolean =>
    table.rows.some((row) =>
      row.cells.some((cell) =>
        cell.nodes.some((cellNode) =>
          cellNode.type === "paragraph"
            ? paragraphNeedsWideLayout(cellNode)
            : tableNeedsWideLayout(cellNode)
        )
      )
    );

  return nodes.some((node) =>
    node.type === "paragraph"
      ? paragraphNeedsWideLayout(node)
      : tableNeedsWideLayout(node)
  );
}

export function sectionNodesNeedFullPageFooterOverlay(
  nodes: DocModel["nodes"]
): boolean {
  const paragraphNeedsFullPageFooterOverlay = (
    paragraph: ParagraphNode
  ): boolean => paragraphHasPageAnchoredAbsoluteFloatingImage(paragraph);
  const tableNeedsFullPageFooterOverlay = (table: TableNode): boolean =>
    table.rows.some((row) =>
      row.cells.some((cell) =>
        cell.nodes.some((cellNode) =>
          cellNode.type === "paragraph"
            ? paragraphNeedsFullPageFooterOverlay(cellNode)
            : tableNeedsFullPageFooterOverlay(cellNode)
        )
      )
    );

  return nodes.some((node) =>
    node.type === "paragraph"
      ? paragraphNeedsFullPageFooterOverlay(node)
      : tableNeedsFullPageFooterOverlay(node)
  );
}

export function paragraphHasFormField(paragraph: ParagraphNode): boolean {
  return paragraph.children.some((child) => child.type === "form-field");
}

export function paragraphHasCheckboxFormField(paragraph: ParagraphNode): boolean {
  return paragraph.children.some(
    (child) => child.type === "form-field" && child.fieldType === "checkbox"
  );
}

export function paragraphHasVisibleText(paragraph: ParagraphNode): boolean {
  return paragraph.children.some(
    (child) =>
      (child.type === "text" && child.text.trim().length > 0) ||
      (child.type === "form-field" &&
        formFieldDisplayValue(child).trim().length > 0)
  );
}

export function normalizeFloatingTextBoxComparisonText(text: string): string {
  return text
    .replace(/\r/g, "")
    .split("\n")
    .map((line) => line.replace(/[\s\u00a0]+/g, " ").trim())
    .filter((line) => line.length > 0)
    .join("\n")
    .trim();
}

export function floatingTextBoxVisibleTextFromImage(
  image: ImageRunNode
): string | undefined {
  if (
    !shouldRenderAbsoluteFloatingImage(image) ||
    !image.sourceXml ||
    !/<w:txbxContent\b/i.test(image.sourceXml)
  ) {
    return undefined;
  }

  const paragraphs = syntheticTextBoxParagraphsFromRunXml(image.sourceXml);
  if (paragraphs.length === 0) {
    return undefined;
  }

  const text = paragraphs
    .map((paragraph) =>
      paragraph.segments.map((segment) => segment.text).join("")
    )
    .join("\n");
  const normalized = normalizeFloatingTextBoxComparisonText(text);
  return normalized.length > 0 ? normalized : undefined;
}

export function paragraphVisibleTextIsOnlyAbsoluteFloatingTextBoxContent(
  paragraph: ParagraphNode
): boolean {
  if (!paragraphHasVisibleText(paragraph) || paragraphHasFormField(paragraph)) {
    return false;
  }

  const paragraphVisibleText = normalizeFloatingTextBoxComparisonText(
    paragraph.children
      .filter((child): child is TextRunNode => child.type === "text")
      .map((child) => child.text)
      .join("")
  );
  if (!paragraphVisibleText) {
    return false;
  }

  const floatingTextBoxTexts = paragraph.children
    .filter(
      (child): child is ImageRunNode =>
        child.type === "image" && shouldRenderAbsoluteFloatingImage(child)
    )
    .map((child) => floatingTextBoxVisibleTextFromImage(child))
    .filter((text): text is string => Boolean(text));

  if (floatingTextBoxTexts.length === 0) {
    return false;
  }

  if (floatingTextBoxTexts.includes(paragraphVisibleText)) {
    return true;
  }

  return (
    normalizeFloatingTextBoxComparisonText(floatingTextBoxTexts.join("\n")) ===
    paragraphVisibleText
  );
}

export function paragraphHasOnlyWhitespaceText(paragraph: ParagraphNode): boolean {
  if (paragraphHasImage(paragraph) || paragraphHasFormField(paragraph)) {
    return false;
  }

  return paragraph.children.every((child) => {
    if (child.type !== "text") {
      return false;
    }

    return child.text.replace(/[\s\u00a0]+/g, "").length === 0;
  });
}

export function paragraphHasActiveNumbering(paragraph: ParagraphNode): boolean {
  const numbering = paragraph.style?.numbering;
  return Boolean(
    numbering &&
      Number.isFinite(numbering.numId) &&
      Math.round(numbering.numId) > 0
  );
}

export function paragraphContainsSectionBreakProperties(
  paragraph: ParagraphNode
): boolean {
  return /<w:sectPr\b/i.test(paragraph.sourceXml ?? "");
}

export function paragraphAbsoluteFloatingAnchorsDependOnParagraphFlow(
  paragraph: ParagraphNode
): boolean {
  return paragraph.children.some((child) => {
    if (
      child.type !== "image" ||
      !shouldRenderAbsoluteFloatingImage(child) ||
      child.syntheticTextBox !== true ||
      !floatingTextBoxVisibleTextFromImage(child) ||
      child.floating?.behindDocument !== true
    ) {
      return false;
    }

    const verticalRelativeTo = child.floating?.verticalRelativeTo
      ?.trim()
      .toLowerCase();
    return (
      verticalRelativeTo === undefined ||
      verticalRelativeTo === "" ||
      verticalRelativeTo === "paragraph" ||
      verticalRelativeTo === "line"
    );
  });
}

export function likelyFullPageCoverImageRelativeToContentBox(
  image: ImageRunNode,
  pageContentWidthPx: number,
  pageContentHeightPx: number
): boolean {
  if (!shouldRenderAbsoluteFloatingImage(image) || !image.floating) {
    return false;
  }

  const floating = image.floating;
  if (
    floating.behindDocument !== true ||
    (floating.wrapType ?? "none") !== "none"
  ) {
    return false;
  }

  const widthPx = Math.max(0, Math.round(image.widthPx ?? 0));
  const heightPx = Math.max(0, Math.round(image.heightPx ?? 0));
  const safeContentWidthPx = Math.max(1, Math.round(pageContentWidthPx));
  const safeContentHeightPx = Math.max(1, Math.round(pageContentHeightPx));
  const widthLooksCoverSized = widthPx >= safeContentWidthPx * 0.8;
  const heightLooksCoverSized = heightPx >= safeContentHeightPx * 0.7;
  const verticalRelativeTo = floating.verticalRelativeTo?.trim().toLowerCase();

  return (
    widthLooksCoverSized &&
    heightLooksCoverSized &&
    (verticalRelativeTo === undefined ||
      verticalRelativeTo === "" ||
      verticalRelativeTo === "paragraph" ||
      verticalRelativeTo === "line")
  );
}

export function isLikelyFullPageCoverFloatingImage(
  image: ImageRunNode,
  pageWidthPx: number,
  pageHeightPx: number
): boolean {
  return likelyFullPageCoverImageRelativeToContentBox(
    image,
    pageWidthPx,
    pageHeightPx
  );
}

export function resolveLinkedImageWrapperStyle(params: {
  baseStyle?: Record<string, string | number | undefined>;
  cursor?: Record<string, string | number | undefined>["cursor"];
}): Record<string, string | number | undefined> {
  const { baseStyle, cursor } = params;
  const hasPositionedBaseStyle =
    baseStyle !== undefined && Object.keys(baseStyle).length > 0;

  return {
    // Floating images carry their positioning (float/position/offsets) on the
    // wrapper so the hyperlink anchor occupies the same box as the image.
    ...(hasPositionedBaseStyle ? baseStyle : { display: "inline-block" }),
    lineHeight: 0,
    verticalAlign: "middle",
    ...(cursor !== undefined ? { cursor } : undefined),
  };
}

export function paragraphIsLikelyFullPageCoverArtAnchor(
  paragraph: ParagraphNode,
  pageContentWidthPx: number,
  pageContentHeightPx: number
): boolean {
  return (
    paragraphContainsOnlyAbsoluteFloatingContent(paragraph) &&
    paragraph.children.some(
      (child) =>
        child.type === "image" &&
        likelyFullPageCoverImageRelativeToContentBox(
          child,
          pageContentWidthPx,
          pageContentHeightPx
        )
    )
  );
}

export function pageAnchoredImageLikelyStartsNearPageTop(
  image: ImageRunNode,
  layout: DocumentLayoutMetrics
): boolean {
  if (!shouldRenderAbsoluteFloatingImage(image) || !image.floating) {
    return false;
  }

  const verticalRelativeTo = image.floating.verticalRelativeTo
    ?.trim()
    .toLowerCase();
  if (verticalRelativeTo !== "page" && verticalRelativeTo !== "margin") {
    return false;
  }

  const imageHeightPx = Math.max(0, Math.round(image.heightPx ?? 0));
  const imageWidthPx = Math.max(0, Math.round(image.widthPx ?? 0));
  const coverSized =
    imageWidthPx >= Math.round(layout.pageWidthPx * 0.8) &&
    imageHeightPx >= Math.round(layout.pageHeightPx * 0.7);
  if (!coverSized) {
    return false;
  }

  const topOffsetPx = image.floating.yPx ?? 0;
  return topOffsetPx <= layout.marginsPx.top;
}

export function paragraphActsAsPageAnchoredCoverOverlayHost(
  paragraph: ParagraphNode,
  layout: DocumentLayoutMetrics
): boolean {
  if (!paragraphNeedsPageWidthAnchorHost(paragraph)) {
    return false;
  }

  return paragraph.children.some(
    (child) =>
      child.type === "image" &&
      pageAnchoredImageLikelyStartsNearPageTop(child, layout)
  );
}

export function paragraphStartsNormalFlowContent(paragraph: ParagraphNode): boolean {
  if (!paragraphHasVisibleText(paragraph)) {
    return false;
  }

  if (paragraphVisibleTextIsOnlyAbsoluteFloatingTextBoxContent(paragraph)) {
    return false;
  }

  return !paragraph.children.some(
    (child) =>
      child.type === "image" && shouldRenderAbsoluteFloatingImage(child)
  );
}

export function paragraphParticipatesInLeadingCoverLayout(
  model: DocModel,
  nodeIndex: number,
  pageContentWidthPx: number,
  pageContentHeightPx: number
): boolean {
  let sawLikelyCoverArtAnchor = false;
  for (let probeIndex = 0; probeIndex <= nodeIndex; probeIndex += 1) {
    const probeNode = model.nodes[probeIndex];
    if (!probeNode || probeNode.type !== "paragraph") {
      return false;
    }

    if (paragraphStartsNormalFlowContent(probeNode)) {
      return false;
    }

    if (
      paragraphIsLikelyFullPageCoverArtAnchor(
        probeNode,
        pageContentWidthPx,
        pageContentHeightPx
      )
    ) {
      sawLikelyCoverArtAnchor = true;
    }
  }

  return sawLikelyCoverArtAnchor;
}

export function fullPageCoverImageRenderKey(
  nodeIndex: number,
  childIndex: number
): string {
  return `${nodeIndex}:${childIndex}`;
}

export function paragraphActsAsLeadingCoverLayoutOverlay(
  model: DocModel,
  nodeIndex: number,
  paragraph: ParagraphNode,
  pageContentWidthPx: number,
  pageContentHeightPx: number
): boolean {
  if (
    !paragraphParticipatesInLeadingCoverLayout(
      model,
      nodeIndex,
      pageContentWidthPx,
      pageContentHeightPx
    ) ||
    paragraphIsLikelyFullPageCoverArtAnchor(
      paragraph,
      pageContentWidthPx,
      pageContentHeightPx
    ) ||
    (paragraphHasVisibleText(paragraph) &&
      !paragraphVisibleTextIsOnlyAbsoluteFloatingTextBoxContent(paragraph)) ||
    paragraphHasFormField(paragraph)
  ) {
    return false;
  }

  return paragraph.children.every(
    (child) =>
      child.type === "text" ||
      (child.type === "image" && shouldRenderAbsoluteFloatingImage(child))
  );
}

export function fullPageCoverAbsoluteFloatingImageStyle(
  image: ImageRunNode,
  layout: DocumentLayoutMetrics,
  options?: {
    deltaX?: number;
    deltaY?: number;
    anchorToPageSurface?: boolean;
  }
): Record<string, string | number | undefined> {
  const floating = image.floating;
  const normalizedZIndex = Number.isFinite(floating?.zIndex)
    ? Math.max(
        1,
        Math.min(
          65535,
          Math.round((floating?.zIndex as number) / WORD_IMAGE_Z_INDEX_STEP)
        )
      )
    : 1;
  return {
    position: "absolute",
    left:
      (options?.anchorToPageSurface ? 0 : -layout.marginsPx.left) +
      Math.round(options?.deltaX ?? 0),
    top:
      (options?.anchorToPageSurface ? 0 : -layout.marginsPx.top) +
      Math.round(options?.deltaY ?? 0),
    width: layout.pageWidthPx,
    height: layout.pageHeightPx,
    zIndex: floating?.behindDocument === true ? 0 : normalizedZIndex,
  };
}

export function paragraphLooksLikeCheckboxChoiceRow(
  paragraph: ParagraphNode
): boolean {
  if (paragraph.children.some((child) => child.type === "image")) {
    return false;
  }

  const checkboxCount = paragraph.children.filter(
    (child) => child.type === "form-field" && child.fieldType === "checkbox"
  ).length;
  if (checkboxCount < 2) {
    return false;
  }

  const combinedText = paragraph.children
    .filter((child): child is TextRunNode => child.type === "text")
    .map((child) => child.text)
    .join("");
  if (!combinedText.includes("\t")) {
    return false;
  }

  const normalized = combinedText.replace(/\s+/g, " ").trim().toLowerCase();
  return normalized.includes("yes") && normalized.includes("no");
}
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

// Paragraph geometry: floating-image / cover-image detection helpers.
//
// Upstream editor.tsx:
//   lines 5012-5811  (floating/cover detection)

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
import {
  floatingImageMovesWithText,
  isPageOrMarginAnchoredAbsoluteFloatingImage,
  paragraphHasPageAnchoredAbsoluteFloatingImage,
  shouldRenderAbsoluteFloatingImage,
  shouldRenderWrappedFloatingImage,
} from "./paragraph-geometry-image";

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

// Re-export render-mode predicates from the split module.
export * from "./paragraph-geometry-image";

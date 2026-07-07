// Synthetic text-box SVG rendering: Word text boxes / shapes rendered as SVG.
// Upstream editor.tsx: lines 16414-17031, 17158-17169.

import type { FormFieldRunNode, ImageRunNode, ParagraphAlignment, TextRunNode } from "../../engine/types";
import { twipsToPixels } from "../../viewer/section-layout";
import { xmlAttribute } from "./ooxml-helpers";
import {
  balancedTagXmlBlocks,
  decodeXmlText,
  mergeTextRunStyles,
  parseParagraphAlignmentFromXml,
  parseRunStyleFromRunXml
} from "./xml-parsing";
import {
  type PageFieldKind,
  instructionTextToPageFieldKind,
  instructionTextToStyleRefTarget
} from "./paragraph-toc";
import { formatPageFieldValue } from "./numbering";
import { WORD_IMAGE_Z_INDEX_STEP } from "./constants";
import { estimateTextAdvanceWidthPx } from "./drop-cap";

export function escapeSvgText(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

export function svgDataUri(svg: string): string {
  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
}

export interface SyntheticTextBoxSegment {
  text: string;
  style?: TextRunNode["style"];
}

export interface SyntheticTextBoxParagraph {
  align?: ParagraphAlignment;
  lineHeightPx: number;
  segments: SyntheticTextBoxSegment[];
}

export interface SyntheticTextBoxFrameStyle {
  backgroundColor?: string;
  borderColor?: string;
  borderWidthPx: number;
  paddingLeftPx: number;
  paddingTopPx: number;
  paddingRightPx: number;
  paddingBottomPx: number;
}

export function emuToPixels(value: number | undefined): number | undefined {
  if (!Number.isFinite(value)) {
    return undefined;
  }

  return Math.max(0, Number(((value as number) / 9525).toFixed(3)));
}

export function resolveSyntheticTextBoxFieldText(
  rawText: string,
  fieldKind: PageFieldKind | undefined,
  styleRefTarget: string | undefined,
  pageNumber: number | undefined,
  totalPages: number | undefined,
  pageNumberFormat: string | undefined,
  resolveStyleRefFieldValue?: (target: string) => string | undefined
): string {
  if (fieldKind === "PAGE") {
    const resolvedFieldValue =
      Number.isFinite(pageNumber) && (pageNumber as number) > 0
        ? Math.max(1, Math.round(pageNumber as number))
        : undefined;
    if (Number.isFinite(resolvedFieldValue)) {
      const leadingWhitespace = rawText.match(/^\s*/)?.[0] ?? "";
      const trailingWhitespace = rawText.match(/\s*$/)?.[0] ?? "";
      return `${leadingWhitespace}${formatPageFieldValue(
        Math.round(resolvedFieldValue as number),
        pageNumberFormat
      )}${trailingWhitespace}`;
    }
  }

  if (fieldKind === "NUMPAGES") {
    const resolvedFieldValue =
      Number.isFinite(totalPages) && (totalPages as number) > 0
        ? Math.max(1, Math.round(totalPages as number))
        : undefined;
    if (Number.isFinite(resolvedFieldValue)) {
      const leadingWhitespace = rawText.match(/^\s*/)?.[0] ?? "";
      const trailingWhitespace = rawText.match(/\s*$/)?.[0] ?? "";
      return `${leadingWhitespace}${Math.round(
        resolvedFieldValue as number
      )}${trailingWhitespace}`;
    }
  }

  if (styleRefTarget) {
    const resolved = resolveStyleRefFieldValue?.(styleRefTarget)?.trim();
    if (resolved) {
      const leadingWhitespace = rawText.match(/^\s*/)?.[0] ?? "";
      const trailingWhitespace = rawText.match(/\s*$/)?.[0] ?? "";
      return `${leadingWhitespace}${resolved}${trailingWhitespace}`;
    }
  }

  return rawText;
}

export function syntheticTextBoxParagraphsFromRunXml(
  runXml: string,
  pageNumber?: number,
  totalPages?: number,
  pageNumberFormat?: string,
  resolveStyleRefFieldValue?: (target: string) => string | undefined
): SyntheticTextBoxParagraph[] {
  const textBoxXml = balancedTagXmlBlocks(runXml, "w:txbxContent")[0];
  if (!textBoxXml) {
    return [];
  }

  const resolved: SyntheticTextBoxParagraph[] = [];
  for (const paragraphXml of balancedTagXmlBlocks(textBoxXml, "w:p")) {
    const paragraphPropertiesXml =
      balancedTagXmlBlocks(paragraphXml, "w:pPr")[0] ??
      paragraphXml.match(/<w:pPr\b[^>]*\/?>/i)?.[0] ??
      "";
    const paragraphRunPropertiesXml =
      balancedTagXmlBlocks(paragraphPropertiesXml, "w:rPr")[0] ??
      paragraphPropertiesXml.match(/<w:rPr\b[^>]*\/?>/i)?.[0] ??
      "";
    const paragraphStyle = parseRunStyleFromRunXml(
      `<w:r>${paragraphRunPropertiesXml}</w:r>`
    );
    const spacingTag =
      paragraphPropertiesXml.match(/<w:spacing\b[^>]*\/?>/i)?.[0] ?? "";
    const lineRaw = spacingTag
      ? Number(xmlAttribute(spacingTag, "w:line"))
      : Number.NaN;
    const lineRule = spacingTag
      ? xmlAttribute(spacingTag, "w:lineRule")?.trim().toLowerCase()
      : undefined;
    const lineHeightPx =
      Number.isFinite(lineRaw) &&
      (lineRaw as number) > 0 &&
      lineRule === "exact"
        ? Math.max(1, twipsToPixels(lineRaw as number) ?? 1)
        : undefined;

    const segments: SyntheticTextBoxSegment[] = [];
    const fieldStack: Array<{
      pageFieldKind?: PageFieldKind;
      styleRefTarget?: string;
      inResult: boolean;
      instructionStyle?: TextRunNode["style"];
    }> = [];

    const appendSegment = (
      text: string,
      style?: TextRunNode["style"]
    ): void => {
      if (!text) {
        return;
      }
      const previous = segments[segments.length - 1];
      if (
        previous &&
        JSON.stringify(previous.style ?? {}) === JSON.stringify(style ?? {})
      ) {
        previous.text += text;
        return;
      }
      segments.push({ text, style });
    };

    for (const runBlockXml of balancedTagXmlBlocks(paragraphXml, "w:r")) {
      const runStyle = mergeTextRunStyles(
        paragraphStyle,
        parseRunStyleFromRunXml(runBlockXml)
      );

      const beginCount =
        runBlockXml.match(/<w:fldChar\b[^>]*\bw:fldCharType="begin"[^>]*\/?>/gi)
          ?.length ?? 0;
      for (let index = 0; index < beginCount; index += 1) {
        fieldStack.push({ inResult: false });
      }

      for (const instructionMatch of runBlockXml.matchAll(
        /<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/gi
      )) {
        const pageFieldKind = instructionTextToPageFieldKind(
          instructionMatch[1] ?? ""
        );
        const styleRefTarget = instructionTextToStyleRefTarget(
          instructionMatch[1] ?? ""
        );
        if ((!pageFieldKind && !styleRefTarget) || fieldStack.length === 0) {
          continue;
        }

        for (
          let stackIndex = fieldStack.length - 1;
          stackIndex >= 0;
          stackIndex -= 1
        ) {
          if (
            !fieldStack[stackIndex].pageFieldKind &&
            !fieldStack[stackIndex].styleRefTarget
          ) {
            fieldStack[stackIndex].pageFieldKind = pageFieldKind;
            fieldStack[stackIndex].styleRefTarget = styleRefTarget;
            fieldStack[stackIndex].instructionStyle = mergeTextRunStyles(
              fieldStack[stackIndex].instructionStyle,
              runStyle
            );
            break;
          }
        }
      }

      const separateCount =
        runBlockXml.match(
          /<w:fldChar\b[^>]*\bw:fldCharType="separate"[^>]*\/?>/gi
        )?.length ?? 0;
      for (let index = 0; index < separateCount; index += 1) {
        for (
          let stackIndex = fieldStack.length - 1;
          stackIndex >= 0;
          stackIndex -= 1
        ) {
          if (!fieldStack[stackIndex].inResult) {
            fieldStack[stackIndex].inResult = true;
            break;
          }
        }
      }

      const activeField = (() => {
        for (
          let stackIndex = fieldStack.length - 1;
          stackIndex >= 0;
          stackIndex -= 1
        ) {
          const stackEntry = fieldStack[stackIndex];
          if (
            stackEntry.inResult &&
            (stackEntry.pageFieldKind || stackEntry.styleRefTarget)
          ) {
            return stackEntry;
          }
        }
        return undefined;
      })();

      for (const textMatch of runBlockXml.matchAll(
        /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/gi
      )) {
        const rawText = decodeXmlText(textMatch[1] ?? "");
        const resolvedText = activeField
          ? resolveSyntheticTextBoxFieldText(
              rawText,
              activeField.pageFieldKind,
              activeField.styleRefTarget,
              pageNumber,
              totalPages,
              pageNumberFormat,
              resolveStyleRefFieldValue
            )
          : rawText;
        appendSegment(
          resolvedText,
          activeField?.instructionStyle
            ? mergeTextRunStyles(activeField.instructionStyle, runStyle)
            : runStyle
        );
      }

      const endCount =
        runBlockXml.match(/<w:fldChar\b[^>]*\bw:fldCharType="end"[^>]*\/?>/gi)
          ?.length ?? 0;
      for (let index = 0; index < endCount; index += 1) {
        fieldStack.pop();
      }
    }

    if (segments.length === 0) {
      continue;
    }

    const effectiveFontSizePt =
      segments.find((segment) => segment.style?.fontSizePt)?.style
        ?.fontSizePt ?? 12;
    resolved.push({
      align: parseParagraphAlignmentFromXml(paragraphPropertiesXml),
      lineHeightPx:
        lineHeightPx ??
        Math.max(
          14,
          Math.round((((effectiveFontSizePt ?? 12) * 96) / 72) * 1.24)
        ),
      segments,
    });
  }

  return resolved;
}

export function syntheticTextBoxFrameStyleFromRunXml(
  runXml: string
): SyntheticTextBoxFrameStyle {
  const bodyPrTag =
    balancedTagXmlBlocks(runXml, "wps:bodyPr")[0] ??
    runXml.match(/<wps:bodyPr\b[^>]*\/?>/i)?.[0] ??
    "";
  const lineTag =
    balancedTagXmlBlocks(runXml, "a:ln")[0] ??
    runXml.match(/<a:ln\b[\s\S]*?<\/a:ln>/i)?.[0] ??
    runXml.match(/<a:ln\b[^>]*\/?>/i)?.[0] ??
    "";
  const shapePropsXml =
    balancedTagXmlBlocks(runXml, "wps:spPr")[0] ??
    runXml.match(/<wps:spPr\b[\s\S]*?<\/wps:spPr>/i)?.[0] ??
    "";

  const hasNoFill =
    /<a:noFill\b/i.test(shapePropsXml) || /<wps:noFill\b/i.test(shapePropsXml);
  const lineHasNoFill = /<a:noFill\b/i.test(lineTag);
  const fillColor =
    shapePropsXml.match(
      /<a:solidFill>\s*<a:srgbClr\b[^>]*val="([^"]+)"/i
    )?.[1] ??
    shapePropsXml.match(
      /<a:solidFill>\s*<a:schemeClr\b[^>]*val="([^"]+)"/i
    )?.[1];
  const lineColor =
    lineTag.match(/<a:solidFill>\s*<a:srgbClr\b[^>]*val="([^"]+)"/i)?.[1] ??
    lineTag.match(/<a:solidFill>\s*<a:schemeClr\b[^>]*val="([^"]+)"/i)?.[1];
  const lineWidthEmu = Number(xmlAttribute(lineTag, "w"));
  const resolvedBorderWidthPx =
    !lineTag || lineHasNoFill
      ? 0
      : Number.isFinite(lineWidthEmu) && (lineWidthEmu as number) > 0
      ? Math.max(1, Math.round((lineWidthEmu as number) / 9525))
      : 0;

  return {
    backgroundColor: !hasNoFill && fillColor ? `#${fillColor}` : undefined,
    borderColor:
      resolvedBorderWidthPx > 0 && lineColor ? `#${lineColor}` : undefined,
    borderWidthPx: resolvedBorderWidthPx,
    paddingLeftPx: emuToPixels(Number(xmlAttribute(bodyPrTag, "lIns"))) ?? 6,
    paddingTopPx: emuToPixels(Number(xmlAttribute(bodyPrTag, "tIns"))) ?? 3,
    paddingRightPx: emuToPixels(Number(xmlAttribute(bodyPrTag, "rIns"))) ?? 6,
    paddingBottomPx: emuToPixels(Number(xmlAttribute(bodyPrTag, "bIns"))) ?? 3,
  };
}

export function syntheticTextBoxTextValue(image: ImageRunNode): string | undefined {
  const explicit = image.textBoxText;
  if (typeof explicit === "string") {
    return explicit;
  }

  const paragraphs = image.sourceXml
    ? syntheticTextBoxParagraphsFromRunXml(image.sourceXml)
    : [];
  if (paragraphs.length === 0) {
    return undefined;
  }

  return paragraphs
    .map((paragraph) =>
      paragraph.segments.map((segment) => segment.text).join("")
    )
    .join("\n");
}

export function resolveSyntheticTextBoxParagraphs(
  image: ImageRunNode,
  pageNumber?: number,
  totalPages?: number,
  pageNumberFormat?: string,
  resolveStyleRefFieldValue?: (target: string) => string | undefined
): SyntheticTextBoxParagraph[] {
  const baseParagraphs = image.sourceXml
    ? syntheticTextBoxParagraphsFromRunXml(
        image.sourceXml,
        pageNumber,
        totalPages,
        pageNumberFormat,
        resolveStyleRefFieldValue
      )
    : [];
  const explicitText = image.textBoxText;
  if (typeof explicitText !== "string") {
    return baseParagraphs;
  }

  const lines = explicitText.split(/\r?\n/);
  if (lines.length === 0) {
    return baseParagraphs;
  }

  return lines.map((line, index) => {
    const baseParagraph = baseParagraphs[index] ?? baseParagraphs[0];
    const baseStyle = baseParagraph?.segments[0]?.style;
    return {
      align: baseParagraph?.align,
      lineHeightPx: baseParagraph?.lineHeightPx ?? 18,
      segments: [{ text: line, style: baseStyle }],
    };
  });
}

export function syntheticTextBoxForegroundZIndex(image: ImageRunNode): number {
  const normalizedZIndex = Number.isFinite(image.floating?.zIndex)
    ? Math.max(
        1,
        Math.min(
          65535,
          Math.round(
            (image.floating?.zIndex as number) / WORD_IMAGE_Z_INDEX_STEP
          )
        )
      )
    : 4;

  return Math.max(4, normalizedZIndex);
}

export function syntheticTextBoxSvg(
  image: ImageRunNode,
  pageNumber?: number,
  totalPages?: number,
  pageNumberFormat?: string,
  resolveStyleRefFieldValue?: (target: string) => string | undefined
): string | undefined {
  if (!image.syntheticTextBox || !image.sourceXml) {
    return undefined;
  }

  // Grouped drawings that already import as a combined SVG image + textbox
  // should keep that original synthetic SVG. Rebuilding them here from only
  // the textbox XML drops the picture layer and regresses letterhead/logo art.
  if (syntheticTextBoxContainsPictureLayer(image)) {
    return undefined;
  }

  const paragraphs = resolveSyntheticTextBoxParagraphs(
    image,
    pageNumber,
    totalPages,
    pageNumberFormat,
    resolveStyleRefFieldValue
  );
  if (paragraphs.length === 0) {
    return undefined;
  }

  const safeWidth = Math.max(8, Math.round(image.widthPx ?? 320));
  const estimatedHeight = paragraphs.reduce(
    (sum, paragraph) => sum + paragraph.lineHeightPx,
    6
  );
  const safeHeight = Math.max(8, Math.round(image.heightPx ?? estimatedHeight));
  let cursorY = 0;
  const textBlocks: string[] = [];

  for (const paragraph of paragraphs) {
    const nextCursorY = cursorY + paragraph.lineHeightPx;
    const clampedBaselineY = Math.min(nextCursorY, Math.max(1, safeHeight - 1));
    if (nextCursorY > safeHeight + 1 && textBlocks.length > 0) {
      break;
    }
    cursorY = nextCursorY;

    const fullText = paragraph.segments.map((segment) => segment.text).join("");
    const baseStyle = paragraph.segments.find(
      (segment) => segment.style
    )?.style;
    const baseFontSizePx = Math.max(
      10,
      Math.round(((baseStyle?.fontSizePt ?? 12) * 96) / 72)
    );
    const estimatedTextWidth = Math.max(
      1,
      estimateTextAdvanceWidthPx(fullText, baseStyle)
    );
    const fitScale =
      estimatedTextWidth > safeWidth && estimatedTextWidth > 0
        ? safeWidth / estimatedTextWidth
        : 1;
    const textAlign = paragraph.align ?? "left";
    const anchor =
      textAlign === "center"
        ? "middle"
        : textAlign === "right"
        ? "end"
        : "start";
    const x =
      textAlign === "center"
        ? Math.round(safeWidth / 2)
        : textAlign === "right"
        ? safeWidth
        : 0;
    let segmentOffsetPx = 0;
    const tspans = paragraph.segments
      .map((segment, segmentIndex) => {
        const segmentStyle = mergeTextRunStyles(baseStyle, segment.style);
        const segmentFontSizePx = Math.max(
          10,
          Math.round((((segmentStyle?.fontSizePt ?? 12) * 96) / 72) * fitScale)
        );
        const textDecoration = [
          segmentStyle?.underline ? "underline" : "",
          segmentStyle?.strike ? "line-through" : "",
        ]
          .filter(Boolean)
          .join(" ");
        const attrs = [
          `font-size="${segmentFontSizePx}"`,
          segmentStyle?.fontFamily
            ? `font-family="${escapeSvgText(segmentStyle.fontFamily)}"`
            : "",
          segmentStyle?.bold ? 'font-weight="700"' : "",
          segmentStyle?.italic ? 'font-style="italic"' : "",
          segmentStyle?.color
            ? `fill="${segmentStyle.color}"`
            : 'fill="#000000"',
          textDecoration ? `text-decoration="${textDecoration}"` : "",
        ]
          .filter(Boolean)
          .join(" ");
        const xAttr =
          textAlign === "left" && segmentIndex > 0
            ? ` dx="${Math.round(segmentOffsetPx)}"`
            : segmentIndex > 0
            ? ' dx="0"'
            : "";
        segmentOffsetPx = 0;
        return `<tspan${xAttr} ${attrs}>${escapeSvgText(segment.text)}</tspan>`;
      })
      .join("");

    if (textAlign === "left") {
      let runningWidth = 0;
      const leftTspans = paragraph.segments
        .map((segment) => {
          const segmentStyle = mergeTextRunStyles(baseStyle, segment.style);
          const segmentFontSizePx = Math.max(
            10,
            Math.round(
              (((segmentStyle?.fontSizePt ?? 12) * 96) / 72) * fitScale
            )
          );
          const textDecoration = [
            segmentStyle?.underline ? "underline" : "",
            segmentStyle?.strike ? "line-through" : "",
          ]
            .filter(Boolean)
            .join(" ");
          const attrs = [
            `x="${Math.round(runningWidth)}"`,
            `font-size="${segmentFontSizePx}"`,
            segmentStyle?.fontFamily
              ? `font-family="${escapeSvgText(segmentStyle.fontFamily)}"`
              : "",
            segmentStyle?.bold ? 'font-weight="700"' : "",
            segmentStyle?.italic ? 'font-style="italic"' : "",
            segmentStyle?.color
              ? `fill="${segmentStyle.color}"`
              : 'fill="#000000"',
            textDecoration ? `text-decoration="${textDecoration}"` : "",
          ]
            .filter(Boolean)
            .join(" ");
          const block = `<tspan ${attrs}>${escapeSvgText(
            segment.text
          )}</tspan>`;
          runningWidth +=
            estimateTextAdvanceWidthPx(segment.text, segmentStyle) * fitScale;
          return block;
        })
        .join("");
      textBlocks.push(
        `<text y="${Math.round(
          clampedBaselineY
        )}" text-anchor="${anchor}">${leftTspans}</text>`
      );
      continue;
    }

    textBlocks.push(
      `<text x="${x}" y="${Math.round(
        clampedBaselineY
      )}" text-anchor="${anchor}">${tspans}</text>`
    );
  }

  return svgDataUri(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${safeWidth}" height="${safeHeight}" viewBox="0 0 ${safeWidth} ${safeHeight}">${textBlocks.join(
      ""
    )}</svg>`
  );
}

export function syntheticTextBoxContainsPictureLayer(image: ImageRunNode): boolean {
  if (!image.syntheticTextBox || !image.sourceXml) {
    return false;
  }

  if (/<pic:pic\b|<a:blip\b/i.test(image.sourceXml)) {
    return true;
  }

  if (!/<wpg:wgp\b/i.test(image.sourceXml)) {
    return false;
  }

  const groupedShapes = balancedTagXmlBlocks(image.sourceXml, "wps:wsp");
  return groupedShapes.some((shapeXml) => {
    const shapePropertiesXml =
      balancedTagXmlBlocks(shapeXml, "wps:spPr")[0] ?? "";
    if (!shapePropertiesXml || /<w:txbxContent\b/i.test(shapeXml)) {
      return false;
    }

    const lineXml = balancedTagXmlBlocks(shapePropertiesXml, "a:ln")[0] ?? "";
    return (
      /<a:solidFill\b|<a:gradFill\b|<a:blipFill\b/i.test(shapePropertiesXml) ||
      (Boolean(lineXml) && !/<a:noFill\b/i.test(lineXml))
    );
  });
}

export function syntheticTextBoxActsAsTopAndBottomMasthead(
  image: ImageRunNode
): boolean {
  if (!image.syntheticTextBox || !image.sourceXml) {
    return false;
  }

  return (
    /<wpg:wgp\b/i.test(image.sourceXml) && /<wps:wsp\b/i.test(image.sourceXml)
  );
}


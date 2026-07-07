// Numbering / list label generation, list-type detection, list indent shifts,
// and clone helpers for paragraph/table border styles.
// Upstream editor.tsx: lines 19581-20509.

import type {
  DocModel,
  FormFieldRunNode,
  NumberingDefinitionSet,
  NumberingLevelDefinition,
  ParagraphBorderSet,
  ParagraphBorderStyle,
  ParagraphIndent,
  ParagraphNode,
  TableBorderSet,
  TableBorderStyle,
  TableCellStyle,
  TableRowStyle,
  TableNode,
  TextRunNode
} from "../../engine/types";
import type { DocxListType } from "./editor-types";
import { TWIPS_PER_PIXEL } from "../../viewer/section-layout";
import {
  DEFAULT_LIST_HANGING_TWIPS,
  DEFAULT_SPLIT_PARAGRAPH_AFTER_TWIPS,
  DEFAULT_SPLIT_PARAGRAPH_LINE_TWIPS,
  LIST_LEVEL_STEP_TWIPS,
  MAX_FALLBACK_LIST_LEVEL,
  UNORDERED_LIST_PREFIX_PATTERN,
  ORDERED_LIST_PREFIX_PATTERN,
  LIST_PREFIX_PATTERN,
  ORDERED_LIST_PREFIX_CAPTURE_PATTERN
} from "./constants";
import { cloneTextStyle } from "./text-mutation";
import { clampNumber } from "./zoom-utils";
import { firstRunStyle, paragraphText, tableCellParagraphs } from "./paragraph-inspect";

export function numberToRoman(value: number): string {
  if (!Number.isFinite(value) || value <= 0) {
    return "";
  }

  const numerals: Array<[number, string]> = [
    [1000, "M"],
    [900, "CM"],
    [500, "D"],
    [400, "CD"],
    [100, "C"],
    [90, "XC"],
    [50, "L"],
    [40, "XL"],
    [10, "X"],
    [9, "IX"],
    [5, "V"],
    [4, "IV"],
    [1, "I"],
  ];

  let remaining = Math.floor(value);
  let output = "";

  for (const [base, numeral] of numerals) {
    while (remaining >= base) {
      output += numeral;
      remaining -= base;
    }
  }

  return output;
}

export function numberToLetters(value: number): string {
  if (!Number.isFinite(value) || value <= 0) {
    return "";
  }

  let remaining = Math.floor(value);
  let output = "";
  while (remaining > 0) {
    const offset = (remaining - 1) % 26;
    output = String.fromCharCode(65 + offset) + output;
    remaining = Math.floor((remaining - 1) / 26);
  }

  return output;
}

export interface ParagraphNumberingLabel {
  text?: string;
  fontFamily?: string;
  color?: string;
  style?: TextRunNode["style"];
  imageSrc?: string;
  imageWidthPx?: number;
  imageHeightPx?: number;
  trailingText?: string;
}

export function formatNumberingCounter(
  format: string | undefined,
  value: number
): string {
  if (!Number.isFinite(value) || value <= 0) {
    return "";
  }

  switch ((format ?? "decimal").toLowerCase()) {
    case "decimal":
      return String(value);
    case "upperroman":
      return numberToRoman(value);
    case "lowerroman":
      return numberToRoman(value).toLowerCase();
    case "upperletter":
      return numberToLetters(value);
    case "lowerletter":
      return numberToLetters(value).toLowerCase();
    case "ordinal":
    case "cardinaltext":
    case "ordinaltext":
      return String(value);
    case "bullet":
      return "\u2022";
    case "none":
      return "";
    default:
      return String(value);
  }
}

export function formatPageFieldValue(
  value: number,
  format: string | undefined
): string {
  const formatted = formatNumberingCounter(format, value);
  return formatted.length > 0 ? formatted : String(value);
}

export function findNumberingLevelDefinition(
  numbering: NumberingDefinitionSet,
  numId: number,
  ilvl: number
): NumberingLevelDefinition | undefined {
  const instance = numbering.instances.find((item) => item.numId === numId);
  if (!instance) {
    return undefined;
  }

  const overrideLevel = instance.levelOverrides?.find(
    (item) => item.ilvl === ilvl
  );
  if (overrideLevel) {
    return overrideLevel;
  }

  const abstract = numbering.abstracts.find(
    (item) => item.abstractNumId === instance.abstractNumId
  );
  return abstract?.levels.find((item) => item.ilvl === ilvl);
}

export function numberingLevelHasVisibleMarker(
  level: NumberingLevelDefinition | undefined
): boolean {
  return Boolean(
    level?.pictureBullet?.src || (level?.text && level.text.trim().length > 0)
  );
}

export function numberingLevelIsBulletLike(
  level: NumberingLevelDefinition | undefined
): boolean {
  const format = level?.format?.trim().toLowerCase();
  return (
    format === "bullet" ||
    Boolean(level?.pictureBullet?.src) ||
    isBulletLikeNumberingText(level?.text ?? "", level?.bulletFontFamily)
  );
}

export function numberingAbstractLevelsForNumId(
  numberingDefinitions: NumberingDefinitionSet,
  numId: number
): NumberingLevelDefinition[] {
  const instance = numberingDefinitions.instances.find(
    (item) => item.numId === numId
  );
  if (!instance) {
    return [];
  }

  const abstract = numberingDefinitions.abstracts.find(
    (item) => item.abstractNumId === instance.abstractNumId
  );
  return [...(abstract?.levels ?? []), ...(instance.levelOverrides ?? [])];
}

export function effectiveNumberingNumIdForParagraph(
  paragraph: ParagraphNode,
  numberingDefinitions: NumberingDefinitionSet | undefined
): number | undefined {
  const numbering = paragraph.style?.numbering;
  if (
    !numberingDefinitions ||
    !numbering ||
    !Number.isFinite(numbering.numId) ||
    numbering.numId <= 0
  ) {
    return numbering?.numId;
  }

  const numId = numbering.numId;
  const ilvl = Math.max(0, Math.round(numbering.ilvl ?? 0));
  const currentLevel = findNumberingLevelDefinition(
    numberingDefinitions,
    numId,
    ilvl
  );
  if (
    !numberingLevelIsBulletLike(currentLevel) ||
    numberingLevelHasVisibleMarker(currentLevel)
  ) {
    return numId;
  }

  const currentAbstractLevels = numberingAbstractLevelsForNumId(
    numberingDefinitions,
    numId
  );
  if (
    currentAbstractLevels.length === 0 ||
    currentAbstractLevels.some(
      (level) =>
        !numberingLevelIsBulletLike(level) ||
        numberingLevelHasVisibleMarker(level)
    )
  ) {
    return numId;
  }

  const paragraphTextValue = paragraphText(paragraph).trim();
  if (
    paragraphTextValue.length === 0 ||
    isBulletLikeNumberingText(paragraphTextValue) ||
    /^[\u2022\u25cf\u25cb\u25a0\u25a1\u25aa\u25ab\u25c6\u25c7\u25e6\u2043]/.test(
      paragraphTextValue
    )
  ) {
    return numId;
  }

  const candidate = numberingDefinitions.instances.find((instance) => {
    if (instance.numId === numId) {
      return false;
    }
    const candidateLevel = findNumberingLevelDefinition(
      numberingDefinitions,
      instance.numId,
      ilvl
    );
    if (!candidateLevel) {
      return false;
    }
    const format = candidateLevel.format?.trim().toLowerCase();
    if (
      !format ||
      format === "none" ||
      numberingLevelIsBulletLike(candidateLevel)
    ) {
      return false;
    }
    if (!candidateLevel.text || !/%\d+/.test(candidateLevel.text)) {
      return false;
    }
    if (ilvl > 0) {
      const parentLevel = findNumberingLevelDefinition(
        numberingDefinitions,
        instance.numId,
        ilvl - 1
      );
      if (!parentLevel) {
        return false;
      }
      const parentFormat = parentLevel.format?.trim().toLowerCase();
      if (
        !parentFormat ||
        parentFormat === "none" ||
        numberingLevelIsBulletLike(parentLevel)
      ) {
        return false;
      }
    }
    return true;
  });

  return candidate?.numId ?? numId;
}

export function numberingStartValue(
  numbering: NumberingDefinitionSet,
  numId: number,
  ilvl: number
): number {
  const instance = numbering.instances.find((item) => item.numId === numId);
  const override = instance?.levelStartOverrides?.[String(ilvl)];
  if (Number.isFinite(override)) {
    return Math.max(1, Math.round(override as number));
  }

  const level = findNumberingLevelDefinition(numbering, numId, ilvl);
  if (Number.isFinite(level?.start) && (level?.start as number) > 0) {
    return Math.round(level?.start as number);
  }

  return 1;
}

export function numberingSuffix(level: NumberingLevelDefinition | undefined): string {
  if (level?.suffix === "tab") {
    return "\t";
  }
  if (level?.suffix === "space") {
    return " ";
  }
  return "";
}

const LEGACY_BULLET_GLYPH_FALLBACKS: Record<string, string> = {
  "\uf0a7": "□",
  "\uf0a8": "▪",
  "\uf0b7": "•",
  "\uf0d8": "➢",
};

export function normalizeLegacyBulletGlyphs(
  text: string,
  bulletFontFamily?: string
): string {
  if (!text) {
    return text;
  }

  const normalizedFont = bulletFontFamily?.trim().toLowerCase() ?? "";
  const fontUsesLegacyGlyphs =
    normalizedFont.includes("wingdings") ||
    normalizedFont.includes("symbol") ||
    normalizedFont.includes("courier");
  if (!fontUsesLegacyGlyphs && !/[\uf0a7\uf0a8\uf0b7\uf0d8]/.test(text)) {
    return text;
  }

  return Array.from(text)
    .map((glyph) => {
      const mapped = LEGACY_BULLET_GLYPH_FALLBACKS[glyph];
      if (mapped) {
        return mapped;
      }
      if (glyph === "o" && normalizedFont.includes("courier")) {
        return "◦";
      }
      return glyph;
    })
    .join("");
}

export function buildParagraphNumberingLabels(
  model: DocModel
): Map<string, ParagraphNumberingLabel> {
  const labels = new Map<string, ParagraphNumberingLabel>();
  const numbering = model.metadata.numberingDefinitions;
  if (!numbering) {
    return labels;
  }

  const numberingInstanceByNumId = new Map(
    numbering.instances.map((instance) => [instance.numId, instance])
  );
  const paragraphStyleById = new Map(
    (model.metadata.paragraphStyles ?? []).map((styleDefinition) => [
      styleDefinition.id,
      styleDefinition,
    ])
  );
  const countersByNumId = new Map<number, Array<number | undefined>>();
  const abstractCountersByAbstractNumId = new Map<
    number,
    Array<number | undefined>
  >();

  const registerParagraph = (paragraph: ParagraphNode, key: string): void => {
    const paragraphNumbering = paragraph.style?.numbering;
    if (!paragraphNumbering || paragraphNumbering.numId <= 0) {
      return;
    }

    const numId =
      effectiveNumberingNumIdForParagraph(paragraph, numbering) ??
      paragraphNumbering.numId;
    const ilvl = Math.max(0, paragraphNumbering.ilvl ?? 0);
    const level = findNumberingLevelDefinition(numbering, numId, ilvl);
    const numberingInstance = numberingInstanceByNumId.get(numId);
    const abstractNumId = numberingInstance?.abstractNumId;
    const template = level?.text ?? `%${ilvl + 1}.`;
    if (!template || template.trim().length === 0) {
      return;
    }

    const counters = countersByNumId.get(numId) ?? [];
    const sharedAbstractCounters = Number.isFinite(abstractNumId)
      ? abstractCountersByAbstractNumId.get(abstractNumId as number) ?? []
      : undefined;
    let parentCountersChanged = false;
    for (let index = 0; index < ilvl; index += 1) {
      const abstractCounterValue =
        sharedAbstractCounters && Number.isFinite(sharedAbstractCounters[index])
          ? Math.max(1, Math.round(sharedAbstractCounters[index] as number))
          : undefined;
      if (abstractCounterValue !== undefined) {
        const previousCounter = counters[index];
        if (
          !Number.isFinite(previousCounter) ||
          Math.max(1, Math.round(previousCounter as number)) !==
            abstractCounterValue
        ) {
          parentCountersChanged = true;
        }
        counters[index] = abstractCounterValue;
        continue;
      }
      if (!Number.isFinite(counters[index])) {
        counters[index] = numberingStartValue(numbering, numId, index);
      }
    }

    const currentValue = parentCountersChanged ? undefined : counters[ilvl];
    counters[ilvl] = Number.isFinite(currentValue)
      ? (currentValue as number) + 1
      : numberingStartValue(numbering, numId, ilvl);
    for (let index = ilvl + 1; index < counters.length; index += 1) {
      counters[index] = undefined;
    }
    countersByNumId.set(numId, counters);
    if (sharedAbstractCounters && Number.isFinite(abstractNumId)) {
      for (let index = 0; index <= ilvl; index += 1) {
        sharedAbstractCounters[index] = counters[index];
      }
      for (
        let index = ilvl + 1;
        index < sharedAbstractCounters.length;
        index += 1
      ) {
        sharedAbstractCounters[index] = undefined;
      }
      abstractCountersByAbstractNumId.set(
        abstractNumId as number,
        sharedAbstractCounters
      );
    }

    const label = template.replaceAll(/%(\d+)/g, (_, rawLevel) => {
      const levelIndex = Math.max(0, Number(rawLevel) - 1);
      const levelValue = counters[levelIndex];
      const safeValue = Number.isFinite(levelValue)
        ? Math.max(1, Math.round(levelValue as number))
        : numberingStartValue(numbering, numId, levelIndex);
      const levelFormat = findNumberingLevelDefinition(
        numbering,
        numId,
        levelIndex
      )?.format;
      return formatNumberingCounter(levelFormat, safeValue);
    });

    if (!label || label.trim().length === 0) {
      return;
    }

    const isBullet = (level?.format ?? "").trim().toLowerCase() === "bullet";
    const trailingText = numberingSuffix(level);
    if (isBullet && level?.pictureBullet?.src) {
      labels.set(key, {
        imageSrc: level.pictureBullet.src,
        imageWidthPx: level.pictureBullet.widthPx,
        imageHeightPx: level.pictureBullet.heightPx,
        trailingText,
      });
      return;
    }

    const firstStyledRun = paragraph.children.find(
      (child): child is TextRunNode | FormFieldRunNode =>
        (child.type === "text" || child.type === "form-field") &&
        Boolean(child.style)
    );
    const paragraphStyleRunStyle = paragraph.style?.styleId
      ? paragraphStyleById.get(paragraph.style.styleId)?.runStyle
      : undefined;
    const baseNumberingTextStyle =
      firstStyledRun?.style ?? paragraphStyleRunStyle;
    const numberingTextStyle = level?.runStyle
      ? { ...(baseNumberingTextStyle ?? {}), ...level.runStyle }
      : baseNumberingTextStyle;
    const normalizedLabel = isBullet
      ? normalizeLegacyBulletGlyphs(label, level?.bulletFontFamily)
      : label;
    labels.set(key, {
      text: `${normalizedLabel}${trailingText}`,
      fontFamily: isBullet ? level?.bulletFontFamily : undefined,
      color: isBullet ? level?.bulletColor : undefined,
      style: numberingTextStyle ? { ...numberingTextStyle } : undefined,
    });
  };

  model.nodes.forEach((node, nodeIndex) => {
    if (node.type === "paragraph") {
      registerParagraph(node, `p:${nodeIndex}`);
      return;
    }

    node.rows.forEach((row, rowIndex) => {
      row.cells.forEach((cell, cellIndex) => {
        tableCellParagraphs(cell.nodes).forEach((paragraph, paragraphIndex) => {
          registerParagraph(
            paragraph,
            `t:${nodeIndex}:${rowIndex}:${cellIndex}:${paragraphIndex}`
          );
        });
      });
    });
  });

  return labels;
}


export function isUnorderedListText(text: string): boolean {
  return UNORDERED_LIST_PREFIX_PATTERN.test(text);
}

export function isOrderedListText(text: string): boolean {
  return ORDERED_LIST_PREFIX_PATTERN.test(text);
}

export function isBulletLikeNumberingText(
  text: string,
  bulletFontFamily?: string
): boolean {
  const normalized = normalizeLegacyBulletGlyphs(text, bulletFontFamily).trim();
  if (!normalized || /%[\d]+/.test(normalized)) {
    return false;
  }

  return /^[\u2022\u25cf\u25cb\u25a0\u25a1\u25aa\u25ab\u25c6\u25c7\u25e6\u2043\-–—]+$/.test(
    normalized
  );
}

export function paragraphHasNumbering(paragraph: ParagraphNode): boolean {
  return Boolean(
    paragraph.style?.numbering && paragraph.style.numbering.numId > 0
  );
}

export function paragraphListType(
  paragraph: ParagraphNode,
  numberingDefinitions?: NumberingDefinitionSet
): DocxListType | undefined {
  const numbering = paragraph.style?.numbering;
  if (numbering && numbering.numId > 0) {
    if (numberingDefinitions) {
      const effectiveNumId =
        effectiveNumberingNumIdForParagraph(paragraph, numberingDefinitions) ??
        numbering.numId;
      const level = findNumberingLevelDefinition(
        numberingDefinitions,
        effectiveNumId,
        Math.max(0, numbering.ilvl ?? 0)
      );
      const format = level?.format?.trim().toLowerCase();
      if (
        format === "bullet" ||
        level?.pictureBullet ||
        isBulletLikeNumberingText(level?.text ?? "", level?.bulletFontFamily)
      ) {
        return "unordered";
      }
      if (format && format !== "none") {
        return "ordered";
      }
    }
    // If numbering exists but we cannot resolve a specific list format, treat it as ordered.
    return "ordered";
  }

  const text = paragraphText(paragraph);
  if (isUnorderedListText(text)) {
    return "unordered";
  }
  if (isOrderedListText(text)) {
    return "ordered";
  }
  return undefined;
}

export function paragraphIsList(
  paragraph: ParagraphNode,
  textOverride?: string
): boolean {
  const text = textOverride ?? paragraphText(paragraph);
  return (
    paragraphHasNumbering(paragraph) ||
    isUnorderedListText(text) ||
    isOrderedListText(text)
  );
}

export function stripListPrefix(text: string): string {
  return text.replace(/^\s*(?:•\s+|\d+\.\s+)/, "");
}

export function listPrefixLength(text: string): number {
  return text.match(LIST_PREFIX_PATTERN)?.[0]?.length ?? 0;
}

export function nextOrderedListItemText(
  currentText: string,
  nextItemText: string
): string {
  const match = currentText.match(ORDERED_LIST_PREFIX_CAPTURE_PATTERN);
  const leadingWhitespace = match?.[1] ?? "";
  const currentNumber = Number.parseInt(match?.[2] ?? "", 10);
  const nextNumber = Number.isFinite(currentNumber)
    ? Math.max(1, currentNumber + 1)
    : 1;
  return `${leadingWhitespace}${nextNumber}. ${stripListPrefix(nextItemText)}`;
}

export function textWithListType(text: string, listType: DocxListType): string {
  const normalized = stripListPrefix(text);
  return listType === "unordered" ? `• ${normalized}` : `1. ${normalized}`;
}

export function cloneParagraphBorderStyle(
  border?: ParagraphBorderStyle
): ParagraphBorderStyle | undefined {
  return border ? { ...border } : undefined;
}

export function cloneParagraphBorderSet(
  borders?: ParagraphBorderSet
): ParagraphBorderSet | undefined {
  if (!borders) {
    return undefined;
  }

  return {
    top: cloneParagraphBorderStyle(borders.top),
    right: cloneParagraphBorderStyle(borders.right),
    bottom: cloneParagraphBorderStyle(borders.bottom),
    left: cloneParagraphBorderStyle(borders.left),
    between: cloneParagraphBorderStyle(borders.between),
    bar: cloneParagraphBorderStyle(borders.bar),
  };
}

export function cloneParagraphStyle(
  style?: ParagraphNode["style"]
): ParagraphNode["style"] | undefined {
  if (!style) {
    return undefined;
  }

  return {
    ...style,
    numbering: style.numbering ? { ...style.numbering } : undefined,
    spacing: style.spacing ? { ...style.spacing } : undefined,
    indent: style.indent ? { ...style.indent } : undefined,
    tabStops: style.tabStops
      ? style.tabStops.map((tabStop) => ({ ...tabStop }))
      : undefined,
    borders: cloneParagraphBorderSet(style.borders),
    dropCap: style.dropCap ? { ...style.dropCap } : undefined,
  };
}

export function splitParagraphStyleWithDefaultSpacing(
  style?: ParagraphNode["style"],
  sourceXml?: string
): ParagraphNode["style"] | undefined {
  const clonedStyle = cloneParagraphStyle(style);
  // Keep imported/explicit paragraph styling untouched.
  if (sourceXml || clonedStyle?.styleId) {
    return clonedStyle;
  }

  const spacing = clonedStyle?.spacing;
  const hasExplicitSpacing =
    spacing?.beforeTwips !== undefined ||
    spacing?.afterTwips !== undefined ||
    spacing?.lineTwips !== undefined ||
    spacing?.lineRule !== undefined;
  if (hasExplicitSpacing) {
    return clonedStyle;
  }

  return {
    ...(clonedStyle ?? {}),
    spacing: {
      ...(spacing ?? {}),
      lineRule: "auto",
      lineTwips: DEFAULT_SPLIT_PARAGRAPH_LINE_TWIPS,
      afterTwips: DEFAULT_SPLIT_PARAGRAPH_AFTER_TWIPS,
    },
  };
}

export function cloneTableBoxSpacing(
  spacing?: TableCellStyle["marginTwips"]
): TableCellStyle["marginTwips"] | undefined {
  if (!spacing) {
    return undefined;
  }

  return {
    topTwips: spacing.topTwips,
    rightTwips: spacing.rightTwips,
    bottomTwips: spacing.bottomTwips,
    leftTwips: spacing.leftTwips,
  };
}

export function cloneTableBorderStyle(
  border?: TableBorderStyle
): TableBorderStyle | undefined {
  if (!border) {
    return undefined;
  }

  return {
    type: border.type,
    color: border.color,
    sizeEighthPt: border.sizeEighthPt,
  };
}

export function cloneTableBorderSet(
  borders?: TableBorderSet
): TableBorderSet | undefined {
  if (!borders) {
    return undefined;
  }

  return {
    top: cloneTableBorderStyle(borders.top),
    right: cloneTableBorderStyle(borders.right),
    bottom: cloneTableBorderStyle(borders.bottom),
    left: cloneTableBorderStyle(borders.left),
    insideH: cloneTableBorderStyle(borders.insideH),
    insideV: cloneTableBorderStyle(borders.insideV),
    tl2br: cloneTableBorderStyle(borders.tl2br),
    tr2bl: cloneTableBorderStyle(borders.tr2bl),
  };
}

export function cloneTableCellStyle(
  style?: TableCellStyle
): TableCellStyle | undefined {
  if (!style) {
    return undefined;
  }

  return {
    ...style,
    marginTwips: cloneTableBoxSpacing(style.marginTwips),
    borders: cloneTableBorderSet(style.borders),
  };
}

export function cloneTableRowStyle(style?: TableRowStyle): TableRowStyle | undefined {
  if (!style) {
    return undefined;
  }

  return {
    ...style,
  };
}

export function createEmptyParagraphFromTemplate(
  template?: ParagraphNode
): ParagraphNode {
  const nextTextStyle = cloneTextStyle(firstRunStyle(template));

  return {
    type: "paragraph",
    style: cloneParagraphStyle(template?.style),
    children: [
      {
        type: "text",
        text: "",
        style: nextTextStyle,
      },
    ],
  };
}

export function createEmptyTableCellFromTemplate(
  templateCell?: TableNode["rows"][number]["cells"][number]
): TableNode["rows"][number]["cells"][number] {
  const templateParagraph = templateCell
    ? tableCellParagraphs(templateCell.nodes)[0]
    : undefined;
  const clonedStyle = cloneTableCellStyle(templateCell?.style);
  const normalizedStyle = clonedStyle
    ? {
        ...clonedStyle,
        rowSpan: undefined,
        vMergeContinuation: undefined,
      }
    : undefined;

  return {
    type: "table-cell",
    style: normalizedStyle,
    nodes: [createEmptyParagraphFromTemplate(templateParagraph)],
  };
}

export function resolveMaxNumberingLevel(
  numberingDefinitions: NumberingDefinitionSet | undefined,
  numId: number
): number {
  if (!numberingDefinitions) {
    return MAX_FALLBACK_LIST_LEVEL;
  }

  const instance = numberingDefinitions.instances.find(
    (item) => item.numId === numId
  );
  if (!instance) {
    return MAX_FALLBACK_LIST_LEVEL;
  }

  const abstract = numberingDefinitions.abstracts.find(
    (item) => item.abstractNumId === instance.abstractNumId
  );
  const levels = [
    ...(abstract?.levels ?? []),
    ...(instance.levelOverrides ?? []),
  ];
  if (levels.length === 0) {
    return MAX_FALLBACK_LIST_LEVEL;
  }

  return clampNumber(
    Math.max(
      0,
      ...levels.map((level) => Math.max(0, Math.round(level.ilvl ?? 0)))
    ),
    0,
    MAX_FALLBACK_LIST_LEVEL
  );
}

export function shiftListIndent(
  indent: ParagraphIndent | undefined,
  levelDelta: number
): ParagraphIndent | undefined {
  if (!levelDelta) {
    return indent;
  }

  const nextLeftTwips = Math.max(
    0,
    (indent?.leftTwips ?? 0) + levelDelta * LIST_LEVEL_STEP_TWIPS
  );
  const nextHangingTwips =
    indent?.hangingTwips ??
    (nextLeftTwips > 0 ? DEFAULT_LIST_HANGING_TWIPS : undefined);

  return {
    ...(indent ?? {}),
    leftTwips: nextLeftTwips,
    hangingTwips: nextHangingTwips,
  };
}

export function ensurePrefixListIndent(paragraph: ParagraphNode): boolean {
  const clonedStyle = cloneParagraphStyle(paragraph.style) ?? {};
  const existingIndent = clonedStyle.indent;
  const hasMeaningfulLeftIndent =
    Number.isFinite(existingIndent?.leftTwips) &&
    Math.abs(existingIndent?.leftTwips ?? 0) > 0;
  const hasMeaningfulFirstLine =
    Number.isFinite(existingIndent?.firstLineTwips) &&
    Math.abs(existingIndent?.firstLineTwips ?? 0) > 0;
  const hasMeaningfulHanging =
    Number.isFinite(existingIndent?.hangingTwips) &&
    Math.abs(existingIndent?.hangingTwips ?? 0) > 0;

  if (
    hasMeaningfulLeftIndent ||
    hasMeaningfulFirstLine ||
    hasMeaningfulHanging
  ) {
    return false;
  }

  paragraph.style = {
    ...clonedStyle,
    indent: {
      ...(existingIndent ?? {}),
      leftTwips: LIST_LEVEL_STEP_TWIPS,
      hangingTwips: DEFAULT_LIST_HANGING_TWIPS,
    },
  };
  paragraph.sourceXml = undefined;
  return true;
}

export function clearAutoPrefixListIndent(paragraph: ParagraphNode): boolean {
  const clonedStyle = cloneParagraphStyle(paragraph.style);
  const existingIndent = clonedStyle?.indent;
  if (!clonedStyle || !existingIndent) {
    return false;
  }

  const leftTwips = existingIndent.leftTwips;
  const hangingTwips = existingIndent.hangingTwips;
  const firstLineTwips = existingIndent.firstLineTwips;
  const matchesAutoIndent =
    Number.isFinite(leftTwips) &&
    Math.round(leftTwips ?? 0) === LIST_LEVEL_STEP_TWIPS &&
    Number.isFinite(hangingTwips) &&
    Math.round(hangingTwips ?? 0) === DEFAULT_LIST_HANGING_TWIPS &&
    !Number.isFinite(firstLineTwips);
  if (!matchesAutoIndent) {
    return false;
  }

  const nextIndent: ParagraphIndent = {
    ...existingIndent,
    leftTwips: undefined,
    hangingTwips: undefined,
  };
  const hasRemainingIndent = [
    nextIndent.leftTwips,
    nextIndent.rightTwips,
    nextIndent.firstLineTwips,
    nextIndent.hangingTwips,
  ].some((value) => Number.isFinite(value));

  paragraph.style = {
    ...clonedStyle,
    indent: hasRemainingIndent ? nextIndent : undefined,
  };
  paragraph.sourceXml = undefined;
  return true;
}

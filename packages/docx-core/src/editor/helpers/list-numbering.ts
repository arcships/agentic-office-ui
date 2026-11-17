// List numbering definition resolution for editor list commands.
//
// List commands must not assume that numId 1 is a bullet list and numId 2 is a
// decimal list: imported documents are free to define any numbering (or none).
// These helpers find a matching definition, or create one when the document
// has no suitable numbering definition yet.

import type {
  DocModel,
  NumberingDefinitionSet,
  NumberingLevelDefinition,
} from "../../engine/types";
import type { DocxListType } from "./editor-types";
import {
  DEFAULT_LIST_HANGING_TWIPS,
  LIST_LEVEL_STEP_TWIPS,
  MAX_FALLBACK_LIST_LEVEL,
} from "./constants";
import {
  findNumberingLevelDefinition,
  isBulletLikeNumberingText,
} from "./numbering";

const BULLET_GLYPHS = ["•", "◦", "▪"];

function levelIsBullet(level: NumberingLevelDefinition | undefined): boolean {
  if (!level) {
    return false;
  }
  return (
    level.format?.trim().toLowerCase() === "bullet" ||
    Boolean(level.pictureBullet) ||
    isBulletLikeNumberingText(level.text ?? "", level.bulletFontFamily)
  );
}

function levelMatchesListType(
  level: NumberingLevelDefinition | undefined,
  listType: DocxListType
): boolean {
  if (!level) {
    return false;
  }
  if (listType === "unordered") {
    return levelIsBullet(level);
  }
  const format = level.format?.trim().toLowerCase();
  return Boolean(format && format !== "none" && !levelIsBullet(level));
}

export function resolveListNumberingId(
  definitions: NumberingDefinitionSet | undefined,
  listType: DocxListType
): number | undefined {
  if (!definitions) {
    return undefined;
  }

  const instanceIds = definitions.instances
    .map((instance) => instance.numId)
    .filter((numId) => Number.isFinite(numId) && numId > 0)
    .sort((left, right) => left - right);

  for (const numId of instanceIds) {
    const level = findNumberingLevelDefinition(definitions, numId, 0);
    if (levelMatchesListType(level, listType)) {
      return numId;
    }
  }

  return undefined;
}

function nextNumericId(values: number[]): number {
  return Math.max(0, ...values.filter((value) => Number.isFinite(value))) + 1;
}

function createListLevels(listType: DocxListType): NumberingLevelDefinition[] {
  return Array.from({ length: MAX_FALLBACK_LIST_LEVEL + 1 }, (_, ilvl) => ({
    ilvl,
    start: 1,
    format: listType === "unordered" ? "bullet" : "decimal",
    text:
      listType === "unordered"
        ? BULLET_GLYPHS[ilvl % BULLET_GLYPHS.length]
        : `%${ilvl + 1}.`,
    suffix: "space" as const,
    indent: {
      leftTwips: (ilvl + 1) * LIST_LEVEL_STEP_TWIPS,
      hangingTwips: DEFAULT_LIST_HANGING_TWIPS,
    },
  }));
}

export function ensureListNumberingId(
  model: DocModel,
  listType: DocxListType
): number {
  const existing = resolveListNumberingId(
    model.metadata.numberingDefinitions,
    listType
  );
  if (existing !== undefined) {
    return existing;
  }

  const definitions: NumberingDefinitionSet =
    model.metadata.numberingDefinitions ?? {
      abstracts: [],
      instances: [],
    };
  const abstractNumId = nextNumericId(
    definitions.abstracts.map((abstract) => abstract.abstractNumId)
  );
  const numId = nextNumericId(
    definitions.instances.map((instance) => instance.numId)
  );

  definitions.abstracts.push({
    abstractNumId,
    levels: createListLevels(listType),
  });
  definitions.instances.push({ numId, abstractNumId });
  model.metadata.numberingDefinitions = definitions;
  return numId;
}

// Post-import typography enrichment.
//
// The WASM model builder intentionally keeps the model compact: it flattens
// many `w:rPr` properties (underline/strike variants, caps, run borders,
// per-script fonts, RTL) and drops `w:lvlJc` / `w:suff` from numbering levels.
// The original paragraph XML is still available on imported paragraphs, so we
// recover that vocabulary once, after the model has been built, without
// re-parsing the whole document in JS.

import type {
  DocModel,
  DocNode,
  DocxThemeFonts,
  NumberingDefinitionSet,
  NumberingLevelDefinition,
  ParagraphAlignment,
  ParagraphDirection,
  ParagraphNode,
  TableNode,
  TextRunNode,
  TextStyle
} from "./types";
import type { OoxmlPackage } from "./ooxml-core";
import {
  decodeXmlText,
  extractBalancedTagRanges,
  parseParagraphAlignmentFromParagraphProperties,
  parseParagraphDirectionFromParagraphProperties,
  parseRunTypographyFromRunXml,
  parseThemeFontsFromXml,
  xmlAttribute
} from "./text-typography";

function packagePartContent(
  pkg: OoxmlPackage,
  partName: string
): string | undefined {
  return pkg.parts.get(partName)?.content;
}

function tagBlocks(xml: string, tagName: string): string[] {
  return extractBalancedTagRanges(xml, tagName).map((range) =>
    xml.slice(range.start, range.end)
  );
}

function firstTag(xml: string, tagName: string): string | undefined {
  const range = extractBalancedTagRanges(xml, tagName)[0];
  return range ? xml.slice(range.start, range.end) : undefined;
}

function tagValue(tagXml: string | undefined, attribute = "w:val"): string | undefined {
  return tagXml ? xmlAttribute(tagXml, attribute)?.trim() : undefined;
}

function applyMissingRunStyle(
  run: TextRunNode,
  parsed: TextStyle | undefined
): void {
  if (!parsed) {
    return;
  }

  const current = run.style ?? {};
  const next: TextStyle = { ...current };
  const assign = <K extends keyof TextStyle>(key: K): void => {
    if (current[key] === undefined && parsed[key] !== undefined) {
      next[key] = parsed[key];
    }
  };

  assign("fontFamily");
  assign("fontFamilyEastAsia");
  assign("fontFamilyComplexScript");
  assign("fontSizePt");
  assign("bold");
  assign("italic");
  assign("color");
  assign("highlight");
  assign("caps");
  assign("smallCaps");
  assign("characterSpacingTwips");
  assign("characterScalePercent");
  assign("direction");
  assign("verticalAlign");
  assign("runBorder");
  assign("textShadow");
  assign("textOutline");

  if (current.underline === undefined && parsed.underline !== undefined) {
    next.underline = parsed.underline;
  }
  if (current.underlineStyle === undefined && parsed.underlineStyle !== undefined) {
    next.underlineStyle = parsed.underlineStyle;
  }
  if (current.strike === undefined && parsed.strike !== undefined) {
    next.strike = parsed.strike;
  }
  if (current.strikeStyle === undefined && parsed.strikeStyle !== undefined) {
    next.strikeStyle = parsed.strikeStyle;
  }

  if (Object.keys(next).length > 0) {
    run.style = next;
  }
}

function runTextFromXml(runXml: string): string {
  const pattern =
    /<w:(t|delText)\b[^>]*>([\s\S]*?)<\/w:\1>|<w:tab\b[^>]*\/?>|<w:(?:br|cr)\b[^>]*\/?>|<w:(?:footnoteReference|endnoteReference)\b[^>]*\/?>/gi;
  let text = "";

  for (const match of runXml.matchAll(pattern)) {
    if (match[2] !== undefined) {
      text += decodeXmlText(match[2]);
      continue;
    }
    const tagXml = match[0] ?? "";
    if (/^<w:tab\b/i.test(tagXml)) {
      text += "\t";
    } else if (/^<w:(?:br|cr)\b/i.test(tagXml)) {
      text += "\n";
    }
  }

  return text;
}

const RUN_TYPOGRAPHY_TAG_PATTERN =
  /<w:(?:rFonts|u|dstrike|strike|caps|smallCaps|spacing|w|bdr|effect|shadow|outline|rtl)\b/i;

function enrichParagraphRuns(
  paragraph: ParagraphNode,
  themeFonts: DocxThemeFonts | undefined
): void {
  const sourceXml = paragraph.sourceXml;
  if (!sourceXml || !RUN_TYPOGRAPHY_TAG_PATTERN.test(sourceXml)) {
    return;
  }

  const textRuns = paragraph.children.filter(
    (child): child is TextRunNode => child.type === "text"
  );
  if (textRuns.length === 0) {
    return;
  }

  const runXmls = tagBlocks(sourceXml, "w:r");
  if (runXmls.length !== textRuns.length) {
    return;
  }

  // Only enrich when the XML run sequence and the model text sequence agree
  // exactly.  Tracked-change wrappers, text boxes and field instructions can
  // make the two sequences differ; in that case the conservative behaviour is
  // to leave the model untouched.
  for (let index = 0; index < runXmls.length; index += 1) {
    if (runTextFromXml(runXmls[index]) !== textRuns[index].text) {
      return;
    }
  }

  for (let index = 0; index < runXmls.length; index += 1) {
    applyMissingRunStyle(
      textRuns[index],
      parseRunTypographyFromRunXml(runXmls[index], { themeFonts })
    );
  }
}

function enrichParagraph(
  paragraph: ParagraphNode,
  themeFonts: DocxThemeFonts | undefined
): void {
  const sourceXml = paragraph.sourceXml;
  if (!sourceXml) {
    return;
  }

  const pPrXml = firstTag(sourceXml, "w:pPr");
  if (pPrXml) {
    const style = paragraph.style ?? {};
    let changed = false;
    if (style.align === undefined) {
      const align = parseParagraphAlignmentFromParagraphProperties(pPrXml);
      if (align) {
        style.align = align;
        changed = true;
      }
    }
    if (style.direction === undefined) {
      const direction = parseParagraphDirectionFromParagraphProperties(pPrXml);
      if (direction) {
        style.direction = direction;
        changed = true;
      }
    }
    if (changed) {
      paragraph.style = style;
    }
  }

  enrichParagraphRuns(paragraph, themeFonts);
}

function walkNodes(
  nodes: DocNode[],
  themeFonts: DocxThemeFonts | undefined
): void {
  for (const node of nodes) {
    if (node.type === "paragraph") {
      enrichParagraph(node, themeFonts);
      continue;
    }
    walkTable(node, themeFonts);
  }
}

function walkTable(table: TableNode, themeFonts: DocxThemeFonts | undefined): void {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      for (const node of cell.nodes) {
        if (node.type === "paragraph") {
          enrichParagraph(node, themeFonts);
        } else {
          walkTable(node, themeFonts);
        }
      }
    }
  }
}

function ensureNumberingLevel(
  levels: NumberingLevelDefinition[],
  ilvl: number
): NumberingLevelDefinition {
  let level = levels.find((candidate) => candidate.ilvl === ilvl);
  if (!level) {
    level = { ilvl };
    levels.push(level);
  }
  return level;
}

function applyNumberingLevelXml(
  level: NumberingLevelDefinition,
  levelXml: string
): void {
  if (level.suffix === undefined) {
    const suffix = tagValue(firstTag(levelXml, "w:suff"));
    if (suffix === "tab" || suffix === "space" || suffix === "nothing") {
      level.suffix = suffix;
    }
  }

  if (level.justification === undefined) {
    const raw = tagValue(firstTag(levelXml, "w:lvlJc"))?.toLowerCase();
    const alignment: ParagraphAlignment | undefined =
      raw === "left" ||
      raw === "center" ||
      raw === "right" ||
      raw === "start" ||
      raw === "end"
        ? raw
        : undefined;
    if (alignment) {
      level.justification = alignment;
    }
  }

  if (level.start === undefined) {
    const start = Number(tagValue(firstTag(levelXml, "w:start")));
    if (Number.isFinite(start) && start > 0) {
      level.start = Math.round(start);
    }
  }

  if (!level.format) {
    const format = tagValue(firstTag(levelXml, "w:numFmt"));
    if (format) {
      level.format = format;
    }
  }

  if (!level.text) {
    const text = tagValue(firstTag(levelXml, "w:lvlText"));
    if (text) {
      level.text = text;
    }
  }
}

/**
 * Fill numbering metadata that the WASM builder omits (`w:lvlJc`, `w:suff`,
 * level start/format fallbacks).  Existing model values win.
 */
export function enrichNumberingDefinitionsFromXml(
  definitions: NumberingDefinitionSet | undefined,
  numberingXml: string | undefined
): void {
  if (!definitions || !numberingXml) {
    return;
  }

  const abstractById = new Map(
    definitions.abstracts.map((abstract) => [abstract.abstractNumId, abstract])
  );
  for (const abstractXml of tagBlocks(numberingXml, "w:abstractNum")) {
    const abstractId = Number(
      xmlAttribute(abstractXml.slice(0, abstractXml.indexOf(">") + 1), "w:abstractNumId")
    );
    const abstract = abstractById.get(abstractId);
    if (!abstract) {
      continue;
    }

    for (const levelXml of tagBlocks(abstractXml, "w:lvl")) {
      const ilvl = Number(
        xmlAttribute(levelXml.slice(0, levelXml.indexOf(">") + 1), "w:ilvl")
      );
      if (!Number.isFinite(ilvl)) {
        continue;
      }
      applyNumberingLevelXml(ensureNumberingLevel(abstract.levels, ilvl), levelXml);
    }
  }

  const instanceById = new Map(
    definitions.instances.map((instance) => [instance.numId, instance])
  );
  for (const numXml of tagBlocks(numberingXml, "w:num")) {
    const numId = Number(
      xmlAttribute(numXml.slice(0, numXml.indexOf(">") + 1), "w:numId")
    );
    const instance = instanceById.get(numId);
    if (!instance) {
      continue;
    }

    const overrides = instance.levelOverrides ?? (instance.levelOverrides = []);
    for (const overrideXml of tagBlocks(numXml, "w:lvlOverride")) {
      const ilvl = Number(
        xmlAttribute(
          overrideXml.slice(0, overrideXml.indexOf(">") + 1),
          "w:ilvl"
        )
      );
      if (!Number.isFinite(ilvl)) {
        continue;
      }
      const startOverride = Number(
        tagValue(firstTag(overrideXml, "w:startOverride"))
      );
      if (Number.isFinite(startOverride) && startOverride > 0) {
        const startOverrides =
          instance.levelStartOverrides ??
          (instance.levelStartOverrides = {});
        if (startOverrides[String(ilvl)] === undefined) {
          startOverrides[String(ilvl)] = Math.round(startOverride);
        }
      }
      const levelXml = firstTag(overrideXml, "w:lvl");
      if (levelXml) {
        applyNumberingLevelXml(ensureNumberingLevel(overrides, ilvl), levelXml);
      }
    }
  }
}

export function enrichDocModelTypography(
  model: DocModel,
  pkg: OoxmlPackage
): void {
  const themeFonts = parseThemeFontsFromXml(
    packagePartContent(pkg, "word/theme/theme1.xml")
  );
  if (themeFonts) {
    model.metadata.themeFonts = themeFonts;
  }

  enrichNumberingDefinitionsFromXml(
    model.metadata.numberingDefinitions,
    packagePartContent(pkg, "word/numbering.xml")
  );

  walkNodes(model.nodes, themeFonts);
}

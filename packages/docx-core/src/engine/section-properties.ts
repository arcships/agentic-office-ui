// Page-setup properties inside a section's `<w:sectPr>`.
//
// A section's page geometry lives in two child elements: `<w:pgSz>` (paper
// width/height in twips plus `@w:orient`) and `<w:pgMar>` (margins and the
// header/footer/gutter distances). This module reads those values and rewrites
// them in place, leaving every other `sectPr` child (header/footer references,
// `w:type`, columns, borders, page numbering) untouched so an edit cannot drop
// a section's references.

export type DocxPageOrientation = "portrait" | "landscape";

/** Page margins in twips (1/1440 inch), matching OOXML `w:pgMar`. */
export interface DocxPageMargins {
  topTwips: number;
  rightTwips: number;
  bottomTwips: number;
  leftTwips: number;
  headerTwips: number;
  footerTwips: number;
  gutterTwips: number;
}

/** Margin patch: omitted fields keep their current value. */
export type DocxPageMarginsUpdate = Partial<DocxPageMargins>;

/** Paper size in twips, matching OOXML `w:pgSz`. */
export interface DocxPageSize {
  widthTwips: number;
  heightTwips: number;
  orientation: DocxPageOrientation;
}

/** Paper-size patch: omitted fields keep their current value. */
export type DocxPageSizeUpdate = Partial<DocxPageSize>;

export type DocxPageSizePreset =
  | "letter"
  | "legal"
  | "tabloid"
  | "a3"
  | "a4"
  | "a5"
  | "b5";

/** Body column layout of a section (`<w:cols>`). */
export interface DocxSectionColumns {
  /** Text columns on the section's pages, at least 1. */
  count: number;
  /** Space between two adjacent columns in twips (OOXML `w:space`). */
  spaceTwips: number;
  /** Explicit per-column widths in twips, when the section declares them. */
  widthsTwips?: number[];
}

/** Column patch: omitted fields keep their current value. */
export interface DocxSectionColumnsUpdate {
  /** Number of columns; 1 keeps the text in one body column. */
  count?: number;
  /** Space between columns in twips (OOXML `w:space`). */
  spaceTwips?: number;
}

/**
 * Common paper sizes in portrait twips. Values match Word's own `w:pgSz`
 * output for each preset.
 */
export const DOCX_PAGE_SIZE_PRESETS: Readonly<
  Record<DocxPageSizePreset, { widthTwips: number; heightTwips: number }>
> = Object.freeze({
  letter: { widthTwips: 12240, heightTwips: 15840 },
  legal: { widthTwips: 12240, heightTwips: 20160 },
  tabloid: { widthTwips: 15840, heightTwips: 24480 },
  a3: { widthTwips: 16838, heightTwips: 23811 },
  a4: { widthTwips: 11906, heightTwips: 16838 },
  a5: { widthTwips: 8391, heightTwips: 11906 },
  b5: { widthTwips: 9979, heightTwips: 14175 },
});

export const TWIPS_PER_INCH = 1440;

/** Word's default spacing between section columns: half an inch. */
export const DEFAULT_DOCX_COLUMN_SPACE_TWIPS = 720;

/** Word's default: US Letter with 1in margins and 0.5in header/footer. */
export const DEFAULT_DOCX_PAGE_MARGINS: DocxPageMargins = Object.freeze({
  topTwips: 1440,
  rightTwips: 1440,
  bottomTwips: 1440,
  leftTwips: 1440,
  headerTwips: 708,
  footerTwips: 708,
  gutterTwips: 0,
});

export const DEFAULT_DOCX_PAGE_SIZE: DocxPageSize = Object.freeze({
  widthTwips: 12240,
  heightTwips: 15840,
  orientation: "portrait" as DocxPageOrientation,
});

/**
 * Section properties used when a document carries no `w:sectPr` at all. Same
 * fallback the section-break commands use.
 */
export const DEFAULT_SECTION_PROPERTIES_XML =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

/**
 * Child order of CT_SectPr (ECMA-376 §17.6.17). Word repairs documents whose
 * children are out of order, so new elements are always inserted at their
 * schema position.
 */
const SECTION_CHILD_ORDER = [
  "w:headerReference",
  "w:footerReference",
  "w:footnotePr",
  "w:endnotePr",
  "w:type",
  "w:pgSz",
  "w:pgMar",
  "w:paperSrc",
  "w:pgBorders",
  "w:lnNumType",
  "w:pgNumType",
  "w:cols",
  "w:formProt",
  "w:vAlign",
  "w:noEndnote",
  "w:titlePg",
  "w:textDirection",
  "w:bidi",
  "w:rtlGutter",
  "w:docGrid",
  "w:printerSettings",
  "w:sectPrChange",
] as const;

const ATTRIBUTE_PATTERN = /([A-Za-z_][\w:.-]*)\s*=\s*"([^"]*)"/g;

function normalizeTwips(value: unknown, fallback: number): number {
  const numeric = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(numeric)) return fallback;
  return Math.max(0, Math.round(numeric));
}

function readElement(sectionPropertiesXml: string, tagName: string): string | undefined {
  const pattern = new RegExp(
    `<${tagName}\\b[^>]*?(?:/>|>[\\s\\S]*?</${tagName}>)`,
    "i",
  );
  return sectionPropertiesXml.match(pattern)?.[0];
}

function readAttributes(tagXml: string | undefined): Array<[string, string]> {
  if (!tagXml) return [];
  const attributes: Array<[string, string]> = [];
  ATTRIBUTE_PATTERN.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = ATTRIBUTE_PATTERN.exec(tagXml)) !== null) {
    attributes.push([match[1], match[2]]);
  }
  return attributes;
}

function readAttribute(tagXml: string | undefined, attribute: string): string | undefined {
  return readAttributes(tagXml).find(([name]) => name === attribute)?.[1];
}

function renderElement(
  tagName: string,
  attributes: Array<[string, string]>,
): string {
  const rendered = attributes
    .map(([name, value]) => `${name}="${value}"`)
    .join(" ");
  return rendered ? `<${tagName} ${rendered}/>` : `<${tagName}/>`;
}

/**
 * Merge `values` into `attributes` in place: existing names keep their
 * position (and are overwritten), new names are appended, and an explicit
 * `undefined` removes the attribute.
 */
function mergeAttributes(
  attributes: Array<[string, string]>,
  values: Record<string, string | number | undefined>,
): Array<[string, string]> {
  const merged = [...attributes];
  for (const [name, rawValue] of Object.entries(values)) {
    const index = merged.findIndex(([existing]) => existing === name);
    if (rawValue === undefined || rawValue === null) {
      if (index >= 0) merged.splice(index, 1);
      continue;
    }
    const value = String(rawValue);
    if (index >= 0) merged[index] = [name, value];
    else merged.push([name, value]);
  }
  return merged;
}

function sectionChildInsertIndex(sectionPropertiesXml: string, tagName: string): number {
  const order = SECTION_CHILD_ORDER.indexOf(
    tagName as (typeof SECTION_CHILD_ORDER)[number],
  );
  if (order >= 0) {
    for (const laterTag of SECTION_CHILD_ORDER.slice(order + 1)) {
      const match = sectionPropertiesXml.match(new RegExp(`<${laterTag}\\b`, "i"));
      if (match?.index !== undefined) return match.index;
    }
  }
  const closeTag = sectionPropertiesXml.search(/<\/w:sectPr>\s*$/i);
  if (closeTag >= 0) return closeTag;
  return sectionPropertiesXml.length;
}

/**
 * Insert or update one child element of a `w:sectPr` fragment.
 *
 * The fragment is accepted as-is (including a self-closing `<w:sectPr/>`);
 * everything outside the touched element is preserved byte for byte.
 */
export function upsertSectionChildElement(
  sectionPropertiesXml: string,
  tagName: string,
  values: Record<string, string | number | undefined>,
): string {
  const xml = /^<w:sectPr\b/i.test(sectionPropertiesXml.trim())
    ? sectionPropertiesXml
    : `<w:sectPr>${sectionPropertiesXml}</w:sectPr>`;

  const existing = readElement(xml, tagName);
  if (existing) {
    const merged = renderElement(tagName, mergeAttributes(readAttributes(existing), values));
    return xml.replace(existing, merged);
  }

  // A self-closing sectPr has no children yet: open it, append, close it.
  const selfClosing = xml.match(/<w:sectPr\b[^>]*\/>/i);
  if (selfClosing) {
    const openTag = selfClosing[0].replace(/\/>\s*$/, ">");
    const children = SECTION_CHILD_ORDER.filter((child) => child === tagName);
    return xml.replace(
      selfClosing[0],
      `${openTag}${children
        .map((child) => renderElement(child, mergeAttributes([], values)))
        .join("")}</w:sectPr>`,
    );
  }

  const insertAt = sectionChildInsertIndex(xml, tagName);
  return (
    xml.slice(0, insertAt) +
    renderElement(tagName, mergeAttributes([], values)) +
    xml.slice(insertAt)
  );
}

/** Read `w:pgMar` with Word's defaults for every missing attribute. */
export function readSectionPageMargins(
  sectionPropertiesXml?: string,
): DocxPageMargins {
  const tag = sectionPropertiesXml
    ? readElement(sectionPropertiesXml, "w:pgMar")
    : undefined;
  return {
    topTwips: normalizeTwips(
      readAttribute(tag, "w:top"),
      DEFAULT_DOCX_PAGE_MARGINS.topTwips,
    ),
    rightTwips: normalizeTwips(
      readAttribute(tag, "w:right"),
      DEFAULT_DOCX_PAGE_MARGINS.rightTwips,
    ),
    bottomTwips: normalizeTwips(
      readAttribute(tag, "w:bottom"),
      DEFAULT_DOCX_PAGE_MARGINS.bottomTwips,
    ),
    leftTwips: normalizeTwips(
      readAttribute(tag, "w:left"),
      DEFAULT_DOCX_PAGE_MARGINS.leftTwips,
    ),
    headerTwips: normalizeTwips(
      readAttribute(tag, "w:header"),
      DEFAULT_DOCX_PAGE_MARGINS.headerTwips,
    ),
    footerTwips: normalizeTwips(
      readAttribute(tag, "w:footer"),
      DEFAULT_DOCX_PAGE_MARGINS.footerTwips,
    ),
    gutterTwips: normalizeTwips(
      readAttribute(tag, "w:gutter"),
      DEFAULT_DOCX_PAGE_MARGINS.gutterTwips,
    ),
  };
}

/**
 * Swap a page size whose dimensions contradict its orientation, the way Word
 * stores landscape: `w:w` is the long edge and `w:h` the short one.
 */
export function normalizePageSizeOrientation(pageSize: DocxPageSize): DocxPageSize {
  const widthTwips = normalizeTwips(pageSize.widthTwips, DEFAULT_DOCX_PAGE_SIZE.widthTwips);
  const heightTwips = normalizeTwips(pageSize.heightTwips, DEFAULT_DOCX_PAGE_SIZE.heightTwips);
  const orientation: DocxPageOrientation =
    pageSize.orientation === "landscape" ? "landscape" : "portrait";

  if (orientation === "landscape" && widthTwips < heightTwips) {
    return { widthTwips: heightTwips, heightTwips: widthTwips, orientation };
  }
  if (orientation === "portrait" && widthTwips > heightTwips) {
    return { widthTwips: heightTwips, heightTwips: widthTwips, orientation };
  }
  return { widthTwips, heightTwips, orientation };
}

/** Read `w:pgSz` with orientation resolved from `w:orient` and the edges. */
export function readSectionPageSize(sectionPropertiesXml?: string): DocxPageSize {
  const tag = sectionPropertiesXml ? readElement(sectionPropertiesXml, "w:pgSz") : undefined;
  const widthTwips = normalizeTwips(
    readAttribute(tag, "w:w"),
    DEFAULT_DOCX_PAGE_SIZE.widthTwips,
  );
  const heightTwips = normalizeTwips(
    readAttribute(tag, "w:h"),
    DEFAULT_DOCX_PAGE_SIZE.heightTwips,
  );
  const explicit = readAttribute(tag, "w:orient")?.trim().toLowerCase();
  const orientation: DocxPageOrientation =
    explicit === "landscape"
      ? "landscape"
      : explicit === "portrait"
        ? "portrait"
        : widthTwips > heightTwips
          ? "landscape"
          : "portrait";

  return normalizePageSizeOrientation({ widthTwips, heightTwips, orientation });
}

/** Resolve a preset name to a page size in the requested orientation. */
export function resolveDocxPageSizePreset(
  preset: DocxPageSizePreset,
  orientation: DocxPageOrientation = "portrait",
): DocxPageSize {
  const size = DOCX_PAGE_SIZE_PRESETS[preset] ?? DOCX_PAGE_SIZE_PRESETS.letter;
  return normalizePageSizeOrientation({
    widthTwips: size.widthTwips,
    heightTwips: size.heightTwips,
    orientation,
  });
}

/**
 * Match a page size against the known presets (either orientation). Returns
 * `undefined` for custom sizes.
 */
export function matchDocxPageSizePreset(
  pageSize: DocxPageSize,
): DocxPageSizePreset | undefined {
  const normalized = normalizePageSizeOrientation(pageSize);
  const shortEdge = Math.min(normalized.widthTwips, normalized.heightTwips);
  const longEdge = Math.max(normalized.widthTwips, normalized.heightTwips);
  return (Object.keys(DOCX_PAGE_SIZE_PRESETS) as DocxPageSizePreset[]).find((preset) => {
    const candidate = DOCX_PAGE_SIZE_PRESETS[preset];
    return (
      Math.min(candidate.widthTwips, candidate.heightTwips) === shortEdge &&
      Math.max(candidate.widthTwips, candidate.heightTwips) === longEdge
    );
  });
}

/** Rewrite `<w:pgMar>` in a `w:sectPr` fragment. */
export function writeSectionPageMargins(
  sectionPropertiesXml: string,
  margins: DocxPageMarginsUpdate,
): string {
  const current = readSectionPageMargins(sectionPropertiesXml);
  const merged: DocxPageMargins = {
    topTwips: margins.topTwips === undefined ? current.topTwips : margins.topTwips,
    rightTwips: margins.rightTwips === undefined ? current.rightTwips : margins.rightTwips,
    bottomTwips: margins.bottomTwips === undefined ? current.bottomTwips : margins.bottomTwips,
    leftTwips: margins.leftTwips === undefined ? current.leftTwips : margins.leftTwips,
    headerTwips: margins.headerTwips === undefined ? current.headerTwips : margins.headerTwips,
    footerTwips: margins.footerTwips === undefined ? current.footerTwips : margins.footerTwips,
    gutterTwips: margins.gutterTwips === undefined ? current.gutterTwips : margins.gutterTwips,
  };

  return upsertSectionChildElement(sectionPropertiesXml, "w:pgMar", {
    "w:top": normalizeTwips(merged.topTwips, current.topTwips),
    "w:right": normalizeTwips(merged.rightTwips, current.rightTwips),
    "w:bottom": normalizeTwips(merged.bottomTwips, current.bottomTwips),
    "w:left": normalizeTwips(merged.leftTwips, current.leftTwips),
    "w:header": normalizeTwips(merged.headerTwips, current.headerTwips),
    "w:footer": normalizeTwips(merged.footerTwips, current.footerTwips),
    "w:gutter": normalizeTwips(merged.gutterTwips, current.gutterTwips),
  });
}

/** Rewrite `<w:pgSz>` in a `w:sectPr` fragment. */
export function writeSectionPageSize(
  sectionPropertiesXml: string,
  pageSize: DocxPageSizeUpdate,
): string {
  const current = readSectionPageSize(sectionPropertiesXml);
  const merged = normalizePageSizeOrientation({
    widthTwips: pageSize.widthTwips ?? current.widthTwips,
    heightTwips: pageSize.heightTwips ?? current.heightTwips,
    orientation: pageSize.orientation ?? current.orientation,
  });
  // Word only writes `w:orient` for landscape (portrait is the default), but an
  // existing attribute is kept explicit so a document that declares it stays
  // explicit after the edit.
  const existingOrientation = readAttribute(
    readElement(sectionPropertiesXml, "w:pgSz"),
    "w:orient",
  );
  const orientationAttribute =
    merged.orientation === "landscape" || existingOrientation !== undefined
      ? merged.orientation
      : undefined;

  return upsertSectionChildElement(sectionPropertiesXml, "w:pgSz", {
    "w:w": merged.widthTwips,
    "w:h": merged.heightTwips,
    "w:orient": orientationAttribute,
    // A stale paper code would name a different size than the new edges.
    "w:code": undefined,
  });
}

/** Read `<w:cols>` with Word's one-column, half-inch-gap defaults. */
export function readSectionColumns(
  sectionPropertiesXml?: string,
): DocxSectionColumns {
  const tag = sectionPropertiesXml
    ? readElement(sectionPropertiesXml, "w:cols")
    : undefined;
  const rawCount = readAttribute(tag, "w:num");
  const parsedCount = rawCount === undefined ? 1 : Math.round(Number(rawCount));
  const count = Number.isFinite(parsedCount) ? Math.max(1, parsedCount) : 1;
  const widthsTwips = [...String(tag ?? "").matchAll(/<w:col\b[^>]*\/?>/gi)]
    .map((match) => normalizeTwips(readAttribute(match[0], "w:w"), 0))
    .filter((width) => width > 0);

  return {
    count,
    spaceTwips: normalizeTwips(
      readAttribute(tag, "w:space"),
      DEFAULT_DOCX_COLUMN_SPACE_TWIPS,
    ),
    ...(count > 1 && widthsTwips.length === count ? { widthsTwips } : {}),
  };
}

/**
 * Rewrite `<w:cols>` in a `w:sectPr` fragment.
 *
 * Changing the count drops explicit `<w:col>` widths and a stale
 * `w:equalWidth="0"`, so the new count uses equal columns like Word's Columns
 * dialog. Every other `sectPr` child, and `w:sep`, are preserved.
 */
export function writeSectionColumns(
  sectionPropertiesXml: string,
  columns: DocxSectionColumnsUpdate,
): string {
  const current = readSectionColumns(sectionPropertiesXml);
  const count =
    columns.count === undefined
      ? current.count
      : Math.max(1, Math.min(45, Math.round(columns.count)));
  const spaceTwips =
    columns.spaceTwips === undefined
      ? current.spaceTwips
      : normalizeTwips(columns.spaceTwips, current.spaceTwips);

  if (count === current.count && spaceTwips === current.spaceTwips) {
    return sectionPropertiesXml;
  }

  const xml = /^<w:sectPr\b/i.test(sectionPropertiesXml.trim())
    ? sectionPropertiesXml
    : `<w:sectPr>${sectionPropertiesXml}</w:sectPr>`;
  const existing = readElement(xml, "w:cols");
  if (!existing) {
    return upsertSectionChildElement(xml, "w:cols", {
      ...(count > 1 ? { "w:num": count } : {}),
      "w:space": spaceTwips,
    });
  }

  // Replacing the whole element drops explicit `<w:col>` children (and their
  // attributes) from the previous count. A gap-only edit keeps the existing
  // start tag so unequal column widths survive.
  const startTag = existing.match(/^<w:cols\b[^>]*?>/i)?.[0] ?? "";
  if (count === current.count && startTag.length > 0) {
    const updatedStartTag = startTag.endsWith("/>")
      ? renderElement(
          "w:cols",
          mergeAttributes(readAttributes(startTag), { "w:space": spaceTwips }),
        )
      : `<w:cols ${mergeAttributes(readAttributes(startTag), {
          "w:space": spaceTwips,
        })
          .map(([name, value]) => `${name}="${value}"`)
          .join(" ")}>`;
    return xml.replace(startTag, updatedStartTag);
  }

  const separator = startTag.match(/\bw:sep="([^"]*)"/i)?.[1];
  const replacement = renderElement("w:cols", [
    ...(count > 1 ? [["w:num", String(count)] as [string, string]] : []),
    ["w:space", String(spaceTwips)],
    ...(separator === undefined ? [] : [["w:sep", separator] as [string, string]]),
  ]);
  return xml.replace(existing, replacement);
}

/** Insert a header/footer reference (`<w:headerReference>`) into a sectPr. */
export function addSectionPartReference(
  sectionPropertiesXml: string,
  kind: "header" | "footer",
  referenceType: string,
  relationshipId: string,
): string {
  const tagName = kind === "header" ? "w:headerReference" : "w:footerReference";
  const existing = readElement(sectionPropertiesXml, tagName);
  if (existing && readAttribute(existing, "w:type") === referenceType) {
    return sectionPropertiesXml;
  }

  // A second `<w:headerReference>` of the same type is invalid: replace the
  // matching element instead of adding a duplicate.
  const replacement = renderElement(tagName, [
    ["w:type", referenceType],
    ["r:id", relationshipId],
  ]);
  if (existing) {
    const sameType = new RegExp(
      `<${tagName}\\b[^>]*\\bw:type="${referenceType}"[^>]*?(?:/>|>[\\s\\S]*?</${tagName}>)`,
      "i",
    );
    if (sameType.test(sectionPropertiesXml)) {
      return sectionPropertiesXml.replace(sameType, replacement);
    }
  }

  const insertAt = sectionChildInsertIndex(sectionPropertiesXml, tagName);
  return (
    sectionPropertiesXml.slice(0, insertAt) +
    replacement +
    sectionPropertiesXml.slice(insertAt)
  );
}

/**
 * Document font loading.
 *
 * Before the first layout the viewer must ask the browser for every family the
 * document actually uses — body runs, headers/footers, notes, comments,
 * numbering labels and embedded text-box content — and it must await
 * `document.fonts.ready` before measuring. Otherwise the first pagination pass
 * measures a fallback face and the page count changes once the real face
 * arrives.
 *
 * `fontTableEntries()` supplies the families declared in `word/fontTable.xml`;
 * embedded faces arrive as `FontFace` descriptors (see
 * `collectEmbeddedFontFaceDescriptors`) and are installed by
 * `loadEmbeddedDocumentFonts()`.
 */

import type {
  DocModel,
  DocNode,
  FormFieldRunNode,
  NumberingLevelDefinition,
  ParagraphNode,
  TextStyle,
} from "../engine/types";
import { segmentTextByDocxScriptFont } from "./script-fonts";
import { invalidateFontMetrics } from "./font-metrics";
import { fontTableEntries } from "./font-table";

export interface DocumentFontRequest {
  /** Canvas font string at a fixed 16 px probe size. */
  font: string;
  family: string;
  characters: Set<string>;
}

export interface DocumentFontFaceDescriptor {
  family: string;
  style: "normal" | "italic";
  weight: string;
  source: ArrayBuffer;
}

export interface DocumentFontLoadOptions {
  /** Extra families to preload (typically the `w:font` names in fontTable). */
  extraFamilies?: readonly string[];
  /** Embedded `@font-face` data collected from the package. */
  embeddedFonts?: readonly DocumentFontFaceDescriptor[];
  /** Test/host override; defaults to the global document. */
  document?: Document;
}

export interface DocumentFontLoadResult {
  requests: DocumentFontRequest[];
  faces: FontFace[];
  loadedFamilies: string[];
}

function fontFamilyCss(family: string): string {
  const trimmed = family.trim();
  if (trimmed.includes(",")) return trimmed;
  const normalized = trimmed.replace(/^['"]+|['"]+$/g, "");
  if (/^(?:serif|sans-serif|monospace|cursive|fantasy|system-ui)$/i.test(normalized)) {
    return normalized;
  }
  return `"${normalized
    .replace(/\\/g, "\\\\")
    .replace(/"/g, '\\"')
    .replace(/[\n\r\f]/g, " ")}"`;
}

function mergeStyles(
  inherited: TextStyle | undefined,
  own: TextStyle | undefined
): TextStyle | undefined {
  if (!inherited) return own;
  if (!own) return inherited;
  return { ...inherited, ...own };
}

/**
 * Every distinct font request in the document. Runs are split by OOXML script
 * so a mixed Chinese/English/Arabic paragraph requests each slot family and
 * the characters that must be shaped for it.
 */
export function documentFontRequests(
  model: DocModel,
  extraFamilies: readonly string[] = []
): DocumentFontRequest[] {
  const requests = new Map<string, DocumentFontRequest>();
  const seenNodes = new Set<DocNode>();
  const usedNumberingIds = new Set<number>();
  const stylesById = new Map(
    model.metadata.paragraphStyles.map((style) => [style.id, style])
  );
  const defaultStyle = (
    stylesById.get(model.metadata.defaultParagraphStyleId ?? "") ??
    model.metadata.paragraphStyles.find((style) => style.isDefault)
  )?.runStyle;

  const collectText = (text: string, style?: TextStyle): void => {
    const safeText = text || "M";
    for (const segment of segmentTextByDocxScriptFont(safeText, style)) {
      const family = fontFamilyCss(segment.fontFamily ?? "Times New Roman");
      const font = `${style?.italic ? "italic" : "normal"} ${
        style?.bold ? "700" : "400"
      } 16px ${family}`;
      const key = font.toLowerCase();
      let request = requests.get(key);
      if (!request) {
        // Include the strut probe and generated numeric text before layout.
        request = { font, family, characters: new Set(" M0123456789") };
        requests.set(key, request);
      }
      for (const character of segment.text) {
        request.characters.add(character);
        if (style?.caps || style?.smallCaps) {
          for (const upper of character.toUpperCase()) {
            request.characters.add(upper);
          }
        }
      }
    }
  };

  const collectParagraph = (paragraph: ParagraphNode): void => {
    const paragraphStyle = stylesById.get(paragraph.style?.styleId ?? "")?.runStyle;
    const inheritedStyle = paragraphStyle ?? defaultStyle;
    const firstRunStyle = paragraph.children.find(
      (child) => child.type === "text"
    )?.style;
    collectText("M", firstRunStyle ?? inheritedStyle);
    if (paragraph.style?.numbering) {
      usedNumberingIds.add(paragraph.style.numbering.numId);
    }

    for (const child of paragraph.children) {
      if (child.type === "image") {
        if (child.textBoxText) {
          collectText(child.textBoxText, inheritedStyle);
        }
        continue;
      }
      if (child.type === "drawing") {
        for (const textBoxParagraph of child.textBox?.paragraphs ?? []) {
          for (const segment of textBoxParagraph.segments) {
            collectText(
              segment.text,
              mergeStyles(inheritedStyle, segment.style)
            );
          }
        }
        if (child.chart?.title) {
          collectText(child.chart.title, inheritedStyle);
        }
        for (const series of child.chart?.series ?? []) {
          if (series.name) collectText(series.name, inheritedStyle);
        }
        continue;
      }
      const style = mergeStyles(inheritedStyle, child.style);
      if (child.type === "text") {
        collectText(child.text, style);
        continue;
      }
      if (child.type === "math") {
        collectText(child.text, style);
        continue;
      }
      collectFormField(child, style);
    }
  };

  const collectFormField = (
    field: FormFieldRunNode,
    style: TextStyle | undefined
  ): void => {
    if (field.fieldType === "checkbox") {
      collectText(field.checkedSymbol ?? "☒", style);
      collectText(field.uncheckedSymbol ?? "☐", style);
      return;
    }
    collectText(
      field.value ?? field.widget?.text?.defaultText ?? field.placeholder ?? "",
      style
    );
    for (const option of field.options ?? []) {
      collectText(option.displayText, style);
    }
  };

  const visitNodes = (nodes: readonly DocNode[]): void => {
    for (const node of nodes) {
      if (seenNodes.has(node)) continue;
      seenNodes.add(node);
      if (node.type === "paragraph") {
        collectParagraph(node);
        continue;
      }
      for (const row of node.rows) {
        for (const cell of row.cells) {
          visitNodes(cell.nodes);
        }
      }
    }
  };

  visitNodes(model.nodes);
  for (const section of [model.metadata, ...(model.metadata.sections ?? [])]) {
    for (const part of [...section.headerSections, ...section.footerSections]) {
      visitNodes(part.nodes);
    }
  }
  for (const note of [
    ...(model.metadata.footnotes ?? []),
    ...(model.metadata.endnotes ?? []),
  ]) {
    if (note.nodes?.length) visitNodes(note.nodes);
    else if (note.text) collectText(note.text, defaultStyle);
  }
  for (const comment of model.metadata.comments ?? []) {
    if (comment.text) collectText(comment.text, defaultStyle);
  }

  const collectNumbering = (level: NumberingLevelDefinition): void => {
    collectText(
      `${level.text ?? ""}0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ`,
      level.bulletFontFamily
        ? { ...defaultStyle, ...level.runStyle, fontFamily: level.bulletFontFamily }
        : { ...defaultStyle, ...level.runStyle }
    );
  };
  const numbering = model.metadata.numberingDefinitions;
  if (numbering && usedNumberingIds.size > 0) {
    const usedAbstractIds = new Set<number>();
    for (const instance of numbering.instances) {
      if (!usedNumberingIds.has(instance.numId)) continue;
      usedAbstractIds.add(instance.abstractNumId);
      instance.levelOverrides?.forEach(collectNumbering);
    }
    for (const abstract of numbering.abstracts) {
      if (usedAbstractIds.has(abstract.abstractNumId)) {
        abstract.levels.forEach(collectNumbering);
      }
    }
  }

  for (const family of extraFamilies) {
    const normalized = family.trim();
    if (!normalized) continue;
    collectText(" M0123456789", { fontFamily: normalized });
  }

  return [...requests.values()];
}

/**
 * Install embedded faces (`word/fontTable.xml` + `word/fonts/*.odttf`). The
 * returned faces are owned by the caller so a viewer can dispose them when the
 * document changes.
 */
export async function loadEmbeddedDocumentFonts(
  descriptors: readonly DocumentFontFaceDescriptor[],
  targetDocument?: Document
): Promise<FontFace[]> {
  const doc = targetDocument ?? (typeof document !== "undefined" ? document : undefined);
  if (!doc?.fonts || descriptors.length === 0) {
    return [];
  }
  if (typeof FontFace === "undefined") {
    return [];
  }

  const loaded = (
    await Promise.all(
      descriptors.map(async (descriptor): Promise<FontFace | undefined> => {
        try {
          const face = new FontFace(descriptor.family, descriptor.source, {
            style: descriptor.style,
            weight: descriptor.weight,
          });
          await face.load();
          return face;
        } catch {
          return undefined;
        }
      })
    )
  ).filter((face): face is FontFace => Boolean(face));

  for (const face of loaded) {
    try {
      doc.fonts.add(face);
    } catch {
      // A face that another document already registered stays usable.
    }
  }
  if (loaded.length > 0) {
    invalidateFontMetrics();
  }
  return loaded;
}

export function removeDocumentFontFaces(
  faces: readonly FontFace[],
  targetDocument?: Document
): void {
  const doc = targetDocument ?? (typeof document !== "undefined" ? document : undefined);
  if (!doc?.fonts || faces.length === 0) return;
  for (const face of faces) {
    try {
      doc.fonts.delete(face);
    } catch {
      // Best-effort cleanup.
    }
  }
  invalidateFontMetrics();
}

/**
 * Ask the browser for every document family and wait until they settle. Safe
 * in SSR/Worker hosts: it becomes a no-op without a document.
 */
export async function awaitDocumentFonts(
  model: DocModel,
  options: DocumentFontLoadOptions = {}
): Promise<DocumentFontLoadResult> {
  const requests = documentFontRequests(model, options.extraFamilies);
  const doc = options.document ?? (typeof document !== "undefined" ? document : undefined);
  const result: DocumentFontLoadResult = {
    requests,
    faces: [],
    loadedFamilies: requests.map((request) => request.family),
  };
  if (!doc?.fonts) {
    return result;
  }

  result.faces = await loadEmbeddedDocumentFonts(
    options.embeddedFonts ?? [],
    doc
  );

  const fonts = doc.fonts;
  if (typeof fonts.load === "function") {
    await Promise.allSettled(
      requests.map(async ({ font, characters }) => {
        // Distinct code points cover unicode-range faces without retaining runs.
        await fonts.load(font, [...characters].join(""));
      })
    );
  }
  try {
    await fonts.ready;
  } catch {
    // Continue with the available fallback when a face cannot be loaded.
  }
  invalidateFontMetrics();
  return result;
}

/** Families declared by a `word/fontTable.xml` part. */
export function documentFontTableFamilies(
  fontTableXml: string | undefined
): string[] {
  return fontTableXml
    ? fontTableEntries(fontTableXml).map((entry) => entry.family)
    : [];
}

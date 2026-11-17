// Note numbering/placement properties (w:footnotePr / w:endnotePr).
//
// Word stores document-wide note properties in word/settings.xml
// (<w:settings><w:footnotePr>…<w:endnotePr>) and per-section overrides in
// <w:sectPr><w:footnotePr>…<w:endnotePr>. Only the fields that change how
// notes are numbered and where endnotes are placed are modelled here.

import type {
  DocumentNoteNumberingProperties,
  DocumentNoteProperties,
  NoteNumberPosition,
  NoteNumberRestart,
  NoteReferenceKind,
} from "./types";

const NOTE_PROPERTY_TAGS: Record<NoteReferenceKind, string> = {
  footnote: "w:footnotePr",
  endnote: "w:endnotePr",
};

const NOTE_RESTART_VALUES = new Set<NoteNumberRestart>([
  "continuous",
  "eachSect",
  "eachPage",
]);

const NOTE_POSITION_VALUES = new Set<NoteNumberPosition>(["docEnd", "sectEnd"]);

function attributeValue(tagXml: string, attribute: string): string | undefined {
  const escaped = attribute.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = tagXml.match(
    new RegExp(`${escaped}=(?:"([^"]*)"|'([^']*)')`, "i")
  );
  return match?.[1] ?? match?.[2];
}

/** First balanced <tag>…</tag> (or self-closing <tag/>) block in `xml`. */
function firstBalancedBlock(xml: string, tagName: string): string | undefined {
  const escaped = tagName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(`<(/?)${escaped}(?=[\\s>/])[^>]*>`, "gi");
  const stack: number[] = [];

  for (const match of xml.matchAll(pattern)) {
    const fullMatch = match[0] ?? "";
    if (!fullMatch) {
      continue;
    }
    const start = match.index ?? 0;
    const isClosing = match[1] === "/";
    const isSelfClosing = /\/>$/.test(fullMatch);

    if (!isClosing) {
      if (isSelfClosing) {
        if (stack.length === 0) {
          return "";
        }
        continue;
      }
      stack.push(start + fullMatch.length);
      continue;
    }

    const contentStart = stack.pop();
    if (contentStart === undefined) {
      continue;
    }
    if (stack.length === 0) {
      return xml.slice(contentStart, start);
    }
  }

  return undefined;
}

function firstTag(block: string, tagName: string): string | undefined {
  const escaped = tagName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return block.match(new RegExp(`<${escaped}(?=[\\s>/])[^>]*>`, "i"))?.[0];
}

function parseNumberRestart(value: string | undefined): NoteNumberRestart | undefined {
  if (!value) {
    return undefined;
  }
  const normalized = value.trim();
  for (const candidate of NOTE_RESTART_VALUES) {
    if (candidate.toLowerCase() === normalized.toLowerCase()) {
      return candidate;
    }
  }
  return undefined;
}

function parseNumberPosition(value: string | undefined): NoteNumberPosition | undefined {
  if (!value) {
    return undefined;
  }
  const normalized = value.trim();
  for (const candidate of NOTE_POSITION_VALUES) {
    if (candidate.toLowerCase() === normalized.toLowerCase()) {
      return candidate;
    }
  }
  return undefined;
}

function parseNumberStart(value: string | undefined): number | undefined {
  if (!value) {
    return undefined;
  }
  const parsed = Number.parseInt(value.trim(), 10);
  if (!Number.isFinite(parsed) || parsed < 1) {
    return undefined;
  }
  return parsed;
}

/**
 * Parse w:footnotePr / w:endnotePr from a container part (settings.xml or a
 * section's sectPr XML). Returns undefined when the block is absent.
 */
export function parseNoteNumberingProperties(
  xml: string | undefined,
  kind: NoteReferenceKind
): DocumentNoteNumberingProperties | undefined {
  if (!xml) {
    return undefined;
  }

  const block = firstBalancedBlock(xml, NOTE_PROPERTY_TAGS[kind]);
  if (block === undefined) {
    return undefined;
  }

  const properties: DocumentNoteNumberingProperties = {};

  const numFmtTag = firstTag(block, "w:numFmt");
  const numFmt = numFmtTag ? attributeValue(numFmtTag, "w:val")?.trim() : undefined;
  if (numFmt) {
    properties.numFmt = numFmt;
  }

  const numRestartTag = firstTag(block, "w:numRestart");
  const numRestart = parseNumberRestart(
    numRestartTag ? attributeValue(numRestartTag, "w:val") : undefined
  );
  if (numRestart) {
    properties.numRestart = numRestart;
  }

  const numStartTag = firstTag(block, "w:numStart");
  const numStart = parseNumberStart(
    numStartTag ? attributeValue(numStartTag, "w:val") : undefined
  );
  if (numStart !== undefined) {
    properties.numStart = numStart;
  }

  const positionTag = firstTag(block, "w:pos");
  const position = parseNumberPosition(
    positionTag ? attributeValue(positionTag, "w:val") : undefined
  );
  if (position) {
    properties.position = position;
  }

  return Object.keys(properties).length > 0 ? properties : undefined;
}

/** Parse both w:footnotePr and w:endnotePr from one XML part. */
export function parseNoteProperties(
  xml: string | undefined
): DocumentNoteProperties | undefined {
  const footnote = parseNoteNumberingProperties(xml, "footnote");
  const endnote = parseNoteNumberingProperties(xml, "endnote");
  if (!footnote && !endnote) {
    return undefined;
  }
  return { footnote, endnote };
}

/** Merge document-wide note properties with a section-level override. */
export function mergeNoteNumberingProperties(
  base?: DocumentNoteNumberingProperties,
  override?: DocumentNoteNumberingProperties
): DocumentNoteNumberingProperties | undefined {
  if (!base && !override) {
    return undefined;
  }
  const merged: DocumentNoteNumberingProperties = {};
  const numFmt = override?.numFmt ?? base?.numFmt;
  if (numFmt) merged.numFmt = numFmt;
  const numRestart = override?.numRestart ?? base?.numRestart;
  if (numRestart) merged.numRestart = numRestart;
  const numStart = override?.numStart ?? base?.numStart;
  if (numStart !== undefined) merged.numStart = numStart;
  const position = override?.position ?? base?.position;
  if (position) merged.position = position;
  return Object.keys(merged).length > 0 ? merged : undefined;
}

// Re-emit footnote/endnote reference runs when a model is serialized.
//
// A note marker is model data: the text run carries
// `noteReference = { kind, id }` and the note body lives in
// `metadata.footnotes` / `metadata.endnotes`. The WASM serializer writes a
// paragraph's `sourceXml` verbatim (markers included) and otherwise renders the
// paragraph from model fields, which have no slot for the reference element: the
// marker run comes out as an empty `<w:r><w:t/></w:r>`.
//
// The note bodies survive (their parts are carried over from the base package),
// so without this patch the exported file holds footnote text that no run points
// at, and re-importing silently drops the marker. This patch splices the
// reference back in at the model's position, for body paragraphs, table cells and
// nested tables alike.

import type { DocModel, DocNode, ParagraphNode } from "./types";
import {
  applyDocumentXmlEdits,
  collectRunTextAtoms,
  ensurePreserveSpace,
  escapeXmlText,
  extractRunProperties,
  locateBodyParagraphSpans,
  type DocumentXmlEdit,
  type XmlTextAtom,
} from "./document-xml-patch";

export interface SerializedNoteReference {
  /** Offset in the paragraph text where the marker belongs. */
  offset: number;
  kind: "footnote" | "endnote";
  id: number;
}

export interface ParagraphNoteTarget {
  /** Index among every `<w:p>` in the part, in document order (cells included). */
  paragraphIndex: number;
  references: SerializedNoteReference[];
}

const NOTE_REFERENCE_ELEMENT: Record<SerializedNoteReference["kind"], string> = {
  footnote: "w:footnoteReference",
  endnote: "w:endnoteReference",
};

/**
 * Note markers of one paragraph, as offsets into `paragraphText(paragraph)`.
 *
 * A marker belongs after the text of the run that carries it, which is also how
 * a split run keeps its reference on the last piece.
 */
function paragraphNoteReferences(
  paragraph: ParagraphNode
): SerializedNoteReference[] {
  const references: SerializedNoteReference[] = [];
  let offset = 0;

  for (const child of paragraph.children) {
    if (child.type !== "text") {
      continue;
    }

    offset += child.text.length;
    const noteReference = child.noteReference;
    if (!noteReference || !Number.isFinite(noteReference.id)) {
      continue;
    }

    references.push({
      offset,
      kind: noteReference.kind,
      id: noteReference.id,
    });
  }

  return references;
}

function collectNodeParagraphs(
  nodes: readonly DocNode[],
  targets: ParagraphNoteTarget[],
  counter: { value: number }
): void {
  for (const node of nodes) {
    if (node.type === "paragraph") {
      const paragraphIndex = counter.value;
      counter.value += 1;
      // `sourceXml` is emitted verbatim, reference runs included.
      if (node.sourceXml) {
        continue;
      }
      const references = paragraphNoteReferences(node);
      if (references.length > 0) {
        targets.push({ paragraphIndex, references });
      }
      continue;
    }

    for (const row of node.rows) {
      for (const cell of row.cells) {
        collectNodeParagraphs(cell.nodes, targets, counter);
      }
    }
  }
}

/**
 * Paragraphs whose note markers only exist in the model (their `sourceXml` was
 * cleared by an edit) and the references to splice back in.
 */
export function collectUnserializedNoteReferences(
  model: DocModel
): ParagraphNoteTarget[] {
  const targets: ParagraphNoteTarget[] = [];
  const counter = { value: 0 };
  collectNodeParagraphs(model.nodes, targets, counter);
  return targets;
}

export function modelNeedsNoteReferenceSerialization(model: DocModel): boolean {
  return collectUnserializedNoteReferences(model).length > 0;
}

function referenceElement(reference: SerializedNoteReference): string {
  return `<${NOTE_REFERENCE_ELEMENT[reference.kind]} w:id="${reference.id}"/>`;
}

/**
 * Reference elements of this kind/id already present in the serialized paragraph.
 *
 * The serializer renders the marker for paragraphs that live in a table cell but
 * writes an empty text run for body paragraphs, so the patch has to tell the two
 * apart instead of always adding a marker.
 */
function countSerializedReferences(
  content: string,
  reference: SerializedNoteReference
): number {
  const pattern = new RegExp(
    `<${NOTE_REFERENCE_ELEMENT[reference.kind]}\\b[^>]*\\bw:id="${reference.id}"[^>]*\\/>`,
    "g"
  );
  return [...content.matchAll(pattern)].length;
}

/** A standalone run carrying only the reference marker. */
function referenceRun(reference: SerializedNoteReference): string {
  return (
    "<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr>" +
    `${referenceElement(reference)}</w:r>`
  );
}

/**
 * Splices the reference into the run the offset falls inside, cutting the run in
 * two so the surrounding text keeps its own run properties.
 */
function splitRunAtNoteReference(
  content: string,
  atom: XmlTextAtom,
  reference: SerializedNoteReference
): DocumentXmlEdit {
  const runXml = content.slice(atom.runStart, atom.runEnd);
  const runOpenEnd = runXml.indexOf(">") + 1;
  const runOpenTag = runXml.slice(0, runOpenEnd);
  const runProperties = extractRunProperties(runXml);
  const textOpenTag = ensurePreserveSpace(atom.openTag);
  const localOffset = Math.max(
    0,
    Math.min(atom.text.length, reference.offset - atom.textOffset)
  );
  const runForText = (text: string): string =>
    `${runOpenTag}${runProperties}${textOpenTag}${escapeXmlText(text)}${atom.closeTag}</w:r>`;

  return {
    start: atom.runStart,
    end: atom.runEnd,
    replacement:
      runForText(atom.text.slice(0, localOffset)) +
      referenceRun(reference) +
      runForText(atom.text.slice(localOffset)),
  };
}

/**
 * Puts one reference marker into the paragraph body.
 *
 * The serializer already writes an empty run for the marker run, so that run is
 * filled in (keeping its run properties, which carry the superscript) and no
 * stray empty run is left behind. When there is no such run the marker is
 * inserted at the text offset instead.
 */
function insertNoteReferenceIntoParagraphContent(
  content: string,
  reference: SerializedNoteReference
): string {
  const atoms = collectRunTextAtoms(content);
  const element = referenceElement(reference);

  // The marker run: an empty text run sitting exactly at the marker offset.
  const emptyRun = atoms.find(
    (atom) => atom.text.length === 0 && atom.textOffset === reference.offset
  );
  if (emptyRun) {
    const runXml = content.slice(emptyRun.runStart, emptyRun.runEnd);
    const hasRunProperties = /<w:rPr\b/i.test(runXml);
    return applyDocumentXmlEdits(content, [
      {
        start: emptyRun.tagStart,
        end: emptyRun.tagEnd,
        replacement: hasRunProperties
          ? element
          : `<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>${element}`,
      },
    ]);
  }

  if (atoms.length === 0) {
    return applyDocumentXmlEdits(content, [
      { start: content.length, end: content.length, replacement: referenceRun(reference) },
    ]);
  }

  const containingAtom = atoms.find(
    (atom) =>
      reference.offset > atom.textOffset &&
      reference.offset < atom.textOffset + atom.text.length
  );
  if (containingAtom) {
    return applyDocumentXmlEdits(content, [
      splitRunAtNoteReference(content, containingAtom, reference),
    ]);
  }

  // Between runs: attach the marker to the nearest run edge at that offset.
  let index = 0;
  for (const atom of atoms) {
    if (atom.textOffset + atom.text.length <= reference.offset) {
      index = atom.runEnd;
    } else if (atom.textOffset >= reference.offset) {
      index = atom.runStart;
      break;
    }
  }

  return applyDocumentXmlEdits(content, [
    { start: index, end: index, replacement: referenceRun(reference) },
  ]);
}

/**
 * Re-attaches footnote/endnote reference runs to serialized paragraphs whose
 * note markers only exist on the model.
 */
export function applyNoteReferencesToDocumentXml(
  documentXml: string,
  model: DocModel
): string {
  const targets = collectUnserializedNoteReferences(model);
  if (targets.length === 0) {
    return documentXml;
  }

  const spans = locateBodyParagraphSpans(documentXml);
  if (!spans) {
    return documentXml;
  }

  const targetByIndex = new Map(
    targets.map((target) => [target.paragraphIndex, target])
  );
  const edits: DocumentXmlEdit[] = [];

  spans.paragraphs.forEach((span, paragraphIndex) => {
    const target = targetByIndex.get(paragraphIndex);
    if (!target) {
      return;
    }

    const content = documentXml.slice(span.contentStart, span.contentEnd);
    let patched = content;
    const remainingSerialized = new Map<string, number>();
    for (const reference of target.references) {
      const key = `${reference.kind}:${reference.id}`;
      if (!remainingSerialized.has(key)) {
        remainingSerialized.set(key, countSerializedReferences(content, reference));
      }
      const alreadySerialized = remainingSerialized.get(key) ?? 0;
      if (alreadySerialized > 0) {
        // The serializer already wrote this marker (table-cell paragraphs).
        remainingSerialized.set(key, alreadySerialized - 1);
        continue;
      }
      patched = insertNoteReferenceIntoParagraphContent(patched, reference);
    }
    if (patched === content) {
      return;
    }

    edits.push({
      start: span.contentStart,
      end: span.contentEnd,
      replacement: patched,
    });
  });

  return applyDocumentXmlEdits(documentXml, edits);
}

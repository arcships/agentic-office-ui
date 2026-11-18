// Re-attach preserved paragraph-level `<w:sectPr>` fragments when the model is
// serialized.
//
// The WASM serializer emits a paragraph's `sourceXml` verbatim and otherwise
// renders it from model fields, which have no slot for a section break. After
// an edit clears `sourceXml` we therefore patch the serialized body: the
// paragraph's model position stays stable, so insert the preserved `<w:sectPr>`
// into that paragraph's `<w:pPr>`.

import type { DocModel } from "./types";
import { paragraphSectionPropertiesXml } from "./section-break";

interface SectionBreakMarker {
  /** Index among top-level body paragraphs (tables are not counted). */
  paragraphIndex: number;
  sectionPropertiesXml: string;
}

export function collectUnserializedSectionBreakMarkers(
  model: DocModel
): SectionBreakMarker[] {
  const markers: SectionBreakMarker[] = [];
  let paragraphIndex = 0;

  for (const node of model.nodes) {
    if (node.type !== "paragraph") {
      continue;
    }

    const sectionPropertiesXml = paragraphSectionPropertiesXml(node);
    // `sourceXml` wins in the serializer, including its section break.
    if (sectionPropertiesXml && !node.sourceXml?.includes("<w:sectPr")) {
      markers.push({ paragraphIndex, sectionPropertiesXml });
    }
    paragraphIndex += 1;
  }

  return markers;
}

export function modelNeedsSectionBreakSerialization(model: DocModel): boolean {
  return collectUnserializedSectionBreakMarkers(model).length > 0;
}

interface ParagraphEdit {
  start: number;
  end: number;
  replacement: string;
}

interface TargetParagraph {
  marker: SectionBreakMarker;
  startTagStart: number;
  startTagEnd: number;
  selfClosing: boolean;
  endTagStart?: number;
}

const TAG_PATTERN = /<\/?([A-Za-z_][\w:.-]*)(?:\s[^<>]*?)?\s*\/?>/g;

function paragraphEditsForBody(
  bodyXml: string,
  markers: SectionBreakMarker[]
): ParagraphEdit[] {
  const markerByParagraphIndex = new Map(
    markers.map((marker) => [marker.paragraphIndex, marker])
  );
  const targets: TargetParagraph[] = [];
  let tableDepth = 0;
  let paragraphDepth = 0;
  let paragraphIndex = -1;
  let activeTarget: TargetParagraph | undefined;
  let match: RegExpExecArray | null;

  TAG_PATTERN.lastIndex = 0;
  while ((match = TAG_PATTERN.exec(bodyXml)) !== null) {
    const tag = match[0];
    const name = match[1].toLowerCase();
    const closing = tag.startsWith("</");
    const selfClosing = tag.endsWith("/>");

    if (closing) {
      if (name === "w:p") {
        if (
          paragraphDepth === 1 &&
          tableDepth === 0 &&
          activeTarget &&
          activeTarget.endTagStart === undefined
        ) {
          activeTarget.endTagStart = match.index;
          targets.push(activeTarget);
          activeTarget = undefined;
        }
        paragraphDepth = Math.max(0, paragraphDepth - 1);
      } else if (name === "w:tbl") {
        tableDepth = Math.max(0, tableDepth - 1);
      }
      continue;
    }

    const topLevelParagraph =
      name === "w:p" && paragraphDepth === 0 && tableDepth === 0;
    if (topLevelParagraph) {
      paragraphIndex += 1;
      const marker = markerByParagraphIndex.get(paragraphIndex);
      if (marker) {
        const target: TargetParagraph = {
          marker,
          startTagStart: match.index,
          startTagEnd: TAG_PATTERN.lastIndex,
          selfClosing
        };
        if (selfClosing) {
          targets.push(target);
        } else {
          activeTarget = target;
        }
      }
    }

    if (!selfClosing) {
      if (name === "w:p") {
        paragraphDepth += 1;
      } else if (name === "w:tbl") {
        tableDepth += 1;
      }
    }
  }

  const edits: ParagraphEdit[] = [];
  for (const target of targets) {
    if (target.selfClosing) {
      const openTag = bodyXml
        .slice(target.startTagStart, target.startTagEnd)
        .replace(/\/>\s*$/, ">");
      edits.push({
        start: target.startTagStart,
        end: target.startTagEnd,
        replacement: `${openTag}<w:pPr>${target.marker.sectionPropertiesXml}</w:pPr></w:p>`
      });
      continue;
    }

    if (target.endTagStart === undefined) {
      continue;
    }

    const content = bodyXml.slice(target.startTagEnd, target.endTagStart);
    if (/<w:sectPr\b/i.test(content)) {
      continue;
    }

    edits.push({
      start: target.startTagEnd,
      end: target.endTagStart,
      replacement: insertSectionPropertiesIntoParagraphContent(
        content,
        target.marker.sectionPropertiesXml
      )
    });
  }

  return edits;
}

function insertSectionPropertiesIntoParagraphContent(
  content: string,
  sectionPropertiesXml: string
): string {
  const selfClosingPPr = content.match(/^(\s*)<w:pPr\b[^>]*\/>/i);
  if (selfClosingPPr) {
    return (
      selfClosingPPr[1] +
      `<w:pPr>${sectionPropertiesXml}</w:pPr>` +
      content.slice(selfClosingPPr[0].length)
    );
  }

  const pPrOpen = content.match(/^(\s*)<w:pPr\b[^>]*>/i);
  if (pPrOpen) {
    const pPrCloseIndex = content.indexOf("</w:pPr>", pPrOpen[0].length);
    if (pPrCloseIndex >= 0) {
      return (
        content.slice(0, pPrCloseIndex) +
        sectionPropertiesXml +
        content.slice(pPrCloseIndex)
      );
    }
  }

  return `<w:pPr>${sectionPropertiesXml}</w:pPr>${content}`;
}

export function applyPreservedSectionBreaksToDocumentXml(
  documentXml: string,
  model: DocModel
): string {
  const markers = collectUnserializedSectionBreakMarkers(model);
  if (markers.length === 0) {
    return documentXml;
  }

  const bodyOpen = /<w:body\b[^>]*>/i.exec(documentXml);
  if (!bodyOpen || bodyOpen.index === undefined) {
    return documentXml;
  }

  const bodyContentStart = bodyOpen.index + bodyOpen[0].length;
  const bodyCloseIndex = documentXml
    .toLowerCase()
    .lastIndexOf("</w:body>");
  if (bodyCloseIndex < bodyContentStart) {
    return documentXml;
  }

  const bodyXml = documentXml.slice(bodyContentStart, bodyCloseIndex);
  const edits = paragraphEditsForBody(bodyXml, markers);
  if (edits.length === 0) {
    return documentXml;
  }

  let patchedBody = bodyXml;
  for (const edit of edits.sort((left, right) => right.start - left.start)) {
    patchedBody =
      patchedBody.slice(0, edit.start) +
      edit.replacement +
      patchedBody.slice(edit.end);
  }

  return (
    documentXml.slice(0, bodyContentStart) +
    patchedBody +
    documentXml.slice(bodyCloseIndex)
  );
}

import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import type { DocModel } from "./types";
import type { OoxmlPackage } from "./ooxml-core";
import {
  applyCommentRangesToDocumentXml,
  modelNeedsCommentRangeSerialization
} from "./comment-range-serialization";
import {
  applyCommentsToDocxArchive,
  applyCommentsToPackage,
  modelNeedsCommentSerialization,
  stripOrphanCommentMarkers
} from "./comment-serialization";
import {
  applyPreservedSectionBreaksToDocumentXml,
  modelNeedsSectionBreakSerialization
} from "./section-break-serialization";
import {
  applyNoteReferencesToDocumentXml,
  modelNeedsNoteReferenceSerialization
} from "./note-reference-serialization";
import {
  applyFormFieldRequiredToDocumentXml,
  modelNeedsFormFieldRequiredPatch
} from "./form-field-serialization";
import {
  DOCX_CONTENT_TYPES_PART,
  DOCX_DOCUMENT_RELATIONSHIPS_PART,
  applySectionPartContentTypes,
  applySectionPartRelationships,
  collectSectionPartReferences,
  modelHasSectionParts
} from "./section-parts-serialization";
import { prepareModelForSerialization } from "./prepare-model";
import {
  applyPreservedDrawingObjectsToDocumentXml,
  applyPreservedDrawingObjectsToPackage,
  applyPreservedDrawingObjectsToPartContent,
  modelNeedsDrawingObjectSerialization
} from "./drawing-object-serialization";
import {
  applyTableRowMergesToDocumentXml,
  applyTableRowMergesToPackage,
  applyTableRowMergesToPartContent,
  isTableMergeStoryPart,
  modelNeedsTableRowMergeSerialization
} from "./table-merge-serialization";
import {
  applyTableFloatingToDocumentXml,
  applyTableFloatingToPackage,
  applyTableFloatingToPartContent,
  isTableFloatingStoryPart,
  modelNeedsTableFloatingSerialization
} from "./table-floating-serialization";
import { mapsToWasmPackage, wasmModelToDocumentXml, wasmSerializeDocx } from "./wasm";

/**
 * All serialization patches rewrite the same document part, so they run in
 * order: preserved section breaks (they live in `w:pPr`), note reference runs
 * (they add runs but no text), comment range markers derived from character
 * offsets, then required form-field flags. Note references go first so the
 * comment splitter sees the marker run as an opaque run.
 */
function patchDocumentXml(
  documentXml: string,
  model: DocModel,
  renamedRelationshipIds?: Map<string, string>
): string {
  let patched = documentXml;
  // A deleted comment leaves its anchors behind in `sourceXml`; they must go
  // before the range pass re-emits the surviving comments.
  if (model.metadata.comments !== undefined) {
    const knownIds = new Set(model.metadata.comments.map((comment) => comment.id));
    patched = stripOrphanCommentMarkers(patched, knownIds);
  }
  if (modelNeedsSectionBreakSerialization(model)) {
    patched = applyPreservedSectionBreaksToDocumentXml(patched, model);
  }
  if (modelNeedsNoteReferenceSerialization(model)) {
    patched = applyNoteReferencesToDocumentXml(patched, model);
  }
  if (modelNeedsCommentRangeSerialization(model)) {
    patched = applyCommentRangesToDocumentXml(patched, model);
  }
  if (modelNeedsFormFieldRequiredPatch(model)) {
    patched = applyFormFieldRequiredToDocumentXml(patched, model);
  }
  if (renamedRelationshipIds && renamedRelationshipIds.size > 0) {
    patched = applyRenamedRelationshipIds(patched, renamedRelationshipIds);
  }
  // The engine rebuilds an edited table from the model and does not turn
  // `rowSpan`/`vMergeContinuation` back into `w:vMerge`; put the markers back
  // once the rest of the part is final.
  if (modelNeedsTableRowMergeSerialization(model)) {
    patched = applyTableRowMergesToDocumentXml(patched, model);
  }
  // Same rebuild path for `TableStyle.floating`: a table dragged to a position
  // keeps its `w:tblpPr` (drop position, anchors and wrap distances).
  if (modelNeedsTableFloatingSerialization(model)) {
    patched = applyTableFloatingToDocumentXml(patched, model);
  }
  // Shapes, charts, embedded objects and equations live only in the stored
  // markup; a paragraph an edit rebuilt has to get them back.
  if (modelNeedsDrawingObjectSerialization(model)) {
    patched = applyPreservedDrawingObjectsToDocumentXml(patched, model);
  }
  return patched;
}

/**
 * A header/footer reference moved to a free relationship id because the model's
 * id already belonged to another relationship. Rewrite the reference so the
 * document and `word/_rels/document.xml.rels` still agree.
 */
function applyRenamedRelationshipIds(
  documentXml: string,
  renamedRelationshipIds: Map<string, string>
): string {
  let patched = documentXml;
  for (const [from, to] of renamedRelationshipIds) {
    const reference = new RegExp(`(r:id\\s*=\\s*")${from}(")`, "g");
    patched = patched.replace(reference, `$1${to}$2`);
  }
  return patched;
}

interface SectionPartPackagePatch {
  relationshipsXml?: string;
  contentTypesXml?: string;
  renamedRelationshipIds: Map<string, string>;
}

/**
 * Relationships and content types a package needs for its header/footer parts.
 * `undefined` entries mean the package part is unchanged.
 */
function planSectionPartPackagePatch(
  parts: ReadonlyMap<string, { content: string }>,
  model: DocModel
): SectionPartPackagePatch {
  const renamedRelationshipIds = new Map<string, string>();
  if (!modelHasSectionParts(model)) {
    return { renamedRelationshipIds };
  }

  const references = collectSectionPartReferences(model);
  if (references.length === 0) {
    return { renamedRelationshipIds };
  }

  const relationshipsPart = parts.get(DOCX_DOCUMENT_RELATIONSHIPS_PART);
  const contentTypesPart = parts.get(DOCX_CONTENT_TYPES_PART);
  const patch: SectionPartPackagePatch = { renamedRelationshipIds };

  if (relationshipsPart) {
    const relationships = applySectionPartRelationships(
      relationshipsPart.content,
      references
    );
    if (relationships.xml !== relationshipsPart.content) {
      patch.relationshipsXml = relationships.xml;
    }
    relationships.renamedIds.forEach((value, key) => {
      renamedRelationshipIds.set(key, value);
    });
  }

  if (contentTypesPart) {
    const contentTypes = applySectionPartContentTypes(
      contentTypesPart.content,
      references
    );
    if (contentTypes !== contentTypesPart.content) {
      patch.contentTypesXml = contentTypes;
    }
  }

  return patch;
}

function modelNeedsDocumentXmlPatch(model: DocModel): boolean {
  return (
    model.metadata.comments !== undefined ||
    modelNeedsSectionBreakSerialization(model) ||
    modelNeedsNoteReferenceSerialization(model) ||
    modelNeedsCommentRangeSerialization(model) ||
    modelNeedsFormFieldRequiredPatch(model) ||
    modelNeedsTableRowMergeSerialization(model) ||
    modelNeedsTableFloatingSerialization(model) ||
    modelNeedsDrawingObjectSerialization(model)
  );
}

function patchPackageDocumentXml(
  pkg: OoxmlPackage,
  model: DocModel
): OoxmlPackage {
  const documentPart = pkg.parts.get("word/document.xml");
  const sectionParts = planSectionPartPackagePatch(pkg.parts, model);
  if (!documentPart) {
    return applySectionPartPackagePatch(pkg, sectionParts);
  }

  const needsDocumentPatch = modelNeedsDocumentXmlPatch(model);
  const patchedXml = needsDocumentPatch
    ? patchDocumentXml(
        documentPart.content,
        model,
        sectionParts.renamedRelationshipIds
      )
    : applyRenamedRelationshipIds(
        documentPart.content,
        sectionParts.renamedRelationshipIds
      );
  const parts = new Map(pkg.parts);
  if (patchedXml !== documentPart.content) {
    parts.set("word/document.xml", {
      name: documentPart.name,
      content: patchedXml
    });
  }
  return applySectionPartPackagePatch(
    { parts, binaryAssets: pkg.binaryAssets },
    sectionParts
  );
}

function applySectionPartPackagePatch(
  pkg: OoxmlPackage,
  patch: SectionPartPackagePatch
): OoxmlPackage {
  if (!patch.relationshipsXml && !patch.contentTypesXml) {
    return pkg;
  }

  const parts = new Map(pkg.parts);
  if (patch.relationshipsXml) {
    const part = parts.get(DOCX_DOCUMENT_RELATIONSHIPS_PART);
    if (part) {
      parts.set(DOCX_DOCUMENT_RELATIONSHIPS_PART, {
        name: part.name,
        content: patch.relationshipsXml
      });
    }
  }
  if (patch.contentTypesXml) {
    const part = parts.get(DOCX_CONTENT_TYPES_PART);
    if (part) {
      parts.set(DOCX_CONTENT_TYPES_PART, {
        name: part.name,
        content: patch.contentTypesXml
      });
    }
  }
  return { parts, binaryAssets: pkg.binaryAssets };
}

/**
 * The package-level text parts the section-part patch may rewrite, read from an
 * unzipped archive. Images, fonts and other binary parts are never decoded.
 */
function collectPackageTextParts(
  archive: Record<string, Uint8Array>
): Map<string, { content: string }> {
  const parts = new Map<string, { content: string }>();
  for (const name of [DOCX_DOCUMENT_RELATIONSHIPS_PART, DOCX_CONTENT_TYPES_PART]) {
    const bytes = archive[name];
    if (bytes) parts.set(name, { content: strFromU8(bytes) });
  }
  return parts;
}

function patchSerializedDocxBytes(
  bytes: ArrayBuffer,
  model: DocModel
): ArrayBuffer {
  const needsDocumentPatch = modelNeedsDocumentXmlPatch(model);
  const needsComments = modelNeedsCommentSerialization(model);
  const needsSectionParts = modelHasSectionParts(model);
  const needsMerges = modelNeedsTableRowMergeSerialization(model);
  const needsFloating = modelNeedsTableFloatingSerialization(model);
  const needsDrawingObjects = modelNeedsDrawingObjectSerialization(model);
  if (!needsDocumentPatch && !needsComments && !needsSectionParts) {
    return bytes;
  }

  const archive = unzipSync(new Uint8Array(bytes));
  const documentBytes = archive["word/document.xml"];
  if (!documentBytes) {
    return bytes;
  }

  let changed = false;
  if (needsComments) {
    // Comment parts, their relationships and content types are written first;
    // the header/footer patch below then appends to the updated wrapper.
    applyCommentsToDocxArchive(archive, model);
    changed = true;
  }

  const sectionParts = planSectionPartPackagePatch(
    collectPackageTextParts(archive),
    model
  );

  const documentXml = strFromU8(documentBytes);
  const patchedXml = needsDocumentPatch
    ? patchDocumentXml(documentXml, model, sectionParts.renamedRelationshipIds)
    : applyRenamedRelationshipIds(documentXml, sectionParts.renamedRelationshipIds);
  if (patchedXml !== documentXml) {
    archive["word/document.xml"] = strToU8(patchedXml);
    changed = true;
  }
  if (needsMerges) {
    // Vertical merges in header/footer/note stories are rebuilt by the same
    // engine path as body tables, so every story part needs the same patch.
    for (const name of Object.keys(archive)) {
      if (!isTableMergeStoryPart(name)) continue;
      const text = strFromU8(archive[name]);
      const patched = applyTableRowMergesToPartContent(name, text, model);
      if (patched !== text) {
        archive[name] = strToU8(patched);
        changed = true;
      }
    }
  }
  if (needsFloating) {
    // Header and footer tables are rebuilt the same way as body tables.
    for (const name of Object.keys(archive)) {
      if (!isTableFloatingStoryPart(name)) continue;
      const text = strFromU8(archive[name]);
      const patched = applyTableFloatingToPartContent(name, text, model);
      if (patched !== text) {
        archive[name] = strToU8(patched);
        changed = true;
      }
    }
  }
  if (needsDrawingObjects) {
    // The body part was patched above; headers, footers and notes are rebuilt
    // by the same engine path and need the same treatment.
    for (const name of Object.keys(archive)) {
      if (name.toLowerCase() === "word/document.xml") continue;
      if (!/^word\/(?:header|footer)\d+\.xml$|^word\/(?:footnotes|endnotes)\.xml$/i.test(name)) {
        continue;
      }
      const text = strFromU8(archive[name]);
      const patched = applyPreservedDrawingObjectsToPartContent(name, text, model);
      if (patched !== text) {
        archive[name] = strToU8(patched);
        changed = true;
      }
    }
  }
  if (sectionParts.relationshipsXml) {
    archive[DOCX_DOCUMENT_RELATIONSHIPS_PART] = strToU8(sectionParts.relationshipsXml);
    changed = true;
  }
  if (sectionParts.contentTypesXml) {
    archive[DOCX_CONTENT_TYPES_PART] = strToU8(sectionParts.contentTypesXml);
    changed = true;
  }
  if (!changed) {
    return bytes;
  }

  const zipped = zipSync(archive);
  return zipped.buffer.slice(
    zipped.byteOffset,
    zipped.byteOffset + zipped.byteLength
  ) as ArrayBuffer;
}

export async function modelToDocumentXml(
  model: DocModel,
  basePackage?: OoxmlPackage
): Promise<string> {
  const prepared = prepareModelForSerialization(model);
  const documentXml = await wasmModelToDocumentXml(
    prepared,
    basePackage ? mapsToWasmPackage(basePackage) : undefined
  );
  return patchDocumentXml(documentXml, prepared);
}

export async function serializeDocModel(
  model: DocModel,
  basePackage?: OoxmlPackage
): Promise<OoxmlPackage> {
  const prepared = prepareModelForSerialization(model);
  const bytes = await wasmSerializeDocx(
    prepared,
    basePackage ? mapsToWasmPackage(basePackage) : undefined
  );
  const { parseDocx } = await import("./ooxml-core");
  return applyCommentsToPackage(
    applyPreservedDrawingObjectsToPackage(
      applyTableFloatingToPackage(
        applyTableRowMergesToPackage(
          patchPackageDocumentXml(await parseDocx(bytes), prepared),
          prepared
        ),
        prepared
      ),
      prepared
    ),
    prepared
  );
}

export async function serializeDocx(
  model: DocModel,
  basePackage?: OoxmlPackage
): Promise<ArrayBuffer> {
  const prepared = prepareModelForSerialization(model);
  const bytes = await wasmSerializeDocx(
    prepared,
    basePackage ? mapsToWasmPackage(basePackage) : undefined
  );
  return patchSerializedDocxBytes(bytes, prepared);
}

import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import type { DocModel } from "./types";
import type { OoxmlPackage } from "./ooxml-core";
import {
  applyCommentRangesToDocumentXml,
  modelNeedsCommentRangeSerialization
} from "./comment-range-serialization";
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
import { mapsToWasmPackage, wasmModelToDocumentXml, wasmSerializeDocx } from "./wasm";

/**
 * All serialization patches rewrite the same document part, so they run in
 * order: preserved section breaks (they live in `w:pPr`), note reference runs
 * (they add runs but no text), comment range markers derived from character
 * offsets, then required form-field flags. Note references go first so the
 * comment splitter sees the marker run as an opaque run.
 */
function patchDocumentXml(documentXml: string, model: DocModel): string {
  let patched = documentXml;
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
  return patched;
}

function modelNeedsDocumentXmlPatch(model: DocModel): boolean {
  return (
    modelNeedsSectionBreakSerialization(model) ||
    modelNeedsNoteReferenceSerialization(model) ||
    modelNeedsCommentRangeSerialization(model) ||
    modelNeedsFormFieldRequiredPatch(model)
  );
}

function patchPackageDocumentXml(
  pkg: OoxmlPackage,
  model: DocModel
): OoxmlPackage {
  if (!modelNeedsDocumentXmlPatch(model)) {
    return pkg;
  }

  const documentPart = pkg.parts.get("word/document.xml");
  if (!documentPart) {
    return pkg;
  }

  const patchedXml = patchDocumentXml(documentPart.content, model);
  if (patchedXml === documentPart.content) {
    return pkg;
  }

  const parts = new Map(pkg.parts);
  parts.set("word/document.xml", {
    name: documentPart.name,
    content: patchedXml
  });
  return { parts, binaryAssets: pkg.binaryAssets };
}

function patchSerializedDocxBytes(
  bytes: ArrayBuffer,
  model: DocModel
): ArrayBuffer {
  if (!modelNeedsDocumentXmlPatch(model)) {
    return bytes;
  }

  const archive = unzipSync(new Uint8Array(bytes));
  const documentBytes = archive["word/document.xml"];
  if (!documentBytes) {
    return bytes;
  }

  const documentXml = strFromU8(documentBytes);
  const patchedXml = patchDocumentXml(documentXml, model);
  if (patchedXml === documentXml) {
    return bytes;
  }

  archive["word/document.xml"] = strToU8(patchedXml);
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
  const documentXml = await wasmModelToDocumentXml(
    model,
    basePackage ? mapsToWasmPackage(basePackage) : undefined
  );
  return patchDocumentXml(documentXml, model);
}

export async function serializeDocModel(
  model: DocModel,
  basePackage?: OoxmlPackage
): Promise<OoxmlPackage> {
  const bytes = await wasmSerializeDocx(
    model,
    basePackage ? mapsToWasmPackage(basePackage) : undefined
  );
  const { parseDocx } = await import("./ooxml-core");
  return patchPackageDocumentXml(await parseDocx(bytes), model);
}

export async function serializeDocx(
  model: DocModel,
  basePackage?: OoxmlPackage
): Promise<ArrayBuffer> {
  const bytes = await wasmSerializeDocx(
    model,
    basePackage ? mapsToWasmPackage(basePackage) : undefined
  );
  return patchSerializedDocxBytes(bytes, model);
}

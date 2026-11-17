import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import type { DocModel } from "./types";
import type { OoxmlPackage } from "./ooxml-core";
import {
  applyPreservedSectionBreaksToDocumentXml,
  modelNeedsSectionBreakSerialization
} from "./section-break-serialization";
import { mapsToWasmPackage, wasmModelToDocumentXml, wasmSerializeDocx } from "./wasm";

function patchPackageSectionBreaks(
  pkg: OoxmlPackage,
  model: DocModel
): OoxmlPackage {
  if (!modelNeedsSectionBreakSerialization(model)) {
    return pkg;
  }

  const documentPart = pkg.parts.get("word/document.xml");
  if (!documentPart) {
    return pkg;
  }

  const patchedXml = applyPreservedSectionBreaksToDocumentXml(
    documentPart.content,
    model
  );
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
  if (!modelNeedsSectionBreakSerialization(model)) {
    return bytes;
  }

  const archive = unzipSync(new Uint8Array(bytes));
  const documentBytes = archive["word/document.xml"];
  if (!documentBytes) {
    return bytes;
  }

  const documentXml = strFromU8(documentBytes);
  const patchedXml = applyPreservedSectionBreaksToDocumentXml(
    documentXml,
    model
  );
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
  return applyPreservedSectionBreaksToDocumentXml(documentXml, model);
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
  return patchPackageSectionBreaks(await parseDocx(bytes), model);
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

// doc-model aggregation — builds DocModel from OOXML packages/bytes.
// Ported from upstream doc-model/src/index.ts (commit 6f70b92).

import type { OoxmlPackage } from "./ooxml-core";
import {
  mapsToWasmPackage,
  wasmBuildDocModelFromPackage,
  wasmPackageToMaps,
} from "./wasm";
import { parseDocx } from "./ooxml-core";

import { normalizeDocModel } from "./normalize";
import { parseNoteProperties } from "./note-properties";
import { enrichDocModelTypography } from "./model-enrichment";
import type { DocModel, DocumentNoteProperties } from "./types";

export * from "./types";
export { cloneDocModel, cloneDocModelShallow, cloneDocNode } from "./clone";
export { normalizeDocModel } from "./normalize";

const SETTINGS_PART = "word/settings.xml";

/**
 * The WASM model builder does not surface w:footnotePr / w:endnotePr, so the
 * document-wide note properties are read from word/settings.xml here. Section
 * level overrides stay in metadata.sections[].sectionPropertiesXml.
 */
function attachDocumentNoteProperties(
  model: DocModel,
  pkg: OoxmlPackage,
): DocModel {
  if (model.metadata.noteProperties) {
    return model;
  }
  const settingsXml = pkg.parts.get(SETTINGS_PART)?.content;
  const noteProperties: DocumentNoteProperties | undefined =
    parseNoteProperties(settingsXml);
  if (!noteProperties) {
    return model;
  }
  return {
    ...model,
    metadata: { ...model.metadata, noteProperties },
  };
}

export async function buildDocModel(pkg: OoxmlPackage): Promise<DocModel> {
  const wasmPackage = mapsToWasmPackage({
    parts: pkg.parts,
    binaryAssets: pkg.binaryAssets,
  });
  const model = (await wasmBuildDocModelFromPackage(wasmPackage)) as DocModel;
  enrichDocModelTypography(model, pkg);
  return normalizeDocModel(attachDocumentNoteProperties(model, pkg));
}

export async function buildDocModelFromBytes(
  bytes: ArrayBuffer | Uint8Array,
): Promise<{ package: OoxmlPackage; model: DocModel }> {
  const payload = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const buffer = payload.buffer.slice(
    payload.byteOffset,
    payload.byteOffset + payload.byteLength,
  ) as ArrayBuffer;
  const pkg = await parseDocx(buffer);
  const model = await buildDocModel(pkg);
  return { package: pkg, model };
}

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
import { parseDocumentBackground } from "./page-background";
import { attachDocumentDefaults } from "./document-defaults";
import { parseNoteProperties } from "./note-properties";
import { enrichDocModelFormFields } from "./form-field-enrichment";
import { enrichDocModelTypography } from "./model-enrichment";
import type { DocModel, DocumentNoteProperties } from "./types";

export * from "./types";
export {
  cloneDocModel,
  cloneDocModelNodes,
  cloneDocModelSectionLists,
  cloneDocModelShallow,
  cloneDocNode,
} from "./clone";
export { normalizeDocModel } from "./normalize";

const SETTINGS_PART = "word/settings.xml";
const DOCUMENT_PART = "word/document.xml";

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

/**
 * The WASM model builder only surfaces the solid w:background colour. Pattern
 * and shading fills need the raw document XML, so they are attached here
 * before normalization (both the main-thread and Worker import paths call
 * buildDocModel).
 */
function attachDocumentBackground(
  model: DocModel,
  pkg: OoxmlPackage,
): DocModel {
  const documentXml = pkg.parts.get(DOCUMENT_PART)?.content;
  const background = parseDocumentBackground(documentXml);
  if (!background) return model;

  const metadata = { ...model.metadata };
  if (!metadata.documentBackgroundColor && background.color) {
    metadata.documentBackgroundColor = background.color;
  }
  if (background.pattern) {
    metadata.documentBackgroundPattern = background.pattern;
  }
  return { ...model, metadata };
}

export async function buildDocModel(pkg: OoxmlPackage): Promise<DocModel> {
  const wasmPackage = mapsToWasmPackage({
    parts: pkg.parts,
    binaryAssets: pkg.binaryAssets,
  });
  const model = (await wasmBuildDocModelFromPackage(wasmPackage)) as DocModel;
  enrichDocModelFormFields(model);
  enrichDocModelTypography(model, pkg);
  return normalizeDocModel(
    attachDocumentDefaults(
      attachDocumentBackground(attachDocumentNoteProperties(model, pkg), pkg),
      pkg
    )
  );
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

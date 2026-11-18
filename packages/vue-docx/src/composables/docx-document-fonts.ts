// Document font hook shared by the viewer model loader and the editor import
// path. Keeps font loading (and the metric invalidation it triggers) out of the
// individual components.

import {
  awaitDocumentFonts,
  collectEmbeddedFontFaceDescriptors,
  documentFontTableFamilies,
} from "@arcships/docx-core"
import type { DocModel, OoxmlPackage } from "@arcships/docx-core"

/**
 * Ask the browser for every family the imported document uses and install any
 * embedded faces from `word/fontTable.xml`. Never rejects: a missing web font
 * must degrade to the fallback face, not fail the document load.
 */
export async function prepareDocxDocumentFonts(
  model: DocModel,
  pkg?: OoxmlPackage
): Promise<void> {
  try {
    const fontTableXml = pkg?.parts.get("word/fontTable.xml")?.content
    await awaitDocumentFonts(model, {
      extraFamilies: documentFontTableFamilies(fontTableXml),
      embeddedFonts: pkg ? collectEmbeddedFontFaceDescriptors(pkg) : undefined,
    })
  } catch {
    // The fallback face is still measurable; layout remains usable.
  }
}

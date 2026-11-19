// docx-file-types.ts — Shared accepted-file and exported-container policy.
//
// The file picker, drag-and-drop handler and importer must agree on which
// formats this package can actually parse. Word files are OOXML packages;
// OpenDocument text (`.odt` and its template/master/web variants plus flat
// `.fodt`) is read through the OpenDocument importer and is view-only because
// there is no ODF writer. Keep `.doc` (legacy binary Word) and ODF
// spreadsheet/presentation/drawing formats out of the list.

import { OPEN_DOCUMENT_VIEW_ONLY_REASON } from "@arcships/docx-core"

export { OPEN_DOCUMENT_VIEW_ONLY_REASON }

export const DOCX_FILE_EXTENSIONS = [".docx", ".docm", ".dotx", ".dotm"] as const

/** OpenDocument text formats accepted for view-only display. */
export const OPEN_DOCUMENT_FILE_EXTENSIONS = [
  ".odt",
  ".ott",
  ".odm",
  ".otm",
  ".oth",
  ".fodt",
] as const

export const DOCX_FILE_MIME_TYPES = [
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.ms-word.document.macroEnabled.12",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
  "application/vnd.ms-word.template.macroEnabled.12",
] as const

export const OPEN_DOCUMENT_FILE_MIME_TYPES = [
  "application/vnd.oasis.opendocument.text",
  "application/vnd.oasis.opendocument.text-template",
  "application/vnd.oasis.opendocument.text-master",
  "application/vnd.oasis.opendocument.text-master-template",
  "application/vnd.oasis.opendocument.text-web",
  "application/vnd.oasis.opendocument.text-flat-xml",
] as const

/** Every extension the DOCX viewer and editor accept. */
export const SUPPORTED_DOCUMENT_FILE_EXTENSIONS = [
  ...DOCX_FILE_EXTENSIONS,
  ...OPEN_DOCUMENT_FILE_EXTENSIONS,
] as const

/** Every MIME type the DOCX viewer and editor accept. */
export const SUPPORTED_DOCUMENT_FILE_MIME_TYPES = [
  ...DOCX_FILE_MIME_TYPES,
  ...OPEN_DOCUMENT_FILE_MIME_TYPES,
] as const

/** Value for an `<input type="file">` accept attribute. */
export const DOCX_FILE_ACCEPT = [
  ...SUPPORTED_DOCUMENT_FILE_EXTENSIONS,
  ...SUPPORTED_DOCUMENT_FILE_MIME_TYPES,
].join(",")

export const DOCX_UNSUPPORTED_FILE_MESSAGE =
  "Only Word (.docx, .docm, .dotx, .dotm) and OpenDocument text (.odt, .ott, .odm, .otm, .oth, .fodt) files are supported"

const DOCX_EXPORT_MIME_TYPES: Readonly<Record<string, string>> = Object.freeze({
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".docm": "application/vnd.ms-word.document.macroEnabled.12",
  ".dotx": "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
  ".dotm": "application/vnd.ms-word.template.macroEnabled.12",
})

/** Lower-case extension including the leading dot, or undefined. */
export function docxFileExtension(fileName: string): string | undefined {
  const match = /(\.[^.\\/]+)$/.exec(fileName.trim())
  return match?.[1].toLowerCase()
}

/** True for a Word OOXML file name. */
export function isWordFileName(fileName: string): boolean {
  const extension = docxFileExtension(fileName)
  return extension !== undefined &&
    (DOCX_FILE_EXTENSIONS as readonly string[]).includes(extension)
}

/** True for an OpenDocument text file name. */
export function isOpenDocumentTextFileName(fileName: string): boolean {
  const extension = docxFileExtension(fileName)
  return extension !== undefined &&
    (OPEN_DOCUMENT_FILE_EXTENSIONS as readonly string[]).includes(extension)
}

/**
 * File names the viewer opens. The parser decides the format from the bytes;
 * a `.docx` name holding OpenDocument bytes is still handled as OpenDocument.
 */
export function isSupportedDocumentFileName(fileName: string): boolean {
  return isWordFileName(fileName) || isOpenDocumentTextFileName(fileName)
}

/** Backwards-compatible name for {@link isSupportedDocumentFileName}. */
export function isSupportedDocxFileName(fileName: string): boolean {
  return isSupportedDocumentFileName(fileName)
}

/**
 * Downloaded name for an exported document. Word source extensions are
 * preserved so macro/template containers stay labelled correctly; anything
 * else falls back to the generic `.docx` export name. OpenDocument inputs
 * cannot be exported and never reach this helper.
 */
export function docxExportFileName(fileName: string): string {
  const trimmed = fileName.trim()
  const extension = docxFileExtension(trimmed)
  if (
    extension !== undefined &&
    (DOCX_FILE_EXTENSIONS as readonly string[]).includes(extension)
  ) {
    return `${trimmed.slice(0, -extension.length)}-edited${extension}`
  }
  return "edited.docx"
}

/** MIME type matching the exported container extension. */
export function docxExportMimeType(fileName: string): string {
  const extension = docxFileExtension(fileName)
  return extension !== undefined && DOCX_EXPORT_MIME_TYPES[extension] !== undefined
    ? DOCX_EXPORT_MIME_TYPES[extension]
    : DOCX_EXPORT_MIME_TYPES[".docx"]
}

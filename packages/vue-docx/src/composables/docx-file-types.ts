// docx-file-types.ts — Shared accepted-file and exported-container policy.
//
// The file picker, drag-and-drop handler and importer must agree on which Word
// formats this package can actually parse. Keep `.doc` (legacy binary Word)
// out of the list: the DOCX loader only understands OOXML packages.
//
// Exported containers keep their source type: a macro-enabled or template
// document must not be renamed to `.docx` while its macro part or template
// content type is still in the package.

export const DOCX_FILE_EXTENSIONS = [".docx", ".docm", ".dotx", ".dotm"] as const

export const DOCX_FILE_MIME_TYPES = [
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.ms-word.document.macroEnabled.12",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
  "application/vnd.ms-word.template.macroEnabled.12",
] as const

/** Value for an `<input type="file">` accept attribute. */
export const DOCX_FILE_ACCEPT = [...DOCX_FILE_EXTENSIONS, ...DOCX_FILE_MIME_TYPES].join(",")

export const DOCX_UNSUPPORTED_FILE_MESSAGE =
  "Only .docx, .docm, .dotx and .dotm files are supported"

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

export function isSupportedDocxFileName(fileName: string): boolean {
  const extension = docxFileExtension(fileName)
  return extension !== undefined &&
    (DOCX_FILE_EXTENSIONS as readonly string[]).includes(extension)
}

/**
 * Downloaded name for an exported document. Supported source extensions are
 * preserved so macro/template containers stay labelled correctly; anything
 * else falls back to the generic `.docx` export name.
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

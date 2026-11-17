// docx-file-types.ts — Shared accepted-file policy for DOCX import surfaces.
//
// The file picker, drag-and-drop handler and importer must agree on which Word
// formats this package can actually parse. Keep `.doc` (legacy binary Word)
// out of the list: the DOCX loader only understands OOXML packages.

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
 * replaced with `.docx`; anything else falls back to the generic export name.
 */
export function docxExportFileName(fileName: string): string {
  const trimmed = fileName.trim()
  const extension = docxFileExtension(trimmed)
  if (
    extension !== undefined &&
    (DOCX_FILE_EXTENSIONS as readonly string[]).includes(extension)
  ) {
    return `${trimmed.slice(0, -extension.length)}-edited.docx`
  }
  return "edited.docx"
}

// pptx-file-types.ts — Accepted presentation container policy shared by the
// file picker and drag-and-drop import.
//
// Presentations, slide shows and templates all share the OOXML presentation
// package; macro-enabled variants keep their own content type.

export const PPTX_FILE_EXTENSIONS = [
  ".pptx",
  ".pptm",
  ".ppsx",
  ".ppsm",
  ".potx",
  ".potm",
] as const

export const PPTX_FILE_MIME_TYPES = [
  "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  "application/vnd.ms-powerpoint.presentation.macroEnabled.12",
  "application/vnd.openxmlformats-officedocument.presentationml.slideshow",
  "application/vnd.ms-powerpoint.slideshow.macroEnabled.12",
  "application/vnd.openxmlformats-officedocument.presentationml.template",
  "application/vnd.ms-powerpoint.template.macroEnabled.12",
] as const

/** Value for an `<input type="file">` accept attribute. */
export const PPTX_FILE_ACCEPT = [...PPTX_FILE_EXTENSIONS, ...PPTX_FILE_MIME_TYPES].join(",")

/** Lower-case extension including the leading dot, or undefined. */
export function pptxFileExtension(fileName: string): string | undefined {
  const match = /(\.[^.\\/]+)$/.exec(fileName.trim())
  return match?.[1]?.toLowerCase()
}

export function isSupportedPptxFileName(fileName: string): boolean {
  const extension = pptxFileExtension(fileName)
  return extension !== undefined &&
    (PPTX_FILE_EXTENSIONS as readonly string[]).includes(extension)
}

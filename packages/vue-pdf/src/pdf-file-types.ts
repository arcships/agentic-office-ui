// pdf-file-types.ts — Accepted document policy shared by the file picker and
// drag-and-drop import.
//
// PDF is the only container the PDFium renderer opens; anything else is
// rejected before it reaches the loader.

export const PDF_FILE_EXTENSIONS = [".pdf"] as const

export const PDF_FILE_MIME_TYPES = ["application/pdf"] as const

/** Value for an `<input type="file">` accept attribute. */
export const PDF_FILE_ACCEPT = [...PDF_FILE_EXTENSIONS, ...PDF_FILE_MIME_TYPES].join(",")

/** Lower-case extension including the leading dot, or undefined. */
export function pdfFileExtension(fileName: string): string | undefined {
  const match = /(\.[^.\\/]+)$/.exec(fileName.trim())
  return match?.[1]?.toLowerCase()
}

export function isSupportedPdfFileName(fileName: string): boolean {
  const extension = pdfFileExtension(fileName)
  return extension !== undefined &&
    (PDF_FILE_EXTENSIONS as readonly string[]).includes(extension)
}

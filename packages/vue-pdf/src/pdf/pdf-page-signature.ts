import type { PdfRenderPageInfo } from "./pdf-render-runtime"

/**
 * Page signatures identify a page across document revisions. The base part
 * covers the rendered geometry; the optional content part covers the page text
 * so a page whose content was replaced at the same size is not treated as the
 * same page.
 */
export const PDF_PAGE_SIGNATURE_PREFIX = "pdf-page-v1:"
const PDF_PAGE_CONTENT_MARKER = ":text:"

/** FNV-1a over UTF-16 code units, with length included to avoid trivial collisions. */
export function fnv1a32(text: string): string {
  let hash = 0x811c9dc5
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index)
    hash ^= code & 0xff
    hash = Math.imul(hash, 0x01000193)
    hash ^= code >>> 8
    hash = Math.imul(hash, 0x01000193)
  }
  return `fnv1a32:${text.length}:${(hash >>> 0).toString(16).padStart(8, "0")}`
}

/** Digest used for page content signatures. */
export function pdfPageTextHash(text: string): string {
  return fnv1a32(text)
}

export function pdfPageBaseSignature(
  page: Pick<PdfRenderPageInfo, "width" | "height" | "rotation">,
): string {
  return `${PDF_PAGE_SIGNATURE_PREFIX}${page.width}:${page.height}:${page.rotation}`
}

export function pdfPageContentSignature(
  page: Pick<PdfRenderPageInfo, "width" | "height" | "rotation" | "contentHash">,
  contentHash: string | undefined = page.contentHash,
): string | undefined {
  if (!contentHash) return undefined
  return `${pdfPageBaseSignature(page)}${PDF_PAGE_CONTENT_MARKER}${contentHash}`
}

/**
 * Returns the base signature plus, when available, the content signature.
 * Keeping both lets old metadata-only references keep their previous meaning
 * while new references require an unchanged page text digest.
 */
export function pdfPageAncestorKeys(
  page: Pick<PdfRenderPageInfo, "width" | "height" | "rotation" | "contentHash">,
): string[] {
  const base = pdfPageBaseSignature(page)
  const content = pdfPageContentSignature(page)
  return content ? [base, content] : [base]
}

export function isPdfPageSignature(key: string): boolean {
  return key.startsWith(PDF_PAGE_SIGNATURE_PREFIX)
}

export function isPdfPageContentSignature(key: string): boolean {
  return key.startsWith(PDF_PAGE_SIGNATURE_PREFIX) && key.includes(PDF_PAGE_CONTENT_MARKER)
}

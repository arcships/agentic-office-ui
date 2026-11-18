import type { PptxFeatureRecord } from "./types"

/**
 * Right-to-left text in a presentation.
 *
 * The renderer applies `a:pPr/@rtl="1"` to the paragraph element as
 * `direction: rtl`, but it does not expose that decision on the DOM. These
 * helpers mirror the renderer's decision back to `dir="rtl"` so hosts can style
 * and hit-test mirrored paragraphs, and record the parts the renderer does not
 * mirror (multi-column text boxes) as a capability limitation.
 */

const TRUE_VALUES = new Set(["1", "true", "on"])

function elements(root: Document, localName: string): Element[] {
  return Array.from(root.getElementsByTagNameNS("*", localName))
}

function localAttr(node: Element | undefined, name: string): string | undefined {
  if (!node) return undefined
  return Array.from(node.attributes).find((attribute) => attribute.localName === name)?.value
}

function isTrue(value: string | undefined): boolean {
  return value !== undefined && TRUE_VALUES.has(value.trim().toLowerCase())
}

/**
 * Capability records for this slide's right-to-left text.
 *
 * Paragraph direction is strict: the renderer already paints those paragraphs
 * with `direction: rtl`, and `applyPptxParagraphDirection()` forwards it to the
 * DOM. Right-to-left column order (`a:bodyPr/@rtlCol`) is recorded as static:
 * the paragraphs mirror, but the renderer keeps the text-box columns in the
 * source order.
 */
export function collectPptxTextDirectionFeatures(
  document: Document,
  slideIndex: number,
): PptxFeatureRecord[] {
  const features: PptxFeatureRecord[] = []
  const rtlParagraphs = elements(document, "pPr").filter((node) => isTrue(localAttr(node, "rtl")))
  if (rtlParagraphs.length > 0) {
    features.push(Object.freeze({
      id: `text-rtl:${slideIndex}`,
      slideIndex,
      feature: "text:rtl-paragraph",
      disposition: "strict" as const,
    }))
  }
  const rtlColumns = elements(document, "bodyPr").filter((node) => isTrue(localAttr(node, "rtlCol")))
  if (rtlColumns.length > 0) {
    features.push(Object.freeze({
      id: `text-rtl-columns:${slideIndex}`,
      slideIndex,
      feature: "text:rtl-columns",
      disposition: "static" as const,
      reason: "文本框的多栏顺序尚未镜像，段落方向已按 rtl 渲染。",
    }))
  }
  return features
}

/**
 * Mirrors the renderer's paragraph direction onto the DOM.
 *
 * The renderer writes `style.direction` for every paragraph that carries
 * `a:pPr/@rtl`, including inherited list styles. Setting the standard `dir`
 * attribute keeps that decision queryable (CSS, selection and host overlays)
 * without re-parsing the slide XML. Returns the number of RTL paragraphs.
 */
export function applyPptxParagraphDirection(root: ParentNode): number {
  let mirrored = 0
  for (const paragraph of root.querySelectorAll<HTMLElement>("[data-pptx-paragraph-index]")) {
    const direction = paragraph.style?.direction?.trim().toLowerCase()
    if (direction !== "rtl" && direction !== "ltr") continue
    paragraph.setAttribute("dir", direction)
    paragraph.dataset.pptxParagraphDirection = direction
    if (direction === "rtl") mirrored += 1
  }
  return mirrored
}

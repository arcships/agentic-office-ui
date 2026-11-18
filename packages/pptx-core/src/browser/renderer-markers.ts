import { createPptxObjectKey } from "../playback/identity"
import { applyPptxParagraphDirection } from "../playback/text-direction"
import type {
  PptxObjectHyperlink,
  PptxObjectIdentity,
  PptxObjectSource,
} from "../playback/types"

export interface PptxRenderedObjectIndex {
  readonly elements: ReadonlyMap<string, HTMLElement>
  readonly duplicates: readonly string[]
}

function objectSource(partPath: string, slidePath: string): PptxObjectSource {
  if (/\/slideMasters\//u.test(partPath)) return "master"
  if (/\/slideLayouts\//u.test(partPath)) return "layout"
  return partPath === slidePath || /\/slides\//u.test(partPath) ? "slide" : "slide"
}

function groupPath(element: HTMLElement, root: HTMLElement): string[] {
  const result: string[] = []
  let current = element.parentElement
  while (current && current !== root) {
    if (current.dataset.pptxNodeType === "group" && current.dataset.pptxNodeId) {
      result.unshift(current.dataset.pptxNodeId)
    }
    current = current.parentElement
  }
  return result
}

/** The click link is the interactive affordance; a hover link only contributes its tooltip. */
function primaryHyperlink(hyperlinks: readonly PptxObjectHyperlink[]): PptxObjectHyperlink | undefined {
  return hyperlinks.find((hyperlink) => hyperlink.trigger === "click")
    ?? hyperlinks.find((hyperlink) => hyperlink.trigger === "hover")
}

/**
 * Mirror a shape's hyperlink onto its rendered element.
 *
 * PowerPoint shows the `tooltip` attribute (falling back to the URL) on hover,
 * underlines the text and makes the shape a link. The renderer paints the text
 * but cannot run the deck's actions, so the markers carry the same metadata to
 * the DOM: `title` for the tooltip, `data-pptx-hyperlink*` for hosts and
 * keyboard activation, and `role`/`tabindex` when a click link makes the shape
 * interactive.
 */
function applyHyperlinkMarker(element: HTMLElement, hyperlinks: readonly PptxObjectHyperlink[]): void {
  const hyperlink = primaryHyperlink(hyperlinks)
  if (!hyperlink) return
  const { dataset } = element
  dataset.pptxHyperlink = hyperlink.trigger
  dataset.pptxHyperlinkKind = hyperlink.kind
  if (hyperlink.tooltip !== undefined) dataset.pptxHyperlinkTooltip = hyperlink.tooltip
  if (hyperlink.url !== undefined) dataset.pptxHyperlinkUrl = hyperlink.url
  if (hyperlink.targetSlideIndex !== undefined) dataset.pptxHyperlinkSlide = String(hyperlink.targetSlideIndex)
  const label = hyperlink.tooltip ?? hyperlink.url
  // The renderer leaves `title` unset without its own navigation hook; never
  // clobber a label a host or the renderer already wrote.
  if (label && !element.getAttribute("title")) element.setAttribute("title", label)
  if (!element.getAttribute("aria-label") && !element.textContent?.trim() && label) {
    element.setAttribute("aria-label", label)
  }
  // Thumbnails paint the same marked slide inside a <button>; a nested link
  // role or tab stop there would be invalid and steal the thumbnail's focus.
  if (hyperlink.trigger !== "click" || insideThumbnail(element)) return
  if (!element.getAttribute("role")) element.setAttribute("role", "link")
  if (!element.hasAttribute("tabindex")) element.setAttribute("tabindex", "0")
}

function insideThumbnail(element: HTMLElement): boolean {
  if (element.parentElement?.dataset?.pptxThumbnail === "true") return true
  return Boolean(element.closest?.('[data-pptx-thumbnail="true"]'))
}

export function applyPptxObjectMarkers(
  root: HTMLElement,
  slidePath: string,
  objects?: readonly PptxObjectIdentity[],
): PptxRenderedObjectIndex {
  root.dataset.pptxSlidePath = slidePath
  const hyperlinksByKey = new Map<string, readonly PptxObjectHyperlink[]>()
  for (const object of objects ?? []) {
    if (object.hyperlinks?.length) hyperlinksByKey.set(object.key, object.hyperlinks)
  }
  const elements = new Map<string, HTMLElement>()
  const duplicates = new Set<string>()
  for (const element of root.querySelectorAll<HTMLElement>("[data-pptx-node-id]")) {
    const shapeId = element.dataset.pptxNodeId?.trim()
    if (!shapeId) continue
    const partPath = element.dataset.pptxPartPath?.trim() || slidePath
    const source = objectSource(partPath, slidePath)
    const groups = groupPath(element, root)
    const key = createPptxObjectKey({
      slidePath: partPath,
      source,
      shapeId,
      groupPath: groups,
    })
    element.dataset.pptxSlidePath = slidePath
    element.dataset.pptxSource = source
    element.dataset.pptxGroupPath = groups.join("/")
    element.dataset.pptxObjectKey = key
    const hyperlinks = hyperlinksByKey.get(key)
    if (hyperlinks) applyHyperlinkMarker(element, hyperlinks)
    if (elements.has(key)) {
      elements.delete(key)
      duplicates.add(key)
    } else if (!duplicates.has(key)) {
      elements.set(key, element)
    }
  }
  // The renderer already paints RTL paragraphs with `direction: rtl`; expose
  // that decision as `dir` so hosts and selection can follow it.
  applyPptxParagraphDirection(root)
  return {
    elements,
    duplicates: Object.freeze([...duplicates]),
  }
}

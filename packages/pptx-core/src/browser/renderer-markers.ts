import { createPptxObjectKey } from "../playback/identity"
import type { PptxObjectSource } from "../playback/types"

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

export function applyPptxObjectMarkers(
  root: HTMLElement,
  slidePath: string,
): PptxRenderedObjectIndex {
  root.dataset.pptxSlidePath = slidePath
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
    if (elements.has(key)) {
      elements.delete(key)
      duplicates.add(key)
    } else if (!duplicates.has(key)) {
      elements.set(key, element)
    }
  }
  return {
    elements,
    duplicates: Object.freeze([...duplicates]),
  }
}


import type { PptxObjectPropertyState, PptxObjectStateMap, PptxTrackValue } from "../playback/track-compiler"

function numericStyle(value: string, fallback: number): number {
  const number = Number.parseFloat(value)
  return Number.isFinite(number) ? number : fallback
}

function elementState(element: HTMLElement): Readonly<PptxObjectPropertyState> {
  const style = getComputedStyle(element)
  return Object.freeze({
    display: style.visibility !== "hidden" && style.display !== "none",
    opacity: numericStyle(style.opacity, 1),
    "translate-x": 0,
    "translate-y": 0,
    "scale-x": 1,
    "scale-y": 1,
    rotate: 0,
    "clip-path": style.clipPath || "none",
    filter: style.filter || "none",
  })
}

export function capturePptxStaticState(root: HTMLElement): PptxObjectStateMap {
  const result: Record<string, Readonly<PptxObjectPropertyState>> = {}
  const objects = new Map<string, HTMLElement[]>()
  for (const element of root.querySelectorAll<HTMLElement>("[data-pptx-object-key]")) {
    const objectKey = element.dataset.pptxObjectKey
    if (!objectKey) continue
    const list = objects.get(objectKey) ?? []
    list.push(element)
    objects.set(objectKey, list)
  }
  for (const [objectKey, elements] of objects) {
    if (elements.length !== 1) continue
    const element = elements[0]
    result[objectKey] = elementState(element)
    const paragraphs = new Map<string, HTMLElement[]>()
    for (const paragraph of element.querySelectorAll<HTMLElement>("[data-pptx-paragraph-index]")) {
      const index = paragraph.dataset.pptxParagraphIndex
      if (index === undefined) continue
      const list = paragraphs.get(index) ?? []
      list.push(paragraph)
      paragraphs.set(index, list)
    }
    for (const [index, matches] of paragraphs) {
      if (matches.length === 1) result[`${objectKey}#paragraph:${index}-${index}`] = elementState(matches[0])
    }
  }
  return Object.freeze(result)
}

function targetElements(root: HTMLElement): Map<string, HTMLElement[]> {
  const result = new Map<string, HTMLElement[]>()
  const objects = new Map<string, HTMLElement[]>()
  for (const element of root.querySelectorAll<HTMLElement>("[data-pptx-object-key]")) {
    const objectKey = element.dataset.pptxObjectKey
    if (!objectKey) continue
    const list = objects.get(objectKey) ?? []
    list.push(element)
    objects.set(objectKey, list)
  }
  for (const [objectKey, matches] of objects) {
    if (matches.length !== 1) continue
    const element = matches[0]
    result.set(objectKey, matches)
    const paragraphs = [...element.querySelectorAll<HTMLElement>("[data-pptx-paragraph-index]")]
    const paragraphCounts = new Map<string, number>()
    for (const paragraph of paragraphs) {
      const index = paragraph.dataset.pptxParagraphIndex
      if (index !== undefined) paragraphCounts.set(index, (paragraphCounts.get(index) ?? 0) + 1)
    }
    for (const paragraph of paragraphs) {
      const index = paragraph.dataset.pptxParagraphIndex
      if (index !== undefined && paragraphCounts.get(index) === 1) {
        result.set(`${objectKey}#paragraph:${index}-${index}`, [paragraph])
      }
    }
    const uniqueParagraphs = paragraphs.filter((paragraph) => {
      const index = paragraph.dataset.pptxParagraphIndex
      return index !== undefined && paragraphCounts.get(index) === 1
    }).sort((left, right) => Number(left.dataset.pptxParagraphIndex) - Number(right.dataset.pptxParagraphIndex))
    if (uniqueParagraphs.length) {
      for (let start = 0; start < uniqueParagraphs.length; start += 1) {
        for (let end = start + 1; end < uniqueParagraphs.length; end += 1) {
          const startIndex = uniqueParagraphs[start].dataset.pptxParagraphIndex
          const endIndex = uniqueParagraphs[end].dataset.pptxParagraphIndex
          result.set(`${objectKey}#paragraph:${startIndex}-${endIndex}`, uniqueParagraphs.slice(start, end + 1))
        }
      }
    }
  }
  return result
}

function cssValue(value: PptxTrackValue | undefined, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback
}

function applyElementState(element: HTMLElement, state: Readonly<PptxObjectPropertyState>): void {
  if (typeof state.display === "boolean") {
    element.style.visibility = state.display ? "visible" : "hidden"
    element.style.pointerEvents = state.display ? "" : "none"
  }
  if (typeof state.opacity === "number") element.style.opacity = String(state.opacity)
  const translateX = cssValue(state["translate-x"], 0)
  const translateY = cssValue(state["translate-y"], 0)
  if (state["translate-x"] !== undefined || state["translate-y"] !== undefined) {
    element.style.translate = `${translateX}px ${translateY}px`
  }
  const scaleX = cssValue(state["scale-x"], 1)
  const scaleY = cssValue(state["scale-y"], 1)
  if (state["scale-x"] !== undefined || state["scale-y"] !== undefined) {
    element.style.scale = `${scaleX} ${scaleY}`
  }
  if (typeof state.rotate === "number") element.style.rotate = `${state.rotate}deg`
  if (typeof state["clip-path"] === "string") element.style.clipPath = state["clip-path"]
  if (typeof state.filter === "string") element.style.filter = state.filter
  if (typeof state["fill-color"] === "string" && state["fill-color"]) element.style.backgroundColor = state["fill-color"]
  if (typeof state["line-color"] === "string" && state["line-color"]) element.style.borderColor = state["line-color"]
  if (typeof state["text-color"] === "string" && state["text-color"]) element.style.color = state["text-color"]
}

export function applyPptxObjectState(root: HTMLElement, state: PptxObjectStateMap): void {
  const elements = targetElements(root)
  for (const [target, values] of Object.entries(state)) {
    for (const element of elements.get(target) ?? []) applyElementState(element, values)
  }
}

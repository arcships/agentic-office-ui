import { pptxTableParagraphs } from "./table-targets"

const PARAGRAPH_SELECTOR = "[data-pptx-paragraph-index]"
const RUN_SELECTOR = "[data-pptx-text-run]"
const TEXT_NODE = 3

export interface PptxParagraphTextRuns {
  paragraph: HTMLElement
  /** Whole words wrapped as inline-block boxes. */
  words: readonly { element: HTMLElement; start: number; end: number }[]
  /** Individual characters wrapped inside the word boxes. */
  characters: readonly { element: HTMLElement; start: number; end: number }[]
  text: string
}

interface MutableRuns {
  words: { element: HTMLElement; start: number; end: number }[]
  characters: { element: HTMLElement; start: number; end: number }[]
}

function isTextNode(node: ChildNode): boolean {
  return node.nodeType === TEXT_NODE
}

function isRunElement(node: ChildNode): boolean {
  return !isTextNode(node)
    && (node as unknown as HTMLElement).dataset?.pptxTextRun !== undefined
}

function textNodes(root: HTMLElement): Text[] {
  const result: Text[] = []
  const visit = (node: ChildNode): void => {
    if (isTextNode(node)) {
      result.push(node as Text)
      return
    }
    // Do not descend into runs this module created earlier.
    if (isRunElement(node)) return
    for (const child of [...node.childNodes]) visit(child)
  }
  for (const child of [...root.childNodes]) visit(child)
  return result
}

function paragraphText(paragraph: HTMLElement): string {
  return paragraph.textContent ?? ""
}

function markRuns(element: HTMLElement, start: number, end: number): void {
  element.dataset.pptxRunStart = String(start)
  element.dataset.pptxRunEnd = String(end)
  element.style.display = "inline-block"
  if (element.dataset.pptxTextRun === "word") element.style.whiteSpace = "nowrap"
}

interface WrappedText {
  generated: ChildNode[]
  runs: MutableRuns
}

function wrapTextNode(node: Text, offset: number): WrappedText {
  const doc = node.ownerDocument ?? document
  const generated: ChildNode[] = []
  const runs: MutableRuns = { words: [], characters: [] }
  let cursor = offset
  for (const part of node.data.split(/(\s+)/u)) {
    if (!part) continue
    if (/^\s+$/u.test(part)) {
      const space = doc.createTextNode(part)
      generated.push(space)
      cursor += part.length
      continue
    }
    const word = doc.createElement("span")
    word.dataset.pptxTextRun = "word"
    markRuns(word, cursor, cursor + part.length - 1)
    for (const character of Array.from(part)) {
      const span = doc.createElement("span")
      span.dataset.pptxTextRun = "char"
      markRuns(span, cursor, cursor + character.length - 1)
      span.textContent = character
      word.appendChild(span)
      runs.characters.push({ element: span, start: cursor, end: cursor + character.length - 1 })
      cursor += character.length
    }
    generated.push(word)
    runs.words.push({ element: word, start: Number(word.dataset.pptxRunStart), end: Number(word.dataset.pptxRunEnd) })
  }
  return { generated, runs }
}

function existingRuns(paragraph: HTMLElement): PptxParagraphTextRuns {
  const runs: MutableRuns = { words: [], characters: [] }
  for (const element of paragraph.querySelectorAll<HTMLElement>(RUN_SELECTOR)) {
    const start = Number(element.dataset.pptxRunStart)
    const end = Number(element.dataset.pptxRunEnd)
    if (!Number.isFinite(start) || !Number.isFinite(end)) continue
    if (element.dataset.pptxTextRun === "word") runs.words.push({ element, start, end })
    else if (element.dataset.pptxTextRun === "char") runs.characters.push({ element, start, end })
  }
  return { paragraph, ...runs, text: paragraphText(paragraph) }
}

function replaceTextNode(node: Text, generated: readonly ChildNode[]): void {
  const parent = node.parentNode
  if (!parent || !generated.length) return
  for (const child of generated) parent.insertBefore(child, node)
  parent.removeChild(node)
}

/**
 * Splits a rendered paragraph into word boxes that contain character boxes,
 * once per paragraph. The boxes carry the paragraph-local character offsets
 * PowerPoint uses in `p:charRg`, so repeated effect application never has to
 * rebuild or measure the DOM.
 */
export function splitPptxParagraphRuns(paragraph: HTMLElement): PptxParagraphTextRuns {
  if (paragraph.dataset.pptxTextRuns === "split") return existingRuns(paragraph)
  const runs: MutableRuns = { words: [], characters: [] }
  let offset = 0
  for (const node of textNodes(paragraph)) {
    const wrapped = wrapTextNode(node, offset)
    replaceTextNode(node, wrapped.generated)
    runs.words.push(...wrapped.runs.words)
    runs.characters.push(...wrapped.runs.characters)
    offset += node.data.length
  }
  // A host may have pre-split some runs; keep them alongside the new boxes.
  const already = existingRuns(paragraph)
  const words = new Map(runs.words.map((word) => [word.element, word]))
  for (const word of already.words) words.set(word.element, word)
  runs.words = [...words.values()].sort((left, right) => left.start - right.start)
  const characters = new Map(runs.characters.map((item) => [item.element, item]))
  for (const item of already.characters) characters.set(item.element, item)
  runs.characters = [...characters.values()].sort((left, right) => left.start - right.start)
  paragraph.dataset.pptxTextRuns = "split"
  return {
    paragraph,
    words: Object.freeze(runs.words),
    characters: Object.freeze(runs.characters),
    text: paragraphText(paragraph),
  }
}

/**
 * All paragraphs that make up one object's text body. Tables flatten their
 * cells row by row, which is the same order PowerPoint uses for `pRg`.
 */
export function pptxObjectParagraphs(object: HTMLElement): readonly HTMLElement[] {
  const hasTable = object.localName === "table"
    || typeof object.querySelector === "function" && Boolean(object.querySelector("table"))
  if (hasTable) {
    const tableParagraphs = pptxTableParagraphs(object)
    if (tableParagraphs.length) return tableParagraphs
  }
  return Object.freeze([...object.querySelectorAll<HTMLElement>(PARAGRAPH_SELECTOR)])
}

export interface PptxTextBodyLayout {
  paragraphs: readonly HTMLElement[]
  /** Flat character offset where each paragraph starts. */
  starts: readonly number[]
  runs: readonly PptxParagraphTextRuns[]
  text: string
}

const textBodyLayouts = new WeakMap<HTMLElement, PptxTextBodyLayout>()

/** Splits every paragraph of one object and exposes the flat `charRg` space. */
export function layoutPptxTextBody(object: HTMLElement): PptxTextBodyLayout {
  const cached = textBodyLayouts.get(object)
  if (cached) return cached
  const paragraphs = pptxObjectParagraphs(object)
  const runs = paragraphs.map((paragraph) => splitPptxParagraphRuns(paragraph))
  const starts: number[] = []
  let offset = 0
  for (const paragraph of runs) {
    starts.push(offset)
    offset += paragraph.text.length + 1
  }
  const layout: PptxTextBodyLayout = {
    paragraphs,
    starts,
    runs,
    text: runs.map((paragraph) => paragraph.text).join("\r"),
  }
  textBodyLayouts.set(object, layout)
  return layout
}

/**
 * Resolves a flat `p:charRg` range to the boxes that represent it. A range that
 * covers whole words targets the word box, so text shaping and line breaking
 * stay word-like; partial ranges fall back to the character boxes.
 */
export function resolvePptxTextRange(
  object: HTMLElement,
  range: { start: number; end: number },
): readonly HTMLElement[] {
  const start = Math.max(0, Math.min(range.start, range.end))
  const end = Math.max(range.start, range.end)
  const layout = layoutPptxTextBody(object)
  const result: HTMLElement[] = []
  for (let index = 0; index < layout.runs.length; index += 1) {
    const paragraph = layout.runs[index]
    const paragraphStart = layout.starts[index]
    const paragraphEnd = paragraphStart + paragraph.text.length - 1
    if (paragraph.text.length === 0 || end < paragraphStart || start > paragraphEnd) continue
    const localStart = Math.max(0, start - paragraphStart)
    const localEnd = Math.min(paragraph.text.length - 1, end - paragraphStart)
    for (const word of paragraph.words) {
      if (localEnd < word.start || localStart > word.end) continue
      if (localStart <= word.start && localEnd >= word.end) {
        result.push(word.element)
        continue
      }
      for (const character of paragraph.characters) {
        if (character.start < word.start || character.end > word.end) continue
        if (localEnd < character.start || localStart > character.end) continue
        result.push(character.element)
      }
    }
  }
  return Object.freeze([...new Set(result)])
}

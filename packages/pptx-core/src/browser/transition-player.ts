import type { PptxMorphMatch, PptxSlideTransition } from "../playback/types"

interface PptxCapturedObjectGeometry {
  left: number
  top: number
  width: number
  height: number
}

export interface PptxTransitionCapture {
  layer: HTMLElement
  objects: ReadonlyMap<string, PptxCapturedObjectGeometry>
}

export interface PptxTransitionHandle {
  readonly durationMs: number
  apply(positionMs: number): void
  finish(): void
}

interface SavedStyle {
  element: HTMLElement
  opacity: string
  translate: string
  scale: string
  clipPath: string
  willChange: string
  position: string
  zIndex: string
}

function markedElements(root: ParentNode): Map<string, HTMLElement> {
  const result = new Map<string, HTMLElement>()
  for (const element of root.querySelectorAll<HTMLElement>("[data-pptx-object-key]")) {
    const key = element.dataset.pptxObjectKey
    if (key && !result.has(key)) result.set(key, element)
  }
  return result
}

function geometry(element: HTMLElement): PptxCapturedObjectGeometry {
  const rect = element.getBoundingClientRect()
  return { left: rect.left, top: rect.top, width: rect.width, height: rect.height }
}

export function capturePptxTransition(root: HTMLElement): PptxTransitionCapture | null {
  if (!root.firstElementChild) return null
  const layer = document.createElement("div")
  layer.dataset.pptxTransitionLayer = "outgoing"
  for (const child of [...root.children]) layer.appendChild(child.cloneNode(true))
  for (const media of layer.querySelectorAll<HTMLMediaElement>("audio,video")) {
    media.removeAttribute("src")
    for (const source of media.querySelectorAll("source")) source.removeAttribute("src")
  }
  const objects = new Map<string, PptxCapturedObjectGeometry>()
  for (const [key, element] of markedElements(root)) objects.set(key, geometry(element))
  return { layer, objects }
}

function directionVector(direction: string | undefined): { x: number; y: number } {
  switch ((direction ?? "l").toLowerCase()) {
    case "r":
    case "right": return { x: -1, y: 0 }
    case "u":
    case "up": return { x: 0, y: 1 }
    case "d":
    case "down": return { x: 0, y: -1 }
    default: return { x: 1, y: 0 }
  }
}

function wipeClip(direction: string | undefined, progress: number): string {
  const hidden = (1 - progress) * 100
  switch ((direction ?? "l").toLowerCase()) {
    case "r":
    case "right": return `inset(0 0 0 ${hidden}%)`
    case "u":
    case "up": return `inset(${hidden}% 0 0 0)`
    case "d":
    case "down": return `inset(0 0 ${hidden}% 0)`
    default: return `inset(0 ${hidden}% 0 0)`
  }
}

function saveStyle(element: HTMLElement): SavedStyle {
  return {
    element,
    opacity: element.style.opacity,
    translate: element.style.translate,
    scale: element.style.scale,
    clipPath: element.style.clipPath,
    willChange: element.style.willChange,
    position: element.style.position,
    zIndex: element.style.zIndex,
  }
}

function restoreStyle(saved: SavedStyle): void {
  const { element } = saved
  element.style.opacity = saved.opacity
  element.style.translate = saved.translate
  element.style.scale = saved.scale
  element.style.clipPath = saved.clipPath
  element.style.willChange = saved.willChange
  element.style.position = saved.position
  element.style.zIndex = saved.zIndex
}

export type PptxMorphGranularity = "object" | "word"

/**
 * Morph 的粒度来自 `<p:transition><p:morph option="..."/>`：`byWord`/`byChar` 逐词，
 * 其余（包括 `byObject` 和缺失）按对象整体执行。
 */
export function pptxMorphGranularity(transition: PptxSlideTransition | undefined): PptxMorphGranularity {
  const option = (transition?.option ?? "").trim().toLowerCase()
  return option === "byword" || option === "bychar" ? "word" : "object"
}

/** 强匹配与中等匹配都能执行，但都要求唯一；弱匹配只用于调试报告。 */
export function isPptxMorphMatchAnimatable(match: PptxMorphMatch): boolean {
  return match.unique && (match.confidence === "strong" || match.confidence === "medium")
}

export interface PptxMorphWordPair {
  from: number
  to: number
  text: string
}

export interface PptxMorphWordPoint {
  left: number
  top: number
}

/** 逐词 Morph 在某个进度上的位移：进度 0 时回到旧页位置，进度 1 时回到自己的位置。 */
export function pptxMorphWordTranslate(
  from: PptxMorphWordPoint,
  to: PptxMorphWordPoint,
  progress: number,
): { x: number; y: number } {
  const remaining = 1 - Math.max(0, Math.min(1, progress))
  return {
    x: (from.left - to.left) * remaining,
    y: (from.top - to.top) * remaining,
  }
}

/** 逐词 Morph 按词文本配对前后页文字，未在同侧重复使用的词才算一对。 */
export function matchPptxMorphWords(fromText: string, toText: string): readonly PptxMorphWordPair[] {
  const fromWords = fromText.match(/\S+/gu) ?? []
  const toWords = toText.match(/\S+/gu) ?? []
  const used = new Set<number>()
  const pairs: PptxMorphWordPair[] = []
  toWords.forEach((text, to) => {
    const from = fromWords.findIndex((word, index) => !used.has(index) && word === text)
    if (from < 0) return
    used.add(from)
    pairs.push({ from, to, text })
  })
  return Object.freeze(pairs)
}

interface WordSplit {
  words: HTMLElement[]
  restore(): void
}

/** `document.createTreeWalker` 的 SHOW_TEXT 位旗；避免在纯工具函数里依赖全局 NodeFilter。 */
const SHOW_TEXT = 4

/** 把元素内的每个词包成可变换的行内块，返回可撤销的拆分结果。 */
function splitWords(element: HTMLElement): WordSplit | undefined {
  const doc = element.ownerDocument
  const walker = doc.createTreeWalker(element, SHOW_TEXT)
  const textNodes: Text[] = []
  while (walker.nextNode()) {
    const node = walker.currentNode as Text
    if (/\S/u.test(node.data)) textNodes.push(node)
  }
  if (textNodes.length === 0) return undefined
  const records: Array<{ textNode: Text; generated: ChildNode[] }> = []
  const words: HTMLElement[] = []
  for (const textNode of textNodes) {
    const fragment = doc.createDocumentFragment()
    const generated: ChildNode[] = []
    for (const part of textNode.data.split(/(\s+)/u)) {
      if (!part) continue
      if (/^\s+$/u.test(part)) {
        const node = doc.createTextNode(part)
        fragment.appendChild(node)
        generated.push(node)
        continue
      }
      const span = doc.createElement("span")
      span.dataset.pptxMorphWord = ""
      span.style.display = "inline-block"
      span.textContent = part
      fragment.appendChild(span)
      generated.push(span)
      words.push(span)
    }
    textNode.replaceWith(fragment)
    records.push({ textNode, generated })
  }
  return {
    words,
    restore() {
      for (const record of records) {
        const first = record.generated[0]
        if (first?.parentNode) first.parentNode.insertBefore(record.textNode, first)
        for (const node of record.generated) node.remove()
      }
    },
  }
}

interface WordMorphEntry {
  element: HTMLElement
  from: PptxMorphWordPoint
  to: PptxMorphWordPoint
}

/** 逐词 Morph 的位移量：每个词从旧页同名词的位置回到新页自己的位置。 */
function wordMorphEntries(
  fromElement: HTMLElement | undefined,
  toElement: HTMLElement,
): { entries: WordMorphEntry[]; restore(): void } | undefined {
  const toSplit = splitWords(toElement)
  if (!toSplit) return undefined
  const fromSplit = fromElement ? splitWords(fromElement) : undefined
  if (!fromSplit) {
    toSplit.restore()
    return undefined
  }
  const pairs = matchPptxMorphWords(
    fromSplit.words.map((word) => word.textContent ?? "").join(" "),
    toSplit.words.map((word) => word.textContent ?? "").join(" "),
  )
  const entries = pairs.flatMap((pair): WordMorphEntry[] => {
    const fromWord = fromSplit.words[pair.from]
    const toWord = toSplit.words[pair.to]
    if (!fromWord || !toWord) return []
    const fromRect = fromWord.getBoundingClientRect()
    const toRect = toWord.getBoundingClientRect()
    return [{
      element: toWord,
      from: { left: fromRect.left, top: fromRect.top },
      to: { left: toRect.left, top: toRect.top },
    }]
  })
  if (entries.length === 0) {
    toSplit.restore()
    fromSplit.restore()
    return undefined
  }
  return {
    entries,
    restore() {
      toSplit.restore()
      fromSplit.restore()
    },
  }
}

export function mountPptxTransition(
  root: HTMLElement,
  capture: PptxTransitionCapture | null,
  transition: PptxSlideTransition | undefined,
  morphMatches: readonly PptxMorphMatch[] = [],
): PptxTransitionHandle | null {
  if (!capture || !transition || transition.kind === "cut" || transition.durationMs <= 0) return null
  const incoming = root.firstElementChild as HTMLElement | null
  if (!incoming) return null
  const rootStyle = saveStyle(root)
  const incomingStyle = saveStyle(incoming)
  const outgoing = capture.layer
  const outgoingStyle = saveStyle(outgoing)
  const rootRect = root.getBoundingClientRect()
  if (getComputedStyle(root).position === "static") root.style.position = "relative"
  incoming.style.position = "relative"
  incoming.style.zIndex = "1"
  incoming.style.willChange = "opacity, translate, clip-path"
  outgoing.style.position = "absolute"
  outgoing.style.inset = "0"
  outgoing.style.pointerEvents = "none"
  outgoing.style.zIndex = "2"
  outgoing.style.willChange = "opacity, translate"
  root.appendChild(outgoing)

  const incomingObjects = markedElements(incoming)
  const outgoingObjects = markedElements(outgoing)
  const granularity = pptxMorphGranularity(transition)
  const morphStyles: SavedStyle[] = []
  const morphGeometry: Array<{
    element: HTMLElement
    from: PptxCapturedObjectGeometry
    to: PptxCapturedObjectGeometry
  }> = []
  const wordGeometry: WordMorphEntry[] = []
  const wordRestores: Array<() => void> = []
  if (transition.kind.toLowerCase() === "morph") {
    for (const match of morphMatches) {
      if (!isPptxMorphMatchAnimatable(match)) continue
      const from = capture.objects.get(match.from)
      const element = incomingObjects.get(match.to)
      if (!from || !element) continue
      const to = geometry(element)
      if (granularity === "word") {
        const wordMorph = wordMorphEntries(outgoingObjects.get(match.from), element)
        if (wordMorph) {
          wordGeometry.push(...wordMorph.entries)
          wordRestores.push(wordMorph.restore)
          continue
        }
      }
      const style = saveStyle(element)
      morphStyles.push(style)
      element.style.willChange = "translate, scale, opacity"
      morphGeometry.push({ element, from, to })
    }
  }
  let finished = false

  const apply = (positionMs: number) => {
    if (finished) return
    const progress = Math.max(0, Math.min(1, positionMs / transition.durationMs))
    const kind = transition.kind.toLowerCase()
    if (kind === "fade" || kind === "morph") {
      incoming.style.opacity = String(progress)
      outgoing.style.opacity = String(1 - progress)
    } else if (kind === "push") {
      const vector = directionVector(transition.direction)
      incoming.style.translate = `${vector.x * rootRect.width * (1 - progress)}px ${vector.y * rootRect.height * (1 - progress)}px`
      outgoing.style.translate = `${-vector.x * rootRect.width * progress}px ${-vector.y * rootRect.height * progress}px`
    } else if (kind === "wipe") {
      incoming.style.clipPath = wipeClip(transition.direction, progress)
    } else {
      incoming.style.opacity = String(progress)
      outgoing.style.opacity = String(1 - progress)
    }
    if (kind === "morph") {
      for (const item of morphGeometry) {
        const scaleX = item.to.width > 0 ? item.from.width / item.to.width : 1
        const scaleY = item.to.height > 0 ? item.from.height / item.to.height : 1
        item.element.style.translate = `${(item.from.left - item.to.left) * (1 - progress)}px ${(item.from.top - item.to.top) * (1 - progress)}px`
        item.element.style.scale = `${1 + (scaleX - 1) * (1 - progress)} ${1 + (scaleY - 1) * (1 - progress)}`
      }
      for (const word of wordGeometry) {
        const offset = pptxMorphWordTranslate(word.from, word.to, progress)
        word.element.style.translate = `${offset.x}px ${offset.y}px`
      }
    }
  }
  const finish = () => {
    if (finished) return
    finished = true
    outgoing.remove()
    for (const restore of wordRestores) restore()
    restoreStyle(incomingStyle)
    restoreStyle(outgoingStyle)
    for (const saved of morphStyles) restoreStyle(saved)
    restoreStyle(rootStyle)
  }
  apply(0)
  return { durationMs: transition.durationMs, apply, finish }
}

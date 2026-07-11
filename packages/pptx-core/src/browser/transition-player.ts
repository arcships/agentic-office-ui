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
  const morphStyles: SavedStyle[] = []
  const morphGeometry = morphMatches.flatMap((match) => {
    const from = capture.objects.get(match.from)
    const element = incomingObjects.get(match.to)
    if (!from || !element || match.confidence !== "strong" || !match.unique) return []
    const to = geometry(element)
    const style = saveStyle(element)
    morphStyles.push(style)
    element.style.willChange = "translate, scale, opacity"
    return [{ element, from, to }]
  })
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
    }
  }
  const finish = () => {
    if (finished) return
    finished = true
    outgoing.remove()
    restoreStyle(incomingStyle)
    restoreStyle(outgoingStyle)
    for (const saved of morphStyles) restoreStyle(saved)
    restoreStyle(rootStyle)
  }
  apply(0)
  return { durationMs: transition.durationMs, apply, finish }
}

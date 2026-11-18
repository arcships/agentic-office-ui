import type { PptxObjectPropertyState, PptxObjectStateMap, PptxTrackValue } from "../playback/track-compiler"

/**
 * Element targets resolved from the rendered slide. The controller rebuilds this when a
 * slide is rendered or reset, so per-frame state application never queries the DOM.
 */
export interface PptxTrackTargets {
  readonly elements: ReadonlyMap<string, readonly HTMLElement[]>
  /** Representative element; once it is detached the slide was rendered again. */
  readonly probe?: HTMLElement
}

/** The element's own transform, which the animation channels are layered on top of. */
export interface PptxTransformBaseline {
  translateX: number
  translateY: number
  scaleX: number
  scaleY: number
  rotate: number
}

interface Matrix2D {
  a: number
  b: number
  c: number
  d: number
  e: number
  f: number
}

const identityMatrix: Matrix2D = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }

const identityBaseline: PptxTransformBaseline = {
  translateX: 0,
  translateY: 0,
  scaleX: 1,
  scaleY: 1,
  rotate: 0,
}

const transformBaselines = new WeakMap<HTMLElement, PptxTransformBaseline>()

function numericStyle(value: string, fallback: number): number {
  const number = Number.parseFloat(value)
  return Number.isFinite(number) ? number : fallback
}

function round(value: number): number {
  return Math.round(value * 1e4) / 1e4
}

function multiplyMatrix(left: Matrix2D, right: Matrix2D): Matrix2D {
  return {
    a: left.a * right.a + left.c * right.b,
    b: left.b * right.a + left.d * right.b,
    c: left.a * right.c + left.c * right.d,
    d: left.b * right.c + left.d * right.d,
    e: left.a * right.e + left.c * right.f + left.e,
    f: left.b * right.e + left.d * right.f + left.f,
  }
}

function functionMatrix(name: string, args: readonly number[]): Matrix2D | undefined {
  switch (name) {
    case "matrix":
      return args.length >= 6
        ? { a: args[0], b: args[1], c: args[2], d: args[3], e: args[4], f: args[5] }
        : undefined
    case "matrix3d":
      return args.length >= 16
        ? { a: args[0], b: args[1], c: args[4], d: args[5], e: args[12], f: args[13] }
        : undefined
    case "translate":
      return { a: 1, b: 0, c: 0, d: 1, e: args[0] ?? 0, f: args[1] ?? 0 }
    case "translatex":
      return { a: 1, b: 0, c: 0, d: 1, e: args[0] ?? 0, f: 0 }
    case "translatey":
      return { a: 1, b: 0, c: 0, d: 1, e: 0, f: args[0] ?? 0 }
    case "scale": {
      const x = args[0] ?? 1
      return { a: x, b: 0, c: 0, d: args[1] ?? x, e: 0, f: 0 }
    }
    case "scalex":
      return { a: args[0] ?? 1, b: 0, c: 0, d: 1, e: 0, f: 0 }
    case "scaley":
      return { a: 1, b: 0, c: 0, d: args[0] ?? 1, e: 0, f: 0 }
    case "rotate": {
      const radians = (args[0] ?? 0) * Math.PI / 180
      return {
        a: Math.cos(radians),
        b: Math.sin(radians),
        c: -Math.sin(radians),
        d: Math.cos(radians),
        e: 0,
        f: 0,
      }
    }
    default:
      return undefined
  }
}

function parseTransformMatrix(value: string): Matrix2D {
  let matrix = identityMatrix
  for (const match of value.matchAll(/([a-z0-9]+)\(([^)]*)\)/giu)) {
    const name = match[1].toLowerCase()
    const args = [...match[2].matchAll(/[+-]?(?:\d+(?:\.\d+)?|\.\d+)(?:e[+-]?\d+)?/giu)]
      .map((item) => Number(item[0]))
      .filter((item) => Number.isFinite(item))
    const next = functionMatrix(name, args)
    if (next) matrix = multiplyMatrix(matrix, next)
  }
  return matrix
}

function parseAngle(value: string): number {
  const match = value.match(/([+-]?(?:\d+(?:\.\d+)?|\.\d+))\s*(deg|rad|turn|grad)?/iu)
  if (!match) return 0
  const amount = Number(match[1])
  if (!Number.isFinite(amount)) return 0
  switch ((match[2] ?? "deg").toLowerCase()) {
    case "rad": return amount * 180 / Math.PI
    case "turn": return amount * 360
    case "grad": return amount * 0.9
    default: return amount
  }
}

function parseNumbers(value: string): number[] {
  return [...value.matchAll(/[+-]?(?:\d+(?:\.\d+)?|\.\d+)(?:e[+-]?\d+)?/giu)]
    .map((match) => Number(match[0]))
    .filter((item) => Number.isFinite(item))
}

function isUnset(value: string | undefined): boolean {
  return !value || value === "none"
}

/**
 * Reads the transform the renderer already applied to the element. Animations write the
 * individual translate/scale/rotate properties, which compose with `transform`, so this
 * baseline is what keeps those channels relative to the rendered rest state.
 */
function transformBaseline(style: CSSStyleDeclaration): PptxTransformBaseline {
  const matrix = parseTransformMatrix(style.transform ?? "")
  const scaleX = Math.hypot(matrix.a, matrix.b)
  const determinant = matrix.a * matrix.d - matrix.b * matrix.c
  const baseline: PptxTransformBaseline = {
    translateX: matrix.e,
    translateY: matrix.f,
    scaleX,
    scaleY: scaleX > 0 ? determinant / scaleX : Math.hypot(matrix.c, matrix.d),
    rotate: Math.atan2(matrix.b, matrix.a) * 180 / Math.PI,
  }
  if (!isUnset(style.translate)) {
    const values = parseNumbers(style.translate)
    baseline.translateX += values[0] ?? 0
    baseline.translateY += values[1] ?? 0
  }
  if (!isUnset(style.scale)) {
    const values = parseNumbers(style.scale)
    const x = values[0] ?? 1
    baseline.scaleX *= x
    baseline.scaleY *= values[1] ?? x
  }
  if (!isUnset(style.rotate)) baseline.rotate += parseAngle(style.rotate)
  return baseline
}

function elementState(element: HTMLElement): Readonly<PptxObjectPropertyState> {
  const style = getComputedStyle(element)
  const baseline = transformBaseline(style)
  transformBaselines.set(element, baseline)
  return Object.freeze({
    display: style.visibility !== "hidden" && style.display !== "none",
    opacity: numericStyle(style.opacity, 1),
    "translate-x": round(baseline.translateX),
    "translate-y": round(baseline.translateY),
    "scale-x": round(baseline.scaleX),
    "scale-y": round(baseline.scaleY),
    rotate: round(baseline.rotate),
    "clip-path": style.clipPath || "none",
    filter: style.filter || "none",
  })
}

/**
 * Resolves every animation target key once per render: shapes by object key, then
 * paragraph runs and paragraph ranges. Duplicate keys keep all of their elements.
 */
export function createPptxTrackTargets(root: HTMLElement): PptxTrackTargets {
  const objects = new Map<string, HTMLElement[]>()
  for (const element of root.querySelectorAll<HTMLElement>("[data-pptx-object-key]")) {
    const objectKey = element.dataset.pptxObjectKey
    if (!objectKey) continue
    const list = objects.get(objectKey) ?? []
    list.push(element)
    objects.set(objectKey, list)
  }
  const elements = new Map<string, readonly HTMLElement[]>()
  for (const [objectKey, matches] of objects) {
    elements.set(objectKey, Object.freeze([...matches]))
    const paragraphs = new Map<number, HTMLElement[]>()
    for (const element of matches) {
      for (const paragraph of element.querySelectorAll<HTMLElement>("[data-pptx-paragraph-index]")) {
        const index = Number(paragraph.dataset.pptxParagraphIndex)
        if (!Number.isInteger(index)) continue
        const list = paragraphs.get(index) ?? []
        list.push(paragraph)
        paragraphs.set(index, list)
      }
    }
    const indices = [...paragraphs.keys()].sort((left, right) => left - right)
    for (const index of indices) {
      elements.set(`${objectKey}#paragraph:${index}-${index}`, Object.freeze(paragraphs.get(index) ?? []))
    }
    for (let start = 0; start < indices.length; start += 1) {
      for (let end = start + 1; end < indices.length; end += 1) {
        const range: HTMLElement[] = []
        for (let index = start; index <= end; index += 1) {
          range.push(...(paragraphs.get(indices[index]) ?? []))
        }
        elements.set(`${objectKey}#paragraph:${indices[start]}-${indices[end]}`, Object.freeze(range))
      }
    }
  }
  return Object.freeze({ elements, probe: elements.values().next().value?.[0] })
}

function currentTargets(root: HTMLElement, targets: PptxTrackTargets | undefined): PptxTrackTargets {
  if (targets && (targets.probe === undefined || targets.probe.isConnected !== false)) return targets
  return createPptxTrackTargets(root)
}

export function capturePptxStaticState(
  root: HTMLElement,
  targets?: PptxTrackTargets,
): PptxObjectStateMap {
  const index = currentTargets(root, targets)
  const result: Record<string, Readonly<PptxObjectPropertyState>> = {}
  const states = new Map<HTMLElement, Readonly<PptxObjectPropertyState>>()
  const stateFor = (element: HTMLElement): Readonly<PptxObjectPropertyState> => {
    const cached = states.get(element)
    if (cached) return cached
    const state = elementState(element)
    states.set(element, state)
    return state
  }
  for (const [target, elements] of index.elements) {
    if (!elements.length) continue
    const paragraphRange = /#paragraph:(\d+)-(\d+)$/u.exec(target)
    if (paragraphRange && paragraphRange[1] !== paragraphRange[2]) {
      // A combined range is a lookup key for range effects; its static state is
      // the leaf paragraphs' state. Keep capturing element baselines, but do not
      // replay a representative leaf state that would undo leaf-level hiding.
      for (const element of elements) stateFor(element)
      result[target] = Object.freeze({})
      continue
    }
    for (const element of elements) stateFor(element)
    // Duplicate keys share one state; the first rendered element settles it.
    result[target] = stateFor(elements[0])
  }
  return Object.freeze(result)
}

function cssValue(value: PptxTrackValue | undefined, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback
}

function ratio(value: number, baseline: number): number {
  return baseline === 0 ? value : value / baseline
}

function applyElementState(element: HTMLElement, state: Readonly<PptxObjectPropertyState>): void {
  const baseline = transformBaselines.get(element) ?? identityBaseline
  if (typeof state.display === "boolean") {
    element.style.visibility = state.display ? "visible" : "hidden"
    element.style.pointerEvents = state.display ? "" : "none"
  }
  if (typeof state.opacity === "number") element.style.opacity = String(state.opacity)
  if (state["translate-x"] !== undefined || state["translate-y"] !== undefined) {
    const translateX = cssValue(state["translate-x"], baseline.translateX) - baseline.translateX
    const translateY = cssValue(state["translate-y"], baseline.translateY) - baseline.translateY
    element.style.translate = `${round(translateX)}px ${round(translateY)}px`
  }
  if (state["scale-x"] !== undefined || state["scale-y"] !== undefined) {
    const scaleX = ratio(cssValue(state["scale-x"], baseline.scaleX), baseline.scaleX)
    const scaleY = ratio(cssValue(state["scale-y"], baseline.scaleY), baseline.scaleY)
    element.style.scale = `${round(scaleX)} ${round(scaleY)}`
  }
  if (typeof state.rotate === "number") {
    element.style.rotate = `${round(state.rotate - baseline.rotate)}deg`
  }
  if (typeof state["clip-path"] === "string") element.style.clipPath = state["clip-path"]
  if (typeof state.filter === "string") element.style.filter = state.filter
  if (typeof state["fill-color"] === "string" && state["fill-color"]) element.style.backgroundColor = state["fill-color"]
  if (typeof state["line-color"] === "string" && state["line-color"]) element.style.borderColor = state["line-color"]
  if (typeof state["text-color"] === "string" && state["text-color"]) element.style.color = state["text-color"]
}

const paragraphRangeTarget = /#paragraph:(\d+)-(\d+)$/u

/** 应用顺序：对象级 → 段落范围（从宽到窄）→ 单段，具体的状态后写。 */
function targetOrder(target: string): number {
  const match = paragraphRangeTarget.exec(target)
  return match ? 1_000_000 - (Number(match[2]) - Number(match[1])) : -1
}

export function applyPptxObjectState(
  root: HTMLElement,
  state: PptxObjectStateMap,
  targets?: PptxTrackTargets,
): void {
  const index = currentTargets(root, targets)
  const entries = Object.entries(state)
  // 范围状态是聚合，不能盖掉已经按单段准备好的时间零状态。
  const ordered = entries.some(([target]) => targetOrder(target) > 0)
    ? [...entries].sort((left, right) => targetOrder(left[0]) - targetOrder(right[0]))
    : entries
  for (const [target, values] of ordered) {
    for (const element of index.elements.get(target) ?? []) applyElementState(element, values)
  }
}

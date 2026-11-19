import type { NormalizedRect } from "@arcships/office-interaction"
import { OfficeRegionCaptureError } from "./region-capture"

const SKIPPED_TAGS = new Set(["SCRIPT", "IFRAME", "OBJECT", "EMBED", "LINK", "META", "NOSCRIPT", "STYLE"])
const SKIPPED_STYLE_PROPERTIES = new Set([
  "animation",
  "animation-delay",
  "animation-direction",
  "animation-duration",
  "animation-fill-mode",
  "animation-iteration-count",
  "animation-name",
  "animation-play-state",
  "animation-timing-function",
  "cursor",
  "pointer-events",
  "transition",
  "transition-delay",
  "transition-duration",
  "transition-property",
  "transition-timing-function",
  "will-change",
])
const SKIPPED_ATTRIBUTES = new Set(["srcset", "sizes", "loading", "decoding", "fetchpriority", "crossorigin"])

export interface RasterizeElementRegionOptions {
  element: Element
  region: NormalizedRect
  outputWidth: number
  outputHeight: number
  background?: string | null
  signal?: AbortSignal
  createCanvas?: (width: number, height: number) => HTMLCanvasElement
  /** Reuse a canvas and context already allocated by `captureRegion`. */
  drawInto?: { canvas: HTMLCanvasElement; context: CanvasRenderingContext2D }
}

interface LayoutMetrics {
  width: number
  height: number
  viewX: number
  viewY: number
  viewWidth: number
  viewHeight: number
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return
  throw new OfficeRegionCaptureError("ABORTED", "region capture was aborted")
}

function layoutMetrics(element: Element, region: NormalizedRect): LayoutMetrics {
  const bounds = element.getBoundingClientRect()
  const elementWithOffsets = element as Element & { offsetWidth?: number; offsetHeight?: number }
  const width = elementWithOffsets.offsetWidth && elementWithOffsets.offsetWidth > 0
    ? elementWithOffsets.offsetWidth
    : bounds.width
  const height = elementWithOffsets.offsetHeight && elementWithOffsets.offsetHeight > 0
    ? elementWithOffsets.offsetHeight
    : bounds.height
  if (!(width > 0) || !(height > 0)) {
    throw new OfficeRegionCaptureError("CAPTURE_UNSUPPORTED", "capture element is not laid out")
  }
  return {
    width,
    height,
    viewX: region.x * width,
    viewY: region.y * height,
    viewWidth: region.width * width,
    viewHeight: region.height * height,
  }
}

function copyComputedStyle(
  view: Window & typeof globalThis,
  original: Element,
  clone: Element,
  isRoot: boolean,
): void {
  const computed = view.getComputedStyle(original)
  const style = (clone as HTMLElement).style
  if (!style) return
  for (let index = 0; index < computed.length; index += 1) {
    const property = computed.item(index)
    if (!property || SKIPPED_STYLE_PROPERTIES.has(property)) continue
    const value = computed.getPropertyValue(property)
    if (!value) continue
    style.setProperty(property, value)
  }
  if (isRoot) {
    // The clone is placed at (0, 0) inside a foreignObject. Absolute offsets and
    // the wrapper's own translate would move it out of the viewBox.
    style.setProperty("position", "relative")
    style.setProperty("inset", "auto")
    style.setProperty("top", "0px")
    style.setProperty("right", "auto")
    style.setProperty("bottom", "auto")
    style.setProperty("left", "0px")
    style.setProperty("margin", "0px")
    style.setProperty("float", "none")
    style.setProperty("transform", "none")
    style.setProperty("transform-origin", "0 0")
    style.setProperty("zoom", "1")
  }
}

function copyAttributes(original: Element, clone: Element): void {
  for (const attribute of Array.from(original.attributes)) {
    if (SKIPPED_ATTRIBUTES.has(attribute.name.toLowerCase())) continue
    if (attribute.name === "style") continue
    try {
      clone.setAttribute(attribute.name, attribute.value)
    } catch {
      // Some reflected attributes (for example `value` on certain inputs) are
      // read-only in XML; copying the value explicitly below is enough.
    }
  }
}

function imageDataUrl(view: Window & typeof globalThis, image: HTMLImageElement): string {
  const source = image.currentSrc || image.src
  if (!source || source.startsWith("data:")) return source
  if (!image.complete || image.naturalWidth <= 0) return source
  try {
    const canvas = view.document.createElement("canvas")
    canvas.width = image.naturalWidth
    canvas.height = image.naturalHeight
    const context = canvas.getContext("2d")
    if (!context) return source
    context.drawImage(image, 0, 0)
    return canvas.toDataURL("image/png")
  } catch {
    // Cross-origin images taint the helper canvas. Keep the original URL; the
    // SVG image renderer may still resolve same-origin blobs and data URLs.
    return source
  }
}

function canvasDataUrl(canvas: HTMLCanvasElement): string | null {
  try {
    return canvas.toDataURL("image/png")
  } catch {
    return null
  }
}

function cloneForRaster(view: Window & typeof globalThis, original: Element, isRoot: boolean): Element | null {
  if (SKIPPED_TAGS.has(original.tagName.toUpperCase())) return null
  let clone: Element
  if (typeof HTMLCanvasElement !== "undefined" && original instanceof HTMLCanvasElement) {
    const source = canvasDataUrl(original)
    if (source) {
      const image = view.document.createElement("img")
      image.src = source
      clone = image
    } else {
      clone = original.cloneNode(false) as Element
    }
  } else if (typeof HTMLImageElement !== "undefined" && original instanceof HTMLImageElement) {
    const image = view.document.createElement("img")
    const source = imageDataUrl(view, original)
    if (source) image.setAttribute("src", source)
    clone = image
  } else if (typeof HTMLInputElement !== "undefined" && original instanceof HTMLInputElement) {
    const input = original.cloneNode(false) as HTMLInputElement
    input.value = original.value
    input.checked = original.checked
    clone = input
  } else if (typeof HTMLTextAreaElement !== "undefined" && original instanceof HTMLTextAreaElement) {
    const textarea = original.cloneNode(false) as HTMLTextAreaElement
    textarea.value = original.value
    textarea.textContent = original.value
    clone = textarea
  } else {
    clone = original.cloneNode(false) as Element
  }
  copyAttributes(original, clone)
  copyComputedStyle(view, original, clone, isRoot)
  for (const child of Array.from(original.childNodes)) {
    if (child.nodeType === 3) {
      clone.appendChild(view.document.createTextNode(child.textContent ?? ""))
    } else if (child.nodeType === 1) {
      const childClone = cloneForRaster(view, child as Element, false)
      if (childClone) clone.appendChild(childClone)
    }
  }
  return clone
}

function resolveBackground(view: Window & typeof globalThis, element: Element, explicit: string | null | undefined): string | null {
  if (explicit !== undefined) return explicit
  const computed = view.getComputedStyle(element)
  return computed.backgroundColor && computed.backgroundColor !== "rgba(0, 0, 0, 0)"
    ? computed.backgroundColor
    : null
}

async function loadSvgImage(view: Window & typeof globalThis, svg: string): Promise<HTMLImageElement> {
  if (typeof view.Image !== "function") {
    throw new OfficeRegionCaptureError("CAPTURE_UNSUPPORTED", "Image constructor is unavailable")
  }
  const image = new view.Image()
  image.decoding = "sync"
  const url = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`
  await new Promise<void>((resolve, reject) => {
    image.onload = () => resolve()
    image.onerror = () => reject(new OfficeRegionCaptureError(
      "CAPTURE_FAILED",
      "browser could not rasterize the surface DOM",
    ))
    image.src = url
  })
  try {
    await image.decode?.()
  } catch {
    // Some browsers reject decode() for SVG images that already fired `load`.
  }
  return image
}

function createCanvasFor(
  options: RasterizeElementRegionOptions,
  width: number,
  height: number,
): HTMLCanvasElement {
  if (options.createCanvas) return options.createCanvas(width, height)
  if (typeof document !== "undefined" && typeof document.createElement === "function") {
    const canvas = document.createElement("canvas")
    canvas.width = width
    canvas.height = height
    return canvas
  }
  const OffscreenCanvasConstructor = (globalThis as { OffscreenCanvas?: typeof OffscreenCanvas }).OffscreenCanvas
  if (typeof OffscreenCanvasConstructor === "function") {
    return new OffscreenCanvasConstructor(width, height) as unknown as HTMLCanvasElement
  }
  throw new OfficeRegionCaptureError("CAPTURE_UNSUPPORTED", "canvas is unavailable for DOM capture")
}

/**
 * Rasterize a normalized region of a live DOM element into a canvas.
 *
 * This is a dependency-free DOM-to-bitmap fallback: computed styles are
 * inlined, canvas children are converted to data URLs, and same-origin images
 * are inlined when the browser allows it. Cross-origin images, external
 * background images, web fonts that are not installed, and pseudo-elements are
 * not embedded; canvas-backed surfaces should pass a `raster` source instead.
 */
export async function rasterizeElementRegion(
  options: RasterizeElementRegionOptions,
): Promise<HTMLCanvasElement> {
  throwIfAborted(options.signal)
  const ownerDocument = options.element.ownerDocument
  const view = ownerDocument?.defaultView
  if (!ownerDocument || !view || typeof view.Image !== "function") {
    throw new OfficeRegionCaptureError(
      "CAPTURE_UNSUPPORTED",
      "DOM capture requires a browser document with an Image constructor",
    )
  }
  try {
    await view.document.fonts?.ready
  } catch {
    // Font loading is best-effort; continue with whatever metrics are ready.
  }
  throwIfAborted(options.signal)
  const layout = layoutMetrics(options.element, options.region)
  const clone = cloneForRaster(view, options.element, true)
  if (!clone) throw new OfficeRegionCaptureError("CAPTURE_UNSUPPORTED", "surface element cannot be cloned")
  clone.setAttribute("xmlns", "http://www.w3.org/1999/xhtml")
  if (typeof XMLSerializer === "undefined") {
    throw new OfficeRegionCaptureError("CAPTURE_UNSUPPORTED", "XMLSerializer is unavailable for DOM capture")
  }
  const serialized = new XMLSerializer().serializeToString(clone)
  const svg = [
    `<svg xmlns="http://www.w3.org/2000/svg" width="${options.outputWidth}" height="${options.outputHeight}"`,
    ` viewBox="${layout.viewX} ${layout.viewY} ${layout.viewWidth} ${layout.viewHeight}">`,
    `<foreignObject x="0" y="0" width="${layout.width}" height="${layout.height}">`,
    serialized,
    "</foreignObject></svg>",
  ].join("")
  const image = await loadSvgImage(view, svg)
  throwIfAborted(options.signal)
  const canvas = options.drawInto?.canvas
    ?? createCanvasFor(options, options.outputWidth, options.outputHeight)
  if (!canvas) throw new OfficeRegionCaptureError("CAPTURE_UNSUPPORTED", "canvas is unavailable for DOM capture")
  if (canvas.width !== options.outputWidth) canvas.width = options.outputWidth
  if (canvas.height !== options.outputHeight) canvas.height = options.outputHeight
  const context = options.drawInto?.context ?? canvas.getContext("2d") as CanvasRenderingContext2D | null
  if (!context) throw new OfficeRegionCaptureError("CAPTURE_UNSUPPORTED", "2D canvas context is unavailable")
  const background = resolveBackground(view, options.element, options.background)
  if (background) {
    context.save()
    context.fillStyle = background
    context.fillRect(0, 0, options.outputWidth, options.outputHeight)
    context.restore()
  }
  context.drawImage(image, 0, 0, options.outputWidth, options.outputHeight)
  return canvas
}

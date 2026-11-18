// Reusable page-geometry state for the block/word overlay components.
//
// The geometry is measured from the mounted viewer DOM, so it has to be
// refreshed when the viewer repaginates, zooms, loads fonts or resizes. All
// observers are optional: the component remains usable in environments without
// ResizeObserver/MutationObserver (SSR, unit-test shims).

import { onBeforeUnmount, onMounted, ref, watch, type Ref } from "vue"
import {
  resolveDocxBlockGeometries,
  type DocxBlockGeometry,
  type DocumentBlock,
} from "@arcships/docx-core"

export interface DocxBlockGeometryState {
  geometries: Ref<Map<string, DocxBlockGeometry>>
  surfaces: Ref<Map<number, HTMLElement>>
  refresh: () => void
}

export interface UseDocxBlockGeometryOptions {
  root: () => HTMLElement | null | undefined
  blocks: () => readonly DocumentBlock[]
  onGeometry?: (geometries: Map<string, DocxBlockGeometry>) => void
}

function collectSurfaces(root: HTMLElement | null | undefined): Map<number, HTMLElement> {
  const surfaces = new Map<number, HTMLElement>()
  if (!root) return surfaces
  for (const surface of root.querySelectorAll<HTMLElement>("[data-docx-page-surface]")) {
    const index = surface.closest<HTMLElement>("[data-docx-page-index]")?.dataset
      .docxPageIndex
    if (index !== undefined) surfaces.set(Number(index), surface)
  }
  return surfaces
}

function geometrySignature(geometries: Map<string, DocxBlockGeometry>): string {
  return [...geometries]
    .map(([id, geometry]) =>
      [
        id,
        geometry.status,
        geometry.issues.map((issue) => `${issue.nodeId}:${issue.reason}`).join(","),
        geometry.fragments
          .map(
            (fragment) =>
              `${fragment.nodeId}@${fragment.pageIndex}:${fragment.boundingBox.left},${fragment.boundingBox.top},${fragment.boundingBox.right},${fragment.boundingBox.bottom}`,
          )
          .join(";"),
      ].join("|"),
    )
    .join("\n")
}

/**
 * Measures the supplied blocks against the mounted viewer DOM and keeps the
 * result fresh while pages mount, repaginate, zoom or resize.
 */
export function useDocxBlockGeometry(
  options: UseDocxBlockGeometryOptions,
): DocxBlockGeometryState {
  const geometries = ref<Map<string, DocxBlockGeometry>>(new Map())
  const surfaces = ref<Map<number, HTMLElement>>(new Map())
  let lastSignature = ""
  let frame: number | ReturnType<typeof setTimeout> | undefined
  let disposed = false
  let mutationObserver: MutationObserver | undefined
  let resizeObserver: ResizeObserver | undefined

  const update = (): void => {
    frame = undefined
    if (disposed) return
    const root = options.root()
    const blocks = options.blocks()
    if (!root || blocks.length === 0) {
      if (
        lastSignature !== "" ||
        geometries.value.size > 0 ||
        surfaces.value.size > 0
      ) {
        lastSignature = ""
        geometries.value = new Map()
        surfaces.value = new Map()
      }
      return
    }
    const next = resolveDocxBlockGeometries(root, blocks)
    const signature = geometrySignature(next)
    const nextSurfaces = collectSurfaces(root)
    if (
      nextSurfaces.size !== surfaces.value.size ||
      [...nextSurfaces].some(([index, surface]) => surfaces.value.get(index) !== surface)
    ) {
      surfaces.value = nextSurfaces
    }
    if (signature !== lastSignature) {
      lastSignature = signature
      geometries.value = next
    }
    options.onGeometry?.(next)
  }

  const schedule = (): void => {
    if (disposed || frame !== undefined) return
    if (typeof requestAnimationFrame === "function") {
      frame = requestAnimationFrame(update)
    } else {
      frame = setTimeout(update, 0)
    }
  }

  const cancelScheduled = (): void => {
    if (frame === undefined) return
    if (typeof cancelAnimationFrame === "function") cancelAnimationFrame(frame as number)
    else clearTimeout(frame as ReturnType<typeof setTimeout>)
    frame = undefined
  }

  const ignoredNode = (node: unknown): boolean => {
    const element = node as {
      nodeType?: number
      matches?: (selector: string) => boolean
      closest?: (selector: string) => unknown
    } | null
    if (!element) return false
    // Teleport anchors are comment nodes; they never change source geometry.
    if (element.nodeType === 8) return true
    if (typeof element.matches !== "function") return false
    if (element.matches("[data-docx-block-highlight],[data-docx-geometry-measure]")) {
      return true
    }
    return Boolean(
      typeof element.closest === "function" &&
        element.closest("[data-docx-block-highlight],[data-docx-geometry-measure]"),
    )
  }

  const startObservers = (): void => {
    const root = options.root()
    if (!root) return
    if (typeof MutationObserver === "function") {
      mutationObserver = new MutationObserver((records) => {
        const relevant = records.some((record) => {
          if (ignoredNode(record.target)) return false
          if (record.type === "attributes" || record.type === "characterData") return true
          return [...record.addedNodes, ...record.removedNodes].some(
            (node) => !ignoredNode(node),
          )
        })
        if (relevant) schedule()
      })
      mutationObserver.observe(root, {
        childList: true,
        characterData: true,
        subtree: true,
        attributes: true,
        attributeFilter: ["style", "class", "data-docx-source-id"],
      })
    }
    if (typeof ResizeObserver === "function") {
      resizeObserver = new ResizeObserver(schedule)
      resizeObserver.observe(root)
      for (const surface of surfaces.value.values()) resizeObserver.observe(surface)
    }
    if (typeof window !== "undefined") window.addEventListener("resize", schedule)
    root.addEventListener?.("scroll", schedule, true)
    root.addEventListener?.("load", schedule, true)
    root.addEventListener?.("error", schedule, true)
    const fonts = typeof document === "undefined" ? undefined : document.fonts
    fonts?.addEventListener?.("loadingdone", schedule)
    void fonts?.ready?.then?.(schedule)
  }

  const stopObservers = (): void => {
    mutationObserver?.disconnect()
    resizeObserver?.disconnect()
    mutationObserver = undefined
    resizeObserver = undefined
    const root = options.root()
    if (typeof window !== "undefined") window.removeEventListener("resize", schedule)
    root?.removeEventListener?.("scroll", schedule, true)
    root?.removeEventListener?.("load", schedule, true)
    root?.removeEventListener?.("error", schedule, true)
    if (typeof document !== "undefined") {
      document.fonts?.removeEventListener?.("loadingdone", schedule)
    }
  }

  watch(options.root, () => {
    lastSignature = "";
    schedule();
  }, { flush: "post" })
  watch(options.blocks, () => schedule(), { flush: "post" })

  onMounted(() => {
    update()
    startObservers()
  })

  onBeforeUnmount(() => {
    disposed = true
    cancelScheduled()
    stopObservers()
  })

  return { geometries, surfaces, refresh: update }
}

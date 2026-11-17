/**
 * Byte-budgeted LRU for rasterized PDF page bitmaps.
 *
 * `PdfSurface` keeps the pages around the viewport rasterized; everything else
 * only stays useful if a short scroll back does not rasterize again. This cache
 * holds those off-viewport bitmaps under a byte budget and hands back the page
 * indexes that must be released when the budget is exceeded, so memory tracks a
 * host-configurable limit instead of the page count.
 *
 * Pure bookkeeping: the caller owns the object URLs and decides what to do with
 * the returned page indexes.
 */

/** Default budget for off-viewport page bitmaps (64 MiB). */
export const DEFAULT_PDF_BITMAP_BUDGET_BYTES = 64 * 1024 * 1024

export interface PdfBitmapEntry {
  pageIndex: number
  bytes: number
}

export interface PdfBitmapCache {
  readonly maxBytes: number
  /** Bytes currently held. */
  readonly bytes: number
  /** Number of held bitmaps. */
  readonly size: number
  /** Held bitmaps, least recently used first. */
  entries(): PdfBitmapEntry[]
  /**
   * Hold (or refresh) the bitmap of `pageIndex`, marking it most recently used.
   * Returns the page indexes that no longer fit the budget and must be
   * released, least recently used first. A bitmap larger than the whole budget
   * is evicted right away.
   */
  retain(pageIndex: number, bytes: number): number[]
  /** Whether the bitmap of `pageIndex` is held. */
  has(pageIndex: number): boolean
  /** Forget one bitmap without treating it as an eviction (it is active again). */
  drop(pageIndex: number): void
  /** Evict until the budget is met; returns the evicted page indexes. */
  trim(): number[]
  /** Forget everything; returns the page indexes that were held. */
  clear(): number[]
}

function toByteCount(value: unknown): number {
  return typeof value === "number" && Number.isFinite(value) && value > 0 ? Math.round(value) : 0
}

/**
 * Create an LRU cache bounded by `maxBytes`. `maxBytes: 0` disables retention,
 * which degrades to releasing every bitmap as soon as it leaves the viewport.
 */
export function createPdfBitmapCache(options: { maxBytes?: number } = {}): PdfBitmapCache {
  const maxBytes = toByteCount(options.maxBytes ?? DEFAULT_PDF_BITMAP_BUDGET_BYTES)
  const bitmaps = new Map<number, number>()
  let usedBytes = 0

  function remove(pageIndex: number): boolean {
    const bytes = bitmaps.get(pageIndex)
    if (bytes === undefined) return false
    bitmaps.delete(pageIndex)
    usedBytes = Math.max(0, usedBytes - bytes)
    return true
  }

  function trim(): number[] {
    const evicted: number[] = []
    while (usedBytes > maxBytes && bitmaps.size > 0) {
      // Map iteration order is recency order: the first key is the LRU one.
      const oldest = bitmaps.keys().next()
      if (oldest.done) break
      remove(oldest.value)
      evicted.push(oldest.value)
    }
    return evicted
  }

  return {
    maxBytes,
    get bytes() {
      return usedBytes
    },
    get size() {
      return bitmaps.size
    },
    entries(): PdfBitmapEntry[] {
      return [...bitmaps].map(([pageIndex, bytes]) => ({ pageIndex, bytes }))
    },
    retain(pageIndex: number, bytes: number): number[] {
      remove(pageIndex)
      const measured = toByteCount(bytes)
      bitmaps.set(pageIndex, measured)
      usedBytes += measured
      return trim()
    },
    has(pageIndex: number): boolean {
      return bitmaps.has(pageIndex)
    },
    drop(pageIndex: number): void {
      remove(pageIndex)
    },
    trim,
    clear(): number[] {
      const held = [...bitmaps.keys()]
      bitmaps.clear()
      usedBytes = 0
      return held
    },
  }
}

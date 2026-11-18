/**
 * `documentPageCount` is the page count Word cached in `docProps/app.xml`.
 * Generated and never-repaginated documents routinely carry a stale value
 * (the demo/legal fixtures store `<Pages>1</Pages>` while Word and LibreOffice
 * render nine pages), so the viewer must treat it as a target to compare
 * against, never as a constraint that may reshape pages.
 *
 * The measure-driven iteration loop (#99) reports whether the target was
 * reached. It never scales a virtual page height, merges trailing pages or
 * otherwise forces the count.
 */
export interface PaginationPageCountTarget {
  /** Cached `documentPageCount`, when it is a usable page number. */
  targetPageCount?: number;
  /** Always false: this target is advisory and never changes pagination. */
  enforced: false;
}

/** Normalize a cached page count into an advisory pagination target. */
export function resolvePaginationPageCountTarget(
  pageCount: unknown
): PaginationPageCountTarget {
  const numeric =
    typeof pageCount === "number"
      ? pageCount
      : typeof pageCount === "string" && pageCount.trim() !== ""
      ? Number(pageCount)
      : Number.NaN;
  if (!Number.isFinite(numeric) || numeric < 1) {
    return { enforced: false };
  }
  return { targetPageCount: Math.max(1, Math.round(numeric)), enforced: false };
}

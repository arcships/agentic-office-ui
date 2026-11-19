/**
 * Reference-stable reuse for pagination results (#99).
 *
 * Repagination runs on every measurement pass. Pages whose segmentation did
 * not change must keep the exact same arrays/objects so the viewer can skip
 * re-rendering them (Vue compares props by reference).
 */
export interface ReuseUnchangedPaginationItemsOptions<T> {
  /** Items from the previous pass, indexed the same way as `next`. */
  previous?: readonly T[];
  /** Items from the new pass. */
  next: readonly T[];
  /** True when the previous and next item describe the same rendered page. */
  isUnchanged: (previous: T, next: T, index: number) => boolean;
}

export interface ReusedPaginationItems<T> {
  /** `next` with unchanged entries replaced by the previous instances. */
  items: T[];
  /** Indexes whose identity changed (or that are new). */
  changedIndexes: number[];
  /** Indexes that reused the previous instance. */
  reusedIndexes: number[];
}

/**
 * Replace entries of `next` with the matching entries of `previous` when they
 * describe the same page. The result keeps order and length of `next`; only
 * unchanged indexes alias the previous pass.
 */
export function reuseUnchangedPaginationItems<T>(
  options: ReuseUnchangedPaginationItemsOptions<T>
): ReusedPaginationItems<T> {
  const items: T[] = new Array(options.next.length);
  const changedIndexes: number[] = [];
  const reusedIndexes: number[] = [];

  for (let index = 0; index < options.next.length; index += 1) {
    const nextItem = options.next[index];
    const previousItem = options.previous?.[index];
    const reused =
      previousItem !== undefined &&
      options.isUnchanged(previousItem, nextItem, index);
    if (reused) {
      items[index] = previousItem as T;
      reusedIndexes.push(index);
    } else {
      items[index] = nextItem;
      changedIndexes.push(index);
    }
  }

  return { items, changedIndexes, reusedIndexes };
}

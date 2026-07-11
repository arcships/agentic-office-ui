export const DEFAULT_DOCX_HISTORY_MAX_BYTES = 32 * 1024 * 1024;

export interface HistoryBudgetOptions<T> {
  maxEntries: number;
  maxBytes: number;
  /** `end` for undo stacks, `start` for redo stacks. */
  newestAt?: "start" | "end";
  estimateBytes?: (entry: T) => number;
}

export interface HistoryBudgetResult<T> {
  entries: T[];
  estimatedBytes: number;
}

/**
 * Estimates retained serializable state without expanding binary buffers into
 * JSON number maps. The result is intentionally conservative and stable.
 */
export function estimateHistoryEntryBytes(value: unknown): number {
  let binaryBytes = 0;
  const seen = new WeakSet<object>();
  try {
    const json = JSON.stringify(value, (_key, candidate: unknown) => {
      if (!candidate || typeof candidate !== "object") return candidate;
      if (candidate instanceof ArrayBuffer) {
        binaryBytes += candidate.byteLength;
        return null;
      }
      if (ArrayBuffer.isView(candidate)) {
        binaryBytes += candidate.byteLength;
        return null;
      }
      if (seen.has(candidate)) return null;
      seen.add(candidate);
      return candidate;
    }) ?? "";
    return Math.min(
      Number.MAX_SAFE_INTEGER,
      new TextEncoder().encode(json).byteLength + binaryBytes,
    );
  } catch {
    return Number.MAX_SAFE_INTEGER;
  }
}

function normalizedPositiveInteger(value: number, fallback: number): number {
  return Number.isFinite(value) && value > 0 ? Math.max(1, Math.floor(value)) : fallback;
}

/**
 * Keeps one contiguous run of the most recent history. A newest entry larger
 * than the byte budget is retained by itself so the latest edit remains
 * undoable; older entries are then evicted.
 */
export function trimHistoryEntriesToBudget<T>(
  input: readonly T[],
  options: HistoryBudgetOptions<T>,
): HistoryBudgetResult<T> {
  const maxEntries = normalizedPositiveInteger(options.maxEntries, 1);
  const maxBytes = normalizedPositiveInteger(options.maxBytes, 1);
  const estimate = options.estimateBytes ?? estimateHistoryEntryBytes;
  const newestAt = options.newestAt ?? "end";
  const candidates = newestAt === "end" ? [...input].reverse() : [...input];
  const retained: T[] = [];
  let estimatedBytes = 0;

  for (const candidate of candidates) {
    if (retained.length >= maxEntries) break;
    const rawBytes = estimate(candidate);
    const candidateBytes = Number.isFinite(rawBytes) && rawBytes >= 0
      ? Math.ceil(rawBytes)
      : Number.MAX_SAFE_INTEGER;
    if (retained.length > 0 && estimatedBytes + candidateBytes > maxBytes) break;
    retained.push(candidate);
    estimatedBytes = Math.min(Number.MAX_SAFE_INTEGER, estimatedBytes + candidateBytes);
  }

  return {
    entries: newestAt === "end" ? retained.reverse() : retained,
    estimatedBytes,
  };
}

export const DEFAULT_DOCX_HISTORY_MAX_BYTES = 32 * 1024 * 1024

interface HistoryBudgetOptions<T> {
  maxEntries: number
  maxBytes: number
  newestAt?: "start" | "end"
  estimateBytes?: (entry: T) => number
}

interface HistoryBudgetResult<T> {
  entries: T[]
  estimatedBytes: number
}

export function estimateHistoryEntryBytes(value: unknown): number {
  let binaryBytes = 0
  const seen = new WeakSet<object>()
  try {
    const json = JSON.stringify(value, (_key, candidate: unknown) => {
      if (!candidate || typeof candidate !== "object") return candidate
      if (candidate instanceof ArrayBuffer) {
        binaryBytes += candidate.byteLength
        return null
      }
      if (ArrayBuffer.isView(candidate)) {
        binaryBytes += candidate.byteLength
        return null
      }
      if (seen.has(candidate)) return null
      seen.add(candidate)
      return candidate
    }) ?? ""
    return Math.min(
      Number.MAX_SAFE_INTEGER,
      new TextEncoder().encode(json).byteLength + binaryBytes,
    )
  } catch {
    return Number.MAX_SAFE_INTEGER
  }
}

export function trimHistoryEntriesToBudget<T>(
  input: readonly T[],
  options: HistoryBudgetOptions<T>,
): HistoryBudgetResult<T> {
  const maxEntries = Number.isFinite(options.maxEntries) && options.maxEntries > 0
    ? Math.max(1, Math.floor(options.maxEntries))
    : 1
  const maxBytes = Number.isFinite(options.maxBytes) && options.maxBytes > 0
    ? Math.max(1, Math.floor(options.maxBytes))
    : 1
  const estimate = options.estimateBytes ?? estimateHistoryEntryBytes
  const newestAt = options.newestAt ?? "end"
  const candidates = newestAt === "end" ? [...input].reverse() : [...input]
  const retained: T[] = []
  let estimatedBytes = 0

  for (const candidate of candidates) {
    if (retained.length >= maxEntries) break
    const rawBytes = estimate(candidate)
    const candidateBytes = Number.isFinite(rawBytes) && rawBytes >= 0
      ? Math.ceil(rawBytes)
      : Number.MAX_SAFE_INTEGER
    if (retained.length > 0 && estimatedBytes + candidateBytes > maxBytes) break
    retained.push(candidate)
    estimatedBytes = Math.min(Number.MAX_SAFE_INTEGER, estimatedBytes + candidateBytes)
  }

  return {
    entries: newestAt === "end" ? retained.reverse() : retained,
    estimatedBytes,
  }
}

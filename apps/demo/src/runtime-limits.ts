/**
 * Pure helpers shared by the runtime demo pages.
 *
 * The demo must be able to demonstrate a limit rejection *before* an uploaded
 * file is read. Keeping the size check here makes the ordering testable without
 * a browser.
 */

export interface StructuredRuntimeError {
  code: string
  message: string
  phase?: string
  limit?: string
  actual?: number
  allowed?: number
}

/** Runtime A in the isolation demo accepts the 36 KB sample. */
export const ISOLATION_SUCCESS_LIMIT_BYTES = 2 * 1024 * 1024

/** Runtime B rejects the same sample before reading it or starting a Worker. */
export const ISOLATION_REJECT_LIMIT_BYTES = 16 * 1024

export function inputSizeLimitError(
  fileSize: number | undefined,
  maxInputBytes: number | undefined,
): StructuredRuntimeError | undefined {
  if (typeof fileSize !== "number" || !Number.isFinite(fileSize) || fileSize < 0) return undefined
  if (typeof maxInputBytes !== "number" || !Number.isFinite(maxInputBytes) || maxInputBytes <= 0) {
    return undefined
  }
  if (fileSize <= maxInputBytes) return undefined
  return {
    code: "LIMIT_EXCEEDED",
    message: `输入大小 ${fileSize} 字节超过允许值 ${maxInputBytes} 字节。`,
    phase: "input",
    limit: "maxInputBytes",
    actual: fileSize,
    allowed: maxInputBytes,
  }
}

export function normalizeRuntimeError(error: unknown): StructuredRuntimeError {
  const value = error instanceof Error ? error : new Error(String(error))
  const details = value as Error & {
    code?: unknown
    phase?: unknown
    limit?: unknown
    actual?: unknown
    allowed?: unknown
  }
  return {
    code: typeof details.code === "string" ? details.code : "PARSE_FAILED",
    message: value.message,
    ...(typeof details.phase === "string" ? { phase: details.phase } : {}),
    ...(typeof details.limit === "string" ? { limit: details.limit } : {}),
    ...(typeof details.actual === "number" ? { actual: details.actual } : {}),
    ...(typeof details.allowed === "number" ? { allowed: details.allowed } : {}),
  }
}

export function isInputLimitError(error: StructuredRuntimeError | null | undefined): boolean {
  return error?.code === "LIMIT_EXCEEDED" && error.limit === "maxInputBytes"
}

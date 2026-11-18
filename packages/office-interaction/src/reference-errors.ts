import type {
  OfficeFormat,
  OfficeReferenceError,
  OfficeReferenceErrorCode,
  ResolveReferenceResult,
} from "./types"

const revisionReasonPattern = /revision|document-id|reflow|quote-unavailable/iu

/**
 * `ResolveReferenceResult` deliberately carries machine reason codes instead of
 * an error enum. This predicate lets surfaces separate "the document moved on"
 * from "the content is missing" without losing those detailed reason codes.
 */
export function isOfficeRevisionReason(reasonCode: string): boolean {
  return revisionReasonPattern.test(reasonCode)
}

export interface OfficeReferenceFailureOptions {
  operation?: "resolve" | "describe" | "scroll"
  referenceId?: string
}

/**
 * Map a failed resolve result onto the declared `OfficeReferenceErrorCode`
 * values. Exact, relocated and ambiguous results are not failures: relocated
 * references were already rewritten by the adapter, and ambiguous references
 * still need a caller decision instead of an error event.
 */
export function officeReferenceErrorForResult(
  format: OfficeFormat,
  result: ResolveReferenceResult,
  options: OfficeReferenceFailureOptions = {},
): OfficeReferenceError | undefined {
  if (result.status === "exact" || result.status === "relocated" || result.status === "ambiguous") {
    return undefined
  }
  const operation = options.operation ?? "resolve"
  const code: OfficeReferenceErrorCode = isOfficeRevisionReason(result.reasonCode)
    ? "REVISION_CHANGED"
    : operation === "describe" ? "DESCRIBE_FAILED" : "RESOLVE_FAILED"
  return {
    code,
    operation,
    format,
    recoverable: true,
    ...(options.referenceId === undefined ? {} : { referenceId: options.referenceId }),
    message: result.reasonCode,
  }
}

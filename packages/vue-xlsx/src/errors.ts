import type { XlsxLoadError, XlsxLoadErrorCode } from "@arcships/xlsx-core"
import { officeMessagesEn } from "@arcships/office-interaction"

export type XlsxLoadErrorMessages = typeof officeMessagesEn.xlsx.errors

/** Resolve a workbook load failure into locale-aware copy. */
export function localizeXlsxLoadError(
  error: Pick<XlsxLoadError, "code">,
  messages: XlsxLoadErrorMessages,
): string {
  switch (error.code as XlsxLoadErrorCode) {
    case "SOURCE_NOT_ALLOWED": return messages.sourceNotAllowed
    case "FETCH_FAILED": return messages.fetchFailed
    case "INVALID_WORKBOOK": return messages.invalidWorkbook
    case "LIMIT_EXCEEDED": return messages.limitExceeded
    case "IMAGE_LIMIT_EXCEEDED": return messages.imageLimitExceeded
    case "INVALID_IMAGE": return messages.invalidImage
    case "IMAGE_DECODE_FAILED": return messages.imageDecodeFailed
    case "TIMEOUT": return messages.timeout
    case "WORKER_UNAVAILABLE": return messages.workerUnavailable
    case "ABORTED": return messages.aborted
    default: return messages.unknown
  }
}

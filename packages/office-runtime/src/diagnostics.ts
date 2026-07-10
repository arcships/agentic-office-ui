import {
  sanitizeOfficeUrl,
  type OfficeFormat,
  type OfficeLoadErrorCode,
  type OfficeSourceKind,
} from "./errors";

export type OfficeDiagnosticType =
  | "load-start"
  | "load-success"
  | "load-error"
  | "load-cancelled"
  | "download";

export interface OfficeDiagnostic {
  type: OfficeDiagnosticType;
  runtimeId: string;
  taskId: string;
  format?: OfficeFormat;
  sourceKind?: OfficeSourceKind;
  phase?: string;
  path?: "worker" | "main-thread" | "network" | "local";
  url?: string;
  bytes?: number;
  durationMs?: number;
  errorCode?: OfficeLoadErrorCode;
}

export function createOfficeDiagnostic(event: OfficeDiagnostic): Readonly<OfficeDiagnostic> {
  return Object.freeze({
    ...event,
    ...(event.url ? { url: sanitizeOfficeUrl(event.url) } : {}),
  });
}

export function emitOfficeDiagnostic(
  callback: ((event: Readonly<OfficeDiagnostic>) => void) | undefined,
  event: OfficeDiagnostic,
): void {
  if (!callback) return;
  try {
    callback(createOfficeDiagnostic(event));
  } catch {
    // Diagnostics must never change the loading result.
  }
}

// Performance tracing helpers for DOCX import. Upstream editor.tsx lines 226-265.
//
// These are diagnostic-only performance.mark/measure wrappers used by the
// composables layer to emit user-timing marks during docx import.

const DOCX_IMPORT_PERFORMANCE_PREFIX = "react-docx.import";

export function markDocxImportPerformance(name: string): void {
  if (
    typeof performance === "undefined" ||
    typeof performance.mark !== "function"
  ) {
    return;
  }

  try {
    performance.mark(name);
  } catch {
    // Performance marks are diagnostic-only.
  }
}

export function measureDocxImportPerformance(
  name: string,
  startMark: string,
  endMark: string
): void {
  if (
    typeof performance === "undefined" ||
    typeof performance.measure !== "function"
  ) {
    return;
  }

  try {
    performance.measure(name, startMark, endMark);
  } catch {
    // A missing mark should not affect import.
  }
}

export function createDocxImportPerformanceTraceName(fileName: string): string {
  const normalizedName = fileName.replace(/[^a-z0-9._-]+/gi, "_").slice(0, 80);
  return `${DOCX_IMPORT_PERFORMANCE_PREFIX}.${Date.now()}.${normalizedName}`;
}

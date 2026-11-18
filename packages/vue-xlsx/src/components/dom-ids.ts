/**
 * Stable DOM ids for the XLSX viewer's ARIA relationships.
 *
 * `aria-controls` and `aria-labelledby` are only useful when the id they name
 * is unique per mounted viewer, so the sequence lives in module scope rather
 * than in `setup()` — a per-instance counter would give two viewers on the
 * same page the same id.
 */

let sequence = 0

export function nextXlsxDomId(prefix: string): string {
  sequence += 1
  return `${prefix}-${sequence}`
}

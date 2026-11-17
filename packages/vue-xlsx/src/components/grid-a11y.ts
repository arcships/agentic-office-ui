/**
 * Stable ids for the XLSX grid's accessibility tree.
 *
 * The grid renders a hidden row/cell for the active cell so
 * `aria-activedescendant` can point at it. The id must be unique per mounted
 * grid, so the sequence lives in module scope rather than in `setup()` — a
 * per-instance counter would give every grid on the page the same id.
 */

let sequence = 0

export function nextGridA11yId(): string {
  sequence += 1
  return `xlsx-grid-a11y-${sequence}`
}

/**
 * Harness shim for `@arcships/xlsx-core`.
 *
 * The XLSX chrome components import the engine only for types (`import type`),
 * which the Vue build erases, plus this one runtime constant. Aliasing the
 * package to this narrow re-export keeps the parity bundle free of the
 * spreadsheet engine while still reading the constant from source.
 */
export { XLSX_MAX_SHEET_NAME_LENGTH } from "../../../../packages/xlsx-core/src/authoring/sheet-rename";

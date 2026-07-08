// Table height estimation and table row slicing.
// Upstream editor.tsx: lines 10541-11590.
// All functions live in line-height-table.ts; this module re-exports them
// under the name expected by the split plan (table-height-estimate).
//
// When line-height-table.ts is eventually renamed, update this barrel.

export * from "./line-height-table";

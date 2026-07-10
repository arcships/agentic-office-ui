/** URL of the package-owned XLSX WASM binary. */
export const bundledXlsxWasmUrl = new URL(
  "./assets/duke_sheets_wasm_bg.wasm",
  import.meta.url,
).href;

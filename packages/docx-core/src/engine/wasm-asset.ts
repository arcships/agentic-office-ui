/**
 * Keep this as a static URL so Vite can discover the source asset when this
 * package is consumed from the workspace. The publish postbuild step rewrites
 * the emitted path to the public dist assets directory.
 */
export const bundledDocxWasmUrl = new URL(
  "./generated/docx_wasm_bg.wasm",
  import.meta.url,
).href;

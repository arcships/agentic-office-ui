import { copyFileSync, mkdirSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, join } from "node:path";

const require = createRequire(new URL("../package.json", import.meta.url));
const dukeEntrypoint = require.resolve("@dukelib/sheets-wasm");
const wasmSource = join(dirname(dukeEntrypoint), "duke_sheets_wasm_bg.wasm");
const distDir = new URL("../dist/", import.meta.url);

mkdirSync(distDir, { recursive: true });
copyFileSync(wasmSource, new URL("duke_sheets_wasm_bg.wasm", distDir));
console.log("Copied duke_sheets_wasm_bg.wasm to dist/");

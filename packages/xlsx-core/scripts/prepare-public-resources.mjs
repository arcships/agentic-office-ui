import { copyFile, mkdir, unlink } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const distDir = fileURLToPath(new URL("../dist/", import.meta.url));
const assetDir = join(distDir, "assets");
const sourceWasm = join(distDir, "duke_sheets_wasm_bg.wasm");

await mkdir(assetDir, { recursive: true });
await copyFile(
  sourceWasm,
  join(assetDir, "duke_sheets_wasm_bg.wasm"),
);
await unlink(sourceWasm);

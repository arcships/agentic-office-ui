import { copyFile, mkdir, readFile, readdir, unlink, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const distDir = fileURLToPath(new URL("../dist/", import.meta.url));
const assetDir = join(distDir, "assets");
const sourceWasm = join(distDir, "docx_wasm_bg.wasm");
const publicWasm = join(assetDir, "docx_wasm_bg.wasm");

async function findJavaScriptFiles(directory) {
  const entries = await readdir(directory, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    const target = join(directory, entry.name);
    if (entry.isDirectory()) {
      files.push(...await findJavaScriptFiles(target));
    } else if (entry.isFile() && entry.name.endsWith(".js")) {
      files.push(target);
    }
  }
  return files;
}

await mkdir(assetDir, { recursive: true });
await copyFile(sourceWasm, publicWasm);
await unlink(sourceWasm);
await Promise.all([
  unlink(join(distDir, "viewer", "docx-import-worker.js")),
  unlink(join(distDir, "viewer", "docx-import-worker.d.ts")),
]);

let rewrittenFiles = 0;
for (const file of await findJavaScriptFiles(distDir)) {
  const source = await readFile(file, "utf8");
  const normalized = source.replace(
    /new URL\(\s*(["'])(?:\.\/generated\/)?docx_wasm_bg\.wasm\1,\s*import\.meta\.url\s*\)/g,
    'new URL("./assets/docx_wasm_bg.wasm", import.meta.url)',
  );
  if (normalized !== source) {
    rewrittenFiles += 1;
    await writeFile(file, normalized);
  }
}

if (rewrittenFiles === 0) {
  throw new Error("DOCX public WASM URL was not found in the build output.");
}

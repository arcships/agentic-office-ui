#!/usr/bin/env node
/**
 * Copy the pinned pdfium wasm next to the reference route entries.
 *
 * The upstream `lib/pdf-thumbnail-utils.ts` fetches the wasm from jsDelivr
 * (`@embedpdf/pdfium@2.15.0`). Route captures must not depend on the network,
 * so this harness serves the same file from its own origin. Keeping the copy
 * in `public/` (git-ignored) makes the wasm version follow the installed
 * package instead of a committed binary.
 */
import { copyFileSync, existsSync, mkdirSync, statSync } from "node:fs";
import path from "node:path";

const root = path.resolve(import.meta.dirname, "..");
const source = path.join(
  root,
  "node_modules/@embedpdf/pdfium/dist/pdfium.wasm",
);
const destination = path.join(root, "public/pdfium.wasm");

if (!existsSync(source)) {
  console.error(
    `pdfium.wasm is missing at ${path.relative(root, source)}; run pnpm install in the workspace root.`,
  );
  process.exit(1);
}

mkdirSync(path.dirname(destination), { recursive: true });
if (
  !existsSync(destination) ||
  statSync(destination).size !== statSync(source).size
) {
  copyFileSync(source, destination);
  console.log(
    `▸ pdfium.wasm ${(statSync(destination).size / 1024 / 1024).toFixed(2)} MiB copied from @embedpdf/pdfium`,
  );
}

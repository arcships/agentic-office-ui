import { readFile, writeFile } from "node:fs/promises";

const workerUrl = new URL("../dist/docx-import-worker.js", import.meta.url);
const worker = await readFile(workerUrl, "utf8");
const normalized = worker.replace(
  /(["'])\.\.\/(chunk-[^"']+\.js)\1/g,
  "$1./$2$1",
);

if (normalized === worker || normalized.includes('../chunk-')) {
  throw new Error(
    "Unable to normalize dist/docx-import-worker.js chunk imports for the published package.",
  );
}

await writeFile(workerUrl, normalized);

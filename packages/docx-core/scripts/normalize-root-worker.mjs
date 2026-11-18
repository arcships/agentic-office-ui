import { readFile, writeFile } from "node:fs/promises";

for (const name of ["docx-import-worker.js", "docx-export-worker.js"]) {
  const workerUrl = new URL(`../dist/${name}`, import.meta.url);
  const worker = await readFile(workerUrl, "utf8");
  const normalized = worker.replace(
    /(["'])\.\.\/(chunk-[^"']+\.js)\1/g,
    "$1./$2$1",
  );

  if (normalized.includes("../chunk-")) {
    throw new Error(
      `Unable to normalize dist/${name} chunk imports for the published package.`,
    );
  }

  if (normalized !== worker) {
    await writeFile(workerUrl, normalized);
  }
}

import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const runtime = await import(
  new URL("../../packages/office-runtime/dist/index.js", import.meta.url).href
);
const requireFromXlsx = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const xlsx = await import(
  pathToFileURL(requireFromXlsx.resolve("@arcships/xlsx-core")).href
);
const docx = await import(
  new URL("../../packages/docx-core/dist/index.js", import.meta.url).href
);
const { zipSync } = await import(
  pathToFileURL(requireFromXlsx.resolve("fflate")).href
);

function declaredPng(width, height, complete = true) {
  const bytes = new Uint8Array(complete ? 45 : 33);
  bytes.set([137, 80, 78, 71, 13, 10, 26, 10]);
  const view = new DataView(bytes.buffer);
  view.setUint32(8, 13);
  bytes.set([73, 72, 68, 82], 12);
  view.setUint32(16, width);
  view.setUint32(20, height);
  bytes[24] = 8;
  bytes[25] = 6;
  if (complete) bytes.set([0, 0, 0, 0, 73, 69, 78, 68, 0, 0, 0, 0], 33);
  return bytes;
}

test("image budget checks declared dimensions before allocating decoded pixels", () => {
  const budget = runtime.createOfficeImageBudget({
    maxImageWidth: 32_768,
    maxImageHeight: 32_768,
    maxSingleImagePixels: 40_000_000,
    maxTotalImagePixels: 100_000_000,
    maxSingleImageBytes: 1024,
    maxTotalImageBytes: 2048,
    maxConcurrentImageDecodes: 2,
  }, "docx");
  assert.throws(
    () => budget.inspectAndReserve(declaredPng(100_000, 100_000, false), "word/media/bomb.png"),
    (error) => error?.code === "IMAGE_LIMIT_EXCEEDED"
      && error?.phase === "image"
      && error?.limit === "maxImageWidth",
  );
  assert.deepEqual(budget.snapshot(), {
    compressedBytes: 0,
    pixels: 0,
    activeDecodes: 0,
    pendingDecodes: 0,
    disposed: false,
  });
  budget.dispose();
});

test("image budget reports malformed images and aggregate pixel excess without partial reservation", () => {
  const budget = runtime.createOfficeImageBudget({
    maxSingleImagePixels: 40_000_000,
    maxTotalImagePixels: 50_000_000,
  }, "xlsx");
  assert.throws(
    () => budget.inspectAndReserve(declaredPng(10, 10, false), "xl/media/broken.png"),
    (error) => error?.code === "INVALID_IMAGE" && error?.phase === "image",
  );
  const metadata = budget.inspectAndReserve(declaredPng(6_000, 6_000), "xl/media/high-res.png");
  assert.equal(metadata.pixels, 36_000_000);
  assert.equal(metadata.decodedBytes, 144_000_000);
  assert.throws(
    () => budget.inspectAndReserve(declaredPng(4_000, 4_000), "xl/media/second.png"),
    (error) => error?.code === "IMAGE_LIMIT_EXCEEDED"
      && error?.limit === "maxTotalImagePixels"
      && error?.actual === 52_000_000,
  );
  assert.equal(budget.snapshot().pixels, 36_000_000);
  budget.dispose();
});

test("compressed image byte limits enforce exact single and aggregate boundaries", () => {
  const image = declaredPng(1, 1);
  assert.equal(image.byteLength, 45);
  const exact = runtime.createOfficeImageBudget({
    maxSingleImageBytes: 45,
    maxTotalImageBytes: 90,
  }, "docx");
  exact.inspectAndReserve(image, "word/media/one.png");
  exact.inspectAndReserve(image, "word/media/two.png");
  assert.equal(exact.snapshot().compressedBytes, 90);
  assert.throws(
    () => exact.inspectAndReserve(image, "word/media/three.png"),
    (error) => error?.code === "IMAGE_LIMIT_EXCEEDED"
      && error?.limit === "maxTotalImageBytes"
      && error?.actual === 135,
  );
  exact.dispose();

  const tooSmall = runtime.createOfficeImageBudget({ maxSingleImageBytes: 44 }, "xlsx");
  assert.throws(
    () => tooSmall.inspectAndReserve(image, "xl/media/one.png"),
    (error) => error?.code === "IMAGE_LIMIT_EXCEEDED"
      && error?.limit === "maxSingleImageBytes"
      && error?.actual === 45,
  );
  tooSmall.dispose();
});

test("decode permits enforce per-instance concurrency and release queued work", async () => {
  const budget = runtime.createOfficeImageBudget({ maxConcurrentImageDecodes: 2 }, "xlsx");
  const releaseFirst = await budget.acquireDecodePermit();
  const releaseSecond = await budget.acquireDecodePermit();
  let thirdStarted = false;
  const third = budget.acquireDecodePermit().then((release) => {
    thirdStarted = true;
    return release;
  });
  await Promise.resolve();
  assert.equal(thirdStarted, false);
  assert.equal(budget.snapshot().pendingDecodes, 1);
  releaseFirst();
  const releaseThird = await third;
  assert.equal(thirdStarted, true);
  assert.equal(budget.snapshot().activeDecodes, 2);
  releaseSecond();
  releaseThird();
  assert.equal(budget.snapshot().activeDecodes, 0);
  budget.dispose();
});

test("disposing an image budget rejects pending decodes and leaves no queue", async () => {
  const budget = runtime.createOfficeImageBudget({ maxConcurrentImageDecodes: 1 }, "docx");
  const release = await budget.acquireDecodePermit();
  const pending = budget.acquireDecodePermit();
  budget.dispose();
  await assert.rejects(
    pending,
    (error) => error?.code === "RUNTIME_DISPOSED" && error?.phase === "image-decode",
  );
  release();
  assert.deepEqual(budget.snapshot(), {
    compressedBytes: 0,
    pixels: 0,
    activeDecodes: 0,
    pendingDecodes: 0,
    disposed: true,
  });
});

test("DOCX and XLSX reject oversized embedded images before browser object URLs", () => {
  const oversized = declaredPng(100_000, 100_000, false);
  assert.throws(
    () => docx.validateDocxImageAssets(
      { binaryAssets: new Map([["word/media/bomb.png", oversized]]) },
      { maxImageWidth: 32_768 },
    ),
    (error) => error?.code === "IMAGE_LIMIT_EXCEEDED" && error?.format === "docx",
  );

  const originalUrl = globalThis.URL;
  let createdObjectUrls = 0;
  globalThis.URL = {
    createObjectURL() {
      createdObjectUrls += 1;
      return "blob:unexpected";
    },
    revokeObjectURL() {},
  };
  try {
    const workbook = zipSync({ "xl/media/bomb.png": oversized });
    assert.throws(
      () => xlsx.parseWorkbookImageAssets(workbook, { maxImageWidth: 32_768 }),
      (error) => error?.code === "IMAGE_LIMIT_EXCEEDED" && error?.format === "xlsx",
    );
    assert.equal(createdObjectUrls, 0);
  } finally {
    globalThis.URL = originalUrl;
  }
});

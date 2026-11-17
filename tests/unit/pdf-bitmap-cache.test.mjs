import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";

// Pure bookkeeping: exercise the TypeScript source directly.
const requireFromPdf = createRequire(
  new URL("../../packages/vue-pdf/package.json", import.meta.url),
);
const typescript = requireFromPdf("typescript");

async function importSource(relativePath) {
  const source = readFileSync(new URL(relativePath, import.meta.url), "utf8");
  return import(`data:text/javascript;base64,${Buffer.from(
    typescript.transpileModule(source, {
      compilerOptions: {
        module: typescript.ModuleKind.ES2022,
        target: typescript.ScriptTarget.ES2022,
      },
    }).outputText,
  ).toString("base64")}`);
}

const {
  DEFAULT_PDF_BITMAP_BUDGET_BYTES,
  createPdfBitmapCache,
} = await importSource("../../packages/vue-pdf/src/pdf/pdf-bitmap-cache.ts");

const MIB = 1024 * 1024;

test("the default bitmap budget is bounded instead of page-count driven", () => {
  assert.equal(DEFAULT_PDF_BITMAP_BUDGET_BYTES, 64 * MIB);
  const cache = createPdfBitmapCache();
  assert.equal(cache.maxBytes, 64 * MIB);
  assert.equal(cache.size, 0);
  assert.equal(cache.bytes, 0);
});

test("bitmaps within the budget are retained without eviction", () => {
  const cache = createPdfBitmapCache({ maxBytes: 10 * MIB });
  assert.deepEqual(cache.retain(0, 2 * MIB), []);
  assert.deepEqual(cache.retain(1, 2 * MIB), []);
  assert.deepEqual(cache.retain(2, 2 * MIB), []);
  assert.equal(cache.size, 3);
  assert.equal(cache.bytes, 6 * MIB);
  assert.deepEqual(
    cache.entries().map((entry) => entry.pageIndex),
    [0, 1, 2],
    "entries are ordered least recently used first",
  );
});

test("exceeding the budget evicts the least recently used bitmap first", () => {
  const cache = createPdfBitmapCache({ maxBytes: 6 * MIB });
  cache.retain(0, 2 * MIB);
  cache.retain(1, 2 * MIB);
  cache.retain(2, 2 * MIB);
  assert.deepEqual(cache.retain(3, 2 * MIB), [0]);
  assert.equal(cache.bytes, 6 * MIB);
  assert.deepEqual(
    cache.entries().map((entry) => entry.pageIndex),
    [1, 2, 3],
  );
});

test("re-retaining a page refreshes its recency", () => {
  const cache = createPdfBitmapCache({ maxBytes: 6 * MIB });
  cache.retain(0, 2 * MIB);
  cache.retain(1, 2 * MIB);
  cache.retain(2, 2 * MIB);
  cache.drop(0);
  cache.retain(0, 2 * MIB);
  assert.deepEqual(cache.retain(3, 2 * MIB), [1], "page 0 re-entered the window, page 1 is now the LRU");
  assert.deepEqual(
    cache.entries().map((entry) => entry.pageIndex),
    [2, 0, 3],
  );
});

test("re-retaining a page updates its footprint instead of double counting", () => {
  const cache = createPdfBitmapCache({ maxBytes: 10 * MIB });
  cache.retain(7, 4 * MIB);
  assert.equal(cache.has(7), true);
  assert.equal(cache.has(8), false);
  cache.retain(7, 2 * MIB);
  assert.equal(cache.size, 1);
  assert.equal(cache.bytes, 2 * MIB);
  assert.deepEqual(cache.retain(8, 8 * MIB), [], "the smaller refreshed bitmap fits");
  assert.equal(cache.bytes, 10 * MIB);
  cache.drop(7);
  assert.equal(cache.has(7), false);
});

test("dropping an active page frees its budget without reporting eviction", () => {
  const cache = createPdfBitmapCache({ maxBytes: 4 * MIB });
  cache.retain(0, 3 * MIB);
  cache.retain(1, 1 * MIB);
  cache.drop(1);
  assert.equal(cache.size, 1);
  assert.equal(cache.bytes, 3 * MIB);
  cache.drop(42);
  assert.equal(cache.bytes, 3 * MIB, "dropping an absent page is a no-op");
});

test("a bitmap larger than the whole budget is evicted immediately", () => {
  const cache = createPdfBitmapCache({ maxBytes: 2 * MIB });
  assert.deepEqual(cache.retain(0, 5 * MIB), [0]);
  assert.equal(cache.size, 0);
  assert.equal(cache.bytes, 0);
});

test("maxBytes 0 disables retention and releases every bitmap", () => {
  const cache = createPdfBitmapCache({ maxBytes: 0 });
  assert.deepEqual(cache.retain(0, 1), [0]);
  assert.deepEqual(cache.retain(1, 1), [1]);
  assert.equal(cache.size, 0);
});

test("clear reports the held pages in LRU order and resets the budget", () => {
  const cache = createPdfBitmapCache({ maxBytes: 10 * MIB });
  cache.retain(0, 1 * MIB);
  cache.retain(4, 1 * MIB);
  cache.retain(9, 1 * MIB);
  cache.drop(0);
  cache.retain(0, 1 * MIB);
  assert.deepEqual(cache.clear(), [4, 9, 0]);
  assert.equal(cache.size, 0);
  assert.equal(cache.bytes, 0);
  assert.deepEqual(cache.retain(1, 1 * MIB), [], "the budget is usable again after a clear");
});

test("non-finite footprints are ignored instead of corrupting the budget", () => {
  const cache = createPdfBitmapCache({ maxBytes: 4 * MIB });
  cache.retain(0, Number.NaN);
  cache.retain(1, -5);
  assert.equal(cache.bytes, 0);
  assert.equal(cache.size, 2);
  assert.deepEqual(
    cache.entries().map((entry) => entry.bytes),
    [0, 0],
  );
});

test("a scrolling session stays inside the byte budget", () => {
  const pageBytes = 3 * MIB;
  const cache = createPdfBitmapCache({ maxBytes: 64 * MIB });
  let released = 0;
  for (let pageIndex = 0; pageIndex < 200; pageIndex++) {
    // Three visible pages enter, three leave on every scroll step.
    for (let offset = 0; offset < 3; offset++) {
      const leaving = pageIndex * 3 + offset;
      const entering = leaving + 1;
      released += cache.retain(leaving, pageBytes).length;
      cache.drop(entering);
    }
    assert.ok(cache.bytes <= cache.maxBytes, `budget exceeded at page ${pageIndex}`);
  }
  assert.ok(cache.size < 200, "the cache holds a bounded slice of the document");
  assert.ok(released > 0, "the LRU actually evicts once the budget is reached");
});

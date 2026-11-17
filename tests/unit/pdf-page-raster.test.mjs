import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";

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
  DEFAULT_MAX_PAGE_RASTER_PIXELS,
  PDF_RASTER_BYTES_PER_PIXEL,
  planPageRaster,
} = await importSource("../../packages/vue-pdf/src/pdf/pdf-page-raster.ts");

test("pages below the pixel budget render at the requested scale", () => {
  const plan = planPageRaster({
    displayWidth: 612,
    displayHeight: 792,
    zoom: 1,
    dpr: 2,
    maxPixels: 4_000_000,
  });
  assert.equal(plan.capped, false);
  assert.equal(plan.scale, 1);
  assert.equal(plan.dpr, 2);
  assert.equal(plan.pixelWidth, 1224);
  assert.equal(plan.pixelHeight, 1584);
  assert.equal(plan.pixels, 1224 * 1584);
  assert.equal(plan.bytes, plan.pixels * PDF_RASTER_BYTES_PER_PIXEL);
});

test("the default per-page budget is 8 MP", () => {
  assert.equal(DEFAULT_MAX_PAGE_RASTER_PIXELS, 8_000_000);
  const plan = planPageRaster({
    displayWidth: 4000,
    displayHeight: 4000,
    zoom: 2,
    dpr: 1,
  });
  assert.equal(plan.capped, true);
  assert.ok(plan.pixels <= DEFAULT_MAX_PAGE_RASTER_PIXELS);
});

test("high zoom pages rasterize at a capped pixel budget", () => {
  // Letter at 400 % on a 2x display is 32 MP before the cap.
  const plan = planPageRaster({
    displayWidth: 2448,
    displayHeight: 3168,
    zoom: 4,
    dpr: 2,
    maxPixels: 8_000_000,
  });
  assert.equal(plan.capped, true);
  assert.ok(plan.pixels <= 8_000_000);
  assert.ok(plan.scale < 4, "the engine is asked for less than the full zoom");
  assert.ok(plan.scale > 0);
  assert.equal(plan.dpr, 2, "the display scale is untouched; only the raster shrinks");
  const aspect = plan.pixelWidth / plan.pixelHeight;
  const expected = 2448 / 3168;
  assert.ok(Math.abs(aspect - expected) < 0.01, `aspect ratio drifted: ${aspect} vs ${expected}`);
});

test("the raster size stays at the pixel budget as zoom grows", () => {
  const atOne = planPageRaster({ displayWidth: 3000, displayHeight: 4000, zoom: 1, dpr: 1, maxPixels: 1_000_000 });
  const atTwo = planPageRaster({ displayWidth: 6000, displayHeight: 8000, zoom: 2, dpr: 1, maxPixels: 1_000_000 });
  assert.ok(atOne.pixels <= 1_000_000 && atTwo.pixels <= 1_000_000);
  assert.ok(
    Math.abs(atTwo.pixels - atOne.pixels) <= atOne.pixels * 0.02,
    "both zooms land on the same raster size",
  );
  assert.ok(
    atTwo.scale < atOne.scale * 1.05,
    "a further zoom cannot grow the raster, so the engine scale stays put",
  );
});

test("a device pixel ratio above 1 consumes the pixel budget", () => {
  const cssOnly = planPageRaster({ displayWidth: 1000, displayHeight: 1500, zoom: 1, dpr: 1, maxPixels: 2_000_000 });
  const hidpi = planPageRaster({ displayWidth: 1000, displayHeight: 1500, zoom: 1, dpr: 2, maxPixels: 2_000_000 });
  assert.equal(cssOnly.capped, false);
  assert.equal(cssOnly.pixels, 1_500_000);
  assert.equal(hidpi.capped, true);
  assert.ok(hidpi.pixels <= 2_000_000);
  assert.ok(hidpi.scale < cssOnly.scale);
});

test("rotated pages swap display dimensions without changing the budget", () => {
  const portrait = planPageRaster({
    displayWidth: 1000,
    displayHeight: 2000,
    zoom: 1,
    dpr: 2,
    maxPixels: 2_000_000,
  });
  const landscape = planPageRaster({
    displayWidth: 2000,
    displayHeight: 1000,
    zoom: 1,
    dpr: 2,
    maxPixels: 2_000_000,
  });
  assert.equal(portrait.pixels, landscape.pixels);
  assert.equal(portrait.scale, landscape.scale);
});

test("degenerate inputs never produce NaN geometry", () => {
  const unmeasured = planPageRaster({ displayWidth: 0, displayHeight: 0, zoom: 1, dpr: 1 });
  assert.deepEqual(unmeasured, {
    scale: 1,
    dpr: 1,
    pixelWidth: 0,
    pixelHeight: 0,
    pixels: 0,
    bytes: 0,
    capped: false,
  });
  const broken = planPageRaster({
    displayWidth: 500,
    displayHeight: 700,
    zoom: Number.NaN,
    dpr: 0,
    maxPixels: Number.NaN,
  });
  assert.equal(broken.scale, 1);
  assert.equal(broken.dpr, 1);
  assert.ok(Number.isFinite(broken.pixels));
  assert.ok(broken.pixels <= DEFAULT_MAX_PAGE_RASTER_PIXELS);
});

test("an extremely narrow page still honours a tiny budget", () => {
  const plan = planPageRaster({
    displayWidth: 10_000,
    displayHeight: 10,
    zoom: 1,
    dpr: 1,
    maxPixels: 100,
  });
  assert.equal(plan.capped, true);
  assert.ok(plan.pixels <= 100, `budget overshot: ${plan.pixels}`);
  assert.ok(plan.pixelWidth >= 1 && plan.pixelHeight >= 1);
});

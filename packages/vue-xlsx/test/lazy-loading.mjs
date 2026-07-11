import assert from "node:assert/strict";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import test from "node:test";

const ROOT = path.resolve(import.meta.dirname, "../../..");
const PACKAGE_ROOT = path.join(ROOT, "packages/vue-xlsx");

function read(relativePath) {
  return readFileSync(path.join(ROOT, relativePath), "utf8");
}

test("vue-xlsx publishes explicit chart, map, and WebGL entries", () => {
  const manifest = JSON.parse(read("packages/vue-xlsx/package.json"));
  for (const feature of ["chart", "map", "webgl"]) {
    const entry = manifest.exports[`./${feature}`];
    assert.ok(entry, `missing ./${feature} export`);
    assert.ok(entry.import.startsWith("./dist/"));
    assert.ok(entry.types.startsWith("./dist/"));
  }
});

test("default renderer facade uses dynamic feature imports", () => {
  const facade = read("packages/vue-xlsx/src/optional/lazy-renderers.ts");
  assert.match(facade, /import\(["']\.\/chart["']\)/);
  assert.match(facade, /import\(["']\.\/webgl["']\)/);
  assert.match(facade, /data-testid["']?: ["']xlsx-optional-renderer-loading["']/);

  const rootEntry = read("packages/vue-xlsx/src/index.ts");
  assert.doesNotMatch(rootEntry, /export \{[^}]+\} from ["']\.\/render["']/s);
});

test("demo routes do not statically import page components", () => {
  const main = read("apps/demo/src/main.ts");
  assert.doesNotMatch(main, /^import .*\.\/pages\/.*\.vue["'];?$/m);
  assert.match(main, /import\(["']\.\/pages\/XlsxViewerPage\.vue["']\)/);
});

test("formal build keeps features separate and drops legacy public WASM copies", () => {
  const dist = path.join(PACKAGE_ROOT, "dist");
  for (const fileName of ["index.js", "chart.js", "map.js", "webgl.js"]) {
    assert.ok(existsSync(path.join(dist, fileName)), `missing built ${fileName}`);
  }
  const rootBundle = readFileSync(path.join(dist, "index.js"), "utf8");
  assert.match(rootBundle, /import\(["']\.\/chart\.js["']\)/);

  const demoAssets = path.join(ROOT, "apps/demo/dist/assets");
  const mainBundle = readdirSync(demoAssets)
    .filter((name) => /^index-.*\.js$/.test(name))
    .map((name) => readFileSync(path.join(demoAssets, name), "utf8"))
    .join("\n");
  assert.doesNotMatch(mainBundle, /xlsx-grid/);
  assert.equal(existsSync(path.join(ROOT, "apps/demo/dist/docx_wasm_bg.wasm")), false);
  assert.equal(existsSync(path.join(ROOT, "apps/demo/dist/duke_sheets_wasm_bg.wasm")), false);
});

// Issue #182 — L2 product-route parity contract.
//
// L1 (`ui:parity`) compares component cells the harness composes itself; it
// cannot prove that `pnpm dev` shows the same thing as the upstream docs route.
// These assertions pin the mechanism that closes that gap:
//
//   - `tools/parity-fixtures/routes.json` stays a complete, self-consistent
//     route matrix (region selectors, landmarks, state, exclusions, classified
//     differences, pinned fixture hash);
//   - the subject route carries the region marker, pins the upstream region
//     height and reads the capture state from the URL without changing the
//     shipped default;
//   - the reference app really renders the pinned upstream wrapper (R2) with
//     pinned public dependencies, a second Vite entry and a local pdfium wasm;
//   - the harness scripts, pnpm entry point and docs exist and stay linked.
//
// The pixel comparison itself needs Chromium and runs through
// `pnpm ui:parity:routes`; these checks are the fast contract around it.

import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";

const ROOT = new URL("../../", import.meta.url);
const rootPath = path.resolve(new URL(".", ROOT).pathname);

function read(relativePath) {
  return readFileSync(new URL(relativePath, ROOT), "utf8");
}

function readJson(relativePath) {
  return JSON.parse(read(relativePath));
}

const routes = readJson("tools/parity-fixtures/routes.json");
const referencePackage = readJson("tools/extend-reference/package.json");
const rootPackage = readJson("package.json");

const ALLOWED_EXCLUSION_CATEGORIES = new Set([
  "demo-scaffolding",
  "product-component",
  "absent-upstream",
]);
const ALLOWED_DIFFERENCE_CATEGORIES = new Set([
  "product-component",
  "demo-scaffolding",
  "absent-upstream",
]);

test("route matrix declares the pilot with a pinned upstream commit", () => {
  assert.equal(routes.schemaVersion, 1);
  assert.match(routes.referenceCommit, /^[0-9a-f]{40}$/);
  assert.ok(routes.routes.length >= 1);
  const ids = routes.routes.map((route) => route.id);
  assert.equal(new Set(ids).size, ids.length, "route ids must be unique");
  assert.ok(ids.includes("pdf-viewer"), "the #182 pilot route must exist");
  assert.deepEqual(routes.themes.sort(), ["dark", "light"]);
  assert.ok(routes.viewports.some((viewport) => viewport.id === "1440x900"));
  assert.ok(routes.viewports.some((viewport) => viewport.id === "390x844"));
  assert.equal(routes.thresholds.geometryTolerancePx, 0);
  assert.equal(routes.thresholds.maxDiffPixelRatio, 0.001);
});

test("route entries define the product region, landmarks and state", () => {
  for (const route of routes.routes) {
    assert.ok(["blocking", "report-only"].includes(route.gateStatus), route.id);
    assert.ok(route.reference?.url.includes("{theme}"), `${route.id}: reference url needs {theme}`);
    assert.ok(route.subject?.url.includes("{theme}"), `${route.id}: subject url needs {theme}`);
    assert.ok(route.reference?.regionSelector, `${route.id}: reference region selector`);
    assert.ok(route.subject?.regionSelector, `${route.id}: subject region selector`);
    assert.ok(
      (route.reference?.ready || []).length > 0,
      `${route.id}: reference readiness selector`,
    );
    assert.ok((route.subject?.ready || []).length > 0, `${route.id}: subject readiness selector`);
    assert.ok(route.region?.heightPx > 0, `${route.id}: pinned region height`);
    assert.ok(route.region?.heightSource, `${route.id}: region height source`);
    assert.ok(route.fixture, `${route.id}: fixture id`);
    assert.ok(routes.fixtures[route.fixture], `${route.id}: unknown fixture ${route.fixture}`);

    const landmarkNames = Object.keys(route.reference.landmarks || {}).sort();
    assert.deepEqual(
      landmarkNames,
      Object.keys(route.subject.landmarks || {}).sort(),
      `${route.id}: both sides must expose the same landmarks`,
    );
    assert.ok(landmarkNames.includes("root"), `${route.id}: root landmark`);
    for (const [name, tolerance] of Object.entries(route.landmarkTolerance || {})) {
      assert.ok(landmarkNames.includes(name), `${route.id}: tolerance for unknown landmark ${name}`);
      assert.ok(tolerance >= 0, `${route.id}: tolerance must be >= 0`);
    }
    for (const landmark of landmarkNames) {
      assert.ok(
        route.landmarkTolerance?.[landmark] !== undefined,
        `${route.id}: landmark ${landmark} needs an explicit tolerance`,
      );
    }
  }
});

test("the capture state matches the URLs the harness opens", () => {
  for (const route of routes.routes) {
    const state = route.state || {};
    const urls = `${route.reference.url} ${route.subject.url}`;
    if (state.zoom !== undefined) {
      assert.ok(
        urls.includes(`zoom=${state.zoom}`),
        `${route.id}: both URLs must request zoom=${state.zoom}`,
      );
    }
    if (state.statusBar === false) {
      assert.ok(
        route.subject.url.includes("statusbar=off"),
        `${route.id}: a capture without the optional status bar must pass statusbar=off`,
      );
    }
    if (state.sidebar === "closed") {
      assert.ok(
        (route.subject.prepare || []).some(
          (action) => action.action === "click" && action.selector,
        ),
        `${route.id}: closing the sidebar needs an explicit prepare click`,
      );
    }
    const docName = new URL(route.reference.url, "https://example.invalid")
      .searchParams.get("doc")
      ?.split("/")
      .pop();
    assert.ok(
      docName && routes.fixtures[route.fixture].path.endsWith(docName),
      `${route.id}: the doc served to the reference must be the pinned fixture`,
    );
  }
});

test("exclusions and differences use the documented taxonomy", () => {
  for (const route of routes.routes) {
    for (const exclusion of route.exclusions || []) {
      assert.ok(["reference", "subject"].includes(exclusion.scope), exclusion.selector);
      assert.ok(exclusion.selector, "an exclusion needs a selector");
      assert.ok(exclusion.reason, "an exclusion needs a reason");
      assert.ok(
        ALLOWED_EXCLUSION_CATEGORIES.has(exclusion.category),
        `unknown exclusion category ${exclusion.category}`,
      );
    }
    for (const difference of route.differences || []) {
      assert.ok(
        ALLOWED_DIFFERENCE_CATEGORIES.has(difference.category),
        `unknown difference category ${difference.category}`,
      );
      assert.equal(typeof difference.blocking, "boolean", difference.id);
      assert.ok(["open", "fixed", "accepted"].includes(difference.status), difference.id);
      assert.ok(difference.evidence, `${difference.id}: needs file:line evidence`);
      if (difference.category === "demo-scaffolding") {
        assert.equal(difference.blocking, false, `${difference.id}: scaffolding never blocks`);
      }
    }
    const exclusions = route.exclusions || [];
    assert.ok(
      exclusions.some((exclusion) => exclusion.scope === "subject"),
      `${route.id}: the demo shell must be excluded explicitly`,
    );
    assert.ok(
      exclusions.some((exclusion) => exclusion.scope === "reference"),
      `${route.id}: the docs preview shell must be excluded explicitly`,
    );
  }
});

test("the fixture is byte-pinned and exists on both sides", () => {
  for (const [id, fixture] of Object.entries(routes.fixtures)) {
    const absolute = path.join(rootPath, fixture.path);
    assert.ok(existsSync(absolute), `${id}: ${fixture.path} is missing`);
    const digest = createHash("sha256").update(readFileSync(absolute)).digest("hex");
    assert.equal(digest, fixture.sha256, `${id}: sha256 drifted; update routes.json deliberately`);
    assert.equal(readFileSync(absolute).byteLength, fixture.bytes, `${id}: byte length drifted`);
  }
});

test("the demo route pins the upstream region height and reads capture state", () => {
  const page = read("apps/demo/src/pages/PdfViewerPage.vue");
  assert.match(page, /data-route-region="viewer"/, "region marker is missing");
  assert.match(page, /data-testid="pdf-viewer-region"/, "region test id is missing");
  assert.match(page, /height:\s*560px/, "region height must be pinned to 560px");
  assert.doesNotMatch(page, /height:\s*80vh/, "the viewport-derived height must be gone");
  assert.match(page, /readRouteCapture/, "the page must read the capture state");
  assert.match(page, /:default-zoom="defaultZoom"/);
  assert.match(page, /:show-status-bar="showStatusBar"/);

  const capture = read("apps/demo/src/lib/route-capture.ts");
  for (const parameter of ["sample", "zoom", "statusbar"]) {
    assert.ok(capture.includes(parameter), `route capture must read ?${parameter}=`);
  }
  assert.match(capture, /statusBar:\s*true/, "the default state must stay unchanged");
});

test("the reference app renders the pinned upstream wrapper through a route entry", () => {
  assert.ok(existsSync(new URL("tools/extend-reference/route.html", ROOT)));
  const html = read("tools/extend-reference/route.html");
  assert.match(html, /src="\/src\/route-main\.tsx"/);

  const routeMain = read("tools/extend-reference/src/route-main.tsx");
  assert.match(routeMain, /RouteApp/);

  const route = read("tools/extend-reference/src/routes/PdfViewerRoute.tsx");
  assert.match(route, /from "@\/components\/extend\/pdf-viewer"/);
  assert.match(route, /className="h-\[560px\]"/, "the upstream preview height must be kept");
  assert.match(route, /data-slot="component-preview"/, "the docs preview shell must be reproduced");

  const vite = read("tools/extend-reference/vite.config.ts");
  assert.match(vite, /alias/);
  assert.match(vite, /routeEntries/);
  assert.match(vite, /src\/upstream/);
});

test("the vendored upstream tree is complete and attributed", () => {
  const required = [
    "tools/extend-reference/src/upstream/components/extend/pdf-viewer.tsx",
    "tools/extend-reference/src/upstream/components/extend/document-viewer-sidebar.tsx",
    "tools/extend-reference/src/upstream/components/icon-placeholder.tsx",
    "tools/extend-reference/src/upstream/components/ui/button.tsx",
    "tools/extend-reference/src/upstream/components/ui/dropdown-menu.tsx",
    "tools/extend-reference/src/upstream/components/ui/input.tsx",
    "tools/extend-reference/src/upstream/components/ui/popover.tsx",
    "tools/extend-reference/src/upstream/components/ui/scroll-area.tsx",
    "tools/extend-reference/src/upstream/components/ui/select.tsx",
    "tools/extend-reference/src/upstream/components/ui/separator.tsx",
    "tools/extend-reference/src/upstream/components/ui/spinner.tsx",
    "tools/extend-reference/src/upstream/components/ui/tooltip.tsx",
    "tools/extend-reference/src/upstream/registry/new-york-v4/ui/dropdown-menu.tsx",
    "tools/extend-reference/src/upstream/registry/new-york-v4/ui/popover.tsx",
    "tools/extend-reference/src/upstream/registry/new-york-v4/ui/select.tsx",
    "tools/extend-reference/src/upstream/lib/utils.ts",
    "tools/extend-reference/src/upstream/lib/config.ts",
    "tools/extend-reference/src/upstream/lib/pdf-thumbnail-utils.ts",
  ];
  for (const file of required) {
    assert.ok(existsSync(new URL(file, ROOT)), `${file} is missing`);
    const content = read(file);
    assert.ok(
      content.startsWith("// Copied from extend-hq/ui @ " + routes.referenceCommit),
      `${file} is missing the attribution header`,
    );
  }

  const thumbnailUtils = read("tools/extend-reference/src/upstream/lib/pdf-thumbnail-utils.ts");
  assert.match(thumbnailUtils, /new URL\("\/pdfium\.wasm", window\.location\.href\)/);
  assert.doesNotMatch(thumbnailUtils, /cdn\.jsdelivr\.net/, "the wasm must be served locally");

  const attribution = readJson("docs/upstream/attribution.json");
  const entry = attribution.index.find((row) =>
    row.target === "tools/extend-reference/src/upstream/",
  );
  assert.ok(entry, "docs/upstream/attribution.json must index the vendored tree");
  assert.equal(entry.status, "copied");
  assert.ok(
    read("docs/upstream/extend-ui.md").includes(entry.upstreamFile),
    "the human index must list the vendored tree",
  );
});

test("reference dependencies are pinned to the upstream public versions", () => {
  const dependencies = referencePackage.dependencies;
  assert.equal(dependencies["@embedpdf/core"], "2.15.0");
  for (const plugin of [
    "plugin-document-manager",
    "plugin-interaction-manager",
    "plugin-render",
    "plugin-rotate",
    "plugin-scroll",
    "plugin-search",
    "plugin-selection",
    "plugin-thumbnail",
    "plugin-tiling",
    "plugin-viewport",
    "plugin-zoom",
  ]) {
    assert.equal(dependencies[`@embedpdf/${plugin}`], "2.15.0", `@embedpdf/${plugin}`);
  }
  assert.equal(dependencies["pdf-lib"], "1.17.1");
  assert.equal(dependencies["@base-ui/react"], "1.4.1");
  assert.doesNotMatch(
    JSON.stringify(dependencies),
    /@extend-ai\//,
    "product/reference apps must not depend on private packages",
  );
  assert.equal(referencePackage.scripts.build, "node scripts/copy-pdfium.mjs && vite build");
  assert.ok(existsSync(new URL("tools/extend-reference/scripts/copy-pdfium.mjs", ROOT)));
});

test("the harness entry point, gate modes and docs stay wired", () => {
  assert.equal(rootPackage.scripts["ui:parity:routes"], "node scripts/ci/route-parity.mjs");
  assert.ok(
    rootPackage.scripts["ui:parity:routes:report"].includes("--gate-mode off"),
    "the report-only entry point must stay explicit",
  );
  const harness = read("scripts/ci/route-parity.mjs");
  assert.match(harness, /route-parity\.py/);
  assert.match(harness, /--filter", "demo\.\.\."/, "the subject build must include its workspace deps");
  assert.match(harness, /buildReferenceApp\("@aou\/extend-reference"\)/);

  const script = read("scripts/ci/route-parity.py");
  assert.match(script, /--gate-mode/);
  assert.match(script, /choices=\["all", "ready", "off"\]/);
  assert.match(script, /routes-matrix\.json/);
  assert.match(script, /fixture .* sha256 mismatch|sha256 mismatch/);
  assert.match(script, /dominant_colors/);
  assert.match(script, /chrome/);

  const routesFile = path.join(rootPath, "tools/parity-fixtures/routes.json");
  assert.ok(existsSync(routesFile));

  const docs = read("docs/testing/ui-parity-routes.md");
  assert.match(docs, /Product region/);
  assert.match(docs, /R2/);
  assert.match(docs, /report-only/);
  assert.ok(read("docs/INDEX.md").includes("testing/ui-parity-routes.md"));
  assert.ok(read("docs/testing/ui-parity-ci.md").includes("ui-parity-routes.md"));
});

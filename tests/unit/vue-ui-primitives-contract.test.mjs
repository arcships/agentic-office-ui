/**
 * Issue #157 — primitive layer contract.
 *
 * Static assertions for the parts that the runtime suite cannot see:
 * dependency shape, no-Tailwind and token-only styling rules, public entry
 * wiring, and the bundle budget. The component suite covers behaviour.
 */

import assert from "node:assert/strict";
import { readFileSync, readdirSync, statSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { gzipSync } from "node:zlib";

const ROOT = new URL("../../", import.meta.url);
const UI_DIR = new URL("packages/vue-ui/src/components/ui/", ROOT);
const UI_CSS = new URL("packages/vue-ui/src/styles/ui.css", ROOT);
const packageJson = JSON.parse(
  readFileSync(new URL("packages/vue-ui/package.json", ROOT), "utf8"),
);
const viteConfig = readFileSync(
  new URL("packages/vue-ui/vite.config.ts", ROOT),
  "utf8",
);
const indexSource = readFileSync(
  new URL("packages/vue-ui/src/index.ts", ROOT),
  "utf8",
);
const uiCss = readFileSync(UI_CSS, "utf8");
const uiCssCode = uiCss.replace(/\/\*[\s\S]*?\*\//g, "");
const budget = JSON.parse(
  readFileSync(
    new URL("scripts/ci/vue-ui-primitives-size-budget.json", ROOT),
    "utf8",
  ),
);

function read(relative) {
  return readFileSync(new URL(relative, ROOT), "utf8");
}

function primitiveFiles() {
  return readdirSync(UI_DIR)
    .filter((name) => name.endsWith(".vue"))
    .map((name) => ({
      name,
      source: readFileSync(new URL(name, UI_DIR), "utf8"),
    }));
}

const PRIMITIVES = [
  "OfficeButton.vue",
  "OfficeIconButton.vue",
  "OfficeInput.vue",
  "OfficeTextarea.vue",
  "OfficeSelect.vue",
  "OfficeDropdownMenu.vue",
  "OfficeContextMenu.vue",
  "OfficeDialog.vue",
  "OfficeSheet.vue",
  "OfficePopover.vue",
  "OfficeTooltip.vue",
  "OfficeTabs.vue",
  "OfficeToggle.vue",
  "OfficeBadge.vue",
  "OfficeCard.vue",
  "OfficeSeparator.vue",
  "OfficeSkeleton.vue",
  "OfficeSpinner.vue",
  "OfficeScrollArea.vue",
  "OfficeToastProvider.vue",
];

test("reka-ui is a runtime dependency and stays external in the library build", () => {
  assert.equal(typeof packageJson.dependencies["reka-ui"], "string");
  assert.match(packageJson.dependencies["reka-ui"], /^\^2\.11/u);
  assert.equal(typeof packageJson.dependencies["@lucide/vue"], "string");
  assert.match(viteConfig, /"reka-ui"/u);
  assert.match(viteConfig, /external:\s*\[/u);
  for (const forbidden of ["tailwindcss", "tailwind-merge", "clsx", "prettier-plugin-tailwindcss"]) {
    assert.equal(
      packageJson.dependencies[forbidden],
      undefined,
      `${forbidden} must not be a runtime dependency`,
    );
  }
  assert.equal(
    packageJson.dependencies["class-variance-authority"],
    undefined,
    "the variant helper is local; no extra runtime dependency",
  );
});

test("the public entry exports every primitive and keeps the existing entries", () => {
  assert.match(indexSource, /export \* from "\.\/components\/ui"/u);
  assert.match(indexSource, /import "\.\/styles\/ui\.css"/u);
  assert.equal(packageJson.exports["."].types, "./dist/index.d.ts");
  assert.equal(packageJson.exports["./style.css"], "./dist/style.css");
  assert.equal(packageJson.exports["./theme.css"], "./dist/theme.css");
  assert.equal(packageJson.exports["./tokens.css"], "./dist/tokens.css");

  const barrel = read("packages/vue-ui/src/components/ui/index.ts");
  for (const file of PRIMITIVES) {
    const exportName = file.replace(".vue", "");
    assert.match(
      barrel,
      new RegExp(`export \\{ default as ${exportName} \\}`),
      `${exportName} must be exported from the ui barrel`,
    );
  }
  assert.match(barrel, /export \{ useOfficeToast \}/u);
  assert.match(barrel, /OfficeVariant/u);
  assert.match(barrel, /OfficeButtonSize/u);
});

test("all 20 primitive files exist and carry a data-slot marker", () => {
  const files = primitiveFiles();
  assert.equal(files.length, PRIMITIVES.length + 1, "unexpected extra SFC in ui/");
  for (const name of PRIMITIVES) {
    const source = read(`packages/vue-ui/src/components/ui/${name}`);
    if (name !== "OfficeToastProvider.vue") {
      assert.match(source, /data-slot=/u, `${name} must set data-slot`);
    }
    if (name !== "OfficeToastProvider.vue") {
      assert.match(source, /class\?:/u, `${name} must accept a class override`);
    }
    assert.doesNotMatch(
      source,
      /@tailwind|@apply|preflight/u,
      `${name} must not use Tailwind`,
    );
  }
});

test("the variant and size contract is complete", () => {
  for (const variant of [
    "default",
    "secondary",
    "ghost",
    "outline",
    "destructive",
    "link",
  ]) {
    assert.match(
      uiCss,
      new RegExp(`\\.office-button--${variant}\\b`),
      `missing button variant ${variant}`,
    );
  }
  for (const size of ["xs", "sm", "lg", "icon", "icon-sm", "icon-lg"]) {
    assert.match(
      uiCss,
      new RegExp(`\\.office-button--${size}\\b`),
      `missing button size ${size}`,
    );
  }
  for (const size of ["xs", "sm", "lg"]) {
    assert.match(
      uiCss,
      new RegExp(`\\.office-input--${size}\\b`),
      `missing input size ${size}`,
    );
    assert.match(
      uiCss,
      new RegExp(`\\.office-select__trigger--${size}\\b`),
      `missing select size ${size}`,
    );
  }
});

test("primitives use tokens instead of Tailwind utilities or literal visuals", () => {
  assert.doesNotMatch(uiCssCode, /@tailwind|@apply|preflight/u);
  assert.doesNotMatch(uiCssCode, /#[0-9a-f]{3,8}\b/iu, "no literal hex colours");
  assert.doesNotMatch(uiCssCode, /\brgba?\(/u, "no literal rgb colours");
  assert.doesNotMatch(
    uiCssCode,
    /font-size:(?![ \t]*(?:var\(|calc\())/u,
    "font-size must come from a token",
  );
  assert.doesNotMatch(
    uiCssCode,
    /border-radius:[ \t]*[\d.]+(?:px|rem)/u,
    "radius must come from a token",
  );
  assert.doesNotMatch(
    uiCssCode,
    /(?:^|\n)\s*(?:html|body|:root|\*)\s*\{/u,
    "the primitive stylesheet must not ship a global reset",
  );
  assert.match(uiCssCode, /var\(--office-radius-lg\)/u);
  assert.match(uiCssCode, /var\(--office-text-sm\)/u);
  assert.match(uiCssCode, /var\(--office-control-height\)/u);
  assert.match(uiCssCode, /var\(--office-shadow-xs\)/u);

  for (const { name, source } of primitiveFiles()) {
    assert.doesNotMatch(source, /#[0-9a-f]{3,8}\b/iu, `${name} literal colour`);
    assert.doesNotMatch(source, /\brgba?\(/u, `${name} literal rgb colour`);
    assert.doesNotMatch(
      source,
      /@lucide\/vue/u,
      `${name} must use OfficeIcon, not the icon pack directly`,
    );
  }
});

test("touch targets and focus-visible rings are part of the interaction contract", () => {
  assert.match(uiCss, /@media \(pointer: coarse\)/u);
  assert.match(uiCss, /--office-ui-touch-target:\s*2\.75rem/u);
  assert.match(uiCss, /var\(--office-ui-touch-target\)/u);
  assert.match(uiCss, /:focus-visible/u);
  assert.match(uiCss, /var\(--office-ring-width\)/u);
  assert.match(uiCss, /prefers-reduced-motion: reduce/u);
});

test("the new primitive layer stays inside the 35 KB gzip budget", () => {
  const dist = new URL("packages/vue-ui/dist/", ROOT);
  const js = readFileSync(new URL("index.js", dist));
  const css = readFileSync(new URL("style.css", dist));
  const jsGzip = gzipSync(js, { level: 9 }).length;
  const cssGzip = gzipSync(css, { level: 9 }).length;
  const total = jsGzip + cssGzip;
  const delta = total - budget.baselineGzipBytes;

  assert.ok(
    delta >= 0,
    `baseline (${budget.baselineGzipBytes}) exceeds the measured build (${total})`,
  );
  assert.ok(
    delta <= budget.budgetBytes,
    `primitive layer added ${delta} gzip bytes; budget is ${budget.budgetBytes}`,
  );
  assert.ok(
    total < 60_000,
    `unexpectedly large vue-ui bundle (${total} gzip bytes)`,
  );

  // reka-ui stays external: its runtime source must not be inlined.
  assert.match(js.toString("utf8"), /from "reka-ui"/u);
  assert.doesNotMatch(js.toString("utf8"), /reka-dialog-content/u);
  assert.doesNotMatch(js.toString("utf8"), /@floating-ui/u);

  assert.ok(statSync(new URL("style.css", dist)).size > 0);
});

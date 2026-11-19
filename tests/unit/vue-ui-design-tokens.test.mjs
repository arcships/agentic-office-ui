// Issue #156 — design token contract.
//
// The CSS is the product here: these assertions pin the token names, the
// Extend/shadcn values, the scoping rules and the package wiring. They read the
// source files so they run without a build (the component suite covers the
// built artifacts and runtime behaviour).

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

const ROOT = new URL("../../", import.meta.url);

function read(relativePath) {
  return readFileSync(new URL(relativePath, ROOT), "utf8");
}

function stripComments(source) {
  return source.replace(/\/\*[\s\S]*?\*\//g, "");
}

/** Return the declarations of the first block whose header is `selector {`. */
function findBlock(source, selector) {
  const marker = `${selector} {`;
  const start = source.indexOf(marker);
  if (start < 0) return null;
  const open = source.indexOf("{", start);
  let depth = 0;
  for (let index = open; index < source.length; index += 1) {
    if (source[index] === "{") depth += 1;
    else if (source[index] === "}") {
      depth -= 1;
      if (depth === 0) return source.slice(open + 1, index);
    }
  }
  return null;
}

function parseDeclarations(block) {
  const declarations = new Map();
  for (const part of block.split(";")) {
    const colon = part.indexOf(":");
    if (colon < 0) continue;
    const name = part.slice(0, colon).trim();
    if (!name.startsWith("--")) continue;
    declarations.set(name, part.slice(colon + 1).replace(/\s+/g, " ").trim());
  }
  return declarations;
}

function tokensOf(source, selector) {
  const block = findBlock(source, selector);
  assert.ok(block, `missing CSS block: ${selector}`);
  return parseDeclarations(block);
}

const tokensSource = read("packages/vue-ui/src/styles/tokens.css");
const themeSource = read("packages/vue-ui/src/styles/theme.css");
const packageJson = JSON.parse(read("packages/vue-ui/package.json"));
const viteConfig = read("packages/vue-ui/vite.config.ts");

const light = tokensOf(stripComments(tokensSource), "[data-office-theme]");
const compact = tokensOf(
  stripComments(tokensSource),
  '[data-office-theme][data-office-density="compact"]',
);
const cleanedTheme = stripComments(themeSource);
const dark = tokensOf(cleanedTheme, '[data-office-theme="dark"]');
const systemMedia = findBlock(cleanedTheme, "@media (prefers-color-scheme: dark)");
assert.ok(systemMedia, "theme.css must ship a prefers-color-scheme: dark block");
const systemDark = tokensOf(systemMedia, '[data-office-theme="system"]');

const REQUIRED_TOKENS = [
  "background",
  "foreground",
  "card",
  "card-foreground",
  "popover",
  "popover-foreground",
  "primary",
  "primary-foreground",
  "secondary",
  "secondary-foreground",
  "muted",
  "muted-foreground",
  "accent",
  "accent-foreground",
  "destructive",
  "destructive-foreground",
  "border",
  "input",
  "ring",
  "sidebar",
  "sidebar-foreground",
  "sidebar-primary",
  "sidebar-primary-foreground",
  "sidebar-accent",
  "sidebar-accent-foreground",
  "sidebar-border",
  "sidebar-ring",
  "success",
  "success-foreground",
  "warning",
  "warning-foreground",
  "info",
  "info-foreground",
];

test("semantic tokens cover the required contract in both light and dark", () => {
  for (const token of REQUIRED_TOKENS) {
    const name = `--office-${token}`;
    assert.ok(light.has(name), `tokens.css light scope is missing ${name}`);
    assert.ok(dark.has(name), `theme.css dark scope is missing ${name}`);
    assert.ok(systemDark.has(name), `prefers-color-scheme dark scope is missing ${name}`);
  }
});

test("radius, ring and elevation match the pinned Extend reference", () => {
  assert.equal(light.get("--office-radius"), "0.625rem");
  assert.equal(light.get("--office-radius-sm"), "calc(var(--office-radius) * 0.6)");
  assert.equal(light.get("--office-radius-md"), "calc(var(--office-radius) * 0.8)");
  assert.equal(light.get("--office-radius-lg"), "var(--office-radius)");
  assert.equal(light.get("--office-radius-xl"), "calc(var(--office-radius) * 1.4)");
  assert.equal(light.get("--office-radius-2xl"), "calc(var(--office-radius) * 1.8)");
  assert.equal(light.get("--office-radius-3xl"), "calc(var(--office-radius) * 2.2)");
  assert.equal(light.get("--office-radius-4xl"), "calc(var(--office-radius) * 2.6)");

  assert.equal(light.get("--office-ring-width"), "2px");
  assert.equal(light.get("--office-ring-offset-width"), "1px");
  assert.equal(light.get("--office-border-width"), "1px");

  // Tailwind v4 shadow-xs / shadow-sm, the utilities used by Extend's Button.
  assert.equal(light.get("--office-shadow-xs"), "0 1px 2px 0 rgb(0 0 0 / 0.05)");
  assert.equal(
    light.get("--office-shadow-sm"),
    "0 1px 3px 0 rgb(0 0 0 / 0.1), 0 1px 2px -1px rgb(0 0 0 / 0.1)",
  );
});

test("light colours match Extend globals.css", () => {
  assert.equal(light.get("--office-background"), "oklch(1 0 0)");
  assert.equal(light.get("--office-foreground"), "oklch(26.9% 0 0)"); // neutral-800
  assert.equal(light.get("--office-primary"), "oklch(26.9% 0 0)"); // neutral primary
  assert.equal(light.get("--office-primary-foreground"), "oklch(98.5% 0 0)"); // neutral-50
  assert.equal(light.get("--office-secondary"), "color-mix(in oklab, black 4%, transparent)");
  assert.equal(light.get("--office-muted"), "color-mix(in oklab, black 4%, transparent)");
  assert.equal(light.get("--office-accent"), "color-mix(in oklab, black 4%, transparent)");
  assert.equal(light.get("--office-border"), "color-mix(in oklab, black 8%, transparent)"); // 8% alpha border
  assert.equal(light.get("--office-input"), "color-mix(in oklab, black 10%, transparent)");
  assert.equal(light.get("--office-ring"), "oklch(70.8% 0 0)"); // neutral-400
  assert.equal(light.get("--office-muted-foreground"), "color-mix(in srgb, oklch(55.6% 0 0) 90%, black)");
  assert.equal(light.get("--office-destructive"), "oklch(63.7% 0.237 25.331)"); // red-500
  assert.equal(light.get("--office-destructive-foreground"), "oklch(50.5% 0.213 27.518)"); // red-700
  assert.equal(light.get("--office-sidebar"), "oklch(98.5% 0 0)");
  assert.equal(light.get("--office-sidebar-border"), "color-mix(in oklab, black 6%, transparent)");
  assert.equal(light.get("--office-success"), "oklch(69.6% 0.17 162.48)");
  assert.equal(light.get("--office-success-foreground"), "oklch(50.8% 0.118 165.612)");
  assert.equal(light.get("--office-warning"), "oklch(76.9% 0.188 70.08)");
  assert.equal(light.get("--office-warning-foreground"), "oklch(55.5% 0.163 48.998)");
  assert.equal(light.get("--office-info"), "oklch(62.3% 0.214 259.815)");
  assert.equal(light.get("--office-info-foreground"), "oklch(48.8% 0.243 264.376)");
  assert.equal(light.get("--office-chart-1"), "oklch(80.9% 0.105 251.813)");
  assert.equal(light.get("--office-chart-5"), "oklch(42.4% 0.199 265.638)");
});

test("dark colours match Extend globals.css", () => {
  assert.equal(dark.get("--office-background"), "oklch(14.5% 0 0)");
  assert.equal(dark.get("--office-foreground"), "oklch(97% 0 0)"); // neutral-100
  assert.equal(dark.get("--office-primary"), "oklch(97% 0 0)");
  assert.equal(dark.get("--office-primary-foreground"), "oklch(26.9% 0 0)");
  assert.equal(dark.get("--office-secondary"), "color-mix(in oklab, white 4%, transparent)");
  assert.equal(dark.get("--office-muted"), "color-mix(in oklab, white 4%, transparent)");
  assert.equal(dark.get("--office-accent"), "color-mix(in oklab, white 4%, transparent)");
  assert.equal(dark.get("--office-border"), "color-mix(in oklab, white 6%, transparent)");
  assert.equal(dark.get("--office-input"), "color-mix(in oklab, white 8%, transparent)");
  assert.equal(dark.get("--office-ring"), "oklch(55.6% 0 0)"); // neutral-500
  assert.equal(dark.get("--office-destructive-foreground"), "oklch(70.4% 0.191 22.216)");
  assert.equal(dark.get("--office-success-foreground"), "oklch(76.5% 0.177 163.223)");
  assert.equal(dark.get("--office-warning-foreground"), "oklch(82.8% 0.189 84.429)");
  assert.equal(dark.get("--office-info-foreground"), "oklch(70.7% 0.165 254.624)");
});

test("system dark and explicit dark cannot drift", () => {
  assert.deepEqual(
    Object.fromEntries(systemDark),
    Object.fromEntries(dark),
    "data-office-theme=system dark values must equal data-office-theme=dark",
  );
  assert.match(cleanedTheme, /\[data-office-theme="dark"\]\s*\{[^}]*color-scheme:\s*dark/u);
  assert.match(systemMedia, /color-scheme:\s*dark/u);
});

test("theme switching is scoped and never touches the host globals", () => {
  for (const [name, source] of [
    ["tokens.css", tokensSource],
    ["theme.css", themeSource],
  ]) {
    const css = stripComments(source);
    assert.doesNotMatch(css, /:root\b/u, `${name} must not declare :root tokens`);
    assert.doesNotMatch(css, /(^|\n)\s*html\s*\{/u, `${name} must not style html`);
    assert.doesNotMatch(css, /(^|\n)\s*body\s*\{/u, `${name} must not style body`);
    assert.doesNotMatch(css, /(^|\n)\s*\*\s*\{/u, `${name} must not apply a global reset`);
    assert.doesNotMatch(css, /\.dark\b/u, `${name} must not rely on a global .dark class`);
    assert.doesNotMatch(css, /@font-face/u, `${name} must not bundle font files`);
    assert.doesNotMatch(css, /url\(/u, `${name} must not reference external assets`);
  }

  assert.match(themeSource, /@import "\.\/tokens\.css";/u);
  assert.match(themeSource, /\[data-office-theme="dark"\]/u);
  assert.match(themeSource, /@media \(prefers-color-scheme: dark\)/u);
  assert.match(themeSource, /\[data-office-theme="system"\]/u);
  assert.match(themeSource, /\[data-office-theme\] ::selection/u);

  // Light is the default carried by tokens.css; every explicit scope still
  // opts the browser's native controls into the right scheme.
  assert.match(cleanedTheme, /\[data-office-theme="light"\][\s\S]{0,80}color-scheme:\s*light/u);
});

test("font tokens default to Geist and stay host-overridable", () => {
  assert.match(light.get("--office-font-sans"), /^Geist,/u);
  assert.match(light.get("--office-font-sans"), /sans-serif$/u);
  assert.match(light.get("--office-font-mono"), /^"Geist Mono",/u);
  assert.match(light.get("--office-font-mono"), /monospace$/u);
  assert.equal(light.get("--office-text-sm"), "0.875rem");
  assert.equal(light.get("--office-leading-sm"), "1.25rem");
});

test("density tokens switch control metrics without changing colour", () => {
  assert.equal(light.get("--office-control-height"), "2.25rem");
  assert.equal(compact.get("--office-control-height"), "2rem");
  assert.equal(compact.get("--office-control-height-sm"), "1.75rem");
  assert.equal(compact.get("--office-control-padding-inline"), "0.625rem");
  assert.equal(compact.get("--office-control-gap"), "0.375rem");
  assert.equal(compact.has("--office-background"), false);
});

test("new component code uses tokens instead of bare visual values", () => {
  for (const file of [
    "packages/vue-ui/src/components/OfficeIcon.vue",
    "packages/vue-ui/src/components/OfficeThemeProvider.vue",
  ]) {
    const source = read(file);
    const style = source.match(/<style[^>]*>([\s\S]*?)<\/style>/u)?.[1] ?? "";
    assert.doesNotMatch(style, /#[0-9a-f]{3,8}\b/iu, `${file} must not hard-code hex colours`);
    assert.doesNotMatch(style, /\brgba?\(/u, `${file} must not hard-code rgb colours`);
    assert.doesNotMatch(style, /font-size\s*:/u, `${file} must not hard-code font-size`);
    assert.doesNotMatch(style, /border-radius\s*:\s*[\d.]+(?:px|rem)/u, `${file} must not hard-code radius`);
  }
  assert.match(read("packages/vue-ui/src/components/OfficeIcon.vue"), /var\(--office-icon-size/u);
});

test("icon adapter only names the icons it ships", () => {
  const source = read("packages/vue-ui/src/icons.ts");
  assert.match(source, /from "@lucide\/vue"/u);
  assert.doesNotMatch(source, /import\s+\*\s+as/u, "never import the whole icon pack");
  assert.match(source, /OFFICE_ICONS = Object\.freeze\(/u);
});

test("package exports theme.css and keeps style.css backward compatible", () => {
  assert.equal(packageJson.exports["./style.css"], "./dist/style.css");
  assert.equal(packageJson.exports["./theme.css"], "./dist/theme.css");
  assert.equal(packageJson.exports["./tokens.css"], "./dist/tokens.css");
  assert.deepEqual(packageJson.files, ["dist"]);
  assert.ok(packageJson.sideEffects.includes("./dist/style.css"));
  assert.ok(packageJson.sideEffects.includes("./dist/theme.css"));
  assert.ok(packageJson.sideEffects.includes("./dist/tokens.css"));
  assert.equal(typeof packageJson.dependencies["@lucide/vue"], "string");
  assert.match(packageJson.dependencies["@lucide/vue"], /^\^?\d+\./u);
  assert.match(viteConfig, /"@lucide\/vue"/u, "@lucide/vue must stay external");
  assert.match(viteConfig, /officeThemeStyles/u);
  assert.match(viteConfig, /dist\/theme\.css/u);
});

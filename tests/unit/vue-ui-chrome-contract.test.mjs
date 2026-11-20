/**
 * Issue #158 — shared Office chrome contract.
 *
 * Static assertions for the parts the runtime suite cannot see: the toolbar /
 * sidebar geometry contract, token-only styling, public entry wiring and the
 * shared catalog keys. Component behaviour lives in
 * `tests/component/vue-ui-chrome.test.mjs`.
 */

import assert from "node:assert/strict"
import { readFileSync } from "node:fs"
import test from "node:test"
import { gzipSync } from "node:zlib"

const ROOT = new URL("../../", import.meta.url)
const CHROME_DIR = new URL("packages/vue-ui/src/components/chrome/", ROOT)
const CHROME_CSS = new URL("packages/vue-ui/src/styles/chrome.css", ROOT)
const UI_CSS = new URL("packages/vue-ui/src/styles/ui.css", ROOT)

const chromeCss = readFileSync(CHROME_CSS, "utf8")
const chromeCssCode = chromeCss.replace(/\/\*[\s\S]*?\*\//g, "")
const uiCss = readFileSync(UI_CSS, "utf8")
const tokensCss = readFileSync(
  new URL("packages/vue-ui/src/styles/tokens.css", ROOT),
  "utf8",
)
const packageJson = JSON.parse(
  readFileSync(new URL("packages/vue-ui/package.json", ROOT), "utf8"),
)
const contract = JSON.parse(
  readFileSync(
    new URL("scripts/ci/public-api-contract.json", ROOT),
    "utf8",
  ),
)

function read(relative) {
  return readFileSync(new URL(relative, ROOT), "utf8")
}

function stripComments(source) {
  return source.replace(/\/\*[\s\S]*?\*\//g, "")
}

/** Return the declarations of the first block whose header is `selector {`. */
function findBlock(source, selector) {
  const marker = `${selector} {`
  const start = source.indexOf(marker)
  if (start < 0) return null
  const open = source.indexOf("{", start)
  let depth = 0
  for (let index = open; index < source.length; index += 1) {
    if (source[index] === "{") depth += 1
    else if (source[index] === "}") {
      depth -= 1
      if (depth === 0) return source.slice(open + 1, index)
    }
  }
  return null
}

const CHROME_FILES = [
  "DocumentViewerSidebar.vue",
  "DocumentViewerSidebarSkeleton.vue",
  "OfficeEmptyState.vue",
  "OfficeErrorState.vue",
  "OfficeLoadingState.vue",
  "OfficePageControl.vue",
  "OfficeSearchPopover.vue",
  "OfficeStatusBar.vue",
  "OfficeToolbar.vue",
  "OfficeToolbarGroup.vue",
  "OfficeToolbarSeparator.vue",
  "OfficeZoomControl.vue",
]

test("every chrome component exists, accepts a class override and carries stable hooks", () => {
  for (const name of CHROME_FILES) {
    const source = read(`packages/vue-ui/src/components/chrome/${name}`)
    const exportName = name.replace(".vue", "")
    assert.match(
      source,
      new RegExp(`export interface ${exportName}Props`),
      `${name} must export its props interface`,
    )
    assert.match(source, /data-slot=/u, `${name} must set data-slot`)
    assert.match(source, /data-testid=/u, `${name} must set data-testid`)
    assert.match(source, /class\?:/u, `${name} must accept a class override`)
    assert.doesNotMatch(
      source,
      /@tailwind|@apply|preflight/u,
      `${name} must not use Tailwind`,
    )
    assert.doesNotMatch(
      source,
      /@lucide\/vue/u,
      `${name} must use OfficeIcon, not the icon pack directly`,
    )
    assert.doesNotMatch(source, /#[0-9a-f]{3,8}\b/iu, `${name} literal colour`)
    assert.doesNotMatch(source, /\brgba?\(/u, `${name} literal rgb colour`)
  }
})

test("the toolbar geometry matches the Extend reference", () => {
  const toolbar = findBlock(stripComments(chromeCss), ".office-toolbar")
  assert.ok(toolbar, "missing .office-toolbar block")
  assert.match(toolbar, /min-block-size:\s*var\(--office-ui-toolbar-min-height\)/u)
  assert.match(
    toolbar,
    /--office-ui-toolbar-min-height:\s*3rem/u,
    "min-height must be 48px",
  )
  assert.match(toolbar, /padding:\s*var\(--office-ui-toolbar-padding-block\)/u)
  assert.match(toolbar, /var\(--office-ui-toolbar-padding-inline\)/u)
  assert.match(toolbar, /--office-ui-toolbar-padding-block:\s*var\(--office-ui-space-2\)/u)
  assert.match(toolbar, /--office-ui-toolbar-padding-inline:\s*var\(--office-ui-space-3\)/u)
  assert.match(toolbar, /gap:\s*var\(--office-ui-toolbar-gap\)/u)
  assert.match(toolbar, /--office-ui-toolbar-gap:\s*var\(--office-ui-space-2\)/u)
  assert.match(toolbar, /border-block-end:\s*var\(--office-border-width\) solid var\(--office-border\)/u)
  assert.match(toolbar, /background:\s*var\(--office-background\)/u)
  assert.match(toolbar, /flex-wrap:\s*wrap/u)
  assert.match(toolbar, /min-inline-size:\s*0/u)

  for (const gap of ["sm", "md"]) {
    assert.match(
      stripComments(chromeCss),
      new RegExp(`\\.office-toolbar__group--gap-${gap} \\{`),
      `missing group gap ${gap}`,
    )
  }
  assert.match(
    stripComments(chromeCss),
    /\.office-toolbar__group--gap-sm \{\s*gap:\s*var\(--office-ui-space-1\)/u,
    "compact group gap must be 4px",
  )
  assert.match(
    stripComments(chromeCss),
    /\.office-toolbar__group--gap-md \{\s*gap:\s*var\(--office-ui-space-2\)/u,
    "wide group gap must be 8px",
  )
  assert.match(
    stripComments(chromeCss),
    /@media \(max-width: 30rem\)[\s\S]*?\.office-toolbar__trailing \{\s*flex:\s*1 1 100%/u,
    "the trailing cluster must take its own row at 390px",
  )
})

test("the sidebar contract is 160px with inline/overlay animation and ResizeObserver auto mode", () => {
  const sidebar = findBlock(stripComments(chromeCss), ".office-document-sidebar")
  assert.ok(sidebar, "missing .office-document-sidebar block")
  assert.match(sidebar, /--office-ui-sidebar-width:\s*10rem/u, "default width is 160px")
  assert.match(sidebar, /inline-size:\s*var\(--office-ui-sidebar-width\)/u)
  assert.match(sidebar, /--office-ui-sidebar-duration:\s*200ms/u)
  assert.match(sidebar, /margin-inline-start var\(--office-ui-sidebar-duration\) ease-out/u)
  assert.match(sidebar, /translate var\(--office-ui-sidebar-duration\) ease-out/u)
  assert.match(sidebar, /background:\s*var\(--office-sidebar\)/u)
  assert.match(sidebar, /border-inline-end:\s*var\(--office-border-width\) solid var\(--office-sidebar-border\)/u)

  assert.match(
    stripComments(chromeCss),
    /\.office-document-sidebar--overlay \{\s*position:\s*absolute/u,
  )
  assert.match(
    stripComments(chromeCss),
    /\.office-document-sidebar--inline\[data-sidebar-open="false"\] \{\s*margin-inline-start:\s*calc\(-1 \* var\(--office-ui-sidebar-width\)\)/u,
  )
  assert.match(
    stripComments(chromeCss),
    /\.office-document-sidebar--overlay\[data-sidebar-open="false"\] \{[\s\S]*?translate:\s*-100% 0/u,
  )
  assert.match(
    stripComments(chromeCss),
    /\[data-transitions-ready="false"\] \{\s*transition:\s*none/u,
  )
  assert.match(chromeCss, /prefers-reduced-motion: reduce/u)

  const sidebarSource = read("packages/vue-ui/src/components/chrome/DocumentViewerSidebar.vue")
  assert.match(sidebarSource, /breakpoint:\s*768/u)
  assert.match(sidebarSource, /mode:\s*"auto"/u)
  assert.match(sidebarSource, /useElementWidth/u)
  assert.match(sidebarSource, /container\?: HTMLElement \| null/u)
  assert.match(sidebarSource, /requestAnimationFrame/u)
  assert.match(sidebarSource, /data-sidebar-mode/u)
  assert.match(sidebarSource, /data-sidebar-open/u)
})

test("the chrome stylesheet stays token-only and ships no reset", () => {
  assert.doesNotMatch(chromeCssCode, /@tailwind|@apply|preflight/u)
  assert.doesNotMatch(chromeCssCode, /#[0-9a-f]{3,8}\b/iu, "no literal hex colours")
  assert.doesNotMatch(chromeCssCode, /\brgba?\(/u, "no literal rgb colours")
  assert.doesNotMatch(
    chromeCssCode,
    /font-size:(?![ \t]*(?:var\(|calc\())/u,
    "font-size must come from a token",
  )
  assert.doesNotMatch(
    chromeCssCode,
    /border-radius:[ \t]*[\d.]+(?:px|rem)/u,
    "radius must come from a token",
  )
  assert.doesNotMatch(
    chromeCssCode,
    /(?:^|\n)\s*(?:html|body|:root|\*)\s*\{/u,
    "the chrome stylesheet must not ship a global reset",
  )
  for (const token of [
    "--office-background",
    "--office-border",
    "--office-foreground",
    "--office-sidebar",
    "--office-sidebar-border",
    "--office-text-sm",
    "--office-radius-md",
  ]) {
    assert.match(chromeCssCode, new RegExp(`var\\(${token}\\)`), `missing ${token}`)
  }
})

test("IconButton keeps the 28/32px desktop sizes and a 44px coarse hit area", () => {
  const icon = findBlock(stripComments(uiCss), ".office-button--size-icon")
  const iconSmall = findBlock(stripComments(uiCss), ".office-button--size-icon-sm")
  assert.ok(icon)
  assert.ok(iconSmall)
  assert.match(icon, /inline-size:\s*var\(--office-ui-height-default\)/u)
  assert.match(icon, /height:\s*var\(--office-ui-height-default\)/u)
  assert.match(iconSmall, /inline-size:\s*var\(--office-ui-height-sm\)/u)
  assert.match(iconSmall, /height:\s*var\(--office-ui-height-sm\)/u)
  assert.match(tokensCss, /--office-control-height-sm:\s*2rem/u, "icon is 32px")
  assert.match(tokensCss, /--office-control-height-xs:\s*1\.75rem/u, "icon-sm is 28px")
  assert.match(uiCss, /--office-ui-touch-target:\s*2\.75rem/u)
  assert.match(uiCss, /@media \(pointer: coarse\)/u)
})

test("the public entries export every chrome component and the width composable", () => {
  const barrel = read("packages/vue-ui/src/components/chrome/index.ts")
  const index = read("packages/vue-ui/src/index.ts")
  assert.match(index, /export \* from "\.\/components\/chrome"/u)
  assert.match(index, /import "\.\/styles\/chrome\.css"/u)
  assert.match(index, /export \{ useElementWidth \}/u)
  for (const name of CHROME_FILES) {
    const exportName = name.replace(".vue", "")
    assert.match(
      barrel,
      new RegExp(`export \\{ default as ${exportName} \\}`),
      `${exportName} must be exported from the chrome barrel`,
    )
    assert.match(
      barrel,
      new RegExp(`export type \\{ ${exportName}Props \\}`),
      `${exportName}Props must be exported`,
    )
  }
  for (const typeName of [
    "OfficeSearchStatus",
    "OfficeSidebarMode",
    "OfficeSidebarResolvedMode",
    "OfficeToolbarGap",
    "OfficeZoomChange",
    "OfficeZoomFitMode",
  ]) {
    assert.match(barrel, new RegExp(typeName), `missing ${typeName}`)
  }
  assert.equal(packageJson.exports["./style.css"], "./dist/style.css")
})

test("SearchPopover exposes an inline mode with stable host hooks", () => {
  const source = read("packages/vue-ui/src/components/chrome/OfficeSearchPopover.vue")
  assert.match(source, /inline\?: boolean/u)
  assert.match(source, /testIdPrefix\?: string/u)
  assert.match(source, /inputType\?: string/u)
  assert.match(source, /statusText\?: string/u)
  assert.match(source, /office-search-popover__fields/u)
  assert.match(source, /office-toggle--outline/u)
})

test("the shared catalog ships chrome and zoom labels in English and zh-CN", () => {
  const i18n = read("packages/office-interaction/src/i18n.ts")
  for (const key of [
    "toolbar",
    "sidebar",
    "sidebarToggle",
    "statusBar",
    "fileName",
    "error",
    "warning",
    "retry",
    "emptyTitle",
    "emptyDescription",
    "errorTitle",
    "errorDescription",
    "loadingTitle",
    "loadingDescription",
  ]) {
    const occurrences = i18n.match(new RegExp(`\\n\\s+${key}:`, "gu")) ?? []
    assert.ok(occurrences.length >= 2, `chrome message ${key} must exist in en and zh-CN`)
  }
  for (const key of ["fitWidth", "fitPage", "actualSize", "percent"]) {
    const occurrences = i18n.match(new RegExp(`\\n\\s+${key}:`, "gu")) ?? []
    assert.ok(occurrences.length >= 2, `zoom message ${key} must exist in en and zh-CN`)
  }
  assert.match(i18n, /chrome: OfficeChromeMessages/u)
})

test("the chrome layer stays inside the #157 gzip budget", () => {
  const budget = JSON.parse(
    read("scripts/ci/vue-ui-primitives-size-budget.json"),
  )
  const js = readFileSync(new URL("packages/vue-ui/dist/index.js", ROOT))
  const css = readFileSync(new URL("packages/vue-ui/dist/style.css", ROOT))
  const total = gzipSync(js, { level: 9 }).length + gzipSync(css, { level: 9 }).length
  const delta = total - budget.baselineGzipBytes
  assert.ok(delta >= 0)
  assert.ok(
    delta <= budget.budgetBytes,
    `vue-ui added ${delta} gzip bytes; the #157/#158 budget is ${budget.budgetBytes}`,
  )
})

test("the public API fingerprint is recorded for the vue-ui entry", () => {
  const vueUi = contract.packages.find(
    (entry) => entry.name === "@arcships/vue-ui",
  )
  assert.ok(vueUi, "vue-ui must be part of the public API contract")
  assert.ok(vueUi.rootDeclarationFingerprint.count > 0)
  assert.match(vueUi.rootDeclarationFingerprint.sha256, /^[0-9a-f]{64}$/u)
  const declarations = read("packages/vue-ui/dist/components/chrome/index.d.ts")
  assert.match(declarations, /OfficeToolbar/u)
  assert.match(declarations, /DocumentViewerSidebar/u)
})

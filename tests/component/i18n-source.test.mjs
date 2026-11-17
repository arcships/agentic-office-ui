import assert from "node:assert/strict"
import { readdirSync, readFileSync, statSync } from "node:fs"
import path from "node:path"
import test from "node:test"

const i18n = await import(
  new URL("../../packages/office-interaction/src/i18n.ts", import.meta.url)
)

const ROOT = path.resolve(import.meta.dirname, "../..")
const VIEWER_PACKAGES = ["vue-docx", "vue-xlsx", "vue-pptx", "vue-pdf"]

const VIEWERS = [
  ["vue-docx", "src/components/DocxViewer.vue", "provideDocxMessages"],
  ["vue-docx", "src/components/DocxEditor.vue", "provideDocxMessages"],
  ["vue-docx", "src/components/DocxDocumentSurface.vue", "provideDocxMessages"],
  ["vue-xlsx", "src/components/XlsxViewer.vue", "provideXlsxMessages"],
  ["vue-xlsx", "src/components/XlsxSheetSurface.vue", "provideXlsxMessages"],
  ["vue-pptx", "src/PptxViewer.vue", "providePptxMessages"],
  ["vue-pdf", "src/components/PdfViewer.vue", "providePdfMessages"],
  ["vue-pdf", "src/components/PdfSurface.vue", "providePdfMessages"],
]

const FONT_NAMES = new Set([
  "Calibri",
  "Aptos",
  "Arial",
  "Georgia",
  "Times New Roman",
  "Helvetica",
  "Courier New",
  "fx",
])

function read(relativePath) {
  return readFileSync(path.join(ROOT, relativePath), "utf8")
}

function walk(directory) {
  return readdirSync(directory).flatMap((name) => {
    const absolute = path.join(directory, name)
    return statSync(absolute).isDirectory() ? walk(absolute) : [absolute]
  })
}

function templateOf(source) {
  const match = /<template>([\s\S]*?)<\/template>/u.exec(source)
  assert.ok(match, "component must contain a template")
  return match[1].replace(/<!--[\s\S]*?-->/gu, "")
}

function textNodes(template) {
  return [...template.matchAll(/>([^<>{}]+)</gu)]
    .map((match) => match[1].replace(/\s+/gu, " ").trim())
    .filter(Boolean)
}

function literalAttributes(template) {
  return [...template.matchAll(/(?<!:)\b(aria-label|title|placeholder|alt)="([^"]*)"/gu)]
    .map((match) => [match[1], match[2]])
    .filter(([, value]) => value.trim().length > 0)
}

function leaves(value, prefix = "", result = new Map()) {
  for (const [key, child] of Object.entries(value)) {
    const next = prefix ? `${prefix}.${key}` : key
    if (typeof child === "string") result.set(next, child)
    else if (child && typeof child === "object") leaves(child, next, result)
  }
  return result
}

const catalogKeys = leaves(i18n.officeMessagesEn)

test("every viewer accepts locale/messages props and provides the shared catalog", () => {
  assert.match(
    read("packages/office-interaction/src/index.ts"),
    /export \* from "\.\/i18n"/u,
    "the shared catalog must be exported from the office-interaction root",
  )
  for (const [packageName, relativePath, provider] of VIEWERS) {
    const source = read(`packages/${packageName}/${relativePath}`)
    const propsSource = source.includes("defineProps<PdfViewerProps>()")
      ? read("packages/vue-pdf/src/types.ts")
      : source
    assert.match(
      propsSource,
      /locale\?:\s*(?:OfficeLocale|import\()/u,
      `${relativePath} must expose a locale prop`,
    )
    assert.match(
      propsSource,
      /messages\?:\s*(?:OfficeMessagesOverride|import\()/u,
      `${relativePath} must expose a messages prop`,
    )
    assert.match(
      source,
      new RegExp(`${provider}\\(`, "u"),
      `${relativePath} must provide the resolved message catalog`,
    )
  }

  for (const packageName of VIEWER_PACKAGES) {
    const helper = read(`packages/${packageName}/src/i18n.ts`)
    assert.match(helper, /OFFICE_I18N_INJECTION_KEY/u, `${packageName} must share the i18n injection key`)
    assert.match(helper, /OFFICE_LOCALE_INJECTION_KEY/u, `${packageName} must share the locale injection key`)
    assert.match(helper, /resolveOfficeMessages/u, `${packageName} must resolve locale + overrides`)
  }
})

test("templates contain no hard-coded labels or translatable copy", () => {
  for (const packageName of VIEWER_PACKAGES) {
    for (const absolutePath of walk(path.join(ROOT, "packages", packageName, "src"))) {
      if (!absolutePath.endsWith(".vue")) continue
      const relativePath = path.relative(ROOT, absolutePath)
      const template = templateOf(readFileSync(absolutePath, "utf8"))

      assert.equal(
        literalAttributes(template).length,
        0,
        `${relativePath} must bind labels/titles/placeholders/alt text to the message catalog: ${JSON.stringify(literalAttributes(template))}`,
      )
      assert.doesNotMatch(template, /[\u4e00-\u9fff]/u, `${relativePath} must not hard-code Chinese template copy`)

      for (const text of textNodes(template)) {
        if (!/[A-Za-z]{2,}/u.test(text)) continue
        assert.ok(
          FONT_NAMES.has(text),
          `${relativePath} must not hard-code user-facing copy: ${JSON.stringify(text)}`,
        )
      }
    }
  }
})

test("templates only reference message keys that exist in the shared catalog", () => {
  for (const packageName of [...VIEWER_PACKAGES, "vue-ui"]) {
    for (const absolutePath of walk(path.join(ROOT, "packages", packageName, "src"))) {
      if (!absolutePath.endsWith(".vue")) continue
      const relativePath = path.relative(ROOT, absolutePath)
      const source = readFileSync(absolutePath, "utf8")
      for (const match of source.matchAll(/\bm(?:\.value)?\.((?:common|docx|xlsx|pptx|pdf)(?:\.[A-Za-z0-9_]+)+)/gu)) {
        assert.ok(
          catalogKeys.has(match[1]),
          `${relativePath} references missing message key ${match[1]}`,
        )
      }
    }
  }
})

test("the previously hard-coded strings now resolve through the catalog", () => {
  const expectations = [
    ["vue-xlsx", "src/components/XlsxToolbar.vue", "Search workbook", "m.xlsx.toolbar.search"],
    ["vue-xlsx", "src/components/XlsxSheetTabs.vue", "上一个工作表", "m.xlsx.sheetTabs.previous"],
    ["vue-xlsx", "src/components/XlsxViewer.vue", "打开一个 XLSX 文件开始查看", "m.xlsx.viewer.empty"],
    ["vue-docx", "src/components/DocxViewerToolbar.vue", "修订", "m.docx.toolbar.trackedChanges"],
    ["vue-docx", "src/components/DocxNotesSection.vue", "脚注", "m.docx.notes.footnote"],
    ["vue-pptx", "src/PptxViewer.vue", "上一页", "m.pptx.viewer.previousSlide"],
    ["vue-pptx", "src/PptxThumbnail.vue", "隐藏", "m.pptx.thumbnail.hidden"],
    ["vue-pdf", "src/components/PdfViewer.vue", "搜索 PDF 文本", "m.pdf.viewer.searchLabel"],
    ["vue-pdf", "src/components/PdfSurface.vue", "选择 PDF 页面区域", "m.pdf.surface.selectionRegion"],
  ]
  for (const [packageName, relativePath, legacy, key] of expectations) {
    const source = read(`packages/${packageName}/${relativePath}`)
    assert.doesNotMatch(source, new RegExp(legacy.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&"), "u"))
    assert.ok(source.includes(key), `${relativePath} must use ${key}`)
  }

  const findBar = read("packages/vue-ui/src/components/OfficeFindBar.vue")
  assert.ok(findBar.includes("m.common.search.previous"))
  assert.ok(findBar.includes("m.common.search.next"))
  assert.ok(findBar.includes("m.common.search.close"))
  assert.doesNotMatch(findBar, /aria-label="上一个搜索结果"/u)
})

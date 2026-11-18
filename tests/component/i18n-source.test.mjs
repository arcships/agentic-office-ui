import assert from "node:assert/strict"
import { readdirSync, readFileSync, statSync } from "node:fs"
import path from "node:path"
import test from "node:test"

const i18n = await import(
  new URL("../../packages/office-interaction/src/i18n.ts", import.meta.url)
)

const ROOT = path.resolve(import.meta.dirname, "../..")
const VIEWER_PACKAGES = ["vue-docx", "vue-xlsx", "vue-pptx", "vue-pdf"]
const SCANNED_PACKAGES = [...VIEWER_PACKAGES, "vue-ui"]

const VIEWERS = [
  ["vue-docx", "src/components/DocxViewer.vue", "provideDocxMessages"],
  ["vue-docx", "src/components/DocxEditor.vue", "provideDocxMessages"],
  ["vue-docx", "src/components/DocxDocumentSurface.vue", "provideDocxMessages"],
  ["vue-xlsx", "src/components/XlsxViewer.vue", "provideXlsxMessages"],
  ["vue-xlsx", "src/components/XlsxSheetSurface.vue", "provideXlsxMessages"],
  ["vue-pptx", "src/PptxViewer.vue", "providePptxMessages"],
  ["vue-pptx", "src/PptxStage.vue", "providePptxMessages"],
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

const ALLOWED_EXPRESSION_LITERALS = new Set([
  "true",
  "false",
  "region",
  "content",
  "object",
  "button",
  "checkbox",
  "radio",
  "page",
  "number",
  "text",
  "date",
  "general",
  "currency",
  "percent",
  "none",
  "single",
  "auto",
  "footnote",
  "endnote",
  "document.docx",
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

function styleOf(source) {
  return [...source.matchAll(/<style[^>]*>([\s\S]*?)<\/style>/gu)]
    .map((match) => match[1])
    .join("\n")
}

function scriptOf(source) {
  return [...source.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/gu)]
    .map((match) => match[1])
    .join("\n")
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

function boundLabelAttributes(template) {
  return [...template.matchAll(/:(aria-label|title|placeholder|alt)="([^"]*)"/gu)]
    .map((match) => [match[1], match[2]])
}

function mustacheExpressions(template) {
  return [...template.matchAll(/\{\{([\s\S]*?)\}\}/gu)].map((match) => match[1])
}

function stringLiterals(expression) {
  return [...expression.matchAll(/(['"])((?:\\.|(?!\1)[^\\])*)\1/gu)]
    .map((match) => match[2])
    .filter((value) => value.length > 0)
}

function isTranslatableLiteral(value) {
  if (/[\u4e00-\u9fff]/u.test(value)) return true
  if (!/[A-Za-z]{3,}/u.test(value)) return false
  return !ALLOWED_EXPRESSION_LITERALS.has(value)
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

  for (const packageName of SCANNED_PACKAGES) {
    const helper = read(`packages/${packageName}/src/i18n.ts`)
    assert.match(helper, /OFFICE_I18N_INJECTION_KEY/u, `${packageName} must share the i18n injection key`)
    assert.match(helper, /OFFICE_LOCALE_INJECTION_KEY/u, `${packageName} must share the locale injection key`)
    assert.match(helper, /resolveOfficeMessages/u, `${packageName} must resolve locale + overrides`)
  }
})

test("the XLSX locale also selects number and date formatting", () => {
  // #104's acceptance covers text, aria-labels and formatted values: the same
  // `locale` prop is the fallback locale for numbers and dates, and the number
  // rules expose the two named helpers from the upstream gap report.
  const viewer = read("packages/vue-xlsx/src/components/XlsxViewer.vue")
  assert.match(
    viewer,
    /props\.documentLocale\s*\?\?\s*\(props\.locale/u,
    "XlsxViewer must fall back from documentLocale to the viewer locale",
  )
  assert.match(viewer, /setDocumentLocale\?\./u, "XlsxViewer must forward the resolved locale")

  const surface = read("packages/vue-xlsx/src/components/XlsxSheetSurface.vue")
  assert.match(
    surface,
    /props\.documentLocale\s*\?\?\s*\(props\.locale/u,
    "XlsxSheetSurface must fall back from documentLocale to the viewer locale",
  )

  const controller = read("packages/vue-xlsx/src/composables/useXlsxViewerController.ts")
  assert.match(controller, /locale:\s*configuredLocale/u, "the controller must accept a locale option")
  assert.match(controller, /get locale\(\)/u, "the controller must expose its locale")

  const numberLocale = read("packages/vue-xlsx/src/composables/number-locale.ts")
  assert.match(numberLocale, /export function localizeCellNumber\(/u)
  assert.match(numberLocale, /export function parseLocaleNumber\(/u)
})

test("templates contain no hard-coded labels or translatable copy", () => {
  for (const packageName of SCANNED_PACKAGES) {
    for (const absolutePath of walk(path.join(ROOT, "packages", packageName, "src"))) {
      if (!absolutePath.endsWith(".vue")) continue
      const relativePath = path.relative(ROOT, absolutePath)
      const source = readFileSync(absolutePath, "utf8")
      const template = templateOf(source)

      assert.equal(
        literalAttributes(template).length,
        0,
        `${relativePath} must bind labels/titles/placeholders/alt text to the message catalog: ${JSON.stringify(literalAttributes(template))}`,
      )
      assert.doesNotMatch(template, /[\u4e00-\u9fff]/u, `${relativePath} must not hard-code Chinese template copy`)

      for (const [attribute, expression] of boundLabelAttributes(template)) {
        const literals = stringLiterals(expression).filter(isTranslatableLiteral)
        assert.equal(
          literals.length,
          0,
          `${relativePath} :${attribute} binding contains translatable literals: ${JSON.stringify(literals)}`,
        )
      }

      for (const expression of mustacheExpressions(template)) {
        const literals = stringLiterals(expression).filter(isTranslatableLiteral)
        assert.equal(
          literals.length,
          0,
          `${relativePath} mustache interpolation contains translatable literals: ${JSON.stringify(literals)}`,
        )
      }

      for (const text of textNodes(template)) {
        if (!/[A-Za-z]{2,}/u.test(text)) continue
        assert.ok(
          FONT_NAMES.has(text),
          `${relativePath} must not hard-code user-facing copy: ${JSON.stringify(text)}`,
        )
      }

      const style = styleOf(source)
      assert.doesNotMatch(style, /[\u4e00-\u9fff]/u, `${relativePath} must not hard-code Chinese style copy`)
      assert.doesNotMatch(
        style,
        /content\s*:\s*["'][^"']*[A-Za-z]{3,}/u,
        `${relativePath} must not inject user-facing copy through CSS content`,
      )

      for (const literal of stringLiterals(scriptOf(source))) {
        assert.doesNotMatch(
          literal,
          /[\u4e00-\u9fff]/u,
          `${relativePath} script contains hard-coded Chinese copy: ${JSON.stringify(literal)}`,
        )
      }
    }
  }
})

test("TypeScript sources contain no hard-coded translatable copy", () => {
  const forbidden = [
    "Hide comments",
    "Show comments",
    "Cannot change the list",
    "Copied to clipboard",
    "Pasted from clipboard",
    "Applied bullet list",
    "Applied numbered list",
    "Export cancelled",
    "Exporting DOCX",
    "Exported DOCX",
    "Failed to load file",
    "Discard unsaved changes?",
    "Unable to load PPTX",
    "PPTX preview session",
    "PPTX open request",
  ]
  for (const packageName of SCANNED_PACKAGES) {
    for (const absolutePath of walk(path.join(ROOT, "packages", packageName, "src"))) {
      if (!absolutePath.endsWith(".ts")) continue
      const relativePath = path.relative(ROOT, absolutePath)
      // This module parses localized CJK date input; its regex and example
      // comment legitimately contain Chinese characters.
      if (relativePath.endsWith("number-locale.ts")) continue
      const source = readFileSync(absolutePath, "utf8")
      for (const literal of stringLiterals(source)) {
        assert.doesNotMatch(
          literal,
          /[\u4e00-\u9fff]/u,
          `${relativePath} contains hard-coded Chinese copy: ${JSON.stringify(literal)}`,
        )
        for (const phrase of forbidden) {
          assert.ok(
            !literal.includes(phrase),
            `${relativePath} contains hard-coded status copy: ${JSON.stringify(literal)}`,
          )
        }
      }
    }
  }
})

test("TypeScript status, label and error assignments use catalog keys", () => {
  function looksLikeCopy(value) {
    if (/[\u4e00-\u9fff]/u.test(value)) return true
    if (!/[A-Za-z]/u.test(value)) return false
    if (/^[A-Z0-9_]+$/u.test(value)) return false
    if (/^[A-Z][A-Za-z0-9]*(?:Error|Type|Code)$/u.test(value)) return false
    if (/(?:Calibri|Arial|Helvetica|Times|sans-serif|serif|monospace|system-ui)/u.test(value)) return false
    if (/^[a-z][A-Za-z0-9_.:/%-]*$/u.test(value)) return false
    return /\s/u.test(value) || /^[A-Z]/u.test(value)
  }

  const patterns = [
    /\b(?:ctx\.)?status(?:\.value)?\s*=\s*(`[^`]*`|"[^"]*"|'[^']*')/gu,
    /\b(?:status|message|label|ariaLabel|title)\s*:\s*(`[^`]*`|"[^"]*"|'[^']*')/gu,
    /\b(?:refuse|unsupported)\(\s*("[^"]*"|'[^']*')/gu,
    /\?\?\s*(`[^`]*`|"[^"]*"|'[^']*')/gu,
    /\|\|\s*(`[^`]*`|"[^"]*"|'[^']*')/gu,
  ]

  for (const packageName of SCANNED_PACKAGES) {
    for (const absolutePath of walk(path.join(ROOT, "packages", packageName, "src"))) {
      const isTs = absolutePath.endsWith(".ts") || absolutePath.endsWith(".tsx")
      const isVue = absolutePath.endsWith(".vue")
      if (!isTs && !isVue) continue
      const relativePath = path.relative(ROOT, absolutePath)
      if (relativePath.endsWith("number-locale.ts")) continue
      const source = readFileSync(absolutePath, "utf8")
      const scanned = isVue ? scriptOf(source) : source
      for (const pattern of patterns) {
        for (const match of scanned.matchAll(pattern)) {
          const literal = stringLiterals(match[1])[0] ?? ""
          assert.ok(
            !looksLikeCopy(literal),
            `${relativePath} hard-codes a status/label/message literal: ${JSON.stringify(literal)}`,
          )
        }
      }
    }
  }

  const pdfPolicy = read("packages/vue-pdf/src/pdf-url-policy.ts")
  for (const literal of stringLiterals(pdfPolicy)) {
    assert.ok(
      !looksLikeCopy(literal),
      `pdf-url-policy.ts hard-codes an error message: ${JSON.stringify(literal)}`,
    )
  }
})

test("TSX and Vue scripts keep JSX copy and aria labels in the catalog", () => {
  function looksLikeCopy(value) {
    if (/[\u4e00-\u9fff]/u.test(value)) return true
    if (!/[A-Za-z]{2,}/u.test(value)) return false
    if (/^[A-Z0-9_]+$/u.test(value)) return false
    if (/^[a-z][A-Za-z0-9_.:/%-]*$/u.test(value)) return false
    return /\s/u.test(value) || /^[A-Z]/u.test(value)
  }

  for (const packageName of SCANNED_PACKAGES) {
    for (const absolutePath of walk(path.join(ROOT, "packages", packageName, "src"))) {
      const relativePath = path.relative(ROOT, absolutePath)
      if (absolutePath.endsWith(".tsx")) {
        const source = readFileSync(absolutePath, "utf8")
        for (const match of source.matchAll(/aria-label=\{([^}]*)\}/gu)) {
          for (const literal of stringLiterals(match[1])) {
            assert.ok(
              !looksLikeCopy(literal),
              `${relativePath} hard-codes an aria-label literal: ${JSON.stringify(literal)}`,
            )
          }
        }
        for (const match of source.matchAll(/>\s*\{`([^`]*)`\}\s*</gu)) {
          const staticText = match[1].replace(/\$\{[^}]*\}/gu, "")
          assert.ok(
            !looksLikeCopy(staticText),
            `${relativePath} hard-codes JSX copy: ${JSON.stringify(match[1])}`,
          )
        }
        for (const match of source.matchAll(/\blabel\s*:\s*(`[^`]*`|"[^"]*"|'[^']*')/gu)) {
          const literal = stringLiterals(match[1])[0] ?? ""
          assert.ok(
            !looksLikeCopy(literal),
            `${relativePath} hard-codes a label: ${JSON.stringify(literal)}`,
          )
        }
      }

      if (absolutePath.endsWith(".vue")) {
        const source = readFileSync(absolutePath, "utf8")
        const script = scriptOf(source)
        for (const match of script.matchAll(/\bconst\s+\w*(?:label|title|aria|text|message)\w*\s*=\s*(`[^`]*`|"[^"]*"|'[^']*')/giu)) {
          const literal = stringLiterals(match[1])[0] ?? ""
          assert.ok(
            !looksLikeCopy(literal),
            `${relativePath} hard-codes a UI constant: ${JSON.stringify(literal)}`,
          )
        }
      }
    }
  }
})

test("templates only reference message keys that exist in the shared catalog", () => {
  for (const packageName of SCANNED_PACKAGES) {
    for (const absolutePath of walk(path.join(ROOT, "packages", packageName, "src"))) {
      if (!absolutePath.endsWith(".vue")) continue
      const relativePath = path.relative(ROOT, absolutePath)
      const source = readFileSync(absolutePath, "utf8")
      for (const match of source.matchAll(/\bm(?:\.value)?\.((?:common|docx|xlsx|pptx|pdf|vueUi)(?:\.[A-Za-z0-9_]+)+)/gu)) {
        const key = match[1]
        const valid = catalogKeys.has(key)
          || [...catalogKeys.keys()].some((candidate) => candidate.startsWith(`${key}.`))
        assert.ok(
          valid,
          `${relativePath} references missing message key ${key}`,
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
    ["vue-docx", "src/components/DocxTrackedChangeGutter.vue", "Hide comments", "m.docx.gutter.hideComments"],
    ["vue-docx", "src/components/DocxTrackedChangeGutter.vue", "Show comments", "m.docx.gutter.showComments"],
    ["vue-docx", "src/components/DocxGutterAnchorMarkers.vue", "Comment ${entry.annotation.ordinal}", "m.docx.gutter.commentMarkerWithAuthor"],
    ["vue-docx", "src/components/DocxFormFieldLayer.vue", "Invalid value.", "m.value.docx.formField.errors.invalidValue"],
    ["vue-docx", "src/composables/form-field-value.ts", "This field is required.", "messages.required"],
    ["vue-docx", "src/composables/useDocxImageWrapMenu.ts", "Move with Text", "messages.moveWithText"],
    ["vue-docx", "src/components/DocxDocumentSurface.vue", "Unable to confirm DOCX reference.", "common.reference.confirmFailed"],
    ["vue-pdf", "src/components/PdfSurface.vue", "Unable to confirm PDF reference.", "common.reference.confirmFailed"],
    ["vue-pptx", "src/PptxStage.vue", "Unable to confirm PPTX reference.", "common.reference.confirmFailed"],
    ["vue-xlsx", "src/components/XlsxSheetSurface.vue", "Unable to confirm XLSX reference.", "common.reference.confirmFailed"],
    ["vue-xlsx", "src/components/XlsxSheetSurface.vue", "label: \"Workbook\"", "descriptor(\"workbook\")"],
    ["vue-xlsx", "src/render/chart-svg-utils.tsx", "Unsupported chart type:", "officeMessagesEn.xlsx.chart.unsupportedType"],
    ["vue-xlsx", "src/render/chart-renderer.tsx", "?? \"Chart\"", "resolveChartAriaLabel"],
    ["vue-xlsx", "src/render/surface-regl.tsx", "?? \"Chart\"", "resolveChartAriaLabel"],
    ["vue-xlsx", "src/render/chart-svg-utils.tsx", "label: series.name ??", "resolveChartSeriesLabel"],
    ["vue-xlsx", "src/render/chart-data.ts", "Series ${", "resolveChartSeriesLabel"],
    ["vue-xlsx", "src/render/chart-bar.tsx", "Series ${", "resolveChartSeriesLabel"],
    ["vue-docx", "src/components/DocxDocumentSurface.vue", "Table ${", "descriptor(\"table\""],
    ["vue-xlsx", "src/components/XlsxSheetSurface.vue", "Row ${", "descriptor(\"row\""],
    ["vue-xlsx", "src/components/XlsxSheetSurface.vue", "Column ${", "descriptor(\"column\""],
    ["vue-xlsx", "src/render/chart-axis.tsx", "Q${", "surfaceSeries"],
    ["vue-docx", "src/render/paragraph-runs-image.ts", "\"Missing image\"", "imageMissingText"],
    ["vue-docx", "src/components/DocxTableHost.vue", "officeMessagesEn.docx.image.alt", "imageAltText: m.value.docx.image.alt"],
    ["vue-docx", "src/components/DocxDocumentSurface.vue", "result.status === \"exact\"", "localizeOfficeResolveResult"],
    ["vue-xlsx", "src/components/XlsxSheetSurface.vue", "result.status === \"exact\"", "localizeOfficeResolveResult"],
    ["vue-pptx", "src/PptxStage.vue", "result.status === \"exact\"", "localizeOfficeResolveResult"],
    ["vue-pdf", "src/pdf/pdf-reference-adapter.ts", "label: \"Document\"", "resolveDescriptorLabels"],
    ["vue-docx", "src/composables/editor-image.ts", "Select an image file", "ctx.statusMessage(\"selectImageFile\")"],
    ["vue-docx", "src/composables/editor-table.ts", "Cannot delete the only table row", "refuse(\"cannotDeleteOnlyTableRow\")"],    ["vue-docx", "src/composables/editor-format.ts", "Unsupported: diagonal paragraph borders", "ctx.statusMessage(\"unsupportedDiagonalParagraphBorders\""],
    ["vue-pdf", "src/pdf-url-policy.ts", "The source is not a verified PDF file.", "messages.invalidPdf"],
    ["vue-pptx", "src/PptxViewer.vue", "上一页", "m.pptx.viewer.previousSlide"],
    ["vue-pptx", "src/PptxThumbnail.vue", "隐藏", "m.pptx.thumbnail.hidden"],
    ["vue-pdf", "src/components/PdfViewer.vue", "搜索 PDF 文本", "m.pdf.viewer.searchLabel"],
    ["vue-pdf", "src/components/PdfSurface.vue", "选择 PDF 页面区域", "m.pdf.surface.selectionRegion"],
    ["vue-ui", "src/components/OfficeRegionSelector.vue", "选定区域", "m.vueUi.regionSelector.label"],
    ["vue-ui", "src/components/OfficeObjectOutlineLayer.vue", "没有可用对象", "m.vueUi.objectOutline.noObjects"],
  ]
  for (const [packageName, relativePath, legacy, key] of expectations) {
    const source = read(`packages/${packageName}/${relativePath}`)
    assert.doesNotMatch(source, new RegExp(legacy.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&"), "u"))
    assert.ok(
      source.includes(key) || source.includes(key.replace(/^m\./u, "m.value.")),
      `${relativePath} must use ${key}`,
    )
  }

  const findBar = read("packages/vue-ui/src/components/OfficeFindBar.vue")
  assert.ok(findBar.includes("m.common.search.previous"))
  assert.ok(findBar.includes("m.common.search.next"))
  assert.ok(findBar.includes("m.common.search.close"))
  assert.doesNotMatch(findBar, /aria-label="上一个搜索结果"/u)
})

test("core adapter lowercase kind labels stay covered by the descriptor mapper", () => {
  const pptx = read("packages/pptx-core/src/reference-adapter.ts")
  assert.match(pptx, /`\$\{kind\} \$\{identity\.shapeId\}`/u)
  const mapper = read("packages/office-interaction/src/descriptor-i18n.ts")
  for (const token of ["lowerShape", "lowerGroup", "lowerTable", "lowerChart", "lowerImage", "lowerTextBox"]) {
    assert.ok(mapper.includes(token), `descriptor mapper must cover ${token}`)
  }
  assert.ok(mapper.includes("status === \"ambiguous\""), "descriptor mapper must localize ambiguous candidates")
  assert.ok(mapper.includes("candidates.map"), "ambiguous candidates must be mapped")
})

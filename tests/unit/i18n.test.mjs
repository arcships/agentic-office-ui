import assert from "node:assert/strict"
import test from "node:test"

const i18n = await import(
  new URL("../../packages/office-interaction/src/i18n.ts", import.meta.url)
)

function leafPaths(value, prefix = "", result = []) {
  for (const [key, child] of Object.entries(value)) {
    const path = prefix ? `${prefix}.${key}` : key
    if (typeof child === "string") {
      result.push([path, child])
    } else if (child && typeof child === "object") {
      leafPaths(child, path, result)
    }
  }
  return result
}

function lookup(value, path) {
  let current = value
  for (const segment of path.split(".")) {
    if (!current || typeof current !== "object" || !(segment in current)) return undefined
    current = current[segment]
  }
  return current
}

test("office i18n ships complete English and zh-CN catalogs with matching keys", () => {
  const en = leafPaths(i18n.officeMessagesEn)
  const zh = leafPaths(i18n.officeMessagesZhCN)
  assert.equal(en.length, zh.length, "en and zh-CN catalogs must expose the same number of messages")
  assert.deepEqual(
    en.map(([path]) => path).sort(),
    zh.map(([path]) => path).sort(),
    "en and zh-CN catalogs must expose the same message keys",
  )
  for (const [path, value] of en) {
    assert.equal(typeof value, "string", `en message ${path} must be a string`)
    assert.ok(value.length > 0, `en message ${path} must not be empty`)
    assert.equal(typeof lookup(i18n.officeMessagesZhCN, path), "string", `missing zh-CN message ${path}`)
  }
  for (const group of ["common", "docx", "xlsx", "pptx", "pdf"]) {
    assert.ok(i18n.officeMessagesEn[group], `missing ${group} message group`)
    assert.ok(i18n.officeMessagesZhCN[group], `missing zh-CN ${group} message group`)
  }
})

test("office locale resolution defaults to English and maps Chinese locales to zh-CN", () => {
  assert.equal(i18n.resolveOfficeLocale(undefined), "en")
  assert.equal(i18n.resolveOfficeLocale(null), "en")
  assert.equal(i18n.resolveOfficeLocale(""), "en")
  assert.equal(i18n.resolveOfficeLocale("en-US"), "en")
  assert.equal(i18n.resolveOfficeLocale("zh"), "zh-CN")
  assert.equal(i18n.resolveOfficeLocale("zh-CN"), "zh-CN")
  assert.equal(i18n.resolveOfficeLocale("zh-Hans"), "zh-CN")
  assert.equal(i18n.resolveOfficeLocale("zh_CN"), "zh-CN")
  assert.equal(i18n.resolveOfficeLocale("zh-TW"), "en")
  assert.equal(i18n.resolveOfficeMessages(undefined).xlsx.viewer.empty, "Open an XLSX file to get started")
  assert.equal(i18n.resolveOfficeMessages("zh-CN").xlsx.viewer.empty, "打开一个 XLSX 文件开始查看")
})

test("office message overrides merge per key without mutating the built-in catalogs", () => {
  const resolved = i18n.resolveOfficeMessages("en", {
    xlsx: { toolbar: { search: "Search this workbook" } },
    docx: { viewer: { emptyTitle: "Nothing open" } },
  })
  assert.equal(resolved.xlsx.toolbar.search, "Search this workbook")
  assert.equal(resolved.xlsx.toolbar.upload, "Upload XLSX")
  assert.equal(resolved.docx.viewer.emptyTitle, "Nothing open")
  assert.equal(resolved.docx.viewer.chooseFile, "Choose Word file")
  assert.equal(i18n.officeMessagesEn.xlsx.toolbar.search, "Search workbook")
  assert.equal(i18n.officeMessagesEn.docx.viewer.emptyTitle, "No document open")
})

test("office message placeholders interpolate and preserve unknown placeholders", () => {
  assert.equal(
    i18n.formatOfficeMessage("Page {current} of {total}", { current: 2, total: 9 }),
    "Page 2 of 9",
  )
  assert.equal(
    i18n.formatOfficeMessage("第 {current} 页，共 {total} 页", { current: 3, total: 7 }),
    "第 3 页，共 7 页",
  )
  assert.equal(i18n.formatOfficeMessage("{missing} and {value}", { value: 1 }), "{missing} and 1")
})

test("office date and number formatting follow the active locale", () => {
  const date = new Date("2024-01-15T12:00:00Z")
  const enDate = i18n.formatOfficeDate(date, "en")
  const zhDate = i18n.formatOfficeDate(date, "zh-CN")
  assert.equal(enDate, "Jan 15, 2024")
  assert.equal(zhDate, "2024年1月15日")
  assert.equal(i18n.formatOfficeDate("not-a-date", "en"), "")
  assert.equal(i18n.formatOfficeNumber(1234.5, "en", { minimumFractionDigits: 1 }), "1,234.5")
})

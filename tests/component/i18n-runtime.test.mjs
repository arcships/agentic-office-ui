import assert from "node:assert/strict"
import test from "node:test"

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs"

const {
  DocxEditor,
  DocxFormFieldLayer,
  DocxTrackedChangeGutter,
  DocxViewerRoot,
  useDocxEditor,
  useDocxImageWrapMenu,
  useDocxViewerThumbnails,
} = await importFromDemo("@arcships/vue-docx")
const {
  XlsxViewer,
  checkCellValueAgainstDataValidation,
  localizeDataValidationFeedback,
  useClipboardCopyNotice,
} = await importFromDemo("@arcships/vue-xlsx")
const { PptxViewer, PptxStage } = await importFromDemo("@arcships/vue-pptx")
const { localizePdfErrorMessage, PdfViewer } = await importFromDemo("@arcships/vue-pdf")
const {
  OFFICE_I18N_INJECTION_KEY,
  OFFICE_LOCALE_INJECTION_KEY,
  resolveOfficeMessages,
} = await import(
  new URL("../../packages/office-interaction/src/i18n.ts", import.meta.url)
)
const { createBlankDocumentModel, normalizeDocModel } = await importFromDemo("@arcships/docx-core")

function message(locale) {
  return resolveOfficeMessages(locale)
}

function withLocale(locale, component, props) {
  return vue.defineComponent({
    setup() {
      vue.provide(OFFICE_I18N_INJECTION_KEY, vue.computed(() => message(locale)))
      vue.provide(OFFICE_LOCALE_INJECTION_KEY, vue.computed(() => locale))
      return () => vue.h(component, props)
    },
  })
}

function modelWithComment() {
  return [{
    annotation: {
      id: "comment:1",
      location: { kind: "paragraph", nodeIndex: 0 },
      comment: {
        id: 1,
        author: "Alex Reviewer",
        date: "2026-01-02T00:00:00Z",
        text: "Please confirm.",
      },
      ordinal: 1,
    },
    anchorX: 0,
    anchorY: 0,
    top: 0,
    heightPx: 60,
    connectorLane: 0,
  }]
}

function drawerTab(root) {
  return walk(root).find((node) => node.props?.["data-docx-gutter-drawer-tab"] !== undefined)
}

function docxModelWithComment() {
  const model = createBlankDocumentModel()
  model.metadata.comments = [
    { id: 1, author: "Alex Reviewer", date: "2026-01-02T00:00:00Z", text: "Please confirm." },
  ]
  model.nodes = [{
    type: "paragraph",
    children: [{ type: "text", text: "Alpha beta gamma" }],
    commentRanges: [{ commentId: 1, start: 6, end: 10, anchorText: "beta" }],
  }]
  return normalizeDocModel(model)
}

function formFieldModel(overrides = {}) {
  const model = createBlankDocumentModel()
  model.nodes = [{
    type: "paragraph",
    children: [{
      type: "form-field",
      fieldType: "text",
      required: true,
      value: "",
      ...overrides,
    }],
  }]
  return model
}

test("DOCX gutter drawer tab follows en and zh-CN copy", async () => {
  for (const [locale, titleOpen, label, titleClosed] of [
    ["en", "Hide comments", "Comments", "Show comments"],
    ["zh-CN", "隐藏批注", "批注", "显示批注"],
  ]) {
    const mounted = await mount(
      withLocale(locale, DocxTrackedChangeGutter, {
        positioned: modelWithComment(),
        drawerMode: true,
        pageHeightPx: 1000,
        zoomFactor: 1,
      }),
    )
    try {
      await vue.nextTick()
      const tab = drawerTab(mounted.root)
      assert.ok(tab, `${locale}: drawer tab renders`)
      assert.equal(tab.props.title, titleOpen)
      assert.match(textContent(mounted.root), new RegExp(label, "u"))

      tab.props.onClick({ stopPropagation() {}, preventDefault() {} })
      await vue.nextTick()
      assert.equal(tab.props.title, titleClosed)
    } finally {
      mounted.app.unmount()
    }
  }
})

test("DOCX gutter anchor markers localize their comment label", async () => {
  for (const [locale, expected] of [
    ["en", "Comment 1 — Alex Reviewer"],
    ["zh-CN", "批注 1 — Alex Reviewer"],
  ]) {
    const wrapper = withLocale(locale, DocxViewerRoot, {
      model: docxModelWithComment(),
      zoomScale: 100,
      showComments: true,
    })
    const mounted = await mount(wrapper)
    try {
      await vue.nextTick()
      const marker = walk(mounted.root).find(
        (node) => node.props?.["data-docx-comment-marker"] !== undefined,
      )
      assert.ok(marker, `${locale}: comment marker renders`)
      assert.equal(marker.props.title, expected)
      assert.equal(marker.props["aria-label"], expected)
    } finally {
      mounted.app.unmount()
    }
  }
})

test("DOCX form-field validation alerts localize", async () => {
  for (const [locale, expected] of [
    ["en", "This field is required."],
    ["zh-CN", "此字段为必填项。"],
  ]) {
    const scope = vue.effectScope()
    const controller = scope.run(() => useDocxEditor({ starterModel: formFieldModel() }))
    const mounted = await mount(withLocale(locale, DocxFormFieldLayer, {
      pageIndex: 0,
      controller,
      pageNodeSegments: [{ nodeIndex: 0 }],
    }))
    try {
      await vue.nextTick()
      const input = walk(mounted.root).find(
        (node) => String(node.props?.class ?? "").includes("docx-form-text-input"),
      )
      assert.ok(input, `${locale}: form text input renders`)
      input.props.onBlur({ target: { value: "" } })
      await vue.nextTick()
      const alert = walk(mounted.root).find(
        (node) => String(node.props?.class ?? "").includes("docx-form-error"),
      )
      assert.ok(alert, `${locale}: validation alert renders`)
      assert.equal(alert.props.role, "alert")
      assert.equal(textContent(alert), expected)
    } finally {
      mounted.app.unmount()
      scope.stop()
    }
  }
})

test("DOCX image wrap menu localizes its labels", () => {
  for (const [locale, expected] of [
    ["en", ["Move with Text", "Fix Position on Page", "Edit Wrap Boundary", "More Layout Options...", "In Line with Text"]],
    ["zh-CN", ["随文字移动", "固定在页面位置", "编辑环绕边界", "更多布局选项…", "嵌入文字"]],
  ]) {
    const menu = useDocxImageWrapMenu({
      context: { kind: "image", image: { floating: {} } },
      runAction() {},
    }, message(locale).docx.image)
    assert.ok(menu, `${locale}: image wrap menu resolves`)
    assert.deepEqual(
      [
        menu.positioningOptions[0].label,
        menu.positioningOptions[1].label,
        menu.editWrapBoundaryOption.label,
        menu.moreLayoutOptionsOption.label,
        menu.wrapOptions[0].label,
      ],
      expected,
    )
  }
})

test("XLSX error text and live region localize from the load error code", async () => {
  for (const [locale, expected] of [
    ["en", "The file is not a valid XLSX workbook."],
    ["zh-CN", "该文件不是有效的 XLSX 工作簿。"],
  ]) {
    const error = new Error("raw catalog message")
    const controller = {
      activeSheet: null,
      activeTab: null,
      canZoomIn: false,
      canZoomOut: false,
      displayFileName: "broken.xlsx",
      error,
      isEditingLocked: false,
      isLoading: false,
      isWorkerBacked: false,
      readOnly: false,
      redo() {},
      revision: 1,
      sheets: [],
      sourceError: {
        code: "INVALID_WORKBOOK",
        message: "raw catalog message",
        sourceKind: "file",
      },
      tabs: [],
      undo() {},
      zoomScale: 100,
    }
    const mounted = await mount(XlsxViewer, {
      controller,
      locale,
      showFormulaBar: false,
      showRibbon: false,
    })
    try {
      await vue.nextTick()
      assert.match(textContent(mounted.root), new RegExp(expected, "u"))
      assert.match(textContent(findByTestId(mounted.root, "xlsx-live-status")), new RegExp(expected, "u"))
    } finally {
      mounted.app.unmount()
    }
  }
})

test("a mounted viewer switches locale without remounting", async () => {
  const controller = {
    activeSheet: null,
    activeTab: null,
    canZoomIn: false,
    canZoomOut: false,
    displayFileName: "book.xlsx",
    error: null,
    isEditingLocked: false,
    isLoading: false,
    isWorkerBacked: false,
    readOnly: false,
    redo() {},
    revision: 1,
    sheets: [],
    sourceError: null,
    tabs: [],
    undo() {},
    zoomScale: 100,
  }
  const locale = vue.ref("en")
  const Harness = vue.defineComponent({
    setup() {
      return () => vue.h(XlsxViewer, {
        controller,
        locale: locale.value,
        showFormulaBar: false,
        showRibbon: false,
      })
    },
  })
  const mounted = await mount(Harness)
  try {
    await vue.nextTick()
    assert.match(textContent(mounted.root), /Open an XLSX file to get started/u)
    assert.equal(findByTestId(mounted.root, "xlsx-search-toggle").props["aria-label"], "Search workbook")

    locale.value = "zh-CN"
    await vue.nextTick()
    assert.match(textContent(mounted.root), /打开一个 XLSX 文件开始查看/u)
    assert.equal(findByTestId(mounted.root, "xlsx-search-toggle").props["aria-label"], "搜索工作簿")
  } finally {
    mounted.app.unmount()
  }
})

test("DOCX editor status bar and live region localize command status", async () => {
  for (const [locale, ready, applied, imageStatus, tableStatus, borderStatus] of [
    ["en", "Ready", "Applied bullet list", "Select an image file", "Cannot delete the table: table not found", "Unsupported: diagonal paragraph borders (diagonal-down)"],
    ["zh-CN", "就绪", "已应用项目符号列表", "请选择图片文件", "无法删除表格：未找到表格", "不支持：对角线段落边框（diagonal-down）"],
  ]) {
    const scope = vue.effectScope()
    const controller = scope.run(() => useDocxEditor({
      starterModel: createBlankDocumentModel(),
      statusMessages: () => message(locale).docx.editor.status,
    }))
    const mounted = await mount(DocxEditor, {
      editor: controller,
      locale,
    })
    try {
      await vue.nextTick()
      assert.equal(textContent(findByTestId(mounted.root, "editor-status")), ready)
      assert.match(textContent(findByTestId(mounted.root, "editor-live-status")), new RegExp(ready, "u"))

      controller.toggleList("unordered")
      await vue.nextTick()
      assert.equal(textContent(findByTestId(mounted.root, "editor-status")), applied)
      assert.match(textContent(findByTestId(mounted.root, "editor-live-status")), new RegExp(applied, "u"))

      await controller.insertImageFile({ type: "text/plain", name: "notes.txt" })
      await vue.nextTick()
      assert.equal(textContent(findByTestId(mounted.root, "editor-status")), imageStatus)

      controller.deleteTable(999)
      await vue.nextTick()
      assert.equal(textContent(findByTestId(mounted.root, "editor-status")), tableStatus)

      controller.applyBorderPreset("diagonal-down")
      await vue.nextTick()
      assert.equal(textContent(findByTestId(mounted.root, "editor-status")), borderStatus)
      assert.match(
        textContent(findByTestId(mounted.root, "editor-live-status")),
        new RegExp(borderStatus.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&"), "u"),
      )
    } finally {
      mounted.app.unmount()
      scope.stop()
    }
  }
})

test("DOCX thumbnail placeholder paints localized copy", () => {
  function fakeCanvas() {
    const drawn = []
    const context = {
      setTransform() {},
      fillRect() {},
      strokeRect() {},
      fillText(text) { drawn.push(text) },
      fillStyle: "",
      strokeStyle: "",
      font: "",
      textAlign: "",
      textBaseline: "",
    }
    return {
      canvas: { width: 0, height: 0, style: {}, getContext: () => context },
      drawn,
    }
  }

  for (const [locale, label] of [
    ["en", "Thumbnail unavailable"],
    ["zh-CN", "缩略图暂不可用"],
  ]) {
    const { canvas, drawn } = fakeCanvas()
    const thumbnails = useDocxViewerThumbnails({ locale })
    thumbnails.attachCanvas(0, canvas)
    assert.equal(thumbnails.unavailableReason, message(locale).docx.thumbnail.unavailableReason)
    assert.deepEqual(drawn, [label])
  }
})

test("XLSX calculation, validation and clipboard notices localize", async () => {
  for (const [locale, expected] of [
    ["en", "A formula references a sheet that no longer exists; cached values are shown."],
    ["zh-CN", "公式引用了不存在的工作表，当前显示缓存值。"],
  ]) {
    const controller = {
      activeSheet: null,
      calculationStatus: {
        reason: "unresolved-sheet-refs",
        message: "raw controller message",
      },
      canZoomIn: false,
      canZoomOut: false,
      displayFileName: "book.xlsx",
      error: null,
      isLoading: false,
      readOnly: false,
      redo() {},
      tabs: [],
      undo() {},
      zoomScale: 100,
    }
    const mounted = await mount(XlsxViewer, {
      controller,
      locale,
      showFormulaBar: false,
      showRibbon: false,
    })
    try {
      await vue.nextTick()
      assert.equal(
        textContent(findByTestId(mounted.root, "xlsx-calculation-status")),
        expected,
      )
    } finally {
      mounted.app.unmount()
    }
  }

  const sheet = {
    dataValidations: [{
      validationType: "whole",
      operator: "between",
      value1: "1",
      value2: "10",
      ranges: [{ start: { row: 1, col: 0 }, end: { row: 1, col: 0 } }],
      allowBlank: false,
      errorStyle: "stop",
    }],
  }
  const check = checkCellValueAgainstDataValidation(sheet, { row: 1, col: 0 }, "abc")
  assert.equal(check.accepted, false)
  assert.equal(check.feedback?.defaultMessageKey, "whole")
  assert.equal(
    localizeDataValidationFeedback(check.feedback, message("en").xlsx.validation),
    "Enter a valid whole number.",
  )
  assert.equal(
    localizeDataValidationFeedback(check.feedback, message("zh-CN").xlsx.validation),
    "请输入有效的整数。",
  )

  for (const [locale, expected] of [
    ["en", "Copy failed: this browser does not allow clipboard writes."],
    ["zh-CN", "复制失败：当前浏览器不允许写入剪贴板。"],
  ]) {
    const scope = vue.effectScope()
    const notice = scope.run(() => useClipboardCopyNotice(
      () => ({ copySelectionToClipboard: async () => false }),
      { failureMessage: () => message(locale).xlsx.clipboard.copyFailed },
    ))
    try {
      assert.equal(await notice.copySelection(), false)
      assert.equal(notice.notice.value, expected)
    } finally {
      scope.stop()
    }
  }
})

test("PDF load errors localize from their codes", async () => {
  for (const [locale, invalid, tooLarge, encrypted] of [
    ["en", "The PDF file could not be parsed or rendered.", "PDF file size 12 bytes exceeds the allowed maximum of 10 bytes.", "The PDF document is encrypted and requires a password."],
    ["zh-CN", "PDF 文件无法解析或渲染。", "PDF 文件大小 12 字节超过允许值 10 字节。", "PDF 文档已加密，需要密码。"],
  ]) {
    const messages = message(locale).pdf.errors
    assert.equal(localizePdfErrorMessage({ code: "INVALID_PDF" }, messages), invalid)
    assert.equal(
      localizePdfErrorMessage({ code: "PDF_TOO_LARGE", actual: 12, allowed: 10 }, messages),
      tooLarge,
    )
    assert.equal(localizePdfErrorMessage({ code: "ENCRYPTED" }, messages), encrypted)

    const mounted = await mount(PdfViewer, {
      src: "file:///tmp/not-a-pdf.pdf",
      locale,
      showToolbar: false,
    })
    try {
      await waitFor(
        () => findByTestId(mounted.root, "load-error") !== undefined,
        `${locale}: PDF load error`,
      )
      assert.match(textContent(findByTestId(mounted.root, "load-error")), new RegExp(
        message(locale).pdf.errors.sourceNotAllowed.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&"),
        "u",
      ))
    } finally {
      mounted.app.unmount()
    }
  }
})

function failingSession(error) {
  return {
    document: null,
    currentSlideIndex: 0,
    zoomPercent: 100,
    async open() { throw error },
    async renderSlide() {},
    renderThumbnail() { return null },
    searchText() { return [] },
    async highlightSearchResult() { return null },
    clearSearchHighlights() {},
    async setZoom() {},
    cancel() {},
    dispose() {},
  }
}

test("PPTX document errors localize and PptxStage accepts standalone locale props", async (t) => {
  const originalRect = Object.getOwnPropertyDescriptor(Object.prototype, "getBoundingClientRect")
  Object.defineProperty(Object.prototype, "getBoundingClientRect", {
    configurable: true,
    value: () => ({ left: 0, top: 0, width: 640, height: 480, right: 640, bottom: 480, x: 0, y: 0 }),
  })
  t.after(() => {
    if (originalRect) Object.defineProperty(Object.prototype, "getBoundingClientRect", originalRect)
    else delete Object.prototype.getBoundingClientRect
  })
  for (const [locale, expected] of [
    ["en", "Unable to load PPTX."],
    ["zh-CN", "无法加载 PPTX。"],
  ]) {
    const mounted = await mount(PptxViewer, {
      source: new Uint8Array([1]),
      sessionFactory: () => failingSession("boom"),
      showSearch: false,
      showSidebar: false,
      locale,
    })
    try {
      await waitFor(
        () => findByTestId(mounted.root, "pptx-error") !== undefined,
        `${locale}: PPTX error state`,
      )
      assert.match(textContent(findByTestId(mounted.root, "pptx-error")), new RegExp(expected.replace(/\./gu, "\\."), "u"))
    } finally {
      mounted.app.unmount()
    }
  }

  for (const [locale, expected] of [
    ["en", "Choose slide area"],
    ["zh-CN", "选择幻灯片区域"],
  ]) {
    const mounted = await mount(PptxStage, {
      selectionMode: "region",
      locale,
    })
    try {
      await vue.nextTick()
      const region = walk(mounted.root).find((node) =>
        String(node.props?.class ?? "").includes("pptx-stage__region-layer")
      )
      assert.ok(region, "region layer renders")
      assert.equal(region.props["aria-label"], expected)
    } finally {
      mounted.app.unmount()
    }
  }
})

import assert from "node:assert/strict"
import test from "node:test"

import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
  walk,
} from "./vue-test-renderer.mjs"

const { createBlankDocumentModel } = await importFromDemo("@arcships/docx-core")
const { DocxDocumentSurface } = await importFromDemo("@arcships/vue-docx")
const { XlsxSheetSurface } = await importFromDemo("@arcships/vue-xlsx")
const { PptxStage } = await importFromDemo("@arcships/vue-pptx")

function pointerEvent(currentTarget, clientX, clientY, pointerId = 1, pointerType = "mouse") {
  return {
    button: 0,
    clientX,
    clientY,
    currentTarget,
    isPrimary: true,
    pointerId,
    pointerType,
    preventDefault() {},
  }
}

function enablePointerCapture(element) {
  const captures = new Set()
  element.setPointerCapture = (pointerId) => captures.add(pointerId)
  element.hasPointerCapture = (pointerId) => captures.has(pointerId)
  element.releasePointerCapture = (pointerId) => captures.delete(pointerId)
}

test("DocxDocumentSurface confirms page and touch-region references and resolves them", async () => {
  const confirmations = []
  const resolutions = []
  let api
  let mode
  const Harness = vue.defineComponent({
    setup() {
      const surface = vue.ref(null)
      mode = vue.ref("object")
      vue.watch(surface, (value) => { api = value }, { flush: "sync" })
      return () => vue.h(DocxDocumentSurface, {
        ref: surface,
        model: createBlankDocumentModel(),
        documentId: "surface.docx",
        selectionMode: mode.value,
        emitReferenceCandidates: true,
        onReferenceConfirm: (event) => confirmations.push(event),
        onReferenceResolve: (event) => resolutions.push(event),
      })
    },
  })
  const mounted = await mount(Harness)
  const surface = walk(mounted.root).find((node) => node.props?.class === "docx-document-surface")
  const page = walk(mounted.root).find((node) => node.props?.["data-docx-page-index"] === 0)
  assert.ok(surface)
  assert.ok(page)
  page.dataset = { docxPageIndex: "0" }
  page.closest = (selector) => selector === "[data-docx-page-index]" ? page : null
  page.getBoundingClientRect = () => ({ left: 10, top: 20, width: 900, height: 1200 })
  surface.getBoundingClientRect = () => ({ left: 0, top: 0, width: 920, height: 1240 })
  surface.ownerDocument = { elementsFromPoint: () => [page] }

  surface.props.onClick({ clientX: 100, clientY: 120, detail: 1, shiftKey: false })
  assert.equal(confirmations.at(-1).reference.kind, "page")
  assert.equal(confirmations.at(-1).reference.document.documentId, "surface.docx")
  assert.equal(api.getDocumentRevision().format, "docx")
  const pageResult = await api.resolveReference(confirmations.at(-1).reference)
  assert.equal(pageResult.status, "exact")
  assert.equal(resolutions.at(-1).referenceId, confirmations.at(-1).reference.referenceId)

  mode.value = "region"
  await vue.nextTick()
  const layer = walk(mounted.root).find((node) => node.props?.class === "docx-document-surface__region-layer")
  assert.ok(layer)
  enablePointerCapture(layer)
  layer.props.onPointerdown(pointerEvent(layer, 100, 120, 9, "touch"))
  layer.props.onPointermove(pointerEvent(layer, 300, 420, 9, "touch"))
  layer.props.onPointerup(pointerEvent(layer, 300, 420, 9, "touch"))
  assert.equal(confirmations.at(-1).reference.kind, "region")
  assert.equal(confirmations.at(-1).trigger, "touch")
  assert.equal(confirmations.at(-1).reference.locator.value.space, "page")
  assert.deepEqual(mounted.warnings, [])
  mounted.app.unmount()
})

function xlsxReferenceController() {
  const sheet = {
    name: "Sheet1",
    workbookSheetIndex: 0,
    visibleRows: [],
    visibleCols: [],
    mergedRegions: [],
    colWidths: [],
    rowHeights: [],
    hyperlinks: [],
    comments: [],
    sparklines: [],
  }
  return {
    activeTab: { kind: "worksheet" },
    activeSheet: sheet,
    activeSheetIndex: 0,
    activeTabIndex: 0,
    sheets: [sheet],
    charts: [], images: [], shapes: [], formControls: [], tables: [],
    isChartsLoading: false,
    selectedChartElement: null, selectedChartId: null, selectedImageId: null,
    maxConcurrentImageDecodes: 1,
    zoomScale: 100,
    revision: 1,
    readOnly: true,
    selection: null,
    activeCell: null,
    selectedRangeAddress: null,
    activeCellAddress: null,
    selectedValue: "",
    selectedFormula: "",
    isWorkerBacked: false,
    signal: new AbortController().signal,
    getCellDisplayValue: () => "",
    getCellFormula: () => undefined,
    getActiveWorksheet: () => null,
    setZoomScale(value) { this.zoomScale = value },
    setActiveSheetIndex() {}, selectCell() {}, selectRange() {},
    selectChart() {}, selectChartElement() {}, selectImage() {}, reportImageDecodeError() {},
    undo() {}, redo() {}, clearSelectedCells() {}, copySelectionToClipboard() {}, pasteFromClipboard() {},
    sortTable() {}, resizeColumn() {}, resizeRow() {}, setCellValue() {},
  }
}

test("XlsxSheetSurface confirms cell, row, column, worksheet and sheet-region references", async (t) => {
  const previousWindow = globalThis.window
  const previousResizeObserver = globalThis.ResizeObserver
  globalThis.window = { devicePixelRatio: 1, requestAnimationFrame: () => 1, cancelAnimationFrame() {} }
  globalThis.ResizeObserver = class ResizeObserver { observe() {} disconnect() {} }
  t.after(() => {
    if (previousWindow === undefined) delete globalThis.window
    else globalThis.window = previousWindow
    if (previousResizeObserver === undefined) delete globalThis.ResizeObserver
    else globalThis.ResizeObserver = previousResizeObserver
  })

  const confirmations = []
  let api
  let mode
  const controller = xlsxReferenceController()
  const sheets = controller.sheets
  let sheetReads = 0
  Object.defineProperty(controller, "sheets", {
    configurable: true,
    get() { sheetReads += 1; return sheets },
  })
  const Harness = vue.defineComponent({
    setup() {
      const surface = vue.ref(null)
      mode = vue.ref("object")
      vue.watch(surface, (value) => { api = value }, { flush: "sync" })
      return () => vue.h(XlsxSheetSurface, {
        ref: surface,
        controller,
        documentId: "surface.xlsx",
        selectionMode: mode.value,
        emitReferenceCandidates: true,
        onReferenceConfirm: (event) => confirmations.push(event),
      })
    },
  })
  const mounted = await mount(Harness)
  const surface = findByTestId(mounted.root, "xlsx-sheet-surface")
  const grid = findByTestId(mounted.root, "xlsx-grid")
  assert.ok(surface)
  assert.ok(grid)
  grid.getBoundingClientRect = () => ({ left: 0, top: 0, width: 640, height: 480 })
  grid.scrollLeft = 0
  grid.scrollTop = 0
  grid.focus = () => {}
  surface.getBoundingClientRect = () => ({ left: 0, top: 0, width: 640, height: 480 })
  sheetReads = 0

  for (const [clientX, clientY, kind] of [
    [60, 30, "cell"],
    [60, 10, "column"],
    [10, 30, "row"],
    [10, 10, "worksheet"],
  ]) {
    surface.props.onClick({ clientX, clientY, detail: 1, shiftKey: false })
    assert.equal(confirmations.at(-1).reference.kind, kind)
  }
  assert.equal(sheetReads, 1, "repeated pointer confirmations must reuse the revision-scoped workbook context")
  assert.equal(api.getDocumentRevision().format, "xlsx")
  assert.equal((await api.resolveReference(confirmations[0].reference)).status, "exact")

  mode.value = "region"
  await vue.nextTick()
  const layer = walk(mounted.root).find((node) => node.props?.class === "xlsx-sheet-surface__region-layer")
  assert.ok(layer)
  enablePointerCapture(layer)
  layer.props.onPointerdown(pointerEvent(layer, 60, 30, 5))
  layer.props.onPointermove(pointerEvent(layer, 180, 90, 5))
  layer.props.onPointerup(pointerEvent(layer, 180, 90, 5))
  assert.equal(confirmations.at(-1).reference.kind, "region")
  assert.equal(confirmations.at(-1).reference.locator.value.space, "sheet")
  assert.deepEqual(mounted.warnings, [])
  mounted.app.unmount()
})

test("PptxStage confirms rendered objects and slide regions without owning product state", async () => {
  const confirmations = []
  let api
  let mode
  const Harness = vue.defineComponent({
    setup() {
      const stage = vue.ref(null)
      mode = vue.ref("object")
      vue.watch(stage, (value) => { api = value }, { flush: "sync" })
      return () => vue.h(PptxStage, {
        ref: stage,
        documentId: "surface.pptx",
        selectionMode: mode.value,
        emitReferenceCandidates: true,
        onReferenceConfirm: (event) => confirmations.push(event),
      })
    },
  })
  const mounted = await mount(Harness)
  const stage = findByTestId(mounted.root, "pptx-stage")
  const content = walk(mounted.root).find((node) => node.props?.class === "pptx-stage__content")
  assert.ok(stage)
  assert.ok(content)

  const object = {
    dataset: {
      pptxObjectKey: "slide-1/shape-7",
      pptxNodeId: "7",
      pptxNodeType: "shape",
      pptxSource: "slide",
    },
    getAttribute: (name) => name === "aria-label" ? "Revenue callout" : null,
    querySelectorAll: () => [],
  }
  const slide = {
    dataset: { slideIndex: "0", pptxSlidePath: "ppt/slides/slide1.xml" },
    getBoundingClientRect: () => ({ left: 10, top: 20, width: 1280, height: 720 }),
    querySelector: () => null,
    querySelectorAll: (selector) => selector === "[data-pptx-object-key]" ? [object] : [],
  }
  let slideQueries = 0
  content.querySelectorAll = (selector) => {
    if (selector !== "[data-slide-index]") return []
    slideQueries += 1
    return [slide]
  }
  content.querySelector = () => slide
  const objectTarget = {
    closest: (selector) => selector === "[data-slide-index]" ? slide : object,
  }
  stage.getBoundingClientRect = () => ({ left: 0, top: 0, width: 1300, height: 760 })
  const ownerDocument = { elementFromPoint: () => objectTarget, elementsFromPoint: () => [objectTarget] }
  stage.ownerDocument = ownerDocument
  content.ownerDocument = ownerDocument

  stage.props.onPointermove({ clientX: 110, clientY: 120 })
  stage.props.onPointermove({ clientX: 111, clientY: 121 })
  assert.equal(slideQueries, 1, "pointermove must reuse the revision-scoped rendered-object context")
  stage.props.onClick({ target: objectTarget, detail: 1, shiftKey: false })
  assert.equal(confirmations.at(-1).reference.kind, "shape")
  assert.equal(confirmations.at(-1).reference.document.documentId, "surface.pptx")
  assert.equal(api.getDocumentRevision().format, "pptx")
  assert.equal((await api.resolveReference(confirmations.at(-1).reference)).status, "exact")

  mode.value = "region"
  await vue.nextTick()
  const layer = walk(mounted.root).find((node) => node.props?.class === "pptx-stage__region-layer")
  assert.ok(layer)
  enablePointerCapture(layer)
  layer.props.onPointerdown(pointerEvent(layer, 110, 120, 6))
  layer.props.onPointermove(pointerEvent(layer, 410, 320, 6))
  layer.props.onPointerup(pointerEvent(layer, 410, 320, 6))
  assert.equal(confirmations.at(-1).reference.kind, "region")
  assert.equal(confirmations.at(-1).reference.locator.value.space, "slide")
  assert.deepEqual(mounted.warnings, [])
  mounted.app.unmount()
})

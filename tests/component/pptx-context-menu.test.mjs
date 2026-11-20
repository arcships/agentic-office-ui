// Issue #161 — the PPTX built-in link menu on the shared Reka dropdown.
//
// The rest of the PPTX component suite uses the repository's DOM-free renderer
// and cannot open a portalled, floating-positioned menu. These cases mount the
// same viewer in Happy DOM so the shared `OfficeDropdownMenu` contract
// (pointer anchor, keyboard, focus, dismissal) runs for real.

import assert from "node:assert/strict"
import test, { after } from "node:test"
import { GlobalRegistrator } from "@happy-dom/global-registrator"

GlobalRegistrator.register({ url: "http://localhost/" })

const { importFromDemo, vue } = await import("./vue-test-renderer.mjs")

const { PptxViewer } = await importFromDemo("@arcships/vue-pptx")

after(() => {
  GlobalRegistrator.unregister()
})

const EXTERNAL_KEY = "ppt/slides/slide1.xml|slide|shape:2"
const INTERNAL_KEY = "ppt/slides/slide1.xml|slide|shape:3"
const PLAIN_KEY = "ppt/slides/slide1.xml|slide|shape:6"

function capability() {
  return { approximate: 0, discovered: 0, features: [], static: 0, strict: 0, unparsed: 0 }
}

function linkActions() {
  return [
    {
      id: "action:0:0",
      sourceObjectKey: EXTERNAL_KEY,
      trigger: "click",
      kind: "open-url",
      url: "https://example.com/docs",
      tooltip: "Read the docs",
    },
    {
      id: "action:0:1",
      sourceObjectKey: INTERNAL_KEY,
      trigger: "click",
      kind: "go-to-slide",
      targetSlideIndex: 2,
      tooltip: "Jump",
    },
  ]
}

function playbackDocument() {
  return {
    capability: capability(),
    slides: Array.from({ length: 3 }, (_, index) => ({
      actions: index === 0 ? linkActions() : [],
      capability: capability(),
      hidden: false,
      index,
      media: [],
      morphFromPrevious: [],
      nodes: {},
      objects: [],
    })),
  }
}

function documentWithSlides(count) {
  return {
    fileName: "links.pptx",
    height: 540,
    width: 960,
    slides: Array.from({ length: count }, (_, index) => ({
      hidden: false,
      index,
      number: index + 1,
    })),
    warnings: [],
  }
}

function createBrowseSession() {
  const calls = { dispose: 0, open: [], renderSlide: [] }
  let currentSlideIndex = 0
  const session = {
    document: null,
    playbackDocument: playbackDocument(),
    get currentSlideIndex() { return currentSlideIndex },
    get zoomPercent() { return 100 },
    async open(_source, options) {
      calls.open.push(options)
      session.document = documentWithSlides(3)
      currentSlideIndex = options?.initialSlide ?? 0
      return session.document
    },
    async renderSlide(index) {
      calls.renderSlide.push(index)
      currentSlideIndex = index
    },
    renderThumbnail() { return null },
    searchText() { return [] },
    async highlightSearchResult() { return null },
    clearSearchHighlights() {},
    async setZoom() {},
    cancel() {},
    dispose() { calls.dispose += 1 },
  }
  return { calls, session }
}

async function waitForDom(predicate, message, timeoutMs = 2_000) {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    await vue.nextTick()
    if (predicate()) return
    await new Promise((resolve) => setTimeout(resolve, 5))
  }
  throw new Error(`Timed out waiting for ${message}`)
}

async function mountViewer(props = {}) {
  const { calls, session } = createBrowseSession()
  const menus = []
  const warnings = []
  const container = document.createElement("div")
  document.body.append(container)
  const app = vue.createApp({
    render: () => vue.h(PptxViewer, {
      source: new Uint8Array([9, 9, 9]),
      sessionFactory: () => session,
      showSearch: false,
      showSidebar: false,
      onContextMenu: (context) => menus.push(context),
      ...props,
    }),
  })
  app.config.warnHandler = (message) => warnings.push(message)
  app.mount(container)
  await waitForDom(
    () => container.querySelector('[data-testid="pptx-viewer"]')?.getAttribute("data-state") === "ready",
    "PPTX viewer ready state",
  )
  return {
    app,
    calls,
    container,
    menus,
    session,
    warnings,
    unmount() {
      app.unmount()
      container.remove()
    },
  }
}

/** Dispatch a real right-click on a slide object rendered inside the stage. */
function rightClickObject(container, objectKey, { x = 40, y = 60 } = {}) {
  const stage = container.querySelector('[data-testid="pptx-stage"]')
  const slide = document.createElement("div")
  slide.dataset.slideIndex = "0"
  const object = document.createElement("span")
  object.dataset.pptxObjectKey = objectKey
  slide.append(object)
  stage.append(slide)
  object.dispatchEvent(new window.MouseEvent("contextmenu", {
    bubbles: true,
    cancelable: true,
    clientX: x,
    clientY: y,
  }))
  return { object, slide, stage }
}

function menuElement() {
  return document.querySelector('[role="menu"]')
}

function menuItem(label) {
  return [...document.querySelectorAll('[role="menuitem"]')]
    .find((item) => item.textContent?.includes(label))
}

function installOpener(t) {
  const opened = []
  const original = globalThis.open
  globalThis.open = (url, name, features) => {
    opened.push([url, name, features])
    return null
  }
  t.after(() => {
    if (original === undefined) delete globalThis.open
    else globalThis.open = original
  })
  return opened
}

function installClipboard(t) {
  const written = []
  const original = Object.getOwnPropertyDescriptor(globalThis.navigator, "clipboard")
  Object.defineProperty(globalThis.navigator, "clipboard", {
    configurable: true,
    value: { async writeText(value) { written.push(value) } },
  })
  t.after(() => {
    if (original) Object.defineProperty(globalThis.navigator, "clipboard", original)
    else delete globalThis.navigator.clipboard
  })
  return written
}

test("the link menu renders the shared surface, copies and opens links", async (t) => {
  const opened = installOpener(t)
  const written = installClipboard(t)
  const mounted = await mountViewer()

  rightClickObject(mounted.container, EXTERNAL_KEY)
  await waitForDom(() => menuElement() !== null, "link menu")

  const menu = menuElement()
  assert.equal(menu.getAttribute("aria-label"), "Link actions")
  assert.match(menu.textContent ?? "", /Open link/)
  assert.match(menu.textContent ?? "", /Copy link/)
  await waitForDom(
    () => menu.contains(document.activeElement),
    "focus enters the menu",
  )

  menuItem("Copy link").click()
  await waitForDom(() => written.length === 1, "copy link")
  assert.deepEqual(written, ["https://example.com/docs"])
  assert.equal(menuElement(), null, "selecting an item closes the menu")
  assert.equal(opened.length, 0)

  // Open link routes through the same policy as a click.
  rightClickObject(mounted.container, EXTERNAL_KEY, { x: 12, y: 24 })
  await waitForDom(() => menuElement() !== null, "reopened link menu")
  menuItem("Open link").click()
  await waitForDom(() => opened.length === 1, "open link")
  assert.deepEqual(opened, [["https://example.com/docs", "_blank", "noopener,noreferrer"]])

  mounted.unmount()
  assert.deepEqual(mounted.warnings, [])
})

test("internal jumps get a slide-specific label and navigate", async () => {
  const mounted = await mountViewer()

  rightClickObject(mounted.container, INTERNAL_KEY, { x: 8, y: 16 })
  await waitForDom(() => menuElement() !== null, "internal link menu")
  assert.match(menuElement().textContent ?? "", /Go to slide 3/)
  assert.equal(menuItem("Copy link"), undefined, "internal jumps have no copy value")

  menuItem("Go to slide 3").click()
  await waitForDom(() => mounted.calls.renderSlide.includes(2), "internal jump")
  assert.equal(menuElement(), null)

  mounted.unmount()
  assert.deepEqual(mounted.warnings, [])
})

test("Escape closes the menu and returns focus to the viewer", async () => {
  const mounted = await mountViewer()

  rightClickObject(mounted.container, EXTERNAL_KEY)
  await waitForDom(() => menuElement() !== null, "link menu")
  const menu = menuElement()
  await waitForDom(() => menu.contains(document.activeElement), "focus enters the menu")

  menu.dispatchEvent(new window.KeyboardEvent("keydown", { key: "Escape", bubbles: true }))
  await waitForDom(() => menuElement() === null, "Escape closes the menu")
  await waitForDom(
    () => document.activeElement === mounted.container.querySelector('[data-testid="pptx-viewer"]'),
    "focus returns to the viewer",
  )

  mounted.unmount()
  assert.deepEqual(mounted.warnings, [])
})

test("an unlinked shape reports the host event without opening the menu", async () => {
  const mounted = await mountViewer()

  rightClickObject(mounted.container, PLAIN_KEY, { x: 3, y: 4 })
  await vue.nextTick()
  await vue.nextTick()
  assert.equal(menuElement(), null)
  assert.equal(mounted.menus.length, 1)
  assert.equal(mounted.menus[0].objectKey, PLAIN_KEY)

  mounted.unmount()
  assert.deepEqual(mounted.warnings, [])
})

test("showContextMenu=false keeps the host event and never opens the menu", async () => {
  const mounted = await mountViewer({ showContextMenu: false })

  rightClickObject(mounted.container, EXTERNAL_KEY)
  await vue.nextTick()
  await vue.nextTick()
  assert.equal(menuElement(), null)
  assert.equal(mounted.menus.length, 1)

  mounted.unmount()
  assert.deepEqual(mounted.warnings, [])
})

test("switching slides closes a stale menu", async () => {
  const mounted = await mountViewer()

  rightClickObject(mounted.container, EXTERNAL_KEY)
  await waitForDom(() => menuElement() !== null, "link menu")

  const viewer = mounted.container.querySelector('[data-testid="pptx-stage"]')
  const slide = document.createElement("div")
  slide.dataset.slideIndex = "1"
  viewer.append(slide)
  slide.dispatchEvent(new window.MouseEvent("click", { bubbles: true }))
  await waitForDom(() => menuElement() === null, "menu closes on slide change")

  mounted.unmount()
  assert.deepEqual(mounted.warnings, [])
})

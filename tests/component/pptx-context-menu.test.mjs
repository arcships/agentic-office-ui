// Issue #161 — the PPTX link menu on the shared Reka dropdown.
//
// This is intentionally a component-level test: mount `PptxLinkMenu`, not the
// whole `PptxViewer`. Mounting the full viewer under Happy DOM was the cause of
// a 2.8G+ test process and an OOM-killed terminal. Viewer integration is
// covered by the browser suite; this file pins the menu contract only.

import assert from "node:assert/strict"
import test, { after } from "node:test"
import { GlobalRegistrator } from "@happy-dom/global-registrator"

GlobalRegistrator.register({ url: "http://localhost/" })

const { importFromDemo, vue } = await import("./vue-test-renderer.mjs")
const { PptxLinkMenu } = await importFromDemo("@arcships/vue-pptx")

after(async () => {
  await new Promise((resolve) => setTimeout(resolve, 20))
  GlobalRegistrator.unregister()
})

const ITEMS = [
  { id: "open", label: "Open link", testId: "pptx-link-open" },
  { id: "copy", label: "Copy link", testId: "pptx-link-copy" },
]

async function waitForDom(predicate, message, timeoutMs = 2_000) {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    await vue.nextTick()
    if (predicate()) return
    await new Promise((resolve) => setTimeout(resolve, 5))
  }
  throw new Error(`Timed out waiting for ${message}`)
}

function menuElement() {
  return document.querySelector('[role="menu"]')
}

async function mountMenu(options = {}) {
  const container = document.createElement("div")
  document.body.append(container)
  const events = []
  const restoreFocus = []
  const open = vue.ref(options.open ?? true)
  const app = vue.createApp({
    render: () => vue.h(PptxLinkMenu, {
      open: open.value,
      x: options.x ?? 12,
      y: options.y ?? 24,
      items: options.items ?? ITEMS,
      label: options.label ?? "Link actions",
      restoreFocus: () => {
        restoreFocus.push(true)
        return container
      },
      "onUpdate:open": (value) => {
        open.value = value
        events.push(["open", value])
      },
      onSelect: (id) => {
        events.push(["select", id])
        open.value = false
      },
    }),
  })
  app.mount(container)
  await waitForDom(() => menuElement() !== null, "menu open")
  return { app, container, events, restoreFocus, open }
}

test("renders the shared menu surface with the pointer anchor", async () => {
  const mounted = await mountMenu({ x: 18, y: 36 })
  const menu = menuElement()
  assert.equal(menu.getAttribute("aria-label"), "Link actions")
  assert.match(menu.textContent ?? "", /Open link/)
  assert.match(menu.textContent ?? "", /Copy link/)

  const anchor = mounted.container.querySelector('[data-testid="pptx-context-menu-anchor"]')
  assert.ok(anchor, "anchor exists")
  assert.match(anchor.getAttribute("style") ?? "", /inset-block-start:\s*36px/)
  assert.match(anchor.getAttribute("style") ?? "", /inset-inline-start:\s*18px/)

  mounted.open.value = false
  await vue.nextTick()
  mounted.app.unmount()
  mounted.container.remove()
})

test("selecting an item reports its stable id and closes", async () => {
  const mounted = await mountMenu()
  const copy = menuElement().querySelector('[data-testid="pptx-link-copy"]')
  assert.ok(copy, "copy item rendered")
  copy.click()
  await waitForDom(() => mounted.events.some(([kind]) => kind === "select"), "select event")
  assert.deepEqual(mounted.events.find(([kind]) => kind === "select"), ["select", "copy"])
  await waitForDom(() => menuElement() === null, "menu closes after select")

  mounted.app.unmount()
  mounted.container.remove()
})

test("Escape closes the menu and restores focus", async () => {
  const mounted = await mountMenu()
  const menu = menuElement()
  menu.dispatchEvent(new window.KeyboardEvent("keydown", { key: "Escape", bubbles: true }))
  await waitForDom(
    () => mounted.events.some(([kind, value]) => kind === "open" && value === false),
    "close event",
  )
  await waitForDom(() => mounted.restoreFocus.length > 0, "restore focus callback")

  mounted.open.value = false
  await vue.nextTick()
  mounted.app.unmount()
  mounted.container.remove()
})

test("disabled entries do not emit a selection", async () => {
  const mounted = await mountMenu({
    items: [
      { id: "open", label: "Open link", testId: "pptx-link-open" },
      { id: "copy", label: "Copy link", testId: "pptx-link-copy", disabled: true },
    ],
  })
  const copy = menuElement().querySelector('[data-testid="pptx-link-copy"]')
  assert.ok(copy, "disabled copy item rendered")
  copy.click()
  await vue.nextTick()
  assert.equal(mounted.events.some(([kind]) => kind === "select"), false)

  mounted.open.value = false
  await vue.nextTick()
  mounted.app.unmount()
  mounted.container.remove()
})

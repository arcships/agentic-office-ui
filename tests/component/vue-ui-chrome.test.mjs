// Issue #158 — shared Office chrome runtime coverage.
//
// Happy DOM gives the components the focus, pointer, portal and
// ResizeObserver surface they need. The suite covers toolbar structure,
// page/zoom/search keyboard behaviour, status and state semantics, and the
// inline/overlay sidebar contract.

import assert from "node:assert/strict"
import test, { after } from "node:test"
import { GlobalRegistrator } from "@happy-dom/global-registrator"

GlobalRegistrator.register({ url: "http://localhost/" })

const { createRequire } = await import("node:module")
const { pathToFileURL } = await import("node:url")

const requireFromDemo = createRequire(
  new URL("../../apps/demo/package.json", import.meta.url),
)
const requireFromUi = createRequire(
  new URL("../../packages/vue-ui/package.json", import.meta.url),
)
const vue = await import(pathToFileURL(requireFromDemo.resolve("vue")).href)
const ui = await import(
  pathToFileURL(requireFromDemo.resolve("@arcships/vue-ui")).href
)
const interaction = await import(
  pathToFileURL(requireFromUi.resolve("@arcships/office-interaction")).href
)

const {
  computed,
  createApp,
  defineComponent,
  h,
  nextTick,
  ref,
} = vue

const {
  DocumentViewerSidebar,
  DocumentViewerSidebarSkeleton,
  OfficeEmptyState,
  OfficeErrorState,
  OfficeIconButton,
  OfficeLoadingState,
  OfficePageControl,
  OfficeSearchPopover,
  OfficeStatusBar,
  OfficeToolbar,
  OfficeToolbarGroup,
  OfficeToolbarSeparator,
  OfficeZoomControl,
} = ui

const { OFFICE_I18N_INJECTION_KEY, OFFICE_LOCALE_INJECTION_KEY, officeMessagesEn, officeMessagesZhCN } =
  interaction

after(() => {
  GlobalRegistrator.unregister()
})

function mountRoot(render, { locale = "en" } = {}) {
  const host = document.createElement("div")
  document.body.append(host)
  const app = createApp({ render })
  const messages = locale === "zh-CN" ? officeMessagesZhCN : officeMessagesEn
  app.provide(OFFICE_I18N_INJECTION_KEY, computed(() => messages))
  app.provide(OFFICE_LOCALE_INJECTION_KEY, computed(() => locale))
  app.mount(host)
  const cleanup = () => {
    app.unmount()
    host.remove()
    document.body.innerHTML = ""
  }
  return { app, host, cleanup }
}

function mount(component, { props = {}, slots = {}, locale = "en" } = {}) {
  return mountRoot(() => h(component, props, slots), { locale })
}

async function flush(ms = 30) {
  await nextTick()
  await new Promise((resolve) => setTimeout(resolve, ms))
  await nextTick()
}

const byTestId = (id, root = document) =>
  root.querySelector(`[data-testid="${id}"]`)
const bySlot = (slot, root = document) =>
  root.querySelector(`[data-slot="${slot}"]`)

function click(target) {
  target.dispatchEvent(
    new PointerEvent("pointerdown", { bubbles: true, button: 0 }),
  )
  target.dispatchEvent(
    new PointerEvent("pointerup", { bubbles: true, button: 0 }),
  )
  target.click()
}

// Reka Select opens on pointerdown and treats the following click as the same
// gesture; a pointerup would toggle it closed again.
function openSelect(trigger) {
  trigger.dispatchEvent(
    new PointerEvent("pointerdown", { bubbles: true, button: 0 }),
  )
  trigger.click()
}

function keydown(target, key, init = {}) {
  target.dispatchEvent(
    new KeyboardEvent("keydown", {
      key,
      bubbles: true,
      cancelable: true,
      ...init,
    }),
  )
}

function setInputValue(input, value) {
  input.value = value
  input.dispatchEvent(new Event("input", { bubbles: true }))
}

// ── Toolbar ───────────────────────────────────────────────────────────────

test("Toolbar exposes a labelled toolbar with leading, center and trailing clusters", async (t) => {
  const { cleanup } = mount(OfficeToolbar, {
    props: { label: "Document tools", sticky: true },
    slots: {
      default: () => [
        h(OfficeToolbarGroup, { gap: "md", align: "end", label: "History" }, () => "leading"),
        h(OfficeToolbarSeparator),
      ],
      center: () => h(OfficeToolbarGroup, { gap: "sm" }, () => "center"),
      trailing: () => h(OfficeToolbarGroup, { gap: "sm", grow: true }, () => "trailing"),
    },
  })
  t.after(cleanup)
  await flush(0)

  const toolbar = bySlot("toolbar")
  assert.equal(toolbar.getAttribute("role"), "toolbar")
  assert.equal(toolbar.getAttribute("aria-label"), "Document tools")
  assert.match(toolbar.className, /office-toolbar/)
  assert.match(toolbar.className, /office-toolbar--sticky/)
  assert.equal(toolbar.dataset.testid, "office-toolbar")

  const groups = [...document.querySelectorAll('[data-slot="toolbar-group"]')]
  assert.equal(groups.length, 3)
  assert.match(groups[0].className, /office-toolbar__group--gap-md/)
  assert.match(groups[0].className, /office-toolbar__group--align-end/)
  assert.equal(groups[0].getAttribute("aria-label"), "History")
  assert.match(groups[2].className, /office-toolbar__group--grow/)

  const separator = bySlot("toolbar-separator")
  assert.ok(separator, "separator must render")
  assert.equal(separator.getAttribute("data-orientation"), "vertical")
  assert.match(separator.className, /office-toolbar__separator/)
  assert.equal(byTestId("office-toolbar-separator"), separator)
})

test("Toolbar defaults its label from the shared catalog and supports no-wrap", async (t) => {
  const { cleanup } = mount(OfficeToolbar, { props: { wrap: false } })
  t.after(cleanup)
  await flush(0)

  const toolbar = bySlot("toolbar")
  assert.equal(toolbar.getAttribute("aria-label"), "Document toolbar")
  assert.match(toolbar.className, /office-toolbar--nowrap/)
})

// ── Icon button ───────────────────────────────────────────────────────────

test("IconButton keeps ghost/active/disabled states and 28/32px size classes", async (t) => {
  const { cleanup } = mountRoot(() =>
    h("div", [
      h(OfficeIconButton, {
        label: "More",
        icon: "more-horizontal",
        active: true,
        "data-testid": "icon-active",
      }),
      h(OfficeIconButton, {
        label: "Search",
        icon: "search",
        size: "icon-sm",
        "data-testid": "icon-small",
      }),
      h(OfficeIconButton, {
        label: "Delete",
        icon: "trash-2",
        disabled: true,
        "data-testid": "icon-disabled",
      }),
    ]),
  )
  t.after(cleanup)
  await flush(0)

  const active = byTestId("icon-active")
  assert.match(active.className, /office-button--ghost/)
  assert.match(active.className, /office-button--icon\b/)
  assert.equal(active.getAttribute("aria-pressed"), "true")
  assert.equal(active.dataset.active, "true")
  assert.equal(active.getAttribute("aria-label"), "More")

  const small = byTestId("icon-small")
  assert.match(small.className, /office-button--icon-sm/)
  assert.equal(small.hasAttribute("aria-pressed"), false)

  const disabled = byTestId("icon-disabled")
  assert.equal(disabled.disabled, true)
  assert.equal(disabled.getAttribute("data-disabled"), "true")
  assert.match(disabled.className, /office-button--icon\b/)
})

// ── Page control ──────────────────────────────────────────────────────────

test("PageControl commits typed pages, clamps, and supports keyboard stepping", async (t) => {
  const page = ref(3)
  const changes = []
  const { cleanup } = mountRoot(() =>
    h(OfficePageControl, {
      modelValue: page.value,
      "onUpdate:modelValue": (value) => {
        page.value = value
        changes.push(value)
      },
      totalPages: 12,
    }),
  )
  t.after(cleanup)
  await flush(0)

  const input = byTestId("office-page-input")
  assert.equal(input.value, "3")
  assert.equal(input.getAttribute("aria-label"), "Current page")
  assert.equal(byTestId("office-page-total").textContent.trim(), "of 12")
  assert.equal(byTestId("office-page-control").getAttribute("role"), "group")

  setInputValue(input, "9")
  await flush(0)
  keydown(input, "Enter")
  await flush(0)
  assert.deepEqual(changes, [9])
  assert.equal(page.value, 9)
  assert.equal(input.value, "9")

  setInputValue(input, "99")
  keydown(input, "Enter")
  await flush(0)
  assert.equal(changes.at(-1), 12)

  setInputValue(input, "0")
  keydown(input, "Enter")
  await flush(0)
  assert.equal(changes.at(-1), 1)

  setInputValue(input, "not a page")
  await flush(0)
  keydown(input, "Enter")
  await flush(0)
  assert.equal(input.value, "1", "invalid draft resets to the current page")

  keydown(input, "ArrowUp")
  await flush(0)
  assert.equal(changes.at(-1), 2)
  keydown(input, "ArrowDown")
  await flush(0)
  assert.equal(changes.at(-1), 1)
  keydown(input, "End")
  await flush(0)
  assert.equal(changes.at(-1), 12)
  keydown(input, "Home")
  await flush(0)
  assert.equal(changes.at(-1), 1)
})

test("PageControl resets an Escape draft, commits on blur and supports uncontrolled use", async (t) => {
  const page = ref(4)
  const changes = []
  const { cleanup } = mountRoot(() =>
    h("div", [
      h(OfficePageControl, {
        modelValue: page.value,
        "onUpdate:modelValue": (value) => {
          page.value = value
          changes.push(value)
        },
        totalPages: 8,
        "data-testid": "controlled",
      }),
      h(OfficePageControl, {
        defaultPage: 2,
        totalPages: 8,
        "onUpdate:modelValue": (value) => changes.push(["uncontrolled", value]),
        "data-testid": "uncontrolled",
      }),
    ]),
  )
  t.after(cleanup)
  await flush(0)

  const controlled = byTestId("controlled").querySelector("input")
  setInputValue(controlled, "7")
  await flush(0)
  keydown(controlled, "Escape")
  await flush(0)
  assert.equal(controlled.value, "4", "Escape restores the model value")

  setInputValue(controlled, "6")
  await flush(0)
  controlled.dispatchEvent(new Event("blur", { bubbles: true }))
  await flush(0)
  assert.equal(page.value, 6)

  const uncontrolled = byTestId("uncontrolled").querySelector("input")
  assert.equal(uncontrolled.value, "2")
  setInputValue(uncontrolled, "5")
  await flush(0)
  keydown(uncontrolled, "Enter")
  await flush(0)
  assert.deepEqual(changes.at(-1), ["uncontrolled", 5])
  assert.equal(uncontrolled.value, "5")
})

test("PageControl disables its input for an empty document", async (t) => {
  const { cleanup } = mount(OfficePageControl, {
    props: { modelValue: 1, totalPages: 0 },
  })
  t.after(cleanup)
  await flush(0)

  const input = byTestId("office-page-input")
  assert.equal(input.disabled, true)
  assert.equal(byTestId("office-page-total").textContent.trim(), "of –")
})

// ── Zoom control ─────────────────────────────────────────────────────────

test("ZoomControl steps through presets and emits controlled zoom/fit changes", async (t) => {
  const zoom = ref(100)
  const fitMode = ref("none")
  const changes = []
  const { cleanup } = mountRoot(() =>
    h(OfficeZoomControl, {
      modelValue: zoom.value,
      "onUpdate:modelValue": (value) => (zoom.value = value),
      fitMode: fitMode.value,
      "onUpdate:fitMode": (value) => (fitMode.value = value),
      onChange: (value) => changes.push(value),
    }),
  )
  t.after(cleanup)
  await flush(0)

  assert.equal(byTestId("office-zoom-select").textContent.trim(), "100%")
  click(byTestId("office-zoom-out"))
  await flush(0)
  assert.equal(zoom.value, 75)
  click(byTestId("office-zoom-in"))
  await flush(0)
  assert.equal(zoom.value, 100)
  assert.deepEqual(changes.at(-1), { zoom: 100, fitMode: "none" })

  const trigger = byTestId("office-zoom-select").querySelector('[role="combobox"]')
  openSelect(trigger)
  await flush(60)
  assert.equal(trigger.getAttribute("aria-expanded"), "true")
  const items = [...document.querySelectorAll('[data-slot="select-item"]')]
  const fitWidth = items.find((item) => item.textContent.includes("Fit width"))
  assert.ok(fitWidth, "fit-width preset must be present")
  click(fitWidth)
  await flush(60)
  assert.equal(fitMode.value, "width")
  assert.equal(byTestId("office-zoom-select").textContent.trim(), "Fit width")

  openSelect(trigger)
  await flush(60)
  const actual = [...document.querySelectorAll('[data-slot="select-item"]')]
    .find((item) => item.textContent.includes("Actual size"))
  assert.ok(actual, "actual-size preset must be present")
  click(actual)
  await flush(60)
  assert.equal(zoom.value, 100)
  assert.equal(fitMode.value, "none")
})

test("ZoomControl clamps at its bounds and disables the matching button", async (t) => {
  const zoom = ref(50)
  const { cleanup } = mountRoot(() =>
    h(OfficeZoomControl, {
      modelValue: zoom.value,
      "onUpdate:modelValue": (value) => (zoom.value = value),
      min: 50,
      max: 200,
      options: [50, 100, 200],
    }),
  )
  t.after(cleanup)
  await flush(0)

  assert.equal(byTestId("office-zoom-out").disabled, true)
  assert.equal(byTestId("office-zoom-in").disabled, false)
  click(byTestId("office-zoom-in"))
  await flush(0)
  assert.equal(zoom.value, 100)
  click(byTestId("office-zoom-in"))
  await flush(0)
  assert.equal(zoom.value, 200)
  assert.equal(byTestId("office-zoom-in").disabled, true)
  assert.equal(byTestId("office-zoom-out").disabled, false)
})

test("ZoomControl supports uncontrolled use and the disabled state", async (t) => {
  const changes = []
  const { cleanup } = mountRoot(() =>
    h("div", [
      h(OfficeZoomControl, {
        defaultZoom: 125,
        "onUpdate:modelValue": (value) => changes.push(value),
      }),
      h(OfficeZoomControl, { modelValue: 100, disabled: true }),
    ]),
  )
  t.after(cleanup)
  await flush(0)

  const controls = document.querySelectorAll('[data-testid="office-zoom-control"]')
  const first = controls[0]
  assert.equal(first.querySelector('[data-testid="office-zoom-select"]').textContent.trim(), "125%")
  click(first.querySelector('[data-testid="office-zoom-out"]'))
  await flush(0)
  assert.deepEqual(changes, [100])

  const disabled = controls[1]
  assert.equal(disabled.querySelector('[data-testid="office-zoom-out"]').disabled, true)
  assert.equal(disabled.querySelector('[data-testid="office-zoom-in"]').disabled, true)
  assert.equal(disabled.querySelector('[role="combobox"]').disabled, true)
})

// ── Search popover ────────────────────────────────────────────────────────

test("SearchPopover opens, focuses the input and navigates with Enter/Shift+Enter", async (t) => {
  const events = []
  const query = ref("quarter")
  const { cleanup } = mountRoot(() =>
    h(OfficeSearchPopover, {
      modelValue: query.value,
      "onUpdate:modelValue": (value) => {
        query.value = value
        events.push(["query", value])
      },
      resultCount: 3,
      activeIndex: 1,
      status: "ready",
      onNext: () => events.push(["next"]),
      onPrevious: () => events.push(["previous"]),
      onClose: () => events.push(["close"]),
      "onUpdate:open": (value) => events.push(["open", value]),
    }),
  )
  t.after(cleanup)
  await flush(0)

  const trigger = byTestId("office-search-trigger")
  assert.equal(trigger.getAttribute("aria-expanded"), "false")
  click(trigger)
  await flush(80)

  const popover = byTestId("office-search-popover")
  assert.ok(popover, "popover content must open")
  const input = byTestId("office-search-input")
  assert.equal(document.activeElement, input, "input receives focus on open")
  assert.equal(byTestId("office-search-status").textContent.trim(), "2 of 3")
  assert.equal(byTestId("office-search-previous").disabled, false)
  assert.equal(byTestId("office-search-next").disabled, false)

  setInputValue(input, "quarterly")
  await flush(0)
  assert.deepEqual(events.at(-1), ["query", "quarterly"])

  keydown(input, "Enter")
  await flush(0)
  assert.deepEqual(events.at(-1), ["next"])
  keydown(input, "Enter", { shiftKey: true })
  await flush(0)
  assert.deepEqual(events.at(-1), ["previous"])

  keydown(input, "Escape")
  await flush(80)
  assert.equal(trigger.getAttribute("aria-expanded"), "false")
  assert.ok(events.some((entry) => entry[0] === "close"))
})

test("SearchPopover honours defaultOpen and focuses the query field", async (t) => {
  const { cleanup } = mount(OfficeSearchPopover, {
    props: { defaultOpen: true, modelValue: "quarter", resultCount: 2 },
  })
  t.after(cleanup)
  await flush(80)

  const popover = byTestId("office-search-popover")
  assert.ok(popover, "defaultOpen must render the content")
  assert.equal(document.activeElement, byTestId("office-search-input"))
  assert.equal(byTestId("office-search-trigger").getAttribute("aria-expanded"), "true")
})

test("SearchPopover reports empty/searching/error states and submits a cold query", async (t) => {
  const events = []
  const { cleanup } = mountRoot(() =>
    h(OfficeSearchPopover, {
      modelValue: "missing",
      resultCount: 0,
      status: "ready",
      onSubmit: (value) => events.push(["submit", value]),
      onNext: () => events.push(["next"]),
    }),
  )
  await flush(0)

  click(byTestId("office-search-trigger"))
  await flush(80)
  const input = byTestId("office-search-input")
  assert.equal(byTestId("office-search-status").textContent.trim(), "No results")
  assert.equal(byTestId("office-search-previous").disabled, true)
  assert.equal(byTestId("office-search-next").disabled, true)

  keydown(input, "Enter")
  await flush(0)
  assert.deepEqual(events, [["submit", "missing"]])
  cleanup()

  const searchingEvents = []
  const { cleanup: cleanup2 } = mountRoot(() =>
    h(OfficeSearchPopover, {
      status: "searching",
      modelValue: "quarter",
      onSubmit: (value) => searchingEvents.push(["submit", value]),
    }),
  )
  t.after(cleanup2)
  await flush(0)
  click(byTestId("office-search-trigger"))
  await flush(80)
  assert.equal(byTestId("office-search-status").textContent.trim(), "Searching…")
  assert.equal(byTestId("office-search-next").disabled, true)
})

test("SearchPopover toggles match options and localizes through the Office catalog", async (t) => {
  const flags = []
  const { cleanup } = mountRoot(
    () =>
      h(OfficeSearchPopover, {
        modelValue: "quarter",
        resultCount: 1,
        activeIndex: 0,
        status: "ready",
        "onUpdate:caseSensitive": (value) => flags.push(["case", value]),
        "onUpdate:wholeWord": (value) => flags.push(["word", value]),
        "onUpdate:regex": (value) => flags.push(["regex", value]),
      }),
    { locale: "zh-CN" },
  )
  t.after(cleanup)
  await flush(0)

  const trigger = byTestId("office-search-trigger")
  assert.equal(trigger.getAttribute("aria-label"), "查找")
  click(trigger)
  await flush(80)
  assert.equal(byTestId("office-search-status").textContent.trim(), "第 1 项，共 1 项")

  click(byTestId("office-search-case-sensitive"))
  await flush(0)
  click(byTestId("office-search-whole-word"))
  await flush(0)
  click(byTestId("office-search-regex"))
  await flush(0)
  assert.deepEqual(flags, [
    ["case", true],
    ["word", true],
    ["regex", true],
  ])
})

// ── Status bar ────────────────────────────────────────────────────────────

test("StatusBar renders page, status, file, warning and error regions", async (t) => {
  const { cleanup } = mount(OfficeStatusBar, {
    props: {
      page: 2,
      totalPages: 9,
      status: "Saved",
      fileName: "report.pdf",
      warning: "Some images were optimized",
      error: "Page 5 failed",
    },
  })
  t.after(cleanup)
  await flush(0)

  const bar = byTestId("office-status-bar")
  assert.equal(bar.tagName, "FOOTER")
  assert.equal(bar.getAttribute("aria-label"), "Document status")
  assert.equal(byTestId("office-status-bar-page").textContent.trim(), "Page 2 of 9")
  assert.equal(byTestId("office-status-bar-status").textContent.trim(), "Saved")
  assert.equal(
    byTestId("office-status-bar-file").querySelector(".office-status-bar__file-name").textContent,
    "report.pdf",
  )
  assert.match(byTestId("office-status-bar-file").textContent, /File\s*report\.pdf/)
  assert.equal(byTestId("office-status-bar-warning").getAttribute("role"), "status")
  assert.equal(
    byTestId("office-status-bar-warning").getAttribute("aria-label"),
    "Warning: Some images were optimized",
  )
  assert.match(byTestId("office-status-bar-warning").textContent, /WarningSome images were optimized/)
  assert.equal(byTestId("office-status-bar-error").getAttribute("role"), "alert")
  assert.equal(
    byTestId("office-status-bar-error").getAttribute("aria-label"),
    "Error: Page 5 failed",
  )
  assert.match(byTestId("office-status-bar-error").textContent, /ErrorPage 5 failed/)
  assert.match(bar.className, /office-status-bar--attention/)
})

test("StatusBar defaults to the loading label while busy", async (t) => {
  const { cleanup } = mount(OfficeStatusBar, {
    props: { busy: true, page: 1, totalPages: 4 },
  })
  t.after(cleanup)
  await flush(0)

  const status = byTestId("office-status-bar-status")
  assert.equal(status.textContent.trim(), "Loading…")
  assert.ok(status.querySelector('[data-slot="spinner"]'))
  assert.equal(byTestId("office-status-bar-error"), null)
})

// ── Empty / error / loading states ────────────────────────────────────────

test("EmptyState uses catalog defaults and exposes an actions slot", async (t) => {
  const { cleanup } = mount(OfficeEmptyState, {
    slots: {
      actions: () => h("button", { "data-testid": "empty-action" }, "Open"),
    },
  })
  t.after(cleanup)
  await flush(0)

  const state = byTestId("office-empty-state")
  assert.equal(state.getAttribute("role"), "group")
  assert.match(state.className, /office-state--empty/)
  assert.match(state.textContent, /No document open/)
  assert.match(state.textContent, /Open a file to get started/)
  assert.equal(byTestId("empty-action").textContent, "Open")
  assert.ok(state.querySelector(".office-state__icon svg"))
})

test("EmptyState accepts a custom title, description and compact layout", async (t) => {
  const { cleanup } = mount(OfficeEmptyState, {
    props: {
      title: "Nothing here",
      description: "Drop a workbook",
      compact: true,
    },
  })
  t.after(cleanup)
  await flush(0)

  const state = byTestId("office-empty-state")
  assert.match(state.className, /office-state--compact/)
  assert.match(state.textContent, /Nothing here/)
  assert.match(state.textContent, /Drop a workbook/)
})

test("ErrorState is an alert, falls back to the raw error and emits retry", async (t) => {
  let retries = 0
  const { cleanup } = mount(OfficeErrorState, {
    props: { error: "Worker exploded", onRetry: () => (retries += 1) },
  })
  t.after(cleanup)
  await flush(0)

  const state = byTestId("office-error-state")
  assert.equal(state.getAttribute("role"), "alert")
  assert.match(state.className, /office-state--error/)
  assert.match(state.textContent, /Unable to open this document/)
  assert.match(state.textContent, /Worker exploded/)
  click(byTestId("office-error-retry"))
  await flush(0)
  assert.equal(retries, 1)
})

test("ErrorState can hide the retry button and render a custom action", async (t) => {
  const { cleanup } = mount(OfficeErrorState, {
    props: { showRetry: false, error: "Nope" },
    slots: { actions: () => h("button", { "data-testid": "custom-action" }, "Reload") },
  })
  t.after(cleanup)
  await flush(0)

  assert.equal(byTestId("office-error-retry"), null)
  assert.equal(byTestId("custom-action").textContent, "Reload")
})

test("LoadingState is a busy status with an optional deterministic progress bar", async (t) => {
  const { cleanup } = mount(OfficeLoadingState, {
    props: { progress: 42, description: "Rendering page 3…" },
  })
  t.after(cleanup)
  await flush(0)

  const state = byTestId("office-loading-state")
  assert.equal(state.getAttribute("role"), "status")
  assert.equal(state.getAttribute("aria-busy"), "true")
  assert.match(state.textContent, /Opening document/)
  assert.match(state.textContent, /Rendering page 3/)
  assert.ok(state.querySelector('[data-slot="spinner"]'))

  const progress = byTestId("office-loading-progress")
  assert.equal(progress.getAttribute("role"), "progressbar")
  assert.equal(progress.getAttribute("aria-valuemin"), "0")
  assert.equal(progress.getAttribute("aria-valuemax"), "100")
  assert.equal(progress.getAttribute("aria-valuenow"), "42")
  assert.match(progress.firstElementChild.getAttribute("style"), /inline-size: 42%/)
})

// ── Document sidebar ──────────────────────────────────────────────────────

function measuredContainer(width, height = 600) {
  const element = document.createElement("div")
  Object.defineProperty(element, "getBoundingClientRect", {
    configurable: true,
    value: () => ({
      width,
      height,
      top: 0,
      left: 0,
      right: width,
      bottom: height,
      x: 0,
      y: 0,
      toJSON() {},
    }),
  })
  document.body.append(element)
  return element
}

function installResizeObserver() {
  const original = globalThis.ResizeObserver
  const observers = []
  class TestResizeObserver {
    constructor(callback) {
      this.callback = callback
      this.targets = []
      observers.push(this)
    }
    observe(target) {
      this.targets.push(target)
    }
    unobserve() {}
    disconnect() {}
  }
  globalThis.ResizeObserver = TestResizeObserver
  return {
    observers,
    restore: () => {
      globalThis.ResizeObserver = original
    },
    emit(width) {
      for (const observer of observers) {
        observer.callback(
          observer.targets.map((target) => ({
            target,
            contentRect: { width, height: 600 },
          })),
        )
      }
    },
  }
}

test("DocumentViewerSidebar resolves auto mode from the measured container", async (t) => {
  const resize = installResizeObserver()
  t.after(resize.restore)

  const wide = measuredContainer(900)
  const narrow = measuredContainer(500)
  t.after(() => {
    wide.remove()
    narrow.remove()
  })

  const { cleanup } = mountRoot(() =>
    h("div", [
      h(DocumentViewerSidebar, { container: wide, open: true }, () => "wide"),
      h(DocumentViewerSidebar, { container: narrow, open: true }, () => "narrow"),
    ]),
  )
  t.after(cleanup)
  await flush(80)

  const [wideSidebar, narrowSidebar] = [
    ...document.querySelectorAll('[data-testid="document-viewer-sidebar"]'),
  ]
  assert.equal(wideSidebar.dataset.sidebarMode, "inline")
  assert.equal(wideSidebar.dataset.sidebarOpen, "true")
  assert.equal(wideSidebar.dataset.sidebarInline, "true")
  assert.equal(wideSidebar.getAttribute("style"), "--office-ui-sidebar-width: 160px;")
  assert.equal(wideSidebar.dataset.transitionsReady, "true")
  assert.equal(narrowSidebar.dataset.sidebarMode, "overlay")
  assert.equal(narrowSidebar.dataset.sidebarInline, "false")

  // Cross the breakpoint through the ResizeObserver callback.
  Object.defineProperty(narrow, "getBoundingClientRect", {
    configurable: true,
    value: () => ({ width: 900, height: 600 }),
  })
  resize.emit(900)
  await flush(40)
  assert.equal(narrowSidebar.dataset.sidebarMode, "inline")

  Object.defineProperty(narrow, "getBoundingClientRect", {
    configurable: true,
    value: () => ({ width: 500, height: 600 }),
  })
  resize.emit(500)
  await flush(40)
  assert.equal(narrowSidebar.dataset.sidebarMode, "overlay")
})

test("DocumentViewerSidebar toggles open/closed and closes an overlay with Escape", async (t) => {
  const resize = installResizeObserver()
  t.after(resize.restore)
  const narrow = measuredContainer(500)
  t.after(() => narrow.remove())

  const open = ref(true)
  const updates = []
  const modes = []
  const { cleanup } = mountRoot(() =>
    h(
      DocumentViewerSidebar,
      {
        container: narrow,
        open: open.value,
        "onUpdate:open": (value) => {
          updates.push(value)
          open.value = value
        },
        onModeChange: (value) => modes.push(value),
      },
      () => "body",
    ),
  )
  t.after(cleanup)
  await flush(80)

  const sidebar = byTestId("document-viewer-sidebar")
  assert.equal(sidebar.dataset.sidebarOpen, "true")
  assert.equal(sidebar.hasAttribute("aria-hidden"), false)
  assert.equal(sidebar.hasAttribute("inert"), false)

  document.dispatchEvent(
    new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }),
  )
  await flush(40)
  assert.deepEqual(updates, [false])
  assert.equal(open.value, false)
  assert.equal(sidebar.dataset.sidebarOpen, "false")
  assert.equal(sidebar.getAttribute("aria-hidden"), "true")
  assert.equal(sidebar.hasAttribute("inert"), true)
  assert.deepEqual(modes, ["overlay"], "initial overlay is reported once")
})

test("DocumentViewerSidebar renders header/footer slots, a skeleton and a custom width", async (t) => {
  const resize = installResizeObserver()
  t.after(resize.restore)
  const container = measuredContainer(900)
  t.after(() => container.remove())

  const { cleanup } = mountRoot(() =>
    h(
      DocumentViewerSidebar,
      { container, open: true, width: 200, loading: true },
      {
        header: () => h("span", { "data-testid": "sidebar-header" }, "Pages"),
        default: () => h("span", { "data-testid": "sidebar-content" }, "Never rendered"),
        footer: () => h("span", { "data-testid": "sidebar-footer" }, "6 pages"),
      },
    ),
  )
  t.after(cleanup)
  await flush(80)

  const sidebar = byTestId("document-viewer-sidebar")
  assert.equal(sidebar.getAttribute("style"), "--office-ui-sidebar-width: 200px;")
  assert.equal(byTestId("sidebar-header").textContent, "Pages")
  assert.equal(byTestId("sidebar-footer").textContent, "6 pages")
  assert.equal(byTestId("sidebar-content"), null)
  const skeleton = byTestId("document-viewer-sidebar-skeleton")
  assert.ok(skeleton)
  assert.equal(skeleton.getAttribute("role"), "status")
  assert.equal(
    skeleton.querySelectorAll('[data-slot="sidebar-skeleton-item"]').length,
    3,
  )
  assert.match(sidebar.className, /office-document-sidebar--inline/)
})

test("DocumentViewerSidebar mode prop overrides auto and the skeleton is standalone", async (t) => {
  const wide = measuredContainer(1200)
  t.after(() => wide.remove())

  const { cleanup } = mountRoot(() =>
    h("div", [
      h(DocumentViewerSidebar, { container: wide, mode: "overlay", open: true }, () => "overlay"),
      h(DocumentViewerSidebar, { container: wide, mode: "inline", open: false }, () => "inline"),
      h(DocumentViewerSidebarSkeleton, { items: 2 }),
    ]),
  )
  t.after(cleanup)
  await flush(80)

  const sidebars = [...document.querySelectorAll('[data-testid="document-viewer-sidebar"]')]
  assert.equal(sidebars[0].dataset.sidebarMode, "overlay")
  assert.equal(sidebars[1].dataset.sidebarMode, "inline")
  assert.equal(sidebars[1].dataset.sidebarOpen, "false")
  assert.equal(
    byTestId("document-viewer-sidebar-skeleton").querySelectorAll('[data-slot="sidebar-skeleton-item"]').length,
    2,
  )
})

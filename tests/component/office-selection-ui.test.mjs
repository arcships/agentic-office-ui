import assert from "node:assert/strict"
import test from "node:test"

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  walk,
} from "./vue-test-renderer.mjs"

const {
  OfficeObjectOutlineLayer,
  OfficeRegionSelector,
} = await importFromDemo("@arcships/vue-ui")

function assertRect(actual, expected) {
  for (const key of ["x", "y", "width", "height"]) {
    assert.ok(Math.abs(actual[key] - expected[key]) < 1e-12, `${key}: expected ${expected[key]}, got ${actual[key]}`)
  }
}

test("OfficeObjectOutlineLayer exposes normalized outlines and keyboard candidate intents", async () => {
  const activated = []
  const confirmed = []
  let dismissed = 0
  const hierarchy = []
  const items = [
    {
      id: "chart-2",
      referenceId: "reference-1",
      label: "收入图表",
      kind: "chart",
      rect: { x: 0.1, y: 0.2, width: 0.5, height: 0.4 },
      state: "active",
      reliability: "exact",
    },
    {
      id: "slide-5",
      label: "幻灯片 5",
      kind: "slide",
      rect: { x: 0, y: 0, width: 1, height: 1 },
      reliability: "likely",
    },
  ]
  const mounted = await mount(OfficeObjectOutlineLayer, {
    items,
    activeId: "chart-2",
    onActivate: (item) => activated.push(item.id),
    onConfirm: (item, options) => confirmed.push([item.id, options]),
    onDismiss: () => { dismissed += 1 },
    onNavigateHierarchy: (direction) => hierarchy.push(direction),
  })
  const root = findByTestId(mounted.root, "office-object-outline-layer")
  const chart = walk(root).find((node) => node.props?.["data-outline-id"] === "chart-2")
  assert.equal(root.props.role, "listbox")
  assert.equal(chart.props.style.left, "10%")
  assert.equal(chart.props.style.width, "50%")
  assert.equal(chart.props["aria-selected"], true)

  chart.props.onClick({ altKey: true, shiftKey: true, stopPropagation() {} })
  root.props.onKeydown({ key: "Tab", shiftKey: false, preventDefault() {} })
  root.props.onKeydown({ key: "ArrowRight", preventDefault() {} })
  root.props.onKeydown({ key: "Escape", preventDefault() {} })
  await vue.nextTick()

  assert.deepEqual(confirmed, [["chart-2", { additiveRequested: true, penetrateRequested: true }]])
  assert.deepEqual(activated, ["slide-5"])
  assert.deepEqual(hierarchy, ["child"])
  assert.equal(dismissed, 1)
  assert.match(textContent(root), /收入图表/u)
  assert.deepEqual(mounted.warnings, [])
  mounted.app.unmount()
})

test("OfficeRegionSelector emits normalized drag geometry and keyboard adjustments", async () => {
  const changes = []
  const commits = []
  const cancellations = []
  const mounted = await mount(OfficeRegionSelector, {
    modelValue: null,
    "onUpdate:modelValue": (rect) => changes.push(rect),
    onSelectionCommit: (rect) => commits.push(rect),
    onSelectionCancel: (rect) => cancellations.push(rect),
  })
  const root = findByTestId(mounted.root, "office-region-selector")
  const captures = new Set()
  root.getBoundingClientRect = () => ({ left: 10, top: 20, width: 200, height: 100 })
  root.setPointerCapture = (pointerId) => captures.add(pointerId)
  root.hasPointerCapture = (pointerId) => captures.has(pointerId)
  root.releasePointerCapture = (pointerId) => captures.delete(pointerId)
  const pointer = (clientX, clientY) => ({
    button: 0,
    clientX,
    clientY,
    currentTarget: root,
    isPrimary: true,
    pointerId: 7,
    preventDefault() {},
  })

  root.props.onPointerdown(pointer(30, 40))
  root.props.onPointermove(pointer(130, 80))
  await vue.nextTick()
  assertRect(changes.at(-1), { x: 0.1, y: 0.2, width: 0.5, height: 0.4 })
  assert.ok(findByTestId(mounted.root, "office-region-frame"))
  root.props.onPointerup(pointer(130, 80))
  await vue.nextTick()
  assert.equal(commits.length, 1)
  assertRect(commits[0], { x: 0.1, y: 0.2, width: 0.5, height: 0.4 })
  assert.equal(captures.size, 0)
  assert.equal(findByTestId(mounted.root, "office-region-frame"), undefined)
  assert.deepEqual(cancellations, [])
  assert.deepEqual(mounted.warnings, [])
  mounted.app.unmount()

  const directCommits = []
  const direct = await mount(OfficeRegionSelector, {
    modelValue: null,
    onSelectionCommit: (rect) => directCommits.push(rect),
  })
  const directRoot = findByTestId(direct.root, "office-region-selector")
  directRoot.getBoundingClientRect = () => ({ left: 0, top: 0, width: 100, height: 100 })
  directRoot.setPointerCapture = () => {}
  directRoot.hasPointerCapture = () => false
  const directPointer = (clientX, clientY) => ({
    button: 0,
    clientX,
    clientY,
    currentTarget: directRoot,
    isPrimary: true,
    pointerId: 8,
    preventDefault() {},
  })
  directRoot.props.onPointerdown(directPointer(10, 10))
  directRoot.props.onPointerup(directPointer(60, 70))
  assert.equal(directCommits.length, 1)
  assertRect(directCommits[0], { x: 0.1, y: 0.1, width: 0.5, height: 0.6 })
  direct.app.unmount()

  const keyboardChanges = []
  const keyboardCommits = []
  const keyboard = await mount(OfficeRegionSelector, {
    modelValue: { x: 0.1, y: 0.2, width: 0.3, height: 0.4 },
    keyboardStep: 0.05,
    "onUpdate:modelValue": (rect) => keyboardChanges.push(rect),
    onSelectionCommit: (rect) => keyboardCommits.push(rect),
  })
  const keyboardRoot = findByTestId(keyboard.root, "office-region-selector")
  keyboardRoot.props.onKeydown({ key: "ArrowRight", shiftKey: false, preventDefault() {} })
  keyboardRoot.props.onKeydown({ key: "ArrowDown", shiftKey: true, preventDefault() {} })
  keyboardRoot.props.onKeydown({ key: "Enter", preventDefault() {} })
  assert.deepEqual(keyboardChanges, [
    { x: 0.15000000000000002, y: 0.2, width: 0.3, height: 0.4 },
    { x: 0.1, y: 0.2, width: 0.3, height: 0.45 },
  ])
  assert.deepEqual(keyboardCommits, [{ x: 0.1, y: 0.2, width: 0.3, height: 0.4 }])
  keyboard.app.unmount()
})

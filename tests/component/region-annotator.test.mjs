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

const { OfficeRegionAnnotator } = await importFromDemo("@arcships/vue-ui")

const document = {
  schemaVersion: 1,
  annotations: [
    {
      id: "rect-1",
      kind: "rectangle",
      rect: { x: 0.1, y: 0.2, width: 0.3, height: 0.4 },
      style: { color: "#ef4444", strokeWidth: 0.006 },
    },
    {
      id: "arrow-1",
      kind: "arrow",
      start: { x: 0.8, y: 0.1 },
      end: { x: 0.45, y: 0.55 },
      style: { color: "#2563eb" },
    },
    {
      id: "text-1",
      kind: "text",
      point: { x: 0.2, y: 0.75 },
      text: "See clause 4.2",
      style: { color: "#18181b" },
    },
    {
      id: "highlight-1",
      kind: "highlight",
      rect: { x: 0.05, y: 0.6, width: 0.5, height: 0.1 },
      style: { color: "#fde047", opacity: 0.35 },
    },
  ],
}

test("OfficeRegionAnnotator renders serialized annotations over the captured bitmap", async () => {
  const mounted = await mount(OfficeRegionAnnotator, {
    modelValue: document,
    image: "data:image/png;base64,AAAA",
    width: 800,
    height: 600,
  })
  const root = findByTestId(mounted.root, "office-region-annotator")
  const overlay = findByTestId(mounted.root, "office-region-annotator-overlay")
  assert.ok(root)
  assert.equal(overlay.props.viewBox, "0 0 800 600")
  assert.equal(findByTestId(mounted.root, "office-region-annotation-rect-1").props["data-kind"], "rectangle")
  assert.equal(findByTestId(mounted.root, "office-region-annotation-arrow-1").props["data-kind"], "arrow")
  assert.equal(findByTestId(mounted.root, "office-region-annotation-highlight-1").props["data-kind"], "highlight")
  assert.match(textContent(findByTestId(mounted.root, "office-region-annotation-text-1")), /See clause 4\.2/u)
  assert.equal(findByTestId(mounted.root, "office-region-annotator-count").children[0].text.trim(), "4 annotations")
  assert.deepEqual(mounted.warnings, [])
  mounted.app.unmount()
})

test("OfficeRegionAnnotator emits normalized geometry when dragging a rectangle", async () => {
  const updates = []
  const added = []
  const mounted = await mount(OfficeRegionAnnotator, {
    modelValue: { schemaVersion: 1, annotations: [] },
    image: "data:image/png;base64,AAAA",
    width: 800,
    height: 600,
    "onUpdate:modelValue": (next) => updates.push(next),
    onAnnotationAdd: (annotation) => added.push(annotation),
  })
  const stage = findByTestId(mounted.root, "office-region-annotator-stage")
  stage.getBoundingClientRect = () => ({ left: 0, top: 0, width: 800, height: 600 })
  stage.setPointerCapture = () => {}
  stage.hasPointerCapture = () => false
  const pointer = (clientX, clientY, pointerId = 4) => ({
    button: 0,
    clientX,
    clientY,
    currentTarget: stage,
    isPrimary: true,
    pointerId,
    preventDefault() {},
  })
  stage.props.onPointerdown(pointer(80, 120))
  stage.props.onPointermove(pointer(480, 360))
  stage.props.onPointerup(pointer(480, 360))
  await vue.nextTick()
  assert.equal(updates.length, 1)
  assert.equal(added.length, 1)
  const annotation = updates[0].annotations[0]
  assert.equal(annotation.kind, "rectangle")
  for (const [key, expected] of Object.entries({ x: 0.1, y: 0.2, width: 0.5, height: 0.4 })) {
    assert.ok(Math.abs(annotation.rect[key] - expected) < 1e-12, `${key}: expected ${expected}, got ${annotation.rect[key]}`)
  }
  assert.equal(annotation.style.strokeWidth, 0.006)
  assert.equal(annotation.style.color, "#ef4444")
  assert.deepEqual(mounted.warnings, [])
  mounted.app.unmount()
})

test("OfficeRegionAnnotator owns and revokes Blob image URLs", async () => {
  const created = []
  const revoked = []
  const originalCreate = URL.createObjectURL
  const originalRevoke = URL.revokeObjectURL
  URL.createObjectURL = (blob) => {
    const url = `blob:annotator-${created.length + 1}`
    created.push({ blob, url })
    return url
  }
  URL.revokeObjectURL = (url) => revoked.push(url)
  try {
    const mounted = await mount(OfficeRegionAnnotator, {
      modelValue: { schemaVersion: 1, annotations: [] },
      image: new Blob(["image"], { type: "image/png" }),
      width: 320,
      height: 200,
    })
    assert.equal(created.length, 1)
    assert.equal(findByTestId(mounted.root, "office-region-annotator-image").props.src, created[0].url)
    mounted.app.unmount()
    assert.deepEqual(revoked, [created[0].url])
  } finally {
    URL.createObjectURL = originalCreate
    URL.revokeObjectURL = originalRevoke
  }
})

test("OfficeRegionAnnotator wires its tool palette and keeps the active tool pressed", async () => {
  const mounted = await mount(OfficeRegionAnnotator, {
    modelValue: document,
    image: "data:image/png;base64,AAAA",
    width: 800,
    height: 600,
  })
  const root = findByTestId(mounted.root, "office-region-annotator")
  assert.ok(root)
  // The tool buttons are wired and keep the active tool pressed.
  const arrow = findByTestId(mounted.root, "office-region-annotator-tool-arrow")
  arrow.props.onClick({})
  await vue.nextTick()
  assert.equal(findByTestId(mounted.root, "office-region-annotator-tool-arrow").props["aria-pressed"], true)
  assert.deepEqual(mounted.warnings, [])
  mounted.app.unmount()
})

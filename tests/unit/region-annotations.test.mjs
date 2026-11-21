import assert from "node:assert/strict"
import test from "node:test"

const {
  DEFAULT_REGION_ANNOTATION_MAX_COUNT,
  OfficeRegionAnnotationValidationError,
  createOfficeRegionAnnotationDocument,
  officeRegionAnnotationBounds,
  officeRegionAnnotationFromRegionAnnotation,
  parseOfficeRegionAnnotationDocument,
  stringifyOfficeRegionAnnotationDocument,
} = await import("../../packages/vue-ui/dist/index.js")

const annotations = [
  {
    id: "rect-1",
    kind: "rectangle",
    rect: { x: 0.1, y: 0.2, width: 0.3, height: 0.4 },
    style: { color: "#ef4444", strokeWidth: 0.006, opacity: 0.9 },
  },
  {
    id: "highlight-1",
    kind: "highlight",
    rect: { x: 0.05, y: 0.6, width: 0.5, height: 0.1 },
    style: { color: "#fde047", opacity: 0.35 },
  },
  {
    id: "arrow-1",
    kind: "arrow",
    start: { x: 0.8, y: 0.1 },
    end: { x: 0.45, y: 0.55 },
    style: { color: "#2563eb", strokeWidth: 0.004 },
  },
  {
    id: "text-1",
    kind: "text",
    point: { x: 0.2, y: 0.75 },
    text: "See clause 4.2",
    style: { color: "#18181b", fontSize: 0.05 },
  },
]

function invalid(document, code) {
  assert.throws(
    () => parseOfficeRegionAnnotationDocument(document),
    (error) => error instanceof OfficeRegionAnnotationValidationError && error.code === code,
  )
}

test("annotation documents round-trip through JSON with normalized geometry and style", () => {
  const document = createOfficeRegionAnnotationDocument({ annotations })
  const json = stringifyOfficeRegionAnnotationDocument(document)
  const parsed = parseOfficeRegionAnnotationDocument(json)
  assert.deepEqual(parsed, document)
  assert.deepEqual(parsed.annotations.map((annotation) => annotation.kind), [
    "rectangle",
    "highlight",
    "arrow",
    "text",
  ])
  assert.deepEqual(parsed.annotations[0].rect, { x: 0.1, y: 0.2, width: 0.3, height: 0.4 })
  assert.equal(parsed.annotations[3].text, "See clause 4.2")
})

test("annotation documents preserve capture source metadata", () => {
  const source = {
    kind: "canvas",
    format: "pdf",
    documentId: "contract.pdf",
    revision: "sha256:abc",
    space: "page",
    pageIndex: 1,
    region: { x: 0.1, y: 0.2, width: 0.3, height: 0.4 },
    scale: 2,
    width: 600,
    height: 400,
    devicePixelRatio: 2,
    elementWidth: 300,
    elementHeight: 200,
    sourceWidth: 1200,
    sourceHeight: 800,
    capturedAt: "2026-11-20T00:00:00.000Z",
    mimeType: "image/png",
    byteLength: 128,
  }
  const document = createOfficeRegionAnnotationDocument({ annotations: [], source })
  const parsed = parseOfficeRegionAnnotationDocument(JSON.parse(JSON.stringify(document)))
  assert.deepEqual(parsed.source, source)
})

test("annotation parsing rejects malformed coordinates, styles, and duplicate ids", () => {
  invalid({ schemaVersion: 2, annotations: [] }, "INVALID_DOCUMENT")
  invalid({ schemaVersion: 1 }, "INVALID_DOCUMENT")
  invalid({ schemaVersion: 1, annotations: [{ ...annotations[0], rect: { x: 0.1, y: 0.2, width: 0, height: 0.4 } }] }, "INVALID_COORDINATE")
  invalid({ schemaVersion: 1, annotations: [{ ...annotations[0], rect: { x: 0.9, y: 0.2, width: 0.3, height: 0.4 } }] }, "INVALID_COORDINATE")
  invalid({ schemaVersion: 1, annotations: [{ ...annotations[2], start: { x: 0.2, y: 0.2 }, end: { x: 0.2, y: 0.2 } }] }, "INVALID_COORDINATE")
  invalid({ schemaVersion: 1, annotations: [{ ...annotations[0], style: { strokeWidth: 2 } }] }, "INVALID_STYLE")
  invalid({ schemaVersion: 1, annotations: [{ ...annotations[3], text: "" }] }, "INVALID_ANNOTATION")
  invalid({ schemaVersion: 1, annotations: [{ ...annotations[0], kind: "ellipse" }] }, "INVALID_ANNOTATION")
  invalid({ schemaVersion: 1, annotations: [annotations[0], { ...annotations[0] }] }, "DUPLICATE_ID")
  invalid({ schemaVersion: 1, annotations: new Array(DEFAULT_REGION_ANNOTATION_MAX_COUNT + 1).fill(annotations[0]) }, "INVALID_DOCUMENT")
  assert.throws(() => parseOfficeRegionAnnotationDocument("not json"), /invalid JSON/u)
})

test("annotation bounds cover rectangles, arrows, and text points", () => {
  assert.deepEqual(officeRegionAnnotationBounds(annotations[0]), { x: 0.1, y: 0.2, width: 0.3, height: 0.4 })
  const arrowBounds = officeRegionAnnotationBounds(annotations[2])
  for (const [key, expected] of Object.entries({ x: 0.45, y: 0.1, width: 0.35, height: 0.45 })) {
    assert.ok(Math.abs(arrowBounds[key] - expected) < 1e-12, `${key}: expected ${expected}, got ${arrowBounds[key]}`)
  }
  assert.deepEqual(officeRegionAnnotationBounds(annotations[3]), { x: 0.2, y: 0.75, width: 0, height: 0 })
})

test("format-neutral RegionAnnotation values convert when the annotator supports the kind", () => {
  assert.deepEqual(
    officeRegionAnnotationFromRegionAnnotation({
      id: "legacy-rect",
      kind: "rectangle",
      rect: { x: 0.2, y: 0.3, width: 0.1, height: 0.1 },
      color: "#f97316",
    }),
    {
      id: "legacy-rect",
      kind: "rectangle",
      rect: { x: 0.2, y: 0.3, width: 0.1, height: 0.1 },
      style: { color: "#f97316" },
    },
  )
  assert.equal(officeRegionAnnotationFromRegionAnnotation({
    id: "freehand",
    kind: "freehand",
    points: [{ x: 0, y: 0 }, { x: 1, y: 1 }],
  }), null)
})

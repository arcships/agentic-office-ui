import assert from "node:assert/strict"
import test from "node:test"

const {
  OfficeRegionCaptureError,
  captureRegion,
  computeRegionCaptureGeometry,
  computeRegionRasterCrop,
  releaseRegionCapture,
} = await import("../../packages/vue-ui/dist/index.js")

function fakeCanvas(width, height, calls) {
  const context = {
    calls,
    fillStyle: "",
    save() {},
    restore() {},
    fillRect(...args) {
      calls.fillRect.push(args)
    },
    drawImage(...args) {
      calls.drawImage.push(args)
    },
  }
  return {
    width,
    height,
    getContext() {
      return context
    },
    toBlob(callback, type) {
      calls.mimeType = type
      callback(new Blob([new Uint8Array(32)], { type: type ?? "image/png" }))
    },
  }
}

function captureHarness({ elementWidth = 800, elementHeight = 600, rasterWidth = 1600, rasterHeight = 1200 } = {}) {
  const calls = { drawImage: [], fillRect: [], mimeType: "" }
  const element = {
    nodeType: 1,
    getBoundingClientRect: () => ({ left: 0, top: 0, width: elementWidth, height: elementHeight }),
  }
  const rasterSource = { width: rasterWidth, height: rasterHeight }
  return {
    calls,
    element,
    rasterSource,
    raster: { source: rasterSource, width: rasterWidth, height: rasterHeight },
    createCanvas: (width, height) => fakeCanvas(width, height, calls),
  }
}

const REGION = { x: 0.1, y: 0.2, width: 0.5, height: 0.25 }

test("computeRegionCaptureGeometry maps normalized regions at 1x and 2x", () => {
  const at1x = computeRegionCaptureGeometry({
    region: REGION,
    elementWidth: 800,
    elementHeight: 600,
    scale: 1,
    devicePixelRatio: 2,
  })
  assert.equal(at1x.outputWidth, 400)
  assert.equal(at1x.outputHeight, 150)
  assert.equal(at1x.scale, 1)
  assert.equal(at1x.devicePixelRatio, 2)

  const at2x = computeRegionCaptureGeometry({
    region: REGION,
    elementWidth: 800,
    elementHeight: 600,
    scale: 2,
  })
  assert.equal(at2x.outputWidth, 800)
  assert.equal(at2x.outputHeight, 300)

  const crop = computeRegionRasterCrop(REGION, 1600, 1200)
  assert.deepEqual(crop, { x: 160, y: 240, width: 800, height: 300 })
})

test("captureRegion crops the raster source pixel-exactly at the requested scale", async () => {
  const harness = captureHarness()
  const capture = await captureRegion({
    surface: { getCaptureSurface: () => ({ element: harness.element, raster: harness.raster, format: "pdf", space: "page", pageIndex: 2 }) },
    region: REGION,
    scale: 1,
    createCanvas: harness.createCanvas,
    now: () => new Date("2026-11-20T00:00:00.000Z"),
  })
  assert.equal(capture.width, 400)
  assert.equal(capture.height, 150)
  assert.equal(capture.scale, 1)
  assert.deepEqual(capture.region, REGION)
  assert.deepEqual(capture.source.region, REGION)
  assert.equal(capture.source.format, "pdf")
  assert.equal(capture.source.space, "page")
  assert.equal(capture.source.pageIndex, 2)
  assert.equal(capture.source.kind, "bitmap")
  assert.equal(capture.source.elementWidth, 800)
  assert.equal(capture.source.elementHeight, 600)
  assert.equal(capture.source.sourceWidth, 1600)
  assert.equal(capture.source.sourceHeight, 1200)
  assert.equal(capture.source.capturedAt, "2026-11-20T00:00:00.000Z")
  assert.equal(capture.source.byteLength, 32)
  assert.deepEqual(harness.calls.drawImage[0], [
    harness.rasterSource,
    160,
    240,
    800,
    300,
    0,
    0,
    400,
    150,
  ])
  capture.release()
  releaseRegionCapture(capture)
})

test("captureRegion doubles output pixels at 2x without changing the source crop", async () => {
  const harness = captureHarness()
  const capture = await captureRegion({
    surface: { element: harness.element, raster: harness.raster },
    region: REGION,
    scale: 2,
    createCanvas: harness.createCanvas,
  })
  assert.equal(capture.width, 800)
  assert.equal(capture.height, 300)
  assert.deepEqual(harness.calls.drawImage[0].slice(1), [
    160,
    240,
    800,
    300,
    0,
    0,
    800,
    300,
  ])
  capture.release()
})

test("captureRegion resolves manual page locators through the surface provider", async () => {
  const harness = captureHarness()
  const requests = []
  let createdUrl = null
  const originalCreate = URL.createObjectURL
  const originalRevoke = URL.revokeObjectURL
  const revoked = []
  URL.createObjectURL = (blob) => {
    createdUrl = `blob:region-${blob.size}`
    return createdUrl
  }
  URL.revokeObjectURL = (url) => revoked.push(url)
  try {
    const capture = await captureRegion({
      surface: {
        getCaptureSurface(request) {
          requests.push(request)
          return {
            element: harness.element,
            raster: harness.raster,
            region: request.region.rect,
            format: "docx",
            space: "page",
            pageIndex: request.target?.pageIndex,
          }
        },
      },
      region: { space: "page", pageIndex: 3, rect: REGION },
      scale: 1,
      createCanvas: harness.createCanvas,
    })
    assert.equal(requests.length, 1)
    assert.deepEqual(requests[0].target, { pageIndex: 3 })
    assert.equal(requests[0].scale, 1)
    assert.equal(capture.source.pageIndex, 3)
    assert.equal(capture.url, createdUrl)
    capture.release()
    capture.release()
    assert.deepEqual(revoked, [createdUrl])
  } finally {
    URL.createObjectURL = originalCreate
    URL.revokeObjectURL = originalRevoke
  }
})

test("captureRegion enforces pixel, dimension, and byte limits", async () => {
  const harness = captureHarness({ elementWidth: 8000, elementHeight: 8000, rasterWidth: 8000, rasterHeight: 8000 })
  await assert.rejects(
    captureRegion({
      surface: { element: harness.element, raster: harness.raster },
      region: { x: 0, y: 0, width: 1, height: 1 },
      scale: 2,
      createCanvas: harness.createCanvas,
      maxDimension: 20_000,
    }),
    (error) => error instanceof OfficeRegionCaptureError && error.code === "CAPTURE_LIMIT_EXCEEDED",
  )
  await assert.rejects(
    captureRegion({
      surface: { element: harness.element, raster: harness.raster },
      region: { x: 0, y: 0, width: 0.1, height: 0.1 },
      scale: 1,
      createCanvas: harness.createCanvas,
      maxDimension: 100,
    }),
    (error) => error instanceof OfficeRegionCaptureError && error.code === "CAPTURE_LIMIT_EXCEEDED",
  )
  await assert.rejects(
    captureRegion({
      surface: { element: harness.element, raster: harness.raster },
      region: { x: 0, y: 0, width: 0.1, height: 0.1 },
      scale: 1,
      createCanvas: harness.createCanvas,
      maxBytes: 8,
    }),
    (error) => error instanceof OfficeRegionCaptureError && error.code === "CAPTURE_LIMIT_EXCEEDED",
  )
})

test("captureRegion rejects invalid regions and headless hosts without a canvas", async () => {
  const harness = captureHarness()
  await assert.rejects(
    captureRegion({
      surface: { element: harness.element, raster: harness.raster },
      region: { x: 0.5, y: 0.5, width: 0, height: 0.2 },
      createCanvas: harness.createCanvas,
    }),
    (error) => error instanceof OfficeRegionCaptureError && error.code === "INVALID_REGION",
  )
  await assert.rejects(
    captureRegion({
      surface: { element: harness.element },
      region: REGION,
      scale: 1,
    }),
    (error) => error instanceof OfficeRegionCaptureError && error.code === "CAPTURE_UNSUPPORTED",
  )
  await assert.rejects(
    captureRegion({
      surface: { getCaptureSurface: () => null },
      region: REGION,
      createCanvas: harness.createCanvas,
    }),
    (error) => error instanceof OfficeRegionCaptureError && error.code === "CAPTURE_UNSUPPORTED",
  )
})

import assert from "node:assert/strict";
import { createRequire } from "node:module";
import test from "node:test";
import { pathToFileURL } from "node:url";

const requireFromVueDocx = createRequire(
  new URL("../../packages/vue-docx/package.json", import.meta.url),
);
const core = await import(
  pathToFileURL(requireFromVueDocx.resolve("@arcships/docx-core")).href
);

function blankMetadata() {
  return { ...core.createBlankDocumentModel().metadata };
}

test("thumbnail metadata signature covers paragraph style tables", () => {
  const baseline = blankMetadata();
  const baselineSignature = core.docModelThumbnailMetadataSignature(baseline);
  assert.equal(
    core.docModelThumbnailMetadataSignature({ ...baseline }),
    baselineSignature,
    "an identical metadata clone must keep the same signature",
  );

  const normalStyle = baseline.paragraphStyles.find((style) => style.id === "Normal");
  assert.ok(normalStyle, "blank model must define the Normal style");
  const restyled = {
    ...baseline,
    paragraphStyles: baseline.paragraphStyles.map((style) =>
      style.id === "Normal"
        ? {
            ...style,
            runStyle: { ...style.runStyle, fontSizePt: 18, bold: true },
          }
        : style,
    ),
  };
  assert.notEqual(
    core.docModelThumbnailMetadataSignature(restyled),
    baselineSignature,
    "editing the Normal style must invalidate page thumbnails",
  );

  const defaultStyleChanged = {
    ...baseline,
    defaultParagraphStyleId: "Heading1",
  };
  assert.notEqual(
    core.docModelThumbnailMetadataSignature(defaultStyleChanged),
    baselineSignature,
    "changing the default paragraph style id must invalidate page thumbnails",
  );
});

test("thumbnail metadata signature ignores non-visual metadata fields", () => {
  const baseline = blankMetadata();
  const changed = {
    ...baseline,
    sourceParts: baseline.sourceParts + 7,
    warnings: [...baseline.warnings, "warning text"],
  };
  assert.equal(
    core.docModelThumbnailMetadataSignature(changed),
    core.docModelThumbnailMetadataSignature(baseline),
  );
});

test("node signature distinguishes same-sized binary payloads", () => {
  const first = new Uint8Array(2048).fill(7);
  const second = new Uint8Array(2048).fill(7);
  // Flip a byte in the middle: the old byte-length-only hash collided here.
  second[1024] = 9;
  const firstSignature = core.docNodeContentSignature({ type: "image", data: first });
  const secondSignature = core.docNodeContentSignature({ type: "image", data: second });
  assert.notEqual(firstSignature, secondSignature);

  const alsoDifferent = new Uint8Array(2048).fill(7);
  alsoDifferent[1] = 8;
  assert.notEqual(
    core.docNodeContentSignature({ type: "image", data: alsoDifferent }),
    firstSignature,
  );

  const sameClone = new Uint8Array(first);
  assert.equal(
    core.docNodeContentSignature({ type: "image", data: sameClone }),
    firstSignature,
  );
});

test("node signature hashes media data URIs in full", () => {
  const prefix = "data:image/png;base64,";
  const head = "A".repeat(400);
  const tail = "Z".repeat(2000);
  // The difference sits outside the old head/middle/tail samples.
  const first = `${prefix}${head}B${"D".repeat(2000)}${tail}`;
  const second = `${prefix}${head}C${"D".repeat(2000)}${tail}`;
  assert.equal(first.length, second.length);
  assert.notEqual(
    core.docNodeContentSignature({ type: "image", src: first }),
    core.docNodeContentSignature({ type: "image", src: second }),
  );
});

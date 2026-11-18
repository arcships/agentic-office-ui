import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromVueDocx = createRequire(
  new URL("../../packages/vue-docx/package.json", import.meta.url),
);
const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const { strFromU8, unzipSync } = await import(
  pathToFileURL(requireFromDocxCore.resolve("fflate")).href
);
const core = await import(
  pathToFileURL(requireFromVueDocx.resolve("@arcships/docx-core")).href
);

const FIXTURE = new URL(
  "../../test-data/corpus/docx-svg-pictures.docx",
  import.meta.url,
);

function arrayBufferFromFixture() {
  const bytes = readFileSync(FIXTURE);
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  );
}

function collectImages(model) {
  const images = [];
  const visit = (nodes) => {
    for (const node of nodes) {
      if (node.type === "paragraph") {
        for (const child of node.children) {
          if (child.type === "image") images.push(child);
        }
      } else if (node.type === "table") {
        for (const row of node.rows) {
          for (const cell of row.cells) visit(cell.nodes);
        }
      }
    }
  };
  visit(model.nodes);
  return images;
}

function imageByAlt(images, alt) {
  const image = images.find((entry) => entry.alt === alt);
  assert.ok(image, `fixture must carry an image named ${alt}`);
  return image;
}

function decodeSvgDataUri(source) {
  assert.match(source, /^data:image\/svg\+xml;charset=utf-8,/);
  return decodeURIComponent(source.slice(source.indexOf(",") + 1));
}

function svgDataUri(markup) {
  return `data:image/svg+xml;base64,${Buffer.from(markup, "utf8").toString("base64")}`;
}

const imported = await core.wasmBuildDocModelFromBytes(arrayBufferFromFixture());
const fixturePackage = core.wasmPackageToMaps(imported.package);
const fixtureImages = collectImages(imported.model);
const fixtureArchive = unzipSync(new Uint8Array(arrayBufferFromFixture()));

test("DOCX SVG pictures are detected from relationships and content types", () => {
  assert.equal(fixtureImages.length, 6);
  const icon = imageByAlt(fixtureImages, "SVG icon");
  const chart = imageByAlt(fixtureImages, "SVG chart");
  const floating = imageByAlt(fixtureImages, "Floating SVG");
  const dual = imageByAlt(fixtureImages, "Dual SVG");
  const foreign = imageByAlt(fixtureImages, "Foreign object SVG");
  const filtered = imageByAlt(fixtureImages, "Filtered SVG");

  for (const image of fixtureImages) {
    assert.equal(core.isSvgPicture(image), true, image.alt);
    assert.equal(image.contentType, "image/svg+xml", image.alt);
    assert.match(image.partName, /^word\/media\/.+\.svg$/u, image.alt);
    assert.match(image.src, /^data:image\/svg\+xml/u, image.alt);
  }

  // The anchored size comes from `wp:extent`, not from the SVG's own width;
  // both are equal here so a rasterised fallback could not hide a mismatch.
  assert.equal(icon.widthPx, 96);
  assert.equal(icon.heightPx, 96);
  assert.equal(chart.widthPx, 320);
  assert.equal(chart.heightPx, 200);
  assert.equal(floating.widthPx, 200);
  assert.equal(floating.heightPx, 120);
  assert.equal(floating.floating?.wrapType, "square");

  // Word's dual blip stores a raster fallback plus an `asvg:svgBlip`; the SVG
  // picture must be the one the model exposes.
  assert.equal(dual.partName, "word/media/dual.svg");
  assert.ok(fixturePackage.binaryAssets.has("word/media/fallback.png"));
  assert.ok(fixturePackage.binaryAssets.has("word/media/dual.svg"));

  // The fixture declares the content type the way Word does.
  const contentTypes = strFromU8(fixtureArchive["[Content_Types].xml"]);
  assert.match(contentTypes, /<Default Extension="svg" ContentType="image\/svg\+xml"\/>/u);
  const relationships = strFromU8(fixtureArchive["word/_rels/document.xml.rels"]);
  assert.match(relationships, /Target="media\/icon\.svg"/u);

  for (const image of [icon, chart, floating, dual]) {
    const inspection = core.inspectSvgPicture(image);
    assert.equal(inspection?.supported, true, image.alt);
    assert.equal(inspection?.passthrough, false, image.alt);
    assert.equal(core.svgPictureUnsupportedDiagnostic(image), undefined);
    assert.equal(core.imageUsesPlaceholderFallback(image), false);
  }
});

test("DOCX SVG pictures render as sanitised vector sources at the anchored size", () => {
  const icon = imageByAlt(fixtureImages, "SVG icon");
  const source = core.resolveRenderableImageSource(icon);
  assert.ok(source);
  assert.match(source, /^data:image\/svg\+xml;charset=utf-8,/u);
  assert.doesNotMatch(source, /^data:image\/png/u, "the SVG must not be rasterised");

  const markup = decodeSvgDataUri(source);
  assert.match(markup, /^<svg[\s>]/u);
  assert.match(markup, /url\(#iconFill\)/u);
  assert.match(markup, /<circle/u);
  assert.match(markup, /<path/u);
  // Script, event handlers and external references never reach the DOM.
  assert.doesNotMatch(markup, /<script/iu);
  assert.doesNotMatch(markup, /onload/iu);
  assert.doesNotMatch(markup, /javascript\s*:/iu);
  assert.doesNotMatch(markup, /https?:\/\/(?!www\.w3\.org\/)/u);
  assert.doesNotMatch(markup, /<\?xml/iu);

  const inspection = core.inspectSvgPicture(icon);
  assert.equal(inspection.removedScripts, 1);
  assert.equal(inspection.removedEventHandlers, 1);
  assert.equal(inspection.removedExternalReferences, 2);
  assert.equal(inspection.markup, markup);

  // A chart of bars keeps its vector content too.
  const chartMarkup = decodeSvgDataUri(
    core.resolveRenderableImageSource(imageByAlt(fixtureImages, "SVG chart")),
  );
  assert.match(chartMarkup, /<rect/u);
  assert.match(chartMarkup, /Quarterly chart/u);
});

test("DOCX SVG sanitizer strips scripts and external references but keeps paint", () => {
  const hostile = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE svg PUBLIC "-//W3C//DTD SVG 1.1//EN" "http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd" [
  <!ENTITY xxe SYSTEM "file:///etc/passwd">
]>
<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"
     width="10" height="10" onload="alert(1)">
  <style>@import url("https://example.invalid/theme.css"); .a { fill: url(#keep); }</style>
  <defs><linearGradient id="keep"><stop offset="0" stop-color="#fff"/></linearGradient></defs>
  <script type="text/javascript">alert("script")</script>
  <rect width="10" height="10" fill="url(#keep)" onclick="alert(2)"/>
  <image xlink:href="https://example.invalid/tracker.png" width="4" height="4"/>
  <image href="#keep" width="4" height="4"/>
  <image href="data:image/png;base64,AAAA" width="4" height="4"/>
  <a href="javascript:alert(3)"><text x="1" y="9">link</text></a>
  <rect width="4" height="4" style="fill:url(https://example.invalid/paint.svg)"/>
</svg>`;

  const sanitized = core.sanitizeSvgPictureMarkup(hostile);
  assert.match(sanitized.markup, /^<svg[\s>]/u);
  assert.doesNotMatch(sanitized.markup, /<!DOCTYPE/iu);
  assert.doesNotMatch(sanitized.markup, /<!ENTITY/iu);
  assert.doesNotMatch(sanitized.markup, /<script/iu);
  assert.doesNotMatch(sanitized.markup, /onload/iu);
  assert.doesNotMatch(sanitized.markup, /onclick/iu);
  assert.doesNotMatch(sanitized.markup, /javascript\s*:/iu);
  assert.doesNotMatch(sanitized.markup, /https?:\/\/(?!www\.w3\.org\/)/u);
  assert.doesNotMatch(sanitized.markup, /@import/iu);
  // Internal paint and embedded rasters survive.
  assert.match(sanitized.markup, /url\(#keep\)/u);
  assert.match(sanitized.markup, /href="#keep"/u);
  assert.match(sanitized.markup, /href="data:image\/png;base64,AAAA"/u);
  assert.match(sanitized.markup, /<linearGradient/u);
  assert.equal(sanitized.removedScripts, 1);
  assert.equal(sanitized.removedEventHandlers, 2);
  assert.ok(
    sanitized.removedExternalReferences >= 4,
    `expected every external reference to be counted, got ${sanitized.removedExternalReferences}`,
  );

  // The produced data URI decodes back to exactly the sanitised markup.
  assert.equal(
    decodeSvgDataUri(core.svgPictureDataUri(sanitized.markup)),
    sanitized.markup,
  );
});

test("DOCX SVG pictures with foreignObject or filters fall back with a diagnostic", () => {
  const foreign = imageByAlt(fixtureImages, "Foreign object SVG");
  const filtered = imageByAlt(fixtureImages, "Filtered SVG");

  for (const [image, reason] of [
    [foreign, "foreignObject"],
    [filtered, "filter"],
  ]) {
    assert.equal(core.resolveRenderableImageSource(image), undefined);
    assert.equal(core.imageUsesPlaceholderFallback(image), true);
    assert.equal(
      core.unsupportedImageFallbackLabel(image, image.widthPx, image.heightPx),
      "SVG",
    );
    const inspection = core.inspectSvgPicture(image);
    assert.equal(inspection?.supported, false);
    assert.equal(inspection?.reason, reason);
    assert.equal(inspection?.src, undefined);
    assert.deepEqual(core.svgPictureUnsupportedDiagnostic(image), {
      code: "IMAGE_UNSUPPORTED_FEATURE",
      phase: "image-render",
      reason,
      partName: image.partName,
    });
  }

  // Small icons keep the compact label used by EMF/TIFF placeholders.
  assert.equal(
    core.unsupportedImageFallbackLabel(foreign, 32, 32),
    "s",
  );

  const invalid = {
    type: "image",
    contentType: "image/svg+xml",
    partName: "word/media/invalid.svg",
    src: svgDataUri("<html><body>not svg</body></html>"),
    widthPx: 64,
    heightPx: 64,
  };
  assert.equal(core.isSvgPicture(invalid), true);
  assert.equal(core.inspectSvgPicture(invalid).reason, "invalid-markup");
  assert.equal(core.resolveRenderableImageSource(invalid), undefined);
  assert.equal(
    core.svgPictureUnsupportedDiagnostic(invalid).reason,
    "invalid-markup",
  );

  // A markup string over the bound is refused instead of being sanitised.
  const tooLarge = `<svg>${"x".repeat(core.MAX_SVG_PICTURE_MARKUP_CHARS)}</svg>`;
  const oversized = {
    type: "image",
    contentType: "image/svg+xml",
    src: svgDataUri(tooLarge),
    data: new TextEncoder().encode(tooLarge),
    widthPx: 64,
    heightPx: 64,
  };
  assert.equal(core.inspectSvgPicture(oversized).reason, "too-large");
  assert.equal(core.resolveRenderableImageSource(oversized), undefined);
});

test("DOCX SVG detection falls back to part names and bytes, never to a raster type", () => {
  const markup = '<svg xmlns="http://www.w3.org/2000/svg"><rect width="4" height="4"/></svg>';
  const bytes = new TextEncoder().encode(markup);

  assert.equal(core.isSvgPicture({ partName: "word/media/diagram.SVG" }), true);
  assert.equal(core.isSvgPicture({ data: bytes }), true);
  assert.equal(
    core.isSvgPicture({ data: new TextEncoder().encode(`\ufeff  <?xml version="1.0"?>${markup}`) }),
    true,
  );

  // An explicit raster content type wins over a coincidental byte prefix.
  assert.equal(
    core.isSvgPicture({ contentType: "image/png", data: bytes }),
    false,
  );
  assert.equal(
    core.isSvgPicture({
      contentType: "image/png",
      partName: "word/media/picture.png",
      src: "data:image/png;base64,AAAA",
    }),
    false,
  );

  // Non-SVG images pass through the image pipeline untouched.
  const png = {
    type: "image",
    contentType: "image/png",
    partName: "word/media/picture.png",
    src: "data:image/png;base64,AAAA",
    widthPx: 10,
    heightPx: 10,
  };
  assert.equal(core.isSvgPicture(png), false);
  assert.equal(core.inspectSvgPicture(png), undefined);
  assert.equal(core.resolveRenderableImageSource(png), png.src);
  assert.equal(core.imageUsesPlaceholderFallback(png), false);

  // A linked SVG has no inline markup to sanitise and keeps its URL.
  const linked = {
    type: "image",
    contentType: "image/svg+xml",
    src: "https://example.invalid/logo.svg",
    widthPx: 40,
    heightPx: 40,
  };
  const inspection = core.inspectSvgPicture(linked);
  assert.equal(inspection.supported, true);
  assert.equal(inspection.passthrough, true);
  assert.equal(core.resolveRenderableImageSource(linked), linked.src);
});

test("DOCX SVG inspection is cached across clones that share media bytes", () => {
  const markup = '<svg xmlns="http://www.w3.org/2000/svg"><circle r="4"/></svg>';
  const data = new TextEncoder().encode(markup);
  const first = {
    type: "image",
    contentType: "image/svg+xml",
    src: svgDataUri(markup),
    data,
  };
  const second = { ...first };
  assert.equal(
    core.inspectSvgPicture(first),
    core.inspectSvgPicture(first),
    "the same picture object must reuse its inspection",
  );
  assert.equal(
    core.inspectSvgPicture(first),
    core.inspectSvgPicture(second),
    "cloned models share the media bytes and must share the inspection",
  );

  const srcOnly = { ...first, data: undefined };
  assert.notEqual(
    core.inspectSvgPicture(srcOnly),
    core.inspectSvgPicture(first),
    "a source-only picture is a different cache entry",
  );
  assert.equal(
    core.inspectSvgPicture(srcOnly).markup,
    core.inspectSvgPicture(first).markup,
  );
  assert.equal(
    core.inspectSvgPicture(srcOnly),
    core.inspectSvgPicture({ ...srcOnly }),
    "identical sources must reuse the cached inspection",
  );
});

test("DOCX SVG pictures survive edit, export and reload byte-for-byte", async () => {
  const mediaParts = [...fixturePackage.binaryAssets].filter(
    ([name]) => name.includes("/media/"),
  );
  assert.ok(mediaParts.length >= 7);
  const svgParts = mediaParts.filter(([name]) => name.endsWith(".svg"));
  assert.equal(svgParts.length, 6);

  const edited = core.updateParagraphText(
    imported.model,
    0,
    "SVG roundtrip paragraph",
  );
  const exported = await core.serializeDocx(edited, fixturePackage);
  const archive = unzipSync(new Uint8Array(exported));

  const exportedContentTypes = strFromU8(archive["[Content_Types].xml"]);
  assert.match(
    exportedContentTypes,
    /<Default Extension="svg" ContentType="image\/svg\+xml"\/>/u,
  );

  for (const [name, bytes] of mediaParts) {
    const reimportedBytes = archive[name];
    assert.ok(reimportedBytes, `${name} must survive export`);
    assert.deepEqual(
      [...reimportedBytes],
      [...bytes],
      `${name} bytes must survive export unchanged`,
    );
  }

  const second = await core.wasmBuildDocModelFromBytes(exported);
  const secondPackage = core.wasmPackageToMaps(second.package);
  for (const [name, bytes] of mediaParts) {
    assert.deepEqual(
      [...(secondPackage.binaryAssets.get(name) ?? [])],
      [...bytes],
      `${name} must survive reimport`,
    );
  }

  const secondImages = collectImages(second.model);
  assert.equal(secondImages.length, fixtureImages.length);
  assert.match(
    [...second.model.nodes]
      .flatMap((node) => (node.type === "paragraph" ? node.children : []))
      .filter((child) => child.type === "text")
      .map((child) => child.text)
      .join(" "),
    /SVG roundtrip paragraph/u,
  );

  for (const original of fixtureImages) {
    const reloaded = imageByAlt(secondImages, original.alt);
    assert.equal(reloaded.contentType, "image/svg+xml", original.alt);
    assert.equal(reloaded.partName, original.partName, original.alt);
    const source = core.resolveRenderableImageSource(reloaded);
    if (core.inspectSvgPicture(original).supported) {
      assert.match(source, /^data:image\/svg\+xml/u, original.alt);
      assert.equal(
        decodeSvgDataUri(source),
        core.inspectSvgPicture(original).markup,
        `${original.alt} must sanitise to the same markup after reload`,
      );
    } else {
      assert.equal(source, undefined, original.alt);
      assert.equal(
        core.inspectSvgPicture(reloaded).reason,
        core.inspectSvgPicture(original).reason,
        original.alt,
      );
    }
  }
});

test("DOCX editor-inserted SVG pictures export as SVG media parts", async () => {
  const markup =
    '<svg xmlns="http://www.w3.org/2000/svg" width="40" height="40">' +
    '<rect width="40" height="40" fill="#0f0"/></svg>';
  const data = new TextEncoder().encode(markup);
  const base = core.createBlankDocumentModel();
  const model = {
    ...base,
    nodes: [
      ...base.nodes,
      {
        type: "paragraph",
        children: [
          {
            type: "image",
            src: `data:image/svg+xml;base64,${Buffer.from(markup, "utf8").toString("base64")}`,
            contentType: "image/svg+xml",
            data,
            alt: "inserted.svg",
            widthPx: 40,
            heightPx: 40,
          },
        ],
      },
    ],
  };

  const exported = await core.serializeDocx(model);
  const archive = unzipSync(new Uint8Array(exported));
  const mediaParts = Object.keys(archive).filter((name) =>
    name.startsWith("word/media/"),
  );
  assert.equal(mediaParts.length, 1);
  assert.match(mediaParts[0], /\.svg$/u);
  assert.deepEqual([...archive[mediaParts[0]]], [...data]);
  assert.match(
    strFromU8(archive["[Content_Types].xml"]),
    /<Default Extension="svg" ContentType="image\/svg\+xml"\/>/u,
  );

  const reloaded = await core.wasmBuildDocModelFromBytes(exported);
  const [image] = collectImages(reloaded.model);
  assert.equal(image.contentType, "image/svg+xml");
  assert.equal(core.isSvgPicture(image), true);
  assert.equal(decodeSvgDataUri(core.resolveRenderableImageSource(image)), markup);
});

test("DOCX thumbnails keep SVG sources vector instead of downscaling them", () => {
  const icon = imageByAlt(fixtureImages, "SVG icon");
  const largeSvg = `<svg xmlns="http://www.w3.org/2000/svg">${"<rect/>".repeat(8_000)}</svg>`;
  const largeSvgSource = core.svgPictureDataUri(largeSvg);
  assert.ok(
    largeSvgSource.length > 32_768,
    "the vector source must be large enough to trip the raster downscale threshold",
  );
  assert.equal(
    core.thumbnailImageSourceQualifiesForDownscale(largeSvgSource),
    false,
    "SVG sources must stay vector in thumbnails",
  );
  assert.equal(
    core.thumbnailImageSourceQualifiesForDownscale(
      core.resolveRenderableImageSource(icon),
    ),
    false,
  );
  assert.equal(
    core.thumbnailImageSourceQualifiesForDownscale(
      `data:image/png;base64,${"A".repeat(40_000)}`,
    ),
    true,
    "large raster sources still take the downscale path",
  );
});

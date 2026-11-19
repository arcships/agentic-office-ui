import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { dirname, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";

const rootDir = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const sourceUrl = (relative) => pathToFileURL(resolve(rootDir, relative)).href;
const resolveHookUrl = `data:text/javascript,${encodeURIComponent(`
  export async function resolve(specifier, context, nextResolve) {
    try {
      return await nextResolve(specifier, context);
    } catch (error) {
      if (specifier.startsWith(".") && !specifier.endsWith(".ts")) {
        return nextResolve(specifier + ".ts", context);
      }
      throw error;
    }
  }
`)}`;
const registerHookUrl = `data:text/javascript,${encodeURIComponent(`
  import { register } from "node:module";
  register(${JSON.stringify(resolveHookUrl)}, import.meta.url);
`)}`;

function runNode(args) {
  return spawnSync(process.execPath, args, {
    cwd: rootDir,
    encoding: "utf8",
    env: { ...process.env, NODE_NO_WARNINGS: "1" },
  });
}

// The parser only needs a DOMParser plus slide XML and relationship maps, so the
// test drives it directly instead of building a real deck archive.
const childScript = `
  import assert from "node:assert/strict";
  import { createRequire } from "node:module";

  const requireFromXlsxCore = createRequire(${JSON.stringify(resolve(rootDir, "packages/xlsx-core/package.json"))});
  const { DOMParser: XmlDomParser, Node: XmlNode } = requireFromXlsxCore("@xmldom/xmldom");

  class TestDOMParser extends XmlDomParser {
    parseFromString(xml, type) {
      const document = super.parseFromString(xml, type);
      // xmldom has no querySelector; the parser only uses it for the
      // browser-specific <parsererror> probe.
      Object.defineProperty(document, "querySelector", {
        value: (selector) => selector === "parsererror"
          ? document.getElementsByTagNameNS("*", "parsererror")[0] ?? null
          : null,
      });
      return document;
    }
  }
  globalThis.DOMParser ??= TestDOMParser;
  globalThis.Node ??= XmlNode;

  const { parsePptxPlaybackDocument } = await import(${JSON.stringify(sourceUrl("packages/pptx-core/src/browser/playback-parser.ts"))});
  const { applyPptxObjectMarkers } = await import(${JSON.stringify(sourceUrl("packages/pptx-core/src/browser/renderer-markers.ts"))});

  const NS = [
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"',
  ].join(" ");

  function shape(id, link, text) {
    return '<p:sp><p:nvSpPr>'
      + '<p:cNvPr id="' + id + '" name="Shape ' + id + '"/>'
      + '<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>'
      + '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US">' + link + '</a:rPr>'
      + '<a:t>' + text + '</a:t></a:r></a:p></p:txBody></p:sp>';
  }

  function slideXml(shapes) {
    return '<?xml version="1.0" encoding="UTF-8"?><p:sld ' + NS + '><p:cSld><p:spTree>'
      + shapes.join("") + '</p:spTree></p:cSld></p:sld>';
  }

  const slideOne = slideXml([
    shape(2, '<a:hlinkClick r:id="rIdDocs" tooltip="Read the docs"/>', "Docs"),
    shape(3, '<a:hlinkClick r:id="rIdSlide" action="ppaction://hlinksldjump" tooltip="Jump to next"/>', "Next slide"),
    shape(4, '<a:hlinkClick r:id="rIdScript"/>', "Script"),
    shape(5, '<a:hlinkHover r:id="rIdHover" tooltip="Hover docs"/>', "Hover"),
    shape(6, "", "Plain"),
    shape(7, '<a:hlinkClick r:id="rIdMail"/>', "Mail"),
  ]);
  const slideTwo = slideXml([]);

  const rels = (entries) => new Map(entries.map(([id, target, targetMode]) => [
    id,
    { id, target, ...(targetMode ? { targetMode } : {}) },
  ]));
  const files = {
    slides: new Map([
      ["ppt/slides/slide1.xml", slideOne],
      ["ppt/slides/slide2.xml", slideTwo],
    ]),
  };
  const presentation = {
    width: 960,
    height: 540,
    slides: [
      {
        index: 0,
        slidePath: "ppt/slides/slide1.xml",
        hidden: false,
        rels: rels([
          ["rIdDocs", "https://example.com/docs", "External"],
          ["rIdSlide", "slide2.xml"],
          ["rIdScript", "javascript:alert(1)", "External"],
          ["rIdHover", "https://example.com/hover", "External"],
          ["rIdMail", "mailto:team@example.com", "External"],
        ]),
      },
      { index: 1, slidePath: "ppt/slides/slide2.xml", hidden: false, rels: rels([]) },
    ],
  };

  const parsed = parsePptxPlaybackDocument(files, presentation, {});
  const [first] = parsed.slides;
  const keyOf = (id) => 'ppt/slides/slide1.xml|slide|shape:' + id;
  const actionFor = (id) => first.actions.find((action) => action.sourceObjectKey === keyOf(id));
  const objectFor = (id) => first.objects.find((object) => object.shapeId === String(id));
  const featureFor = (feature, id) => first.capability.features.find(
    (record) => record.feature === feature && record.objectKey === keyOf(id),
  );

  assert.equal(first.actions.length, 5);
  assert.deepEqual(actionFor(2), {
    id: "action:0:0",
    sourceObjectKey: keyOf(2),
    trigger: "click",
    kind: "open-url",
    targetSlideIndex: undefined,
    url: "https://example.com/docs",
    tooltip: "Read the docs",
    rawAction: undefined,
  });
  assert.equal(actionFor(3).kind, "go-to-slide");
  assert.equal(actionFor(3).targetSlideIndex, 1);
  assert.equal(actionFor(3).tooltip, "Jump to next");
  assert.equal(actionFor(4).kind, "unsupported");
  assert.equal(actionFor(4).url, "javascript:alert(1)");
  assert.equal(actionFor(5).trigger, "hover");
  assert.equal(actionFor(5).kind, "open-url");
  assert.equal(actionFor(5).tooltip, "Hover docs");
  assert.equal(actionFor(7).kind, "mailto");
  assert.equal(actionFor(6), undefined);

  assert.deepEqual(objectFor(2).hyperlinks, [{
    trigger: "click",
    kind: "open-url",
    url: "https://example.com/docs",
    tooltip: "Read the docs",
  }]);
  assert.deepEqual(objectFor(3).hyperlinks, [{
    trigger: "click",
    kind: "go-to-slide",
    targetSlideIndex: 1,
    tooltip: "Jump to next",
  }]);
  assert.deepEqual(objectFor(5).hyperlinks, [{
    trigger: "hover",
    kind: "open-url",
    url: "https://example.com/hover",
    tooltip: "Hover docs",
  }]);
  assert.equal(objectFor(6).hyperlinks, undefined);
  assert.equal(Object.isFrozen(objectFor(2).hyperlinks), true);

  const hoverFeature = featureFor("action:open-url", 5);
  assert.equal(hoverFeature.disposition, "static");
  assert.match(hoverFeature.reason, /悬停/u);
  assert.equal(featureFor("action:unsupported", 4).disposition, "static");
  assert.equal(featureFor("action:go-to-slide", 3).disposition, "strict");
  assert.equal(featureFor("action:mailto", 7).disposition, "strict");

  function fakeElement(nodeId, text) {
    const attributes = new Map();
    return {
      attributes,
      dataset: { pptxNodeId: nodeId, pptxPartPath: "ppt/slides/slide1.xml" },
      parentElement: null,
      style: {},
      textContent: text,
      getAttribute(name) { return attributes.get(name) ?? null; },
      setAttribute(name, value) { attributes.set(name, String(value)); },
      hasAttribute(name) { return attributes.has(name); },
    };
  }

  const docs = fakeElement("2", "Docs");
  const hover = fakeElement("5", "Hover");
  const mail = fakeElement("7", "Mail");
  const plain = fakeElement("6", "Plain");
  const titled = fakeElement("2", "Docs");
  titled.setAttribute("title", "Host tooltip");
  const thumbnail = fakeElement("2", "Docs");
  thumbnail.parentElement = { dataset: { pptxThumbnail: "true" } };
  const root = {
    dataset: {},
    querySelectorAll(selector) {
      if (selector === "[data-pptx-node-id]") return [docs, hover, mail, plain, titled, thumbnail];
      if (selector === "[data-pptx-paragraph-index]") return [];
      return [];
    },
  };

  applyPptxObjectMarkers(root, "ppt/slides/slide1.xml", first.objects);

  assert.equal(docs.dataset.pptxHyperlink, "click");
  assert.equal(docs.dataset.pptxHyperlinkKind, "open-url");
  assert.equal(docs.dataset.pptxHyperlinkUrl, "https://example.com/docs");
  assert.equal(docs.dataset.pptxHyperlinkTooltip, "Read the docs");
  assert.equal(docs.getAttribute("title"), "Read the docs");
  assert.equal(docs.getAttribute("role"), "link");
  assert.equal(docs.getAttribute("tabindex"), "0");

  assert.equal(hover.dataset.pptxHyperlink, "hover");
  assert.equal(hover.getAttribute("title"), "Hover docs");
  assert.equal(hover.getAttribute("role"), null, "hover links are not click targets");
  assert.equal(hover.hasAttribute("tabindex"), false);

  // No tooltip: PowerPoint falls back to the target itself.
  assert.equal(mail.getAttribute("title"), "mailto:team@example.com");

  assert.equal(plain.dataset.pptxHyperlink, undefined);
  assert.equal(plain.getAttribute("title"), null);

  // A title written by the renderer or the host wins over the fallback.
  assert.equal(titled.getAttribute("title"), "Host tooltip");

  // Thumbnails live inside a <button>: keep the tooltip, drop the tab stop.
  assert.equal(thumbnail.getAttribute("title"), "Read the docs");
  assert.equal(thumbnail.dataset.pptxHyperlink, "click");
  assert.equal(thumbnail.getAttribute("role"), null);
  assert.equal(thumbnail.hasAttribute("tabindex"), false);

  process.stdout.write(JSON.stringify({
    status: "PASS",
    actions: first.actions.length,
    objects: first.objects.length,
    static: first.capability.static,
  }));
`;

test("PPTX hyperlinks keep tooltips, hover state and object metadata", () => {
  const result = runNode([
    "--experimental-strip-types",
    "--import",
    registerHookUrl,
    "--input-type=module",
    "--eval",
    childScript,
  ]);
  assert.equal(result.status, 0, result.stderr || result.stdout);
  const evidence = JSON.parse(result.stdout);
  assert.equal(evidence.status, "PASS");
  assert.equal(evidence.actions, 5);
  assert.equal(evidence.objects, 6);
  assert.ok(evidence.static >= 2);
});

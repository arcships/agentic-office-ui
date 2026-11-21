import assert from "node:assert/strict";
import { createRequire } from "node:module";
import test from "node:test";

import { importFromDemo } from "../component/vue-test-renderer.mjs";

const requireFromPptxCore = createRequire(
  new URL("../../packages/pptx-core/package.json", import.meta.url),
);
// The PPTX renderer is browser-only. This small DOM stand-in covers the empty
// slides used below; xmldom already ships with the xlsx core package.
const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const { strToU8, zipSync } = requireFromPptxCore("fflate");
const { DOMParser, Document, XMLSerializer } = requireFromXlsxCore("@xmldom/xmldom");

Document.prototype.querySelector = function querySelector(selector) {
  return selector === "parsererror" ? null : undefined;
};

function matchesSelector(element, selector) {
  if (selector.startsWith(".")) return element.classList.contains(selector.slice(1));
  if (selector.startsWith("[")) {
    const match = /^\[([^=\]]+)(?:="?([^"\]]*)"?)?\]$/.exec(selector);
    if (!match) return false;
    const [, name, value] = match;
    const actual = name.startsWith("data-")
      ? element.dataset[name.slice(5).replace(/-([a-z])/g, (_, letter) => letter.toUpperCase())]
      : element.attributes[name];
    return value === undefined ? actual !== undefined : actual === value;
  }
  return element.localName.toLowerCase() === selector.toLowerCase();
}

class FakeElement {
  constructor(localName) {
    this.localName = localName;
    this.tagName = localName.toUpperCase();
    this.nodeType = 1;
    this.children = [];
    this.childNodes = this.children;
    this.parentElement = null;
    this.parentNode = null;
    this.dataset = {};
    this.attributes = {};
    this.textContent = "";
    this.innerHTMLValue = "";
    this.style = {
      cssText: "",
      setProperty(name, value) {
        this[name] = String(value);
      },
      removeProperty(name) {
        delete this[name];
      },
    };
    this.classList = {
      add() {},
      contains() {
        return false;
      },
      remove() {},
      toggle() {},
    };
    this.ownerDocument = globalThis.document;
  }

  get firstElementChild() {
    return this.children[0] ?? null;
  }

  get lastElementChild() {
    return this.children.at(-1) ?? null;
  }

  get innerHTML() {
    return this.innerHTMLValue;
  }

  set innerHTML(value) {
    this.innerHTMLValue = value;
    if (value === "") this.children = [];
    this.childNodes = this.children;
  }

  appendChild(child) {
    if (child.isFragment) {
      for (const item of [...child.children]) this.appendChild(item);
      return child;
    }
    child.parentElement = this;
    child.parentNode = this;
    this.children.push(child);
    return child;
  }

  insertBefore(child, reference) {
    const index = this.children.indexOf(reference);
    if (index < 0) this.children.push(child);
    else this.children.splice(index, 0, child);
    child.parentElement = this;
    child.parentNode = this;
    return child;
  }

  removeChild(child) {
    const index = this.children.indexOf(child);
    if (index >= 0) this.children.splice(index, 1);
    child.parentElement = null;
    child.parentNode = null;
    return child;
  }

  remove() {
    this.parentElement?.removeChild(this);
  }

  replaceChildren(...children) {
    this.children = [];
    for (const child of children) this.appendChild(child);
  }

  setAttribute(name, value) {
    this.attributes[name] = String(value);
    if (name.startsWith("data-")) {
      const key = name.slice(5).replace(/-([a-z])/g, (_, letter) => letter.toUpperCase());
      this.dataset[key] = String(value);
    }
  }

  getAttribute(name) {
    return this.attributes[name] ?? null;
  }

  hasAttribute(name) {
    return name in this.attributes;
  }

  removeAttribute(name) {
    delete this.attributes[name];
  }

  matches(selector) {
    return this.querySelectorAll(selector).includes(this);
  }

  closest(selector) {
    let current = this;
    while (current) {
      if (current.matches(selector)) return current;
      current = current.parentElement;
    }
    return null;
  }

  querySelector(selector) {
    return this.querySelectorAll(selector)[0] ?? null;
  }

  querySelectorAll(selector) {
    const result = [];
    const visit = (node) => {
      for (const child of node.children) {
        if (matchesSelector(child, selector)) result.push(child);
        visit(child);
      }
    };
    visit(this);
    return result;
  }

  getBoundingClientRect() {
    return { bottom: 100, height: 100, left: 0, right: 100, top: 0, width: 100, x: 0, y: 0 };
  }

  getContext() {
    return null;
  }

  addEventListener() {}
  removeEventListener() {}
  dispatchEvent() {}
  focus() {}
  scrollIntoView() {}

  cloneNode() {
    return new FakeElement(this.localName);
  }

  get offsetWidth() {
    return 100;
  }

  get offsetHeight() {
    return 100;
  }

  get clientWidth() {
    return 100;
  }

  get clientHeight() {
    return 100;
  }
}

const documentObject = {
  body: new FakeElement("body"),
  createDocumentFragment() {
    const fragment = new FakeElement("#fragment");
    fragment.isFragment = true;
    return fragment;
  },
  createElement(tag) {
    return new FakeElement(tag);
  },
  createElementNS(_namespace, tag) {
    return new FakeElement(tag);
  },
  createTextNode(text) {
    const node = new FakeElement("#text");
    node.textContent = text;
    return node;
  },
  documentElement: new FakeElement("html"),
  querySelector() {
    return null;
  },
};

globalThis.document = documentObject;
globalThis.window = {
  addEventListener() {},
  innerHeight: 720,
  removeEventListener() {},
};
globalThis.getComputedStyle = () => ({
  clipPath: "none",
  display: "block",
  filter: "none",
  opacity: "1",
  position: "static",
  visibility: "visible",
});
globalThis.requestAnimationFrame = (callback) => {
  callback(performance.now());
  return 1;
};
globalThis.cancelAnimationFrame = () => {};
globalThis.HTMLElement = FakeElement;
globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;

const {
  createPptxDocumentSession,
  createPptxPlaybackController,
} = await importFromDemo("@arcships/pptx-core/browser");

const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships";

const emptySlide = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="${NS_P}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>`;
const textSlide = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="${NS_P}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="1000000" y="1000000"/><a:ext cx="2000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN"/><a:t>切换文字</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`;

function deck({ text = false, slideExtras = {}, notes = {} } = {}) {
  const notesContentTypes = Object.keys(notes)
    .map((index) => `<Override PartName="/ppt/notesSlides/notesSlide${Number(index) + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`)
    .join("");
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>${notesContentTypes}</Types>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NS_REL}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`;
  const presentation = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="${NS_P}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;
  const presentationRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NS_REL}"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>`;
  const emptyRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NS_REL}"></Relationships>`;
  const withExtras = (xml, index) => slideExtras[index]
    ? xml.replace("</p:sld>", `${slideExtras[index]}</p:sld>`)
    : xml;
  const slideRels = (index) => notes[index]
    ? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NS_REL}"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide${index + 1}.xml"/></Relationships>`
    : emptyRels;
  const entries = {
    "[Content_Types].xml": strToU8(contentTypes),
    "_rels/.rels": strToU8(rootRels),
    "ppt/presentation.xml": strToU8(presentation),
    "ppt/_rels/presentation.xml.rels": strToU8(presentationRels),
    "ppt/slides/slide1.xml": strToU8(withExtras(emptySlide, 0)),
    "ppt/slides/slide2.xml": strToU8(withExtras(text ? textSlide : emptySlide, 1)),
    "ppt/slides/_rels/slide1.xml.rels": strToU8(slideRels(0)),
    "ppt/slides/_rels/slide2.xml.rels": strToU8(slideRels(1)),
  };
  for (const [index, paragraphs] of Object.entries(notes)) {
    entries[`ppt/notesSlides/notesSlide${Number(index) + 1}.xml`] = strToU8(notesSlide(paragraphs));
  }
  return zipSync(entries);
}

function notesSlide(paragraphs) {
  const body = paragraphs
    .map((paragraph) => `<a:p><a:r><a:rPr lang="en-US"/><a:t>${paragraph}</a:t></a:r></a:p>`)
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:p="${NS_P}" xmlns:a="${NS_A}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${body}</p:txBody></p:sp></p:spTree></p:cSld></p:notes>`;
}

function createRoot({ failOnRender = false } = {}) {
  const root = new FakeElement("div");
  if (failOnRender) {
    let failed = false;
    Object.defineProperty(root.style, "position", {
      get() {
        return this.positionValue ?? "";
      },
      set(value) {
        if (!failed) {
          failed = true;
          throw new Error("render container rejected the slide");
        }
        this.positionValue = value;
      },
    });
  }
  return root;
}

test("a failed open keeps the previous document and playback model", async () => {
  const session = createPptxDocumentSession(createRoot(), {
    lazyMedia: true,
    lazySlides: true,
    renderMode: "slide",
  });
  try {
    const first = await session.open(deck(), { initialSlide: 0 });
    const playback = session.playbackDocument;
    assert.equal(first.slides.length, 2);
    assert.ok(playback);

    await assert.rejects(
      session.open(new Uint8Array([1, 2, 3])),
      (error) => error.code === "INVALID_SOURCE",
    );

    assert.equal(session.document, first);
    assert.equal(session.playbackDocument, playback);
  } finally {
    session.dispose();
  }
});

test("password-protected OLE/CFB containers report UNSUPPORTED_CONTENT", async () => {
  const session = createPptxDocumentSession(createRoot(), {
    lazyMedia: true,
    lazySlides: true,
    renderMode: "slide",
  });
  try {
    // PowerPoint writes password-protected decks (and legacy .ppt renamed to
    // .pptx) as an OLE Compound File, not an OOXML zip. The core detects the
    // signature and reports the dedicated content code the viewer maps to its
    // "protected presentation" state.
    const ole = new Uint8Array(64);
    ole.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    await assert.rejects(
      session.open(ole),
      (error) => error.code === "UNSUPPORTED_CONTENT",
    );
    assert.equal(session.document, null);
  } finally {
    session.dispose();
  }
});

test("a render failure clears the session instead of keeping a dead document", async () => {
  const session = createPptxDocumentSession(createRoot({ failOnRender: true }), {
    lazyMedia: true,
    lazySlides: true,
    renderMode: "slide",
  });
  try {
    await assert.rejects(session.open(deck(), { initialSlide: 0 }));
    assert.equal(session.document, null);
    assert.equal(session.playbackDocument, null);

    const recovered = await session.open(deck(), { initialSlide: 0 });
    assert.equal(recovered.slides.length, 2);
    assert.equal(session.document, recovered);
  } finally {
    session.dispose();
  }
});

test("search results from a previous document cannot highlight the next one", async () => {
  const root = createRoot();
  const session = createPptxDocumentSession(root, {
    lazyMedia: true,
    lazySlides: true,
    renderMode: "slide",
  });
  try {
    await session.open(deck({ text: true }), { initialSlide: 0 });
    const matches = session.searchText("切换");
    assert.equal(matches.length, 1);
    const liveHighlight = await session.highlightSearchResult(matches[0]);
    assert.ok(liveHighlight);
    liveHighlight.dispose();

    await session.open(deck(), { initialSlide: 0 });
    const staleHighlight = await session.highlightSearchResult(matches[0]);
    assert.equal(staleHighlight, null);
    assert.equal(root.querySelectorAll("[data-pptx-search-highlight]").length, 0);
  } finally {
    session.dispose();
  }
});

test("paragraph entrance state is not undone by combined range targets", () => {
  const root = new FakeElement("div");
  const shape = new FakeElement("div");
  shape.dataset.pptxObjectKey = "shape:1";
  root.appendChild(shape);
  const paragraphs = [0, 1, 2].map((index) => {
    const paragraph = new FakeElement("div");
    paragraph.dataset.pptxParagraphIndex = String(index);
    shape.appendChild(paragraph);
    return paragraph;
  });
  const nodes = {
    root: {
      acceleration: 0,
      autoReverse: false,
      childIds: ["effect-1"],
      conditions: [],
      container: "root",
      deceleration: 0,
      delayMs: 0,
      durationMs: 0,
      fill: "hold",
      id: "root",
      kind: "group",
      restart: "never",
    },
    "effect-1": {
      acceleration: 0,
      autoReverse: false,
      childIds: [],
      conditions: [{ delayMs: 0, event: "after-previous", source: "start" }],
      container: "behavior",
      deceleration: 0,
      delayMs: 0,
      durationMs: 0,
      effect: {
        id: "animation-1",
        kind: "fade-in",
        paragraphRange: { end: 0, start: 0 },
        presetClass: "entr",
        targetObjectKey: "shape:1",
        values: {},
      },
      fill: "hold",
      id: "effect-1",
      kind: "effect",
      restart: "never",
    },
  };
  const controller = createPptxPlaybackController({
    document: {
      capability: {
        approximate: 0,
        discovered: 0,
        features: [],
        static: 0,
        strict: 0,
        unparsed: 0,
      },
      slides: [{
        actions: [],
        capability: {
          approximate: 0,
          discovered: 0,
          features: [],
          static: 0,
          strict: 0,
          unparsed: 0,
        },
        hidden: false,
        index: 0,
        media: [],
        morphFromPrevious: [],
        nodes,
        objects: [],
        rootNodeId: "root",
      }],
    },
    initialSlideIndex: 0,
    renderSlide: async () => {},
    root,
  }, { autoplay: false });

  try {
    assert.deepEqual(
      paragraphs.map((paragraph) => paragraph.style.visibility),
      ["hidden", "visible", "visible"],
    );
  } finally {
    controller.dispose();
  }
});

test("a fresh session stays empty after a failed open", async () => {
  const session = createPptxDocumentSession(createRoot(), {
    lazyMedia: true,
    lazySlides: true,
    renderMode: "slide",
  });
  try {
    await assert.rejects(session.open(new Uint8Array([1, 2, 3])));
    assert.equal(session.document, null);
    assert.equal(session.playbackDocument, null);
  } finally {
    session.dispose();
  }
});

test("speaker notes are parsed into the preview document", async () => {
  const session = createPptxDocumentSession(createRoot(), {
    lazyMedia: true,
    lazySlides: true,
    renderMode: "slide",
  });
  try {
    const loaded = await session.open(
      deck({ notes: { 0: ["Open with the quarter", "Then pause for the chart."] } }),
      { initialSlide: 0 },
    );
    assert.equal(loaded.slides[0].notes, "Open with the quarter\nThen pause for the chart.");
    assert.equal(loaded.slides[1].notes, undefined, "a slide without a notes part reports no notes");
    assert.deepEqual(loaded.warnings, [], "notes are no longer reported as unsupported");
  } finally {
    session.dispose();
  }
});

test("notes parts that break the XML budget fail the open request", async () => {
  const session = createPptxDocumentSession(createRoot(), {
    lazyMedia: true,
    lazySlides: true,
    renderMode: "slide",
    limits: { maxSingleXmlBytes: 512 },
  });
  try {
    await assert.rejects(
      session.open(deck({ notes: { 0: ["x".repeat(2_000)] } }), { initialSlide: 0 }),
      (error) => error.code === "LIMIT_EXCEEDED",
    );
  } finally {
    session.dispose();
  }
});

test("transitions beyond cut/fade/push/wipe reach the capability report", async () => {
  const transitions = {
    0: '<p:transition spd="fast"><p:cover dir="u"/></p:transition>',
    1: '<p:transition spd="fast"><p:split orient="horz" dir="out"/></p:transition>',
  };
  const session = createPptxDocumentSession(createRoot(), {
    lazyMedia: true,
    lazySlides: true,
    renderMode: "slide",
  });
  try {
    await session.open(deck({ slideExtras: transitions }), { initialSlide: 0 });
    const report = session.capabilityReport;
    const dispositions = report.features
      .filter((feature) => feature.feature.startsWith("transition:"))
      .map((feature) => [feature.feature, feature.disposition]);
    assert.deepEqual(dispositions, [
      ["transition:cover", "strict"],
      ["transition:split", "strict"],
    ]);
  } finally {
    session.dispose();
  }
});

test("dissolve is reported as static by default and approximate with approximation: safe", async () => {
  const dissolve = { 0: "<p:transition spd=\"fast\"><p:dissolve/></p:transition>" };
  const strictSession = createPptxDocumentSession(createRoot(), {
    lazyMedia: true,
    lazySlides: true,
    renderMode: "slide",
  });
  const safeSession = createPptxDocumentSession(createRoot(), {
    lazyMedia: true,
    lazySlides: true,
    renderMode: "slide",
    approximation: "safe",
  });
  const transitionDisposition = (session) => session.capabilityReport.features
    .find((feature) => feature.feature === "transition:dissolve")?.disposition;
  try {
    await strictSession.open(deck({ slideExtras: dissolve }), { initialSlide: 0 });
    assert.equal(transitionDisposition(strictSession), "static");
    await safeSession.open(deck({ slideExtras: dissolve }), { initialSlide: 0 });
    assert.equal(transitionDisposition(safeSession), "approximate");
  } finally {
    strictSession.dispose();
    safeSession.dispose();
  }
});

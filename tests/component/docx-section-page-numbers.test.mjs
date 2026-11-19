// Issue #122 — DOCX section page numbering and parity blank pages (surface).
//
// The model-level tracker is covered by tests/unit/docx-section-page-numbers.
// This suite asserts what a reader actually sees:
//
//   1. every page carries the label Word shows (`1 … n`, then `i, ii, …`),
//      including the parity blank page a `w:type="oddPage"` break inserts;
//   2. the blank page is marked as such instead of looking like a lost page;
//   3. `PAGE`, `NUMPAGES` and `SECTIONPAGES` in the header/footer resolve from
//      the same tracker;
//   4. the thumbnail rail labels every page with those numbers;
//   5. a single-section document keeps plain physical numbering.

import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  textContent,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const { DocxThumbnailPanel, DocxViewerRoot } = await importFromDemo(
  "@arcships/vue-docx",
);
const { cloneDocModel, defaultStarterModel } = await importFromDemo(
  "@arcships/docx-core",
);

// 200 x 300 CSS px section geometry, so the fixture paginates quickly.
const SMALL_GEOMETRY_XML =
  '<w:pgSz w:w="3000" w:h="4500"/>' +
  '<w:pgMar w:top="150" w:right="150" w:bottom="150" w:left="150"' +
  ' w:header="150" w:footer="150" w:gutter="0"/>';
const VIEWER_LAYOUT = {
  pageWidth: 200,
  pageHeight: 300,
  margin: 10,
  minLineHeight: 30,
  paragraphSpacing: 10,
};

function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

/** A paragraph that ends its page, so the next node starts a new one. */
function paragraphEndingPage(text) {
  return {
    type: "paragraph",
    children: [{ type: "text", text }],
    sourceXml: `<w:p><w:r><w:t>${text}</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>`,
  };
}

/**
 * One paragraph holding `PAGE / NUMPAGES / SECTIONPAGES` fields, written the way
 * Word writes them: an `fldSimple` instruction plus the cached result text.
 * The rendered children carry the cached values, so field resolution replaces
 * each of them with the value the page resolves.
 */
function pageFieldParagraph({ page = "1", total = "1", section = "1" } = {}) {
  const field = (instruction, cached) =>
    `<w:fldSimple w:instr=" ${instruction} "><w:r><w:t>${cached}</w:t></w:r></w:fldSimple>`;
  return {
    type: "paragraph",
    children: [
      { type: "text", text: page },
      { type: "text", text: "/" },
      { type: "text", text: total },
      { type: "text", text: "/" },
      { type: "text", text: section },
    ],
    sourceXml:
      "<w:p>" +
      field("PAGE", page) +
      "<w:r><w:t>/</w:t></w:r>" +
      field("NUMPAGES", total) +
      "<w:r><w:t>/</w:t></w:r>" +
      field("SECTIONPAGES", section) +
      "</w:p>",
  };
}

function footerSection(partName, nodes) {
  return { partName, referenceType: "default", nodes };
}

/**
 * The acceptance fixture: section 1 numbers with plain `decimal`, section 2
 * restarts at 1 in `lowerRoman` and must start on an odd page. Section 1 ends
 * on page 1, so page 2 becomes a blank parity page and section 2 starts on
 * page 3 — the blank page keeps section 1's footer and numbering.
 */
function twoSectionModel({ firstSectionPages = 1 } = {}) {
  const model = cloneDocModel(defaultStarterModel);
  const firstSectionNodes =
    firstSectionPages > 1
      ? [paragraphEndingPage("Section one cover"), paragraph("Section one body")]
      : [paragraph("Section one cover")];
  model.nodes = [...firstSectionNodes, paragraph("Section two body")];

  const firstFooter = footerSection("word/footer1.xml", [
    pageFieldParagraph({ page: "1", total: "2", section: "1" }),
  ]);
  const secondFooter = footerSection("word/footer2.xml", [
    pageFieldParagraph({ page: "1", total: "2", section: "1" }),
  ]);

  const firstSectionXml =
    "<w:sectPr>" +
    '<w:footerReference w:type="default" r:id="rId11"/>' +
    '<w:type w:val="nextPage"/>' +
    SMALL_GEOMETRY_XML +
    "</w:sectPr>";
  const secondSectionXml =
    "<w:sectPr>" +
    '<w:footerReference w:type="default" r:id="rId12"/>' +
    '<w:type w:val="oddPage"/>' +
    '<w:pgNumType w:fmt="lowerRoman" w:start="1"/>' +
    SMALL_GEOMETRY_XML +
    "</w:sectPr>";

  const sectionStart = firstSectionNodes.length;
  model.metadata.sections = [
    {
      startNodeIndex: 0,
      sectionPropertiesXml: firstSectionXml,
      headerSections: [],
      footerSections: [firstFooter],
    },
    {
      startNodeIndex: sectionStart,
      sectionPropertiesXml: secondSectionXml,
      headerSections: [],
      footerSections: [secondFooter],
    },
  ];
  model.metadata.sectionPropertiesXml = secondSectionXml;
  model.metadata.footerSections = [];
  return model;
}

function singleSectionModel() {
  const model = cloneDocModel(defaultStarterModel);
  model.nodes = [
    paragraphEndingPage("Page one"),
    paragraph("Page two"),
    paragraph("Page two tail"),
  ];
  model.metadata.sections = [
    {
      startNodeIndex: 0,
      sectionPropertiesXml: "<w:sectPr>" + SMALL_GEOMETRY_XML + "</w:sectPr>",
      headerSections: [],
      footerSections: [],
    },
  ];
  model.metadata.sectionPropertiesXml =
    "<w:sectPr>" + SMALL_GEOMETRY_XML + "</w:sectPr>";
  return model;
}

function classNames(node) {
  return String(node?.props?.class || "").split(/\s+/).filter(Boolean);
}

/** The rail paints into a canvas; the fake renderer has no 2D context. */
function installCanvasEnvironment(t) {
  const originalDataset = Object.getOwnPropertyDescriptor(Object.prototype, "dataset");
  const originalGetContext = Object.getOwnPropertyDescriptor(Object.prototype, "getContext");

  Object.defineProperty(Object.prototype, "dataset", {
    configurable: true,
    get() {
      if (!Object.hasOwn(this, "__thumbnailDataset")) {
        Object.defineProperty(this, "__thumbnailDataset", {
          configurable: true,
          value: {},
        });
      }
      return this.__thumbnailDataset;
    },
  });
  Object.defineProperty(Object.prototype, "getContext", {
    configurable: true,
    value(kind) {
      if (this.type !== "canvas" || kind !== "2d") return null;
      if (Object.hasOwn(this, "__thumbnailContext")) return this.__thumbnailContext;
      const noop = () => {};
      const context = {
        setTransform: noop,
        clearRect: noop,
        fillRect: noop,
        strokeRect: noop,
        save: noop,
        restore: noop,
        scale: noop,
        beginPath: noop,
        rect: noop,
        clip: noop,
        fillText: noop,
        moveTo: noop,
        lineTo: noop,
        stroke: noop,
        drawImage: noop,
        measureText: (value) => ({ width: String(value).length * 8 }),
      };
      Object.defineProperty(this, "__thumbnailContext", {
        configurable: true,
        value: context,
      });
      return context;
    },
  });

  t.after(() => {
    if (originalDataset) Object.defineProperty(Object.prototype, "dataset", originalDataset);
    else delete Object.prototype.dataset;
    if (originalGetContext) Object.defineProperty(Object.prototype, "getContext", originalGetContext);
    else delete Object.prototype.getContext;
  });
}

function nodesWithClass(root, className) {
  return walk(root).filter((node) => classNames(node).includes(className));
}

function nodesWithAttribute(root, attribute) {
  return walk(root).filter((node) => node?.props?.[attribute] !== undefined);
}

/** Template-declared props reach the vnode as the kebab-case attribute. */
function vnodeProp(vnode, name) {
  if (!vnode?.props) return undefined;
  const kebab = name.replace(/[A-Z]/g, (letter) => `-${letter.toLowerCase()}`);
  return vnode.props[name] ?? vnode.props[kebab];
}

function findComponentVNodes(vnode, name, result = []) {
  if (!vnode || typeof vnode !== "object") return result;
  if (vnode.type?.__name === name || vnode.type?.name === name) result.push(vnode);
  if (Array.isArray(vnode.children)) {
    for (const child of vnode.children) findComponentVNodes(child, name, result);
  }
  if (vnode.component?.subTree && vnode.component.subTree !== vnode) {
    findComponentVNodes(vnode.component.subTree, name, result);
  }
  return result;
}

async function mountViewer(model) {
  const mounted = await mount(DocxViewerRoot, {
    model,
    layoutOptions: VIEWER_LAYOUT,
  });
  await vue.nextTick();
  await vue.nextTick();
  return mounted;
}

function pageLabels(mounted) {
  return nodesWithAttribute(mounted.root, "data-docx-page-label").map(
    (node) => node.props["data-docx-page-label"],
  );
}

function parityBlankFlags(mounted) {
  return nodesWithAttribute(mounted.root, "data-docx-page-parity-blank").map(
    (node) => node.props["data-docx-page-parity-blank"],
  );
}

function footersOf(mounted) {
  return findComponentVNodes(mounted.app._instance.subTree, "DocxPageFooter");
}

/** Rendered footer text per page (paragraph hosts render their runs as HTML). */
function footerTexts(mounted) {
  return nodesWithAttribute(mounted.root, "data-docx-page-footer").map((footer) =>
    walk(footer)
      .filter((node) => node?.props?.["data-docx-paragraph-host"] === true)
      .map((host) =>
        String(host.props.innerHTML ?? "")
          .replace(/<[^>]*>/g, "")
          .replace(/&nbsp;/g, " ")
          .trim(),
      )
      .join(" "),
  );
}

// ── 1. page labels and the parity blank page ────────────────────────

test("the viewer labels pages per section and marks the parity blank page", async () => {
  const mounted = await mountViewer(twoSectionModel());
  try {
    assert.deepEqual(
      pageLabels(mounted),
      ["1", "2", "i"],
      "decimal pages keep counting through the blank page, then the section restarts in lowerRoman",
    );
    assert.deepEqual(parityBlankFlags(mounted), ["false", "true", "false"]);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("no parity blank page is inserted when the section already starts on an odd page", async () => {
  const mounted = await mountViewer(twoSectionModel({ firstSectionPages: 2 }));
  try {
    assert.deepEqual(pageLabels(mounted), ["1", "2", "i"]);
    assert.deepEqual(parityBlankFlags(mounted), ["false", "false", "false"]);
  } finally {
    mounted.app.unmount();
  }
});

test("a single-section document keeps plain physical page numbers", async () => {
  const mounted = await mountViewer(singleSectionModel());
  try {
    const labels = pageLabels(mounted);
    assert.ok(labels.length >= 2, `expected several pages, got ${labels.length}`);
    assert.deepEqual(
      labels,
      labels.map((_label, index) => String(index + 1)),
    );
    assert.deepEqual(parityBlankFlags(mounted), labels.map(() => "false"));
  } finally {
    mounted.app.unmount();
  }
});

// ── 2. PAGE / NUMPAGES / SECTIONPAGES ───────────────────────────────

test("header/footer fields resolve from the per-section tracker", async () => {
  const mounted = await mountViewer(twoSectionModel());
  try {
    const footers = footersOf(mounted);
    assert.equal(footers.length, 3, "every page renders its section's footer");

    assert.deepEqual(
      footers.map((footer) => vnodeProp(footer, "pageNumber")),
      [1, 2, 1],
      "PAGE uses the section's start number and format",
    );
    assert.deepEqual(
      footers.map((footer) => vnodeProp(footer, "sectionPages")),
      [2, 2, 1],
      "the parity blank page belongs to the section that ends there",
    );
    assert.deepEqual(
      footers.map((footer) => vnodeProp(footer, "pageNumberFormat")),
      [undefined, undefined, "lowerRoman"],
    );

    assert.deepEqual(
      footerTexts(mounted).map((text) => text.replace(/\s+/g, "")),
      ["1/3/2", "2/3/2", "i/3/1"],
      "PAGE / NUMPAGES / SECTIONPAGES render the resolved values",
    );
  } finally {
    mounted.app.unmount();
  }
});

// ── 3. thumbnail rail labels ────────────────────────────────────────

test("the thumbnail rail labels pages with their section numbers", async (t) => {
  installCanvasEnvironment(t);
  const model = twoSectionModel();
  const mounted = await mount(DocxThumbnailPanel, {
    model,
    layoutOptions: VIEWER_LAYOUT,
    showToggle: false,
  });
  await vue.nextTick();
  await vue.nextTick();

  try {
    const items = walk(mounted.root).filter(
      (node) => node?.props?.["data-thumbnail-page-index"] !== undefined,
    );
    assert.deepEqual(
      items.map((item) => item.props["data-thumbnail-page-index"]),
      [0, 1, 2],
    );
    assert.deepEqual(
      items.map((item) => item.props["data-thumbnail-page-label"]),
      ["1", "2", "i"],
    );
    assert.deepEqual(
      items.map((item) => item.props["data-thumbnail-parity-blank"]),
      ["false", "true", "false"],
    );
    assert.equal(
      classNames(items[1]).includes("docx-thumbnail-item--blank"),
      true,
      "the inserted blank page is marked in the rail",
    );
    assert.equal(classNames(items[0]).includes("docx-thumbnail-item--blank"), false);
    assert.deepEqual(
      items.map((item) =>
        textContent(
          walk(item).find((node) => node?.props?.["data-testid"] === "docx-thumbnail-label"),
        ),
      ),
      ["1", "2", "i"],
    );
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("the thumbnail rail keeps physical labels when no section restarts numbering", async (t) => {
  installCanvasEnvironment(t);
  const model = singleSectionModel();
  const mounted = await mount(DocxThumbnailPanel, {
    model,
    layoutOptions: VIEWER_LAYOUT,
    showToggle: false,
  });
  await vue.nextTick();
  await vue.nextTick();

  try {
    const items = walk(mounted.root).filter(
      (node) => node?.props?.["data-thumbnail-page-index"] !== undefined,
    );
    assert.ok(items.length >= 2, `expected several pages, got ${items.length}`);
    assert.deepEqual(
      items.map((item) => item.props["data-thumbnail-page-label"]),
      items.map((_item, index) => String(index + 1)),
    );
  } finally {
    mounted.app.unmount();
  }
});

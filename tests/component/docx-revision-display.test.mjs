// Issue #149: DOCX revision display modes.
//
// The display model changes what the page shows without changing the editing
// model: inline keeps revisions in the text flow, balloons moves deleted runs
// onto gutter cards, and none shows the final text. Export and accept/reject
// keep every revision in all three modes.

import assert from "node:assert/strict";
import test from "node:test";

import { importFromDemo, mount, textContent, vue, walk } from "./vue-test-renderer.mjs";

const {
  DocxPageSurface,
  DocxToolbar,
  DocxViewer,
  DocxViewerRoot,
  displayTextRangeToSourceTextRange,
  paragraphDisplayOffsetToSourceOffset,
  paragraphSourceOffsetToDisplayOffset,
  renderParagraphRuns,
  renderStaticHtml,
  sourceTextRangeToDisplayTextRange,
  useDocxEditor,
} = await importFromDemo("@arcships/vue-docx");
const {
  createBlankDocumentModel,
  modelToDocumentXml,
} = await importFromDemo("@arcships/docx-core");

const PAGE_WIDTH = 816;
const PAGE_HEIGHT = 1056;
const MARGINS = { top: 72, bottom: 72, left: 72, right: 72 };
const PAGE_LAYOUT = {
  pageWidthPx: PAGE_WIDTH,
  pageHeightPx: PAGE_HEIGHT,
  marginsPx: MARGINS,
};

const REVISION_ATTRS =
  'w:id="7" w:author="Ada" w:date="2024-01-01T00:00:00Z"';

const REVISION_PARAGRAPH_XML =
  "<w:p>" +
  '<w:r><w:t xml:space="preserve">Keep </w:t></w:r>' +
  `<w:del ${REVISION_ATTRS}><w:r><w:t>gone</w:t></w:r></w:del>` +
  `<w:ins ${REVISION_ATTRS}><w:r><w:t>new</w:t></w:r></w:ins>` +
  '<w:r><w:t xml:space="preserve"> tail</w:t></w:r>' +
  "</w:p>";

function revisionParagraph() {
  return {
    type: "paragraph",
    children: [
      { type: "text", text: "Keep " },
      { type: "text", text: "gone" },
      { type: "text", text: "new" },
      { type: "text", text: " tail" },
    ],
    sourceXml: REVISION_PARAGRAPH_XML,
  };
}

function revisionModel() {
  return {
    ...createBlankDocumentModel(),
    nodes: [revisionParagraph()],
  };
}

function paragraphText(paragraph) {
  return paragraph.children
    .filter((child) => child.type === "text")
    .map((child) => child.text)
    .join("");
}

function renderRevision(mode) {
  const rendered = renderParagraphRuns(
    revisionParagraph(),
    "revision-display",
    "light",
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    {
      showTrackedChanges: true,
      revisionDisplay: mode,
      emitTextRangeAttributes: true,
    }
  );
  return renderStaticHtml(rendered);
}

function dataNodes(root, attribute) {
  return walk(root).filter((node) => node.props?.[attribute] !== undefined);
}

function paragraphHosts(root) {
  return dataNodes(root, "data-docx-paragraph-host");
}

async function settle() {
  await vue.nextTick();
  await vue.nextTick();
  await vue.nextTick();
}

test("inline keeps deleted and inserted text in the flow with their markup", () => {
  const html = renderRevision("inline");
  assert.match(html, /gone/u);
  assert.match(html, /new/u);
  assert.match(html, /data-docx-tracked-change="deletion"/u);
  assert.match(html, /data-docx-tracked-change="insertion"/u);
});

test("balloons removes deleted text from the body and keeps insertions inline", () => {
  const html = renderRevision("balloons");
  assert.doesNotMatch(html, /gone/u, "deleted text leaves the body");
  assert.match(html, />new</u);
  assert.match(html, /data-docx-tracked-change="insertion"/u);
  assert.doesNotMatch(html, /data-docx-tracked-change="deletion"/u);

  // Offsets follow the display text: "new" starts at 5, " tail" at 8.
  assert.match(html, /data-docx-text-start="5"/u);
  assert.match(html, /data-docx-text-start="8"/u);
});

test("none renders the final text with no inline revision markup", () => {
  const html = renderRevision("none");
  assert.doesNotMatch(html, /gone/u);
  assert.match(html, />new</u);
  assert.doesNotMatch(html, /data-docx-tracked-change=/u);
});

test("hidden revisions render the final text even when inline is the preferred mode", () => {
  const html = renderStaticHtml(
    renderParagraphRuns(
      revisionParagraph(),
      "hidden-revisions",
      "light",
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      { showTrackedChanges: false, revisionDisplay: "inline" }
    )
  );
  assert.doesNotMatch(html, /gone/u);
  assert.match(html, />new</u);
  assert.doesNotMatch(html, /data-docx-tracked-change=/u);
});

test("paragraph offset mapping covers body, table, header, footer and note stories", () => {
  const paragraph = revisionParagraph();
  assert.equal(
    paragraphDisplayOffsetToSourceOffset(paragraph, "balloons", 8),
    12,
    "caret after the inserted word maps to the source offset"
  );
  assert.equal(
    paragraphSourceOffsetToDisplayOffset(paragraph, "balloons", 12),
    8,
    "restoring the source caret maps back to the display offset"
  );

  const tableModel = {
    ...createBlankDocumentModel(),
    nodes: [
      {
        type: "table",
        rows: [
          {
            cells: [{ nodes: [revisionParagraph()] }],
          },
        ],
      },
    ],
  };
  const cellLocation = {
    kind: "table-cell",
    tableIndex: 0,
    rowIndex: 0,
    cellIndex: 0,
    paragraphIndex: 0,
  };
  const displayRange = {
    start: { location: cellLocation, offset: 5 },
    end: { location: cellLocation, offset: 8 },
  };
  const sourceRange = displayTextRangeToSourceTextRange(
    tableModel,
    "balloons",
    displayRange
  );
  assert.equal(sourceRange.start.offset, 9);
  assert.equal(sourceRange.end.offset, 12);
  assert.deepEqual(
    sourceTextRangeToDisplayTextRange(tableModel, "balloons", sourceRange),
    displayRange
  );

  // Header/footer/note paragraphs use the same paragraph-level helpers; they
  // carry story metadata on their container, not on the paragraph itself.
  assert.equal(
    paragraphDisplayOffsetToSourceOffset(paragraph, "none", 8),
    12
  );
  assert.equal(
    paragraphSourceOffsetToDisplayOffset(paragraph, "none", 12),
    8
  );
});

test("renderer omits synthetic w:delText segments outside inline mode", () => {
  const paragraph = {
    type: "paragraph",
    children: [
      { type: "text", text: "Keep " },
      { type: "text", text: " tail" },
    ],
    sourceXml:
      "<w:p><w:r><w:t xml:space=\"preserve\">Keep </w:t></w:r>" +
      `<w:del ${REVISION_ATTRS}><w:r><w:delText>gone</w:delText></w:r></w:del>` +
      '<w:r><w:t xml:space="preserve"> tail</w:t></w:r></w:p>',
  };
  const render = (mode) =>
    renderStaticHtml(
      renderParagraphRuns(
        paragraph,
        "deltext",
        "light",
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        { showTrackedChanges: true, revisionDisplay: mode }
      )
    );

  assert.match(render("inline"), /data-docx-tracked-change="deletion"/u);
  assert.doesNotMatch(render("balloons"), /gone/u);
  assert.doesNotMatch(render("none"), /data-docx-tracked-change=/u);
});

test("DocxViewerRoot defaults to balloons and switches the visible text flow", async () => {
  const model = revisionModel();
  const display = vue.ref("balloons");
  const Harness = vue.defineComponent({
    setup() {
      return () =>
        vue.h(DocxViewerRoot, {
          model,
          showTrackedChanges: true,
          revisionDisplay: display.value,
          zoomScale: 100,
        });
    },
  });
  const mounted = await mount(Harness);
  try {
    await settle();
    const balloonsHtml = paragraphHosts(mounted.root)[0]?.props.innerHTML ?? "";
    assert.doesNotMatch(balloonsHtml, /gone/u);
    assert.match(balloonsHtml, />new</u);

    display.value = "inline";
    await settle();
    const inlineHtml = paragraphHosts(mounted.root)[0]?.props.innerHTML ?? "";
    assert.match(inlineHtml, /gone/u);
    assert.match(inlineHtml, /data-docx-tracked-change="deletion"/u);

    display.value = "none";
    await settle();
    const noneHtml = paragraphHosts(mounted.root)[0]?.props.innerHTML ?? "";
    assert.doesNotMatch(noneHtml, /gone/u);
    assert.match(noneHtml, />new</u);
    assert.doesNotMatch(noneHtml, /data-docx-tracked-change=/u);

    // The source model and its revision markup survive every display change.
    assert.equal(paragraphText(model.nodes[0]), "Keep gonenew tail");
    assert.match(model.nodes[0].sourceXml, /<w:del\b/u);
    const xml = await modelToDocumentXml(model);
    assert.match(xml, /<w:del\b/u);
    assert.match(xml, /gone/u);
  } finally {
    mounted.app.unmount();
  }
});

function trackedChangeFixture({ id, kind, nodeIndex, text }) {
  return {
    id,
    kind,
    author: "Ada",
    date: "2024-01-01T00:00:00Z",
    text,
    nodeIndex,
    location: { kind: "paragraph", nodeIndex },
  };
}

function modelWithParagraphs(texts) {
  const model = createBlankDocumentModel();
  return {
    ...model,
    nodes: texts.map((text) => ({
      type: "paragraph",
      children: [{ type: "text", text }],
    })),
  };
}

async function mountPageSurface(overrides = {}) {
  return mount(DocxPageSurface, {
    pageIndex: 0,
    pageLayout: PAGE_LAYOUT,
    pageNodeSegments: [{ nodeIndex: 0 }],
    model: modelWithParagraphs(["First paragraph"]),
    pageWidthPx: PAGE_WIDTH,
    pageContentWidthPx: PAGE_WIDTH - MARGINS.left - MARGINS.right,
    pageNumber: 1,
    totalPages: 1,
    trackedChangesEnabled: true,
    commentsEnabled: false,
    trackedChanges: [
      trackedChangeFixture({
        id: "change-deletion",
        kind: "deletion",
        nodeIndex: 0,
        text: "gone",
      }),
    ],
    zoomFactor: 1,
    ...overrides,
  });
}

test("only balloons reserves tracked-change cards; inline and none hide them", async () => {
  for (const [mode, expected] of [
    ["balloons", 1],
    ["inline", 0],
    ["none", 0],
  ]) {
    const mounted = await mountPageSurface({ revisionDisplay: mode });
    try {
      const cards = dataNodes(mounted.root, "data-docx-gutter-card-id");
      assert.equal(cards.length, expected, `${mode} tracked-change cards`);
      assert.deepEqual(mounted.warnings, []);
    } finally {
      mounted.app.unmount();
    }
  }
});

test("inline keeps comment cards even though revision cards are hidden", async () => {
  const mounted = await mountPageSurface({
    revisionDisplay: "inline",
    commentsEnabled: true,
    comments: [
      {
        id: "comment-1",
        commentId: 1,
        author: "Alex Reviewer",
        text: "Check this.",
        nodeIndex: 0,
        location: { kind: "paragraph", nodeIndex: 0 },
      },
    ],
  });
  try {
    const cards = dataNodes(mounted.root, "data-docx-gutter-card-id");
    assert.equal(cards.length, 1);
    assert.match(String(cards[0].props.class), /docx-gutter-card--comment/u);
  } finally {
    mounted.app.unmount();
  }
});

test("the viewer toolbar switches from balloons to inline and back", async () => {
  const mounted = await mount(DocxViewer, {
    model: revisionModel(),
    showToolbar: true,
    // The upstream default row has no review-display control; hosts opt in and
    // the choices render inside the DOCX actions menu.
    toolbarReviewControls: true,
  });
  try {
    await settle();
    const html = () => paragraphHosts(mounted.root)[0]?.props.innerHTML ?? "";
    const element = (testId) =>
      walk(mounted.root).find((node) => node.props?.["data-testid"] === testId);
    const openMenu = async () => {
      await element("docx-toolbar-actions").props.onClick();
      await settle();
      assert.ok(element("docx-actions-menu"), "the actions menu must open");
    };
    const activeDisplay = async () => {
      await openMenu();
      const checked = [
        "balloons",
        "inline",
        "none",
      ].filter((value) => element(`docx-revision-display-${value}`)?.props["aria-checked"]);
      await element("docx-toolbar-actions").props.onClick();
      await settle();
      return checked;
    };
    const chooseDisplay = async (value) => {
      await openMenu();
      await element(`docx-revision-display-${value}`).props.onClick();
      await settle();
    };

    assert.deepEqual(await activeDisplay(), ["balloons"]);
    assert.doesNotMatch(html(), /gone/u);

    await chooseDisplay("inline");
    assert.deepEqual(await activeDisplay(), ["inline"]);
    assert.match(html(), /gone/u);
    assert.match(html(), /data-docx-tracked-change="deletion"/u);

    await chooseDisplay("none");
    assert.deepEqual(await activeDisplay(), ["none"]);
    assert.doesNotMatch(html(), /gone/u);
    assert.doesNotMatch(html(), /data-docx-tracked-change=/u);

    await chooseDisplay("balloons");
    assert.deepEqual(await activeDisplay(), ["balloons"]);
    assert.doesNotMatch(html(), /gone/u);
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

test("the editor toolbar changes the revision display mode", async () => {
  const controller = useDocxEditor({
    starterModel: revisionModel(),
    initialShowTrackedChanges: true,
    initialRevisionDisplay: "balloons",
  });
  assert.equal(controller.revisionDisplay, "balloons");
  const mounted = await mount(DocxToolbar, { controller });
  try {
    const select = walk(mounted.root).find(
      (node) => node.props?.["data-testid"] === "editor-revision-display"
    );
    assert.ok(select, "the display selector renders");
    select.props.onChange({ target: { value: "inline" } });
    assert.equal(controller.revisionDisplay, "inline");
    select.props.onChange({ target: { value: "none" } });
    assert.equal(controller.revisionDisplay, "none");
    controller.setRevisionDisplay("balloons");
    await vue.nextTick();
    assert.equal(controller.revisionDisplay, "balloons");
  } finally {
    mounted.app.unmount();
  }
});

test("balloons keeps the source revision visible to accept/reject", async () => {
  const controller = useDocxEditor({
    starterModel: revisionModel(),
    initialShowTrackedChanges: true,
    initialRevisionDisplay: "balloons",
  });
  assert.ok(controller.trackedChanges.some((change) => change.kind === "deletion"));
  const mounted = await mount(DocxViewerRoot, {
    model: controller.model,
    controller,
    editable: true,
    showTrackedChanges: true,
    revisionDisplay: "balloons",
  });
  try {
    await settle();
    const cards = dataNodes(mounted.root, "data-docx-gutter-card-id");
    assert.ok(
      cards.some((card) =>
        String(card.props["data-docx-tracked-change-kind"]).includes("deletion")
      ),
      "the deleted text is still reviewable in the gutter"
    );
    assert.doesNotMatch(
      paragraphHosts(mounted.root)[0]?.props.innerHTML ?? "",
      /gone/u
    );
    assert.deepEqual(mounted.warnings, []);
  } finally {
    mounted.app.unmount();
  }
});

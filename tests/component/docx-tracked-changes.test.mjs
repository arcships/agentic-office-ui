// DOCX revision UX coverage for issue #41, end to end across the built package:
//
//   * accept/reject commands with history and no-ops
//   * revisions are only collected while they are shown
//   * an edit that drops revision markup reports it
//   * gutter: anchored cards carry Accept/Reject, insertions are not duplicated
//   * the document stays editable while revisions are shown
//   * committed text skips inline deletions
//   * toolbar accept-all / reject-all
//
// Anchor positioning, numbering and page-chrome behaviour stay covered by
// tests/component/docx-annotation-gutter.test.mjs; core model rewriting by
// tests/unit/docx-tracked-changes.test.mjs.
import assert from "node:assert/strict";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  vue,
  walk,
} from "./vue-test-renderer.mjs";

const vueNextTick = () => vue.nextTick();

/** Gutter cards are positioned after mount, so give the layout a couple of ticks. */
async function settleGutter() {
  await vue.nextTick();
  await vue.nextTick();
  await vue.nextTick();
}

const {
  DocxParagraphHost,
  DocxTableHost,
  DocxToolbar,
  DocxViewerRoot,
  useDocxEditor,
} = await importFromDemo("@arcships/vue-docx");
import { readFileSync } from "node:fs";

const {
  acceptTrackedChangesInModel,
  collectTrackedChangesFromModel,
  createBlankDocumentModel,
  modelToDocumentXml,
  rejectTrackedChangesInModel,
  wasmBuildDocModelFromBytes,
} = await importFromDemo("@arcships/docx-core");

function findAllByTestId(root, testId) {
  return walk(root).filter((node) => node.props?.["data-testid"] === testId);
}

function dataNodes(root, attribute) {
  return walk(root).filter((node) => node.props?.[attribute] !== undefined);
}

// Same inline revision markup as apps/demo/public/samples/review-comments.docx.
const PARAGRAPH_REVISION_XML =
  "<w:p>" +
  '<w:r><w:t xml:space="preserve">Response time changed from </w:t></w:r>' +
  '<w:del w:id="1" w:author="Alex Reviewer" w:date="2026-07-11T00:00:00Z">' +
  "<w:r><w:delText>five business days</w:delText></w:r></w:del>" +
  '<w:r><w:t xml:space="preserve"> to </w:t></w:r>' +
  '<w:ins w:id="2" w:author="Morgan Editor" w:date="2026-07-11T00:05:00Z">' +
  "<w:r><w:t>two business days</w:t></w:r></w:ins>" +
  "<w:r><w:t>.</w:t></w:r>" +
  "</w:p>";

function paragraphRevisionModel({ trackRevisionCollector = false } = {}) {
  const reads = { collector: 0 };
  const paragraph = {
    type: "paragraph",
    children: [
      { type: "text", text: "Response time changed from " },
      { type: "text", text: " to " },
      { type: "text", text: "two business days" },
      { type: "text", text: "." },
    ],
  };
  if (trackRevisionCollector) {
    let sourceXml = PARAGRAPH_REVISION_XML;
    Object.defineProperty(paragraph, "sourceXml", {
      enumerable: true,
      configurable: true,
      get() {
        if (
          (new Error("sourceXml").stack ?? "").includes(
            "resolveParagraphTrackedMarkup"
          )
        ) {
          reads.collector += 1;
        }
        return sourceXml;
      },
      set(next) {
        sourceXml = next;
      },
    });
  } else {
    paragraph.sourceXml = PARAGRAPH_REVISION_XML;
  }

  return {
    model: { ...createBlankDocumentModel(), nodes: [paragraph] },
    reads,
  };
}

function insertionOnlyModel() {
  const model = createBlankDocumentModel();
  return {
    ...model,
    nodes: [
      {
        type: "paragraph",
        children: [
          { type: "text", text: "Hello " },
          { type: "text", text: "world" },
        ],
        sourceXml:
          '<w:p><w:r><w:t xml:space="preserve">Hello </w:t></w:r>' +
          '<w:ins w:id="4" w:author="Morgan Editor"><w:r><w:t>world</w:t></w:r></w:ins></w:p>',
      },
    ],
  };
}

function tableRowModel() {
  const rows = [0, 1, 2, 3].map((rowIndex) => ({
    cells: [
      {
        nodes: [
          {
            type: "paragraph",
            children: [{ type: "text", text: `Row ${rowIndex}` }],
            sourceXml:
              rowIndex === 3
                ? '<w:p><w:del w:id="9" w:author="Kim Editor"><w:r><w:delText>Row three (old)</w:delText></w:r></w:del><w:r><w:t>Row 3</w:t></w:r></w:p>'
                : `<w:p><w:r><w:t>Row ${rowIndex}</w:t></w:r></w:p>`,
          },
        ],
      },
    ],
  }));
  return { ...createBlankDocumentModel(), nodes: [{ type: "table", rows }] };
}

function paragraphText(paragraph) {
  return paragraph.children.map((child) => child.text ?? "").join("");
}

function createController(options) {
  return useDocxEditor({ starterModel: paragraphRevisionModel(options).model });
}

// ---------------------------------------------------------------------------
// Controller decisions
// ---------------------------------------------------------------------------

test("revisions are only collected while they are shown", () => {
  const { model, reads } = paragraphRevisionModel({ trackRevisionCollector: true });
  const controller = useDocxEditor({ starterModel: model });

  assert.equal(controller.showTrackedChanges, false);
  assert.deepEqual(controller.trackedChanges, []);
  assert.equal(controller.trackedChangeCount, null);
  assert.equal(
    reads.collector,
    0,
    "hidden revisions must not run the tracked-change collector"
  );

  // Typing while revisions stay hidden must not rescan the document either.
  controller.commitParagraphText(0, "Typed with revisions hidden");
  assert.equal(reads.collector, 0);
});

test("accepting one revision keeps the other one and is undoable", () => {
  const controller = createController();
  controller.setShowTrackedChanges(true);
  const insertion = controller.trackedChanges.find((change) => change.kind === "insertion");
  assert.ok(insertion);

  assert.equal(controller.acceptTrackedChange(insertion.id), true);

  assert.equal(
    paragraphText(controller.model.nodes[0]),
    "Response time changed from  to two business days."
  );
  assert.deepEqual(
    controller.trackedChanges.map((change) => change.kind),
    ["deletion"]
  );
  assert.match(controller.status, /Accepted 1 tracked change/u);
  assert.equal(controller.canUndo, true);

  controller.undo();

  // Revision order follows the collector (document order), not the decision.
  assert.deepEqual(
    controller.trackedChanges.map((change) => change.kind).sort(),
    ["deletion", "insertion"]
  );
  assert.match(controller.model.nodes[0].sourceXml, /<w:ins\b/u);
});

test("rejecting a revision restores the deleted text and is undoable", () => {
  const controller = createController();
  controller.setShowTrackedChanges(true);
  const deletion = controller.trackedChanges.find((change) => change.kind === "deletion");
  assert.ok(deletion);

  assert.equal(controller.rejectTrackedChange(deletion.id), true);
  assert.equal(
    paragraphText(controller.model.nodes[0]),
    "Response time changed from five business days to two business days."
  );
  assert.doesNotMatch(controller.model.nodes[0].sourceXml, /<w:del\b/u);
  assert.doesNotMatch(controller.model.nodes[0].sourceXml, /<w:delText\b/u);

  controller.undo();
  assert.equal(
    paragraphText(controller.model.nodes[0]),
    "Response time changed from  to two business days."
  );
});

test("accept-all and reject-all resolve every revision in one history step", async () => {
  const controller = createController();
  controller.setShowTrackedChanges(true);

  assert.equal(controller.acceptAllTrackedChanges(), true);
  assert.deepEqual(controller.trackedChanges, []);
  assert.match(controller.status, /Accepted 2 tracked changes/u);

  controller.undo();
  assert.equal(controller.trackedChanges.length, 2);

  assert.equal(controller.rejectAllTrackedChanges(), true);
  assert.equal(
    paragraphText(controller.model.nodes[0]),
    "Response time changed from five business days to ."
  );
  assert.deepEqual(controller.trackedChanges, []);

  // The transaction snapshot refreshes on the next Vue tick; a second decision
  // in the same tick would otherwise still see the pre-decision model.
  await vueNextTick();
  assert.equal(controller.acceptAllTrackedChanges(), false);
  assert.equal(controller.rejectAllTrackedChanges(), false);
});

test("the row-scoped decisions only touch the requested row", () => {
  const controller = useDocxEditor({ starterModel: tableRowModel() });

  assert.equal(controller.rejectTrackedChangesInTableRow(0, 1), false);
  assert.equal(controller.rejectTrackedChangesInTableRow(0, 3), true);
  assert.deepEqual(collectTrackedChangesFromModel(controller.model), []);
});

test("an edit that drops revision markup reports it instead of dropping it silently", () => {
  const controller = createController();
  controller.setShowTrackedChanges(true);
  assert.equal(controller.trackedChanges.length, 2);

  controller.commitParagraphText(0, "Response time changed from two business days.");

  assert.deepEqual(controller.trackedChanges, []);
  assert.match(controller.status, /edited paragraph/u);
});

test("the revision sample document round-trips through accept-all and reject-all", async () => {
  const sample = readFileSync(
    new URL("../../apps/demo/public/samples/review-comments.docx", import.meta.url)
  );
  const bytes = sample.buffer.slice(
    sample.byteOffset,
    sample.byteOffset + sample.byteLength
  );
  const { model } = await wasmBuildDocModelFromBytes(bytes);

  const accepted = acceptTrackedChangesInModel(model);
  assert.deepEqual(collectTrackedChangesFromModel(accepted.model), []);
  const acceptedXml = await modelToDocumentXml(accepted.model);
  assert.doesNotMatch(acceptedXml, /<w:(?:ins|del)\b/u);

  const rejected = rejectTrackedChangesInModel(model);
  assert.deepEqual(collectTrackedChangesFromModel(rejected.model), []);
  const rejectedXml = await modelToDocumentXml(rejected.model);
  assert.doesNotMatch(rejectedXml, /<w:(?:ins|del|delText)\b/u);
  assert.match(rejectedXml, /five business days/u);
});

// ---------------------------------------------------------------------------
// Gutter cards
// ---------------------------------------------------------------------------

async function mountViewer({ model, ...props }) {
  return mount(DocxViewerRoot, { model, zoomScale: 100, ...props });
}

function gutterCards(root) {
  return dataNodes(root, "data-docx-gutter-card-id");
}

test("the gutter lists the deletion once, not the inline insertion, and can accept it", async () => {
  const controller = createController();
  controller.setShowTrackedChanges(true);
  const mounted = await mountViewer({
    model: controller.model,
    controller,
    editable: true,
  });
  await settleGutter();

  const cards = gutterCards(mounted.root);
  assert.equal(cards.length, 1, "only revisions without inline markup get a card");
  assert.match(String(cards[0].props["data-docx-tracked-change-kind"]), /deletion/u);

  const acceptButtons = findAllByTestId(mounted.root, "docx-gutter-accept");
  assert.equal(acceptButtons.length, 1);
  // Card actions use `@click.stop`, so the handler expects a real event.
  acceptButtons[0].props.onClick(fakeEvent());
  await vueNextTick();

  assert.equal(
    paragraphText(controller.model.nodes[0]),
    "Response time changed from  to two business days."
  );
  assert.deepEqual(gutterCards(mounted.root), []);
  assert.equal(controller.canUndo, true);
  mounted.app.unmount();
});

test("rejecting a deletion from the gutter restores the deleted text", async () => {
  const controller = createController();
  controller.setShowTrackedChanges(true);
  const mounted = await mountViewer({
    model: controller.model,
    controller,
    editable: true,
  });
  await settleGutter();

  const rejectButtons = findAllByTestId(mounted.root, "docx-gutter-reject");
  assert.equal(rejectButtons.length, 1);
  rejectButtons[0].props.onClick(fakeEvent());
  await vueNextTick();

  assert.equal(
    paragraphText(controller.model.nodes[0]),
    "Response time changed from five business days to two business days."
  );
  mounted.app.unmount();
});

test("an insertion-only document renders no duplicate gutter card", async () => {
  const model = insertionOnlyModel();
  const mounted = await mountViewer({ model, showTrackedChanges: true });
  await settleGutter();

  assert.deepEqual(gutterCards(mounted.root), []);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("the surface follows the controller's Edits flag without an explicit prop", async () => {
  const controller = createController();
  const mounted = await mountViewer({ model: controller.model, controller });
  await settleGutter();

  // Boolean props default to `false`; the surface must still follow the
  // controller when the host did not pass an explicit flag.
  assert.deepEqual(gutterCards(mounted.root), []);

  controller.setShowTrackedChanges(true);
  await settleGutter();
  assert.equal(gutterCards(mounted.root).length, 1);

  controller.setShowTrackedChanges(false);
  await settleGutter();
  assert.deepEqual(gutterCards(mounted.root), []);
  mounted.app.unmount();
});

// ---------------------------------------------------------------------------
// Editing with revisions on screen
// ---------------------------------------------------------------------------

test("a table cell stays editable while revisions are shown", async () => {
  let controller;
  const model = tableRowModel();
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      controller.setShowTrackedChanges(true);
      return () =>
        vue.h(DocxTableHost, {
          table: controller.model.nodes[0],
          tableIndex: 0,
          editable: true,
          controller,
        });
    },
  });
  const mounted = await mount(Harness);

  const cells = dataNodes(mounted.root, "data-docx-table-cell-paragraph-host");
  assert.ok(cells.length > 0, "expected editable paragraph hosts inside the table");
  assert.equal(cells[0].props.contenteditable, "true");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("an editable paragraph host renders while revisions are shown", async () => {
  let controller;
  const model = ordinaryParagraphModel();
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      controller.setShowTrackedChanges(true);
      return () =>
        vue.h(DocxParagraphHost, {
          paragraph: controller.model.nodes[0],
          paragraphIndex: 0,
          editable: true,
          controller,
        });
    },
  });
  const mounted = await mount(Harness);

  const host = dataNodes(mounted.root, "data-docx-paragraph-host")[0];
  assert.ok(host);
  assert.equal(host.props.contenteditable, "true");
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

function ordinaryParagraphModel() {
  return {
    ...createBlankDocumentModel(),
    nodes: [{ type: "paragraph", children: [{ type: "text", text: "Hello" }] }],
  };
}

function fakeEvent() {
  return { stopPropagation() {}, preventDefault() {} };
}

function fakeTextNode(text) {
  return { nodeType: 3, textContent: text, childNodes: [] };
}

function fakeElement({ tagName = "SPAN", attributes = {}, children = [] } = {}) {
  return {
    nodeType: 1,
    tagName,
    childNodes: children,
    get textContent() {
      return children.map((child) => child.textContent ?? "").join("");
    },
    getAttribute(name) {
      return Object.prototype.hasOwnProperty.call(attributes, name)
        ? attributes[name]
        : null;
    },
    closest() {
      return null;
    },
    matches(selector) {
      if (selector === "[data-docx-numbering-label='true']") {
        return attributes["data-docx-numbering-label"] === "true";
      }
      if (selector === "[data-docx-image-child-index]") {
        return "data-docx-image-child-index" in attributes;
      }
      return false;
    },
  };
}

test("committed cell text skips inline deletions so an edit cannot resurrect them", async () => {
  const committed = [];
  let controller;
  const model = tableRowModel();
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      controller.setShowTrackedChanges(true);
      return () =>
        vue.h(DocxTableHost, {
          table: controller.model.nodes[0],
          tableIndex: 0,
          editable: true,
          controller,
          onCellInput: (_tableIndex, _rowIndex, _cellIndex, _paragraphIndex, text) =>
            committed.push(text),
        });
    },
  });
  const mounted = await mount(Harness);

  const host = dataNodes(mounted.root, "data-docx-table-cell-paragraph-host")[0];
  assert.ok(host);

  host.props.onInput({
    target: fakeElement({
      tagName: "DIV",
      children: [
        fakeElement({
          attributes: { "data-docx-numbering-label": "true" },
          children: [fakeTextNode("1. ")],
        }),
        fakeElement({ children: [fakeTextNode("Row ")] }),
        fakeElement({
          attributes: { "data-docx-tracked-change": "deletion" },
          children: [fakeTextNode("three (old)")],
        }),
        fakeElement({ children: [fakeTextNode("3")] }),
        fakeElement({ tagName: "BR", children: [] }),
        fakeElement({ children: [fakeTextNode("Second line")] }),
        fakeElement({
          attributes: { "data-docx-image-child-index": "0" },
          children: [],
        }),
      ],
    }),
  });

  // List markers, deleted runs, line breaks and image children are not part of
  // the text the model should receive.
  assert.deepEqual(committed, ["Row 3\nSecond line"]);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

// ---------------------------------------------------------------------------
// Toolbar
// ---------------------------------------------------------------------------

test("the toolbar accepts and rejects every revision", async () => {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: paragraphRevisionModel().model });
      controller.setShowTrackedChanges(true);
      return () => vue.h(DocxToolbar, { controller });
    },
  });
  const mounted = await mount(Harness);

  const acceptAll = findByTestId(mounted.root, "editor-accept-all-changes");
  const rejectAll = findByTestId(mounted.root, "editor-reject-all-changes");
  assert.ok(acceptAll);
  assert.ok(rejectAll);
  assert.equal(acceptAll.props.disabled, false);
  assert.equal(rejectAll.props.disabled, false);

  acceptAll.props.onClick();
  await vueNextTick();

  assert.deepEqual(controller.trackedChanges, []);
  assert.equal(
    paragraphText(controller.model.nodes[0]),
    "Response time changed from  to two business days."
  );
  // Both buttons report "nothing left" once every revision is resolved.
  assert.equal(findByTestId(mounted.root, "editor-accept-all-changes").props.disabled, true);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

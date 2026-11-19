// Issue #98 — DOCX multi-column sections in the rendered viewer and editor.
//
// Covers the rendering/authoring surface end to end:
//
//   1. the corpus fixture renders two columns at 1440 and 390 px widths with
//      Word's page count and text order,
//   2. an editable surface renders real columns and clips line-range
//      fragments instead of duplicating whole paragraphs,
//   3. the toolbar columns control changes the caret's section, re-paginates,
//      reports the status and is undoable,
//   4. the control follows the caret across sections.
//
// Model-level parsing, export/reload and the fixture snapshot live in
// tests/unit/docx-section-columns.test.mjs.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  textContent,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

const { DocxToolbar, DocxViewerRoot, useDocxEditor } =
  await importFromDemo("@arcships/vue-docx");
const core = await importFromDemo("@arcships/docx-core");

const SAMPLES = new URL("../../apps/demo/public/samples/", import.meta.url);
const CORPUS = new URL("../../test-data/corpus/", import.meta.url);

const COLUMN_SECTION_PROPERTIES =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" ' +
  'w:header="708" w:footer="708" w:gutter="0"/>' +
  '<w:cols w:num="2" w:space="720"/>' +
  "</w:sectPr>";

const SINGLE_SECTION_PROPERTIES =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" ' +
  'w:header="708" w:footer="708" w:gutter="0"/>' +
  '<w:cols w:space="720"/>' +
  "</w:sectPr>";

function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

function modelWith(text, sectionProperties = COLUMN_SECTION_PROPERTIES) {
  const model = core.createBlankDocumentModel();
  model.metadata.sectionPropertiesXml = sectionProperties;
  model.metadata.sections = [
    {
      startNodeIndex: 0,
      sectionPropertiesXml: sectionProperties,
      headerSections: [],
      footerSections: [],
    },
  ];
  model.nodes = [paragraph(text)];
  return model;
}

function modelWithSections(paragraphTexts, sectionColumns) {
  const model = core.createBlankDocumentModel();
  model.nodes = paragraphTexts.map(paragraph);
  model.metadata.sections = sectionColumns.map(({ startNodeIndex, columns }) => {
    const sectionPropertiesXml = columns
      ? COLUMN_SECTION_PROPERTIES
      : SINGLE_SECTION_PROPERTIES;
    return {
      startNodeIndex,
      sectionPropertiesXml,
      headerSections: [],
      footerSections: [],
    };
  });
  model.metadata.sectionPropertiesXml =
    model.metadata.sections.at(-1).sectionPropertiesXml;
  return model;
}

function withContainerWidth(t, widthPx) {
  const originalWidth = Object.getOwnPropertyDescriptor(
    Object.prototype,
    "clientWidth",
  );
  const originalHeight = Object.getOwnPropertyDescriptor(
    Object.prototype,
    "clientHeight",
  );
  Object.defineProperty(Object.prototype, "clientWidth", {
    configurable: true,
    get: () => widthPx,
  });
  // The viewer only mounts pages inside its scroll viewport. A tall fake
  // viewport keeps every page mounted so page/column assertions can see them.
  Object.defineProperty(Object.prototype, "clientHeight", {
    configurable: true,
    get: () => 100_000,
  });
  t.after(() => {
    if (originalWidth) {
      Object.defineProperty(Object.prototype, "clientWidth", originalWidth);
    } else {
      delete Object.prototype.clientWidth;
    }
    if (originalHeight) {
      Object.defineProperty(Object.prototype, "clientHeight", originalHeight);
    } else {
      delete Object.prototype.clientHeight;
    }
  });
}

const fixtureCache = new Map();

async function loadFixture(name, directory = SAMPLES) {
  const key = new URL(name, directory).href;
  if (!fixtureCache.has(key)) {
    const bytes = readFileSync(new URL(name, directory));
    fixtureCache.set(
      key,
      core.importDocxBuffer(
        bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
        { useWorker: false },
      ),
    );
  }
  return fixtureCache.get(key);
}

function dataNodes(root, attribute) {
  return walk(root).filter((node) => node.props?.[attribute] !== undefined);
}

function hasClass(node, className) {
  return String(node.props?.class || "")
    .split(/\s+/)
    .filter(Boolean)
    .includes(className);
}

function nodesWithClass(root, className) {
  return walk(root).filter((node) => hasClass(node, className));
}

/** Body column a rendered host sits in, for column-order assertions. */
function columnAncestor(node) {
  let current = node.parent;
  while (current) {
    if (hasClass(current, "docx-page-column")) return current;
    current = current.parent;
  }
  return undefined;
}

function pageWrappers(root) {
  return dataNodes(root, "data-docx-page-wrapper").sort(
    (left, right) =>
      Number(left.props["data-docx-page-index"]) -
      Number(right.props["data-docx-page-index"]),
  );
}

function columnIndexesIn(root) {
  return nodesWithClass(root, "docx-page-column").map((column) =>
    dataNodes(column, "data-docx-paragraph-node-index").map((node) =>
      Number(node.props["data-docx-paragraph-node-index"]),
    ),
  );
}

async function mountViewerAt(t, widthPx) {
  withContainerWidth(t, widthPx);
  const { model } = await loadFixture("docx-two-columns.docx", CORPUS);
  const mounted = await mount(DocxViewerRoot, {
    model,
    zoomScale: 100,
  });
  t.after(() => mounted.app.unmount());
  await waitFor(
    () => pageWrappers(mounted.root).length > 0,
    "the fixture to paginate",
  );
  return mounted;
}

async function mountEditor(model, layoutOptions) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () =>
        vue.h("div", [
          vue.h(DocxToolbar, { controller }),
          vue.h(DocxViewerRoot, {
            model: controller.model,
            controller,
            editable: true,
            ...(layoutOptions ? { layoutOptions } : {}),
          }),
        ]);
    },
  });
  const mounted = await mount(Harness);
  await vue.nextTick();
  return {
    ...mounted,
    get controller() {
      return controller;
    },
  };
}

// ── 1. viewer at the two widths ─────────────────────────────────────

test("the two-column corpus fixture renders two columns at 1440 and 390 px with Word's page count and text order", async (t) => {
  const results = [];
  for (const width of [1440, 390]) {
    const mounted = await mountViewerAt(t, width);
    const pages = pageWrappers(mounted.root);
    const lastPage = pages.at(-1);
    const columnContainer = dataNodes(
      lastPage,
      "data-docx-page-columns",
    )[0];
    assert.ok(columnContainer, `width ${width}: the page exposes columns`);
    assert.equal(
      Number(columnContainer.props["data-docx-page-columns"]),
      2,
      `width ${width}: the fixture section is two-column`,
    );
    const columns = columnIndexesIn(columnContainer);
    assert.deepEqual(
      columns,
      [
        [0, 1, 2, 3],
        [4, 5, 6, 7],
      ],
      `width ${width}: column one is read before column two`,
    );
    // Column widths are derived from the page box, not the viewport, so the
    // same paper renders the same columns at desktop and phone widths.
    const columnWidths = nodesWithClass(columnContainer, "docx-page-column").map(
      (column) => column.props.style.width,
    );
    assert.equal(new Set(columnWidths).size, 1);
    results.push({
      pageCount: pages.length,
      columns,
      columnWidths,
    });
    mounted.app.unmount();
  }

  assert.deepEqual(
    results.map((result) => result.pageCount),
    [1, 1],
    "the fixture keeps Word's one-page pagination at both widths",
  );
  assert.deepEqual(results[0].columns, results[1].columns);
  assert.deepEqual(results[0].columnWidths, results[1].columnWidths);
});

// ── 2. editable columns ─────────────────────────────────────────────

test("the editor renders columns and clips split paragraph fragments", async () => {
  const mounted = await mountEditor(modelWith("word ".repeat(300)));
  try {
    const container = dataNodes(
      mounted.root,
      "data-docx-page-columns",
    )[0];
    assert.ok(container, "the editable surface renders the column container");
    assert.equal(Number(container.props["data-docx-page-columns"]), 2);
    const columns = nodesWithClass(container, "docx-page-column");
    assert.equal(columns.length, 2);

    const hosts = dataNodes(container, "data-docx-paragraph-node-index");
    assert.equal(hosts.length, 2, "the tall paragraph is split into two slices");
    for (const host of hosts) {
      assert.equal(host.props.contenteditable, "true", "both slices stay editable");
      assert.match(String(host.props.style?.height ?? ""), /^\d+px$/);
      assert.equal(
        host.props.style?.overflow,
        "clip",
        "an editable slice must not scroll its hidden lines into view",
      );
    }
    // The two slices are the same paragraph at different line offsets, not two
    // copies of the whole paragraph pushed into the flow.
    assert.deepEqual(
      hosts.map((host) =>
        Number(host.props["data-docx-paragraph-node-index"]),
      ),
      [0, 0],
    );
    assert.deepEqual(
      hosts.map((host) => columns.indexOf(columnAncestor(host))),
      [0, 1],
      "the continuation slice starts column two",
    );
    assert.notEqual(
      hosts[0].props.style.marginTop,
      hosts[1].props.style.marginTop,
    );
  } finally {
    mounted.app.unmount();
  }
});

test("an editable two-column section keeps paragraph order across its page", async () => {
  const mounted = await mountEditor(
    modelWithSections(
      Array.from({ length: 8 }, (_, index) => `Column item ${index + 1}`),
      [{ startNodeIndex: 0, columns: true }],
    ),
  );
  try {
    const container = dataNodes(
      mounted.root,
      "data-docx-page-columns",
    )[0];
    assert.ok(container, "the editor renders a column container");
    assert.deepEqual(
      columnIndexesIn(container),
      [
        [0, 1, 2, 3],
        [4, 5, 6, 7],
      ],
      "the editor flows paragraphs through column one, then column two",
    );
  } finally {
    mounted.app.unmount();
  }
});

// ── 3. toolbar control ──────────────────────────────────────────────

test("the toolbar columns control updates the caret's section and the layout", async () => {
  const mounted = await mountEditor(
    modelWith("One column paragraph", SINGLE_SECTION_PROPERTIES),
  );
  try {
    const { controller } = mounted;
    const select = findByTestId(mounted.root, "editor-columns");
    assert.ok(select, "the toolbar exposes the columns control");
    assert.equal(select.props["aria-label"], "Columns");
    assert.equal(String(select.props.value), "1");

    select.props.onChange({ target: { value: "2" } });
    await vue.nextTick();
    await vue.nextTick();

    const sectionXml = controller.model.metadata.sections[0].sectionPropertiesXml;
    assert.match(sectionXml, /<w:cols w:num="2" w:space="720"\/>/);
    assert.equal(
      controller.model.metadata.sectionPropertiesXml,
      sectionXml,
      "the body-level sectPr the serializer writes mirrors the final section",
    );
    assert.equal(controller.status, "Section columns updated");
    assert.equal(
      Number(
        dataNodes(mounted.root, "data-docx-page-columns")[0]?.props[
          "data-docx-page-columns"
        ],
      ),
      2,
    );

    select.props.onChange({ target: { value: "3" } });
    await vue.nextTick();
    await vue.nextTick();
    assert.match(
      controller.model.metadata.sections[0].sectionPropertiesXml,
      /w:num="3"/,
    );
    assert.equal(
      Number(
        dataNodes(mounted.root, "data-docx-page-columns")[0]?.props[
          "data-docx-page-columns"
        ],
      ),
      3,
    );

    // The toolbar edit is exportable: the serialized document carries the new
    // count and a fresh import reads the section back the same way.
    const exported = await controller.getDocxBytes();
    const reopened = (await core.buildDocModelFromBytes(exported)).model;
    assert.equal(
      core.readSectionColumns(
        reopened.metadata.sections[0].sectionPropertiesXml,
      ).count,
      3,
    );

    // One undo entry per change; redo restores the three-column layout.
    controller.undo();
    await vue.nextTick();
    assert.match(
      controller.model.metadata.sections[0].sectionPropertiesXml,
      /w:num="2"/,
    );
    assert.equal(
      Number(
        dataNodes(mounted.root, "data-docx-page-columns")[0]?.props[
          "data-docx-page-columns"
        ],
      ),
      2,
    );
    controller.redo();
    await vue.nextTick();
    assert.match(
      controller.model.metadata.sections[0].sectionPropertiesXml,
      /w:num="3"/,
    );

    // Back to one column: the column wrapper disappears and the section is a
    // plain single-column body again.
    findByTestId(mounted.root, "editor-columns").props.onChange({
      target: { value: "1" },
    });
    await vue.nextTick();
    await vue.nextTick();
    assert.doesNotMatch(
      controller.model.metadata.sections[0].sectionPropertiesXml,
      /w:num=/,
    );
    assert.equal(dataNodes(mounted.root, "data-docx-page-columns").length, 0);
  } finally {
    mounted.app.unmount();
  }
});

test("the columns control keeps an imported count above three selectable", async () => {
  const fourColumns = COLUMN_SECTION_PROPERTIES.replace(
    'w:num="2"',
    'w:num="4"',
  );
  const mounted = await mountEditor(modelWith("Four column paragraph", fourColumns));
  try {
    const select = findByTestId(mounted.root, "editor-columns");
    assert.equal(String(select.props.value), "4");
    const options = walk(select).filter((node) => node.type === "option");
    assert.deepEqual(
      options.map((option) => Number(option.props.value)),
      [1, 2, 3, 4],
      "the imported count stays in the list next to Word's common counts",
    );
    assert.equal(textContent(options[3]), "4 columns");

    // Picking a common count rewrites the section and re-renders at that count.
    select.props.onChange({ target: { value: "2" } });
    await vue.nextTick();
    await vue.nextTick();
    assert.match(
      mounted.controller.model.metadata.sections[0].sectionPropertiesXml,
      /w:num="2"/,
    );
    assert.equal(
      String(findByTestId(mounted.root, "editor-columns").props.value),
      "2",
    );
  } finally {
    mounted.app.unmount();
  }
});

test("the columns control follows the caret's section in a multi-section document", async () => {
  const mounted = await mountEditor(
    modelWithSections(
      Array.from({ length: 8 }, (_, index) => `Section item ${index + 1}`),
      [
        { startNodeIndex: 0, columns: false },
        { startNodeIndex: 4, columns: true },
      ],
    ),
  );
  try {
    const { controller } = mounted;
    controller.selectParagraph(5);
    await vue.nextTick();
    assert.equal(
      String(findByTestId(mounted.root, "editor-columns").props.value),
      "2",
      "the control reflects the caret's two-column section",
    );

    findByTestId(mounted.root, "editor-columns").props.onChange({
      target: { value: "3" },
    });
    await vue.nextTick();
    await vue.nextTick();
    assert.match(
      controller.model.metadata.sections[1].sectionPropertiesXml,
      /w:num="3"/,
    );
    assert.doesNotMatch(
      controller.model.metadata.sections[0].sectionPropertiesXml,
      /w:num=/,
      "the first section keeps its one-column layout",
    );

    // Moving the caret into the first section retargets the control.
    controller.selectParagraph(0);
    await vue.nextTick();
    assert.equal(
      String(findByTestId(mounted.root, "editor-columns").props.value),
      "1",
    );
    findByTestId(mounted.root, "editor-columns").props.onChange({
      target: { value: "2" },
    });
    await vue.nextTick();
    await vue.nextTick();
    assert.match(
      controller.model.metadata.sections[0].sectionPropertiesXml,
      /w:num="2"/,
    );
    assert.match(
      controller.model.metadata.sections[1].sectionPropertiesXml,
      /w:num="3"/,
      "the later section's edit survives",
    );
  } finally {
    mounted.app.unmount();
  }
});
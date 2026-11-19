// Issue #97 — DOCX page setup dialog and click-to-edit header/footer.
//
// Covers the authoring surface end to end on the public controller:
//
//   1. toolbar dialog → setPageMargins/setPageSize, one undo entry for "Apply"
//   2. header activation creates a missing part and is undoable
//   3. a header edit lands in every page of the section and survives undo/redo
//
// Model-level export/reload for the same commands lives in
// tests/unit/docx-page-setup.test.mjs.

import assert from "node:assert/strict";
import test from "node:test";

import {
  findByTestId,
  importFromDemo,
  mount,
  teleportTarget,
  textContent,
  vue,
  waitFor,
  walk,
} from "./vue-test-renderer.mjs";

const { DocxToolbar, DocxViewerRoot, useDocxEditor } = await importFromDemo(
  "@arcships/vue-docx",
);
const { cloneDocModel, defaultStarterModel } = await importFromDemo(
  "@arcships/docx-core",
);

// Letter portrait with 1in margins and a default header reference.
const SECTION_PROPERTIES =
  '<w:sectPr><w:headerReference w:type="default" r:id="rId9"/>' +
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
  "</w:sectPr>";
// 200 x 300 CSS px, so the layout options below actually paginate.
const SMALL_PAGE_PROPERTIES =
  '<w:pgSz w:w="3000" w:h="4500"/>' +
  '<w:pgMar w:top="150" w:right="150" w:bottom="150" w:left="150" w:header="150" w:footer="150" w:gutter="0"/>';
const SECTION_PROPERTIES_SMALL =
  '<w:sectPr><w:headerReference w:type="default" r:id="rId9"/>' +
  SMALL_PAGE_PROPERTIES +
  "</w:sectPr>";
const SECTION_PROPERTIES_SMALL_WITHOUT_HEADER =
  "<w:sectPr>" + SMALL_PAGE_PROPERTIES + "</w:sectPr>";

// Pagination the test renderer can measure; matches the 3000 x 4500 twips
// (200 x 300 px) section size above, so several paragraphs share a page.
const VIEWER_LAYOUT = {
  pageWidth: 200,
  pageHeight: 300,
  margin: 10,
  minLineHeight: 30,
  paragraphSpacing: 10,
};

/** Synthetic click for `@click.stop` handlers. */
function clickEvent() {
  return { stopPropagation() {}, preventDefault() {} };
}

function paragraph(text) {
  return { type: "paragraph", children: [{ type: "text", text }] };
}

function sectionHeader(text) {
  return {
    partName: "word/header1.xml",
    referenceType: "default",
    nodes: [paragraph(text)],
  };
}

function modelWithParagraphs(count, sectionProperties = SECTION_PROPERTIES) {
  const model = cloneDocModel(defaultStarterModel);
  model.nodes = Array.from({ length: count }, (_, index) =>
    paragraph(`Planning item ${index + 1}`),
  );
  model.metadata.sectionPropertiesXml = sectionProperties;
  return model;
}

function modelWithHeader(count = 45) {
  const model = modelWithParagraphs(count, SECTION_PROPERTIES_SMALL);
  const header = sectionHeader("Original header");
  model.metadata.headerSections = [header];
  model.metadata.sections = [
    {
      startNodeIndex: 0,
      sectionPropertiesXml: SECTION_PROPERTIES_SMALL,
      headerSections: [header],
      footerSections: [],
    },
  ];
  return model;
}

function modelWithoutHeader(count = 45) {
  const model = modelWithParagraphs(count, SECTION_PROPERTIES_SMALL_WITHOUT_HEADER);
  model.metadata.headerSections = [];
  model.metadata.sections = [
    {
      startNodeIndex: 0,
      sectionPropertiesXml: SECTION_PROPERTIES_SMALL_WITHOUT_HEADER,
      headerSections: [],
      footerSections: [],
    },
  ];
  return model;
}

function findByTestIdInBody(testId) {
  return walk(teleportTarget("body")).find(
    (node) => node.props?.["data-testid"] === testId,
  );
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

function storyText(entry) {
  const node = entry?.nodes?.[0];
  if (!node || node.type !== "paragraph") return "";
  return node.children
    .filter((child) => child.type === "text")
    .map((child) => child.text)
    .join("");
}

async function mountEditor(model) {
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
            layoutOptions: VIEWER_LAYOUT,
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

function pageSurfaces(mounted) {
  return findComponentVNodes(mounted.app._instance.subTree, "DocxPageSurface");
}

// ── 1. page-setup dialog ─────────────────────────────────────────────

test("the toolbar opens the page-setup dialog and applies margins and orientation as one undo step", async () => {
  const mounted = await mountEditor(modelWithParagraphs(2));
  try {
    const { controller } = mounted;
    assert.equal(findByTestIdInBody("docx-page-setup-dialog"), undefined);

    const openButton = findByTestId(mounted.root, "editor-page-setup");
    assert.ok(openButton, "the toolbar exposes a page-setup button");
    openButton.props.onClick();
    await vue.nextTick();

    const dialog = findByTestIdInBody("docx-page-setup-dialog");
    assert.ok(dialog, "the dialog is rendered");
    assert.equal(dialog.props.role, "dialog");
    assert.equal(dialog.props["aria-modal"], "true");

    // 1in margins and Letter portrait are the document's current values.
    const topMargin = findByTestIdInBody("docx-page-setup-margin-top");
    assert.equal(topMargin.props.value, "1");

    topMargin.props.onInput({ target: { value: "0.75" } });
    await vue.nextTick();
    findByTestIdInBody("docx-page-setup-paper-size").props.onChange({
      target: { value: "a4" },
    });
    await vue.nextTick();
    findByTestIdInBody("docx-page-setup-landscape").props.onChange({ target: {} });
    await vue.nextTick();
    findByTestIdInBody("docx-page-setup-apply").props.onClick();
    await vue.nextTick();

    assert.equal(findByTestIdInBody("docx-page-setup-dialog"), undefined, "apply closes the dialog");

    const sectionXml = controller.model.metadata.sections[0].sectionPropertiesXml;
    assert.match(sectionXml, /<w:pgMar w:top="1080"/);
    assert.match(sectionXml, /<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/);
    // The body-level sectPr the serializer writes carries the same fragment.
    assert.equal(controller.model.metadata.sectionPropertiesXml, sectionXml);

    // "Apply" is a single undo step even though it changed paper and margins.
    assert.equal(controller.canUndo, true);
    controller.undo();
    await vue.nextTick();
    const restoredXml =
      controller.model.metadata.sections?.[0]?.sectionPropertiesXml ??
      controller.model.metadata.sectionPropertiesXml;
    assert.match(restoredXml, /<w:pgMar w:top="1440"/);
    assert.match(restoredXml, /<w:pgSz w:w="12240" w:h="15840"\/>/);
    assert.equal(controller.canUndo, false);

    controller.redo();
    await vue.nextTick();
    assert.match(
      controller.model.metadata.sections[0].sectionPropertiesXml,
      /<w:pgMar w:top="1080"/,
    );
  } finally {
    mounted.app.unmount();
  }
});

test("the dialog cancels without touching the model", async () => {
  const mounted = await mountEditor(modelWithParagraphs(2));
  try {
    findByTestId(mounted.root, "editor-page-setup").props.onClick();
    await vue.nextTick();
    findByTestIdInBody("docx-page-setup-margin-top").props.onInput({
      target: { value: "2" },
    });
    await vue.nextTick();
    findByTestIdInBody("docx-page-setup-cancel").props.onClick();
    await vue.nextTick();

    assert.equal(findByTestIdInBody("docx-page-setup-dialog"), undefined);
    const sectionXml = mounted.controller.model.metadata.sections?.[0]?.sectionPropertiesXml
      ?? mounted.controller.model.metadata.sectionPropertiesXml;
    assert.match(sectionXml, /<w:pgMar w:top="1440"/);
    assert.equal(mounted.controller.canUndo, false);
  } finally {
    mounted.app.unmount();
  }
});

test("controller page-setup commands reject no-op values and stay undoable", async () => {
  const mounted = await mountEditor(modelWithParagraphs(2));
  try {
    const { controller } = mounted;
    assert.equal(controller.setPageMargins(0, { topTwips: 1440 }), false);
    assert.equal(controller.setPageSize(0, { orientation: "portrait" }), false);

    assert.equal(controller.setPageMargins(0, { leftTwips: 720 }), true);
    assert.equal(controller.canUndo, true);
    assert.match(
      controller.model.metadata.sectionPropertiesXml,
      /<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="720"/,
    );
    assert.equal(controller.setPageSize(0, { orientation: "landscape" }), true);
    assert.match(
      controller.model.metadata.sectionPropertiesXml,
      /w:orient="landscape"/,
    );
    controller.undo();
    controller.undo();
    await vue.nextTick();
    assert.doesNotMatch(
      controller.model.metadata.sectionPropertiesXml,
      /w:orient="landscape"/,
    );
  } finally {
    mounted.app.unmount();
  }
});

// ── 2. header activation ─────────────────────────────────────────────

test("clicking the header band creates the missing part and opens it for editing", async () => {
  const mounted = await mountEditor(modelWithoutHeader(45));
  try {
    const { controller } = mounted;
    assert.equal(controller.model.metadata.sections[0].headerSections.length, 0);

    const activate = findByTestId(mounted.root, "docx-page-header-edit");
    assert.ok(activate, "every editable page offers a header band");
    assert.equal(activate.props["data-docx-section-empty"], "true");
    activate.props.onClick(clickEvent());
    await vue.nextTick();

    const section = controller.model.metadata.sections[0];
    assert.equal(section.headerSections.length, 1);
    assert.equal(section.headerSections[0].partName, "word/header1.xml");
    assert.match(
      section.sectionPropertiesXml,
      /<w:headerReference w:type="default" r:id="rIdHdr1"\/>/,
    );
    // The document-level list the serializer reads knows the part too.
    assert.deepEqual(
      controller.model.metadata.headerSections.map((entry) => entry.partName),
      ["word/header1.xml"],
    );

    // The region is open: the page now renders an editable header.
    assert.ok(findByTestId(mounted.root, "docx-page-header-editing"));
    assert.equal(findByTestId(mounted.root, "docx-page-header-edit").props["data-docx-section-empty"], "false");
    // Creating the part is undoable on its own.
    assert.equal(controller.canUndo, true);
  } finally {
    mounted.app.unmount();
  }
});

// ── 3. header edits reach every page of the section ──────────────────

test("a header edit updates every page of the section and survives undo/redo", async () => {
  const mounted = await mountEditor(modelWithHeader(45));
  try {
    const { controller } = mounted;
    const surfacesBefore = pageSurfaces(mounted);
    assert.ok(surfacesBefore.length >= 2, "the document paginates into several pages");
    assert.ok(
      surfacesBefore.every(
        (surface) => vnodeProp(surface, "headerSection")?.partName === "word/header1.xml",
      ),
      "every page of the section resolves the same header part",
    );

    findByTestId(mounted.root, "docx-page-header-edit").props.onClick(clickEvent());
    await vue.nextTick();
    assert.ok(findByTestId(mounted.root, "docx-page-header-editing"));

    const headerVNodes = findComponentVNodes(mounted.app._instance.subTree, "DocxPageHeader");
    assert.ok(headerVNodes.length >= 2, "each visible page renders the header");
    // The header host commits on blur; the page wires that to the section text
    // command. Emit the same event with the story-local node index.
    headerVNodes[0].props.onTextInput(0, "Updated header");
    await vue.nextTick();

    const sectionHeaderEntry = controller.model.metadata.sections[0].headerSections[0];
    assert.equal(storyText(sectionHeaderEntry), "Updated header");
    assert.equal(
      storyText(controller.model.metadata.headerSections[0]),
      "Updated header",
      "the top-level header list is updated with the section list",
    );

    // Every page of the section reads the same part, so every page shows it.
    const surfacesAfter = pageSurfaces(mounted);
    assert.ok(surfacesAfter.length >= 2);
    for (const surface of surfacesAfter) {
      assert.equal(storyText(vnodeProp(surface, "headerSection")), "Updated header");
    }

    // One undo entry restores the previous header text on every page.
    assert.equal(controller.canUndo, true);
    controller.undo();
    await vue.nextTick();
    const surfacesUndone = pageSurfaces(mounted);
    assert.equal(storyText(vnodeProp(surfacesUndone[0], "headerSection")), "Original header");
    for (const surface of surfacesUndone) {
      assert.equal(storyText(vnodeProp(surface, "headerSection")), "Original header");
    }

    controller.redo();
    await vue.nextTick();
    const surfacesRedone = pageSurfaces(mounted);
    for (const surface of surfacesRedone) {
      assert.equal(storyText(vnodeProp(surface, "headerSection")), "Updated header");
    }
  } finally {
    mounted.app.unmount();
  }
});

test("the header band closes when the page body is clicked", async () => {
  const mounted = await mountEditor(modelWithHeader(45));
  try {
    findByTestId(mounted.root, "docx-page-header-edit").props.onClick(clickEvent());
    await vue.nextTick();
    assert.ok(findByTestId(mounted.root, "docx-page-header-editing"));

    findByTestId(mounted.root, "docx-page-header-done").props.onClick(clickEvent());
    await vue.nextTick();
    assert.equal(findByTestId(mounted.root, "docx-page-header-editing"), undefined);
    assert.equal(findByTestId(mounted.root, "docx-page-footer-edit") !== undefined, true);
    // Nothing was committed, so no history entry was opened.
    assert.equal(mounted.controller.canUndo, false);
  } finally {
    mounted.app.unmount();
  }
});

test("section story text commits ignore body selection ranges", async () => {
  const mounted = await mountEditor(modelWithHeader(45));
  try {
    const { controller } = mounted;
    findByTestId(mounted.root, "docx-page-header-edit").props.onClick(clickEvent());
    await vue.nextTick();
    const before = JSON.stringify(controller.activeTextRange ?? null);

    const headerVNodes = findComponentVNodes(mounted.app._instance.subTree, "DocxPageHeader");
    headerVNodes[0].props.onTextInput(0, "Story text");
    await vue.nextTick();

    assert.equal(JSON.stringify(controller.activeTextRange ?? null), before);
    assert.equal(controller.selection.kind, "paragraph");
    await waitFor(
      () => storyText(controller.model.metadata.sections[0].headerSections[0]) === "Story text",
      "the header text commit reaches the model",
    );
  } finally {
    mounted.app.unmount();
  }
});

// ── 4. localized chrome ──────────────────────────────────────────────

test("the header band and dialog use the message catalog", async () => {
  const { OFFICE_I18N_INJECTION_KEY, resolveOfficeMessages } = await import(
    new URL("../../packages/office-interaction/src/i18n.ts", import.meta.url)
  );
  const messages = resolveOfficeMessages("zh-CN");
  const mounted = await mount(
    vue.defineComponent({
      setup() {
        vue.provide(OFFICE_I18N_INJECTION_KEY, vue.computed(() => messages));
        const controller = useDocxEditor({ starterModel: modelWithHeader(45) });
        return () =>
          vue.h("div", [
            vue.h(DocxToolbar, { controller }),
            vue.h(DocxViewerRoot, {
              model: controller.model,
              controller,
              editable: true,
              layoutOptions: VIEWER_LAYOUT,
            }),
          ]);
      },
    }),
  );
  try {
    await vue.nextTick();
    const band = findByTestId(mounted.root, "docx-page-header-edit");
    assert.equal(textContent(band), messages.docx.headerFooter.header);
    findByTestId(mounted.root, "editor-page-setup").props.onClick();
    await vue.nextTick();
    const dialog = findByTestIdInBody("docx-page-setup-dialog");
    assert.equal(dialog.props["aria-label"], messages.docx.pageSetup.title);
  } finally {
    mounted.app.unmount();
  }
});

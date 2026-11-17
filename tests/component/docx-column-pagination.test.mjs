// DOCX column pagination through the public viewer component.
//
// A two-column section must render two real body columns and split a paragraph
// that is taller than one column into clipped line-range fragments instead of
// stacking everything in one full-width column.

import assert from "node:assert/strict";
import test from "node:test";

import { importFromDemo, mount, vue, waitFor, walk } from "./vue-test-renderer.mjs";

const { DocxViewer } = await importFromDemo("@arcships/vue-docx");
const { createBlankDocumentModel } = await importFromDemo("@arcships/docx-core");

function hasClass(node, name) {
  return String(node?.props?.class || "")
    .split(/\s+/)
    .filter(Boolean)
    .includes(name);
}

function twoColumnModel() {
  const model = createBlankDocumentModel();
  return {
    ...model,
    metadata: {
      ...model.metadata,
      sections: [
        {
          startNodeIndex: 0,
          headerSections: [],
          footerSections: [],
          sectionPropertiesXml:
            '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>' +
            '<w:cols w:num="2" w:space="720"/></w:sectPr>',
        },
      ],
    },
    nodes: [
      {
        type: "paragraph",
        children: [{ type: "text", text: "word ".repeat(300) }],
      },
    ],
  };
}

test("DOCX two-column section renders real columns and clips line-range fragments", async () => {
  const mounted = await mount({
    render() {
      return vue.h("div", { "data-testid": "column-viewer" }, [
        vue.h(DocxViewer, { model: twoColumnModel() }),
      ]);
    },
  });
  await waitFor(() => walk(mounted.root).some((node) => node.props?.["data-docx-page-columns"]));

  const columnContainer = walk(mounted.root).find(
    (node) => node.props?.["data-docx-page-columns"] !== undefined,
  );
  assert.ok(columnContainer, "viewer must expose the column container");
  assert.equal(String(columnContainer.props["data-docx-page-columns"]), "2");

  const columns = walk(columnContainer).filter((node) =>
    hasClass(node, "docx-page-column"),
  );
  assert.equal(columns.length, 2);

  const paragraphHosts = walk(columnContainer).filter(
    (node) => node.props?.["data-docx-paragraph-host"] === true ||
      node.props?.["data-docx-paragraph-host"] === "true",
  );
  assert.ok(
    paragraphHosts.length >= 2,
    "the tall paragraph must be split into at least two line-range fragments",
  );
  assert.ok(
    paragraphHosts.some((host) => {
      const height = host.props?.style?.height;
      return typeof height === "string" && height.endsWith("px");
    }),
    "line-range fragments must be clipped to their rendered slice height",
  );
});

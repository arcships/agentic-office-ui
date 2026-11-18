import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const { strFromU8, strToU8, unzipSync, zipSync } = await import(
  pathToFileURL(requireFromDocxCore.resolve("fflate")).href
);
const core = await import(
  pathToFileURL(requireFromDocxCore.resolve("@arcships/docx-core")).href
);

const SAMPLE_DOCX = new URL(
  "../../apps/demo/public/samples/demo.docx",
  import.meta.url,
);

const SECTION_ONE_XML =
  '<w:sectPr><w:headerReference w:type="default" r:id="rId9"/><w:footerReference w:type="default" r:id="rId10"/><w:pgSz w:w="12240" w:h="15840" w:orient="landscape"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/><w:type w:val="nextPage"/></w:sectPr>';
const SECTION_TWO_XML =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/><w:pgMar w:top="1008" w:right="1152" w:bottom="1008" w:left="1152" w:header="720" w:footer="720" w:gutter="0"/><w:type w:val="nextPage"/></w:sectPr>';
const FINAL_SECTION_XML =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:type w:val="nextPage"/></w:sectPr>';

function paragraph(text, extra = {}) {
  return {
    type: "paragraph",
    children: [{ type: "text", text }],
    ...extra,
  };
}

function header(partName, text) {
  return {
    partName,
    referenceType: "default",
    nodes: [paragraph(text)],
  };
}

function makeModel() {
  return {
    nodes: [
      paragraph("one"),
      paragraph("two"),
      paragraph("three"),
      paragraph("four"),
      paragraph("five"),
    ],
    metadata: {
      sourceParts: 0,
      warnings: [],
      headerSections: [],
      footerSections: [],
      paragraphStyles: [],
      sectionPropertiesXml: FINAL_SECTION_XML,
      sections: [
        {
          startNodeIndex: 0,
          sectionPropertiesXml: SECTION_ONE_XML,
          headerSections: [header("header-one.xml", "section one header")],
          footerSections: [],
        },
        {
          startNodeIndex: 2,
          sectionPropertiesXml: SECTION_TWO_XML,
          headerSections: [header("header-two.xml", "section two header")],
          footerSections: [],
        },
        {
          startNodeIndex: 4,
          sectionPropertiesXml: FINAL_SECTION_XML,
          headerSections: [],
          footerSections: [],
        },
      ],
    },
  };
}

function withPreservedMarkers(model) {
  model.nodes[1].sourceXml = `<w:p><w:pPr>${SECTION_ONE_XML}</w:pPr><w:r><w:t>two</w:t></w:r></w:p>`;
  model.nodes[3].sourceXml = `<w:p><w:pPr>${SECTION_TWO_XML}</w:pPr><w:r><w:t>four</w:t></w:r></w:p>`;
  return core.normalizeDocModel(model);
}

function sampleDocxWithMidSectionBreak() {
  const archive = unzipSync(new Uint8Array(readFileSync(SAMPLE_DOCX)));
  const documentXml = strFromU8(archive["word/document.xml"]);
  const bodyStart = documentXml.indexOf("<w:body>") + "<w:body>".length;
  const firstParagraphEnd =
    documentXml.indexOf("</w:p>", bodyStart) + "</w:p>".length;
  const sectionParagraph = `<w:p><w:pPr>${SECTION_ONE_XML}</w:pPr><w:r><w:t>End of section one</w:t></w:r></w:p>`;
  archive["word/document.xml"] = strToU8(
    documentXml.slice(0, firstParagraphEnd) +
      sectionParagraph +
      documentXml.slice(firstParagraphEnd),
  );
  const bytes = zipSync(archive);
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  );
}

test("normalize promotes a paragraph sectPr into a preserved section break", () => {
  const model = withPreservedMarkers(makeModel());

  assert.match(
    model.nodes[1].sectionBreak.sectionPropertiesXml,
    /w:orient="landscape"/,
  );
  assert.match(
    model.nodes[3].sectionBreak.sectionPropertiesXml,
    /w:orient="portrait"/,
  );
  assert.equal(
    model.nodes[1].sectionBreak.sectionPropertiesXml,
    SECTION_ONE_XML,
  );
});

test("normalize rebuilds missing metadata.sections from preserved markers", () => {
  const model = makeModel();
  delete model.metadata.sections;
  const normalized = withPreservedMarkers(model);

  assert.deepEqual(
    normalized.metadata.sections.map((section) => section.startNodeIndex),
    [0, 2, 4],
  );
  assert.equal(
    normalized.metadata.sections[2].sectionPropertiesXml,
    FINAL_SECTION_XML,
  );
});

test("typing in a section-ending paragraph keeps section metadata and headers", () => {
  const model = withPreservedMarkers(makeModel());
  const edited = core.updateParagraphText(model, 1, "two edited");

  assert.equal(edited.nodes[1].sourceXml, undefined);
  assert.match(
    edited.nodes[1].sectionBreak.sectionPropertiesXml,
    /w:orient="landscape"/,
  );
  assert.deepEqual(
    edited.metadata.sections.map((section) => section.startNodeIndex),
    [0, 2, 4],
  );
  assert.match(
    edited.metadata.sections[0].sectionPropertiesXml,
    /w:orient="landscape"/,
  );
  assert.equal(
    edited.metadata.sections[0].headerSections[0].partName,
    "header-one.xml",
  );
  assert.equal(
    edited.metadata.sections[1].headerSections[0].partName,
    "header-two.xml",
  );
  assert.equal(
    core.sectionBreakAfterParagraphStartsNewPage(edited.nodes[1]),
    true,
  );
  assert.equal(core.paragraphContainsSectionBreakProperties(edited.nodes[1]), true);
  const resolvedSections = core.resolveDocumentSectionsFromMetadata(
    edited.metadata,
  );
  assert.match(resolvedSections[0].sectionPropertiesXml, /w:orient="landscape"/);
  assert.equal(
    resolvedSections[0].headerSections[0].partName,
    "header-one.xml",
  );
  assert.deepEqual(
    [...core.collectDocxHardPageBreakStartNodeIndexes(edited)],
    [2, 4],
  );
});

test("typing in a paragraph that does not end a section leaves the break alone", () => {
  const model = withPreservedMarkers(makeModel());
  const edited = core.updateParagraphText(model, 0, "one edited");

  assert.equal(edited.nodes[1].sectionBreak.sectionPropertiesXml, SECTION_ONE_XML);
  assert.deepEqual(
    edited.metadata.sections.map((section) => section.startNodeIndex),
    [0, 2, 4],
  );
});

test("removeParagraph transfers the section break to the previous paragraph", () => {
  const model = withPreservedMarkers(makeModel());
  const edited = core.removeParagraph(model, 1);

  assert.equal(edited.nodes.length, 4);
  assert.equal(edited.nodes[0].sectionBreak.sectionPropertiesXml, SECTION_ONE_XML);
  assert.equal(edited.nodes[1].sectionBreak, undefined);
  assert.deepEqual(
    edited.metadata.sections.map((section) => section.startNodeIndex),
    [0, 1, 3],
  );
  assert.equal(
    edited.metadata.sections[0].headerSections[0].partName,
    "header-one.xml",
  );
  assert.equal(
    edited.metadata.sections[1].headerSections[0].partName,
    "header-two.xml",
  );
});

test("removeParagraph keeps section starts aligned when the removed node precedes a boundary", () => {
  const model = withPreservedMarkers(makeModel());
  const edited = core.removeParagraph(model, 0);

  assert.equal(edited.nodes[0].sectionBreak.sectionPropertiesXml, SECTION_ONE_XML);
  assert.deepEqual(
    edited.metadata.sections.map((section) => section.startNodeIndex),
    [0, 1, 3],
  );
});

test("insertParagraph shifts section starts after the insertion point", () => {
  const model = withPreservedMarkers(makeModel());
  const edited = core.insertParagraph(model, "new", 2);

  assert.deepEqual(
    edited.metadata.sections.map((section) => section.startNodeIndex),
    [0, 2, 5],
  );
  assert.equal(edited.nodes[1].sectionBreak.sectionPropertiesXml, SECTION_ONE_XML);
  assert.equal(edited.nodes[0].sectionBreak, undefined);
});

test("duplicateParagraph does not duplicate the section boundary", () => {
  const model = withPreservedMarkers(makeModel());
  const edited = core.duplicateParagraph(model, 1);

  assert.equal(edited.nodes.length, 6);
  assert.equal(edited.nodes[1].sectionBreak.sectionPropertiesXml, SECTION_ONE_XML);
  assert.equal(edited.nodes[2].sectionBreak, undefined);
  assert.deepEqual(
    edited.metadata.sections.map((section) => section.startNodeIndex),
    [0, 2, 5],
  );
});

test("insertSectionBreak at the last paragraph creates a new final section", () => {
  const model = core.normalizeDocModel(makeModel());
  const edited = core.insertSectionBreak(model, 4, "continuous");

  assert.equal(edited.nodes.length, 6);
  assert.equal(edited.nodes[4].sectionBreak.sectionPropertiesXml, FINAL_SECTION_XML);
  assert.deepEqual(
    edited.metadata.sections.map((section) => section.startNodeIndex),
    [0, 2, 4, 5],
  );
  assert.match(
    edited.metadata.sections[3].sectionPropertiesXml,
    /w:type w:val="continuous"/,
  );
  assert.match(
    edited.metadata.sectionPropertiesXml,
    /w:type w:val="continuous"/,
  );
  assert.equal(
    core.sectionBreakAfterParagraphStartsNewPage(edited.nodes[4]),
    true,
  );
  assert.equal(
    core.sectionBreakAfterParagraphStartsNewPage(edited.nodes[5]),
    false,
  );
});

test("insertSectionBreak into a non-final section keeps later boundaries intact", () => {
  const model = withPreservedMarkers(makeModel());
  const edited = core.insertSectionBreak(model, 0, "continuous");

  assert.deepEqual(
    edited.metadata.sections.map((section) => section.startNodeIndex),
    [0, 1, 2, 4],
  );
  assert.equal(edited.nodes[0].sectionBreak.sectionPropertiesXml, SECTION_ONE_XML);
  assert.match(
    edited.nodes[1].sectionBreak.sectionPropertiesXml,
    /w:type w:val="continuous"/,
  );
  assert.match(
    edited.metadata.sections[0].sectionPropertiesXml,
    /w:orient="landscape"/,
  );
  assert.match(
    edited.metadata.sections[1].sectionPropertiesXml,
    /w:type w:val="continuous"/,
  );
  assert.equal(edited.metadata.sections[2].sectionPropertiesXml, SECTION_TWO_XML);
  assert.equal(edited.metadata.sectionPropertiesXml, FINAL_SECTION_XML);
});

test("insertSectionBreak on an existing boundary updates that section's type", () => {
  const model = withPreservedMarkers(makeModel());
  const edited = core.insertSectionBreak(model, 1, "continuous");

  assert.equal(edited.nodes.length, 5);
  assert.deepEqual(
    edited.metadata.sections.map((section) => section.startNodeIndex),
    [0, 2, 4],
  );
  assert.match(
    edited.metadata.sections[1].sectionPropertiesXml,
    /w:type w:val="continuous"/,
  );
  assert.match(
    edited.nodes[3].sectionBreak.sectionPropertiesXml,
    /w:type w:val="continuous"/,
  );
  assert.match(
    edited.metadata.sections[0].sectionPropertiesXml,
    /w:type w:val="nextPage"/,
  );
});

test("deleteSectionBreak merges into the following section's setup and headers", () => {
  const model = withPreservedMarkers(makeModel());
  const edited = core.deleteSectionBreak(model, 1);

  assert.equal(edited.nodes[1].sectionBreak, undefined);
  assert.deepEqual(
    edited.metadata.sections.map((section) => section.startNodeIndex),
    [0, 4],
  );
  assert.equal(edited.metadata.sections[0].sectionPropertiesXml, SECTION_TWO_XML);
  assert.equal(
    edited.metadata.sections[0].headerSections[0].partName,
    "header-two.xml",
  );
  assert.equal(edited.nodes[3].sectionBreak.sectionPropertiesXml, SECTION_TWO_XML);
});

test("serialization re-attaches a preserved section break after typing", async () => {
  const { package: pkg, model } = await core.buildDocModelFromBytes(
    sampleDocxWithMidSectionBreak(),
  );

  assert.equal(model.metadata.sections.length, 2);
  assert.match(
    model.metadata.sections[0].sectionPropertiesXml,
    /w:orient="landscape"/,
  );
  assert.equal(model.metadata.sections[0].headerSections.length, 1);

  const edited = core.updateParagraphText(model, 1, "End of section one edited");
  assert.equal(edited.nodes[1].sourceXml, undefined);

  const serialized = await core.serializeDocx(edited, pkg);
  const archive = unzipSync(new Uint8Array(serialized));
  const documentXml = strFromU8(archive["word/document.xml"]);

  assert.equal((documentXml.match(/<w:sectPr\b/gi) ?? []).length, 2);
  assert.match(documentXml, /w:orient="landscape"/);

  const reimported = await core.buildDocModelFromBytes(serialized);
  assert.equal(
    reimported.model.nodes[1].children
      .filter((child) => child.type === "text")
      .map((child) => child.text)
      .join(""),
    "End of section one edited",
  );
  assert.equal(reimported.model.metadata.sections.length, 2);
  assert.match(
    reimported.model.metadata.sections[0].sectionPropertiesXml,
    /w:orient="landscape"/,
  );
  assert.equal(reimported.model.metadata.sections[0].headerSections.length, 1);
});

test("serialization keeps the surviving section after deleting its ending paragraph", async () => {
  const { package: pkg, model } = await core.buildDocModelFromBytes(
    sampleDocxWithMidSectionBreak(),
  );
  const edited = core.removeParagraph(model, 1);

  assert.equal(edited.nodes[0].sectionBreak !== undefined, true);

  const serialized = await core.serializeDocx(edited, pkg);
  const reimported = await core.buildDocModelFromBytes(serialized);

  assert.equal(reimported.model.metadata.sections.length, 2);
  assert.match(
    reimported.model.metadata.sections[0].sectionPropertiesXml,
    /w:orient="landscape"/,
  );
  assert.equal(reimported.model.metadata.sections[0].headerSections.length, 1);
});

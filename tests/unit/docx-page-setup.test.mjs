// Issue #97 — DOCX page setup and header/footer authoring.
//
// Every block names the bullet it covers so the mapping stays checkable:
//
//   1. margins command rewrites pgMar ............... setPageMargins
//   2. paper size/orientation command rewrites pgSz . setPageSize
//   3. addHeaderFooter creates + references a part .. addHeaderFooter
//   4. export/reload round-trip ..................... serialize + rebuild
//   5. header edits reach every page of a section ... resolved sections
//
// The visual multi-page assertion lives in
// tests/component/docx-page-setup.test.mjs.

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const { strFromU8, unzipSync } = await import(
  pathToFileURL(requireFromDocxCore.resolve("fflate")).href
);
const core = await import(
  pathToFileURL(requireFromDocxCore.resolve("@arcships/docx-core")).href
);

const SAMPLES = new URL("../../apps/demo/public/samples/", import.meta.url);

function sampleBytes(name) {
  const bytes = readFileSync(new URL(name, SAMPLES));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function loadSample(name) {
  return core.buildDocModelFromBytes(sampleBytes(name));
}

function documentXmlOf(exported) {
  const archive = unzipSync(new Uint8Array(exported));
  return strFromU8(archive["word/document.xml"]);
}

function sectionPropertiesOf(model, sectionIndex) {
  return core.resolveSectionPropertiesXml(model, sectionIndex);
}

function pgMarOf(xml) {
  return xml.match(/<w:pgMar\b[^>]*>/)?.[0] ?? "";
}

function pgSzOf(xml) {
  return xml.match(/<w:pgSz\b[^>]*>/)?.[0] ?? "";
}

function twipsOf(tag, attribute) {
  const value = tag.match(new RegExp(`\\b${attribute}="(\\d+)"`))?.[1];
  return value === undefined ? undefined : Number(value);
}

// ── 1. margins ───────────────────────────────────────────────────────

test("setPageMargins writes pgMar and preserves every other sectPr child", async () => {
  const { model } = await loadSample("demo.docx");
  const before = sectionPropertiesOf(model, 0);
  const next = core.setPageMargins(model, 0, {
    topTwips: 720,
    rightTwips: 1080,
    bottomTwips: 900,
    leftTwips: 1260,
  });

  const after = sectionPropertiesOf(next, 0);
  const margins = pgMarOf(after);
  assert.equal(twipsOf(margins, "w:top"), 720);
  assert.equal(twipsOf(margins, "w:right"), 1080);
  assert.equal(twipsOf(margins, "w:bottom"), 900);
  assert.equal(twipsOf(margins, "w:left"), 1260);
  // Header/footer distances and the gutter are untouched.
  assert.equal(
    twipsOf(margins, "w:header"),
    twipsOf(pgMarOf(before), "w:header"),
  );
  assert.equal(twipsOf(margins, "w:gutter"), 0);
  // Page size, columns, doc grid and header/footer references survive.
  assert.equal(pgSzOf(after), pgSzOf(before));
  assert.match(after, /<w:headerReference\b/);
  assert.match(after, /<w:footerReference\b/);
  assert.match(after, /<w:cols\b/);
  assert.match(after, /<w:docGrid\b/);
  // The command copies, never mutates.
  assert.equal(sectionPropertiesOf(model, 0), before);
});

test("setPageMargins keeps the document-level sectPr in step with the final section", async () => {
  const { model } = await loadSample("demo.docx");
  const next = core.setPageMargins(model, 0, { leftTwips: 1440 });

  assert.equal(
    next.metadata.sectionPropertiesXml,
    next.metadata.sections[0].sectionPropertiesXml,
  );
  assert.equal(twipsOf(pgMarOf(next.metadata.sectionPropertiesXml), "w:left"), 1440);
  assert.equal(core.parseSectionLayout(next.metadata.sectionPropertiesXml).marginsPx.left, 96);
});

test("setPageMargins edits only the addressed section", async () => {
  const { model } = await loadSample("sections-columns.docx");
  assert.equal(model.metadata.sections.length, 3);
  const secondBefore = model.metadata.sections[1].sectionPropertiesXml;
  const finalBefore = model.metadata.sectionPropertiesXml;

  const next = core.setPageMargins(model, 0, { topTwips: 500 });

  assert.equal(twipsOf(pgMarOf(sectionPropertiesOf(next, 0)), "w:top"), 500);
  assert.equal(next.metadata.sections[1].sectionPropertiesXml, secondBefore);
  assert.equal(next.metadata.sectionPropertiesXml, finalBefore);
  // The paragraph that ends section 1 carries the same fragment, so a later
  // text edit cannot drop it (issue #56 contract).
  const endParagraph = next.nodes[next.metadata.sections[1].startNodeIndex - 1];
  assert.equal(
    endParagraph.sectionBreak?.sectionPropertiesXml,
    next.metadata.sections[0].sectionPropertiesXml,
  );
});

test("setPageMargins clamps out-of-range section indexes", async () => {
  const { model } = await loadSample("demo.docx");
  const next = core.setPageMargins(model, 99, { topTwips: 800 });
  assert.equal(twipsOf(pgMarOf(sectionPropertiesOf(next, 0)), "w:top"), 800);
});

test("setPageMargins returns the input model when nothing changes", async () => {
  const { model } = await loadSample("demo.docx");
  const current = core.readSectionPageMargins(sectionPropertiesOf(model, 0));
  const next = core.setPageMargins(model, 0, { topTwips: current.topTwips });
  assert.equal(next, model);
});

// ── 2. paper size / orientation ──────────────────────────────────────

test("setPageSize writes pgSz with orientation and swaps contradicting edges", async () => {
  const { model } = await loadSample("demo.docx");
  const next = core.setPageSize(model, 0, {
    widthTwips: 11906,
    heightTwips: 16838,
    orientation: "landscape",
  });

  const pageSize = pgSzOf(sectionPropertiesOf(next, 0));
  assert.equal(twipsOf(pageSize, "w:w"), 16838);
  assert.equal(twipsOf(pageSize, "w:h"), 11906);
  assert.match(pageSize, /w:orient="landscape"/);

  const layout = core.parseSectionLayout(next.metadata.sectionPropertiesXml);
  assert.equal(layout.pageWidthPx, Math.round(16838 / 15));
  assert.equal(layout.pageHeightPx, Math.round(11906 / 15));
  assert.ok(layout.pageWidthPx > layout.pageHeightPx);
});

test("setPageSize keeps portrait when the edges already agree", async () => {
  const { model } = await loadSample("demo.docx");
  const next = core.setPageSize(model, 0, {
    widthTwips: 11906,
    heightTwips: 16838,
    orientation: "portrait",
  });
  const pageSize = pgSzOf(sectionPropertiesOf(next, 0));
  assert.equal(twipsOf(pageSize, "w:w"), 11906);
  assert.equal(twipsOf(pageSize, "w:h"), 16838);
  // Portrait is the default, so Word omits `w:orient` for it.
  assert.doesNotMatch(pageSize, /w:orient=/);
});

test("setPageSize keeps an explicit portrait attribute explicit", () => {
  const xml =
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838" w:orient="portrait"/></w:sectPr>';
  const next = core.writeSectionPageSize(xml, { widthTwips: 11906 });
  assert.match(next, /<w:pgSz w:w="11906" w:h="16838" w:orient="portrait"\/>/);
});

test("setPageSize drops a stale paper code", () => {
  const xml =
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840" w:code="1"/></w:sectPr>';
  const next = core.writeSectionPageSize(xml, { widthTwips: 11906, heightTwips: 16838 });
  assert.doesNotMatch(next, /w:code=/);
  assert.match(next, /<w:pgSz w:w="11906" w:h="16838"\/>/);
});

test("page size presets and orientation invert each other", () => {
  const a4Landscape = core.resolveDocxPageSizePreset("a4", "landscape");
  assert.deepEqual(a4Landscape, {
    widthTwips: 16838,
    heightTwips: 11906,
    orientation: "landscape",
  });
  assert.equal(core.matchDocxPageSizePreset(a4Landscape), "a4");
  assert.equal(
    core.matchDocxPageSizePreset({ widthTwips: 1000, heightTwips: 2000, orientation: "portrait" }),
    undefined,
  );
});

// ── 3. header/footer authoring ───────────────────────────────────────

test("addHeaderFooter creates a part, a sectPr reference and keeps both lists in step", async () => {
  const { model } = await loadSample("demo.docx");
  const next = core.addHeaderFooter(model, 0, "header", { referenceType: "first" });

  assert.notEqual(next, model);
  const section = next.metadata.sections[0];
  const created = section.headerSections.find((entry) => entry.referenceType === "first");
  assert.ok(created, "the section owns the new part");
  assert.equal(created.nodes.length, 1);
  assert.equal(created.nodes[0].type, "paragraph");
  assert.equal(core.paragraphText(created.nodes[0]), "");
  // Word numbers parts per namespace, so the demo's header1.xml is skipped.
  assert.equal(created.partName, "word/header2.xml");
  // The top-level list is the serializer's part source.
  assert.ok(
    next.metadata.headerSections.some((entry) => entry.partName === created.partName),
  );

  const xml = sectionPropertiesOf(next, 0);
  const reference = xml.match(/<w:headerReference\b[^>]*w:type="first"[^>]*>/)?.[0] ?? "";
  assert.ok(reference, "the section references the new part");
  const relationshipId = reference.match(/r:id="([^"]+)"/)?.[1];
  assert.ok(relationshipId);
  // The generated id must not collide with the ids already in the document.
  const existingIds = [...xml.matchAll(/r:id="([^"]+)"/g)].map((match) => match[1]);
  assert.equal(existingIds.filter((id) => id === relationshipId).length, 1);
});

test("addHeaderFooter is a no-op for a variant the section already has", async () => {
  const { model } = await loadSample("demo.docx");
  // demo.docx already carries a default header.
  assert.equal(core.addHeaderFooter(model, 0, "header"), model);
});

test("addHeaderFooter adds a footer to a section without one", async () => {
  const blank = core.createBlankDocumentModel();
  const withFooter = core.addHeaderFooter(blank, 0, "footer");
  const section = withFooter.metadata.sections[0];
  assert.equal(section.footerSections.length, 1);
  assert.equal(section.footerSections[0].partName, "word/footer1.xml");
  assert.match(sectionPropertiesOf(withFooter, 0), /<w:footerReference\b[^>]*w:type="default"/);
});

test("a new header is editable through the section-story text command", async () => {
  const { model } = await loadSample("demo.docx");
  const withHeader = core.addHeaderFooter(model, 0, "header", { referenceType: "even" });
  const partName = withHeader.metadata.sections[0].headerSections.at(-1).partName;

  const nodes = core.findDocumentSectionNodes(withHeader, "header", partName);
  assert.ok(nodes, "the part resolves by name");
  const edited = core.updateSectionParagraphTextAtLocation(
    withHeader,
    { region: "header", partName, nodeIndex: 0 },
    "Confidential draft",
  );
  assert.equal(
    core.paragraphTextFromSectionLocation(
      core.findDocumentSectionNodes(edited, "header", partName),
      { nodeIndex: 0 },
    ),
    "Confidential draft",
  );
  // Every section list that names the part was updated too.
  const topLevel = edited.metadata.headerSections.find(
    (entry) => entry.partName === partName,
  );
  assert.equal(core.paragraphText(topLevel.nodes[0]), "Confidential draft");
});

// ── 4. export / reload ───────────────────────────────────────────────

test("margins and orientation round-trip through export and reload", async () => {
  const { package: pkg, model } = await loadSample("demo.docx");
  let next = core.setPageMargins(model, 0, {
    topTwips: 720,
    rightTwips: 1080,
    bottomTwips: 900,
    leftTwips: 1260,
  });
  next = core.setPageSize(next, 0, {
    widthTwips: 11906,
    heightTwips: 16838,
    orientation: "landscape",
  });

  const exported = await core.serializeDocx(next, pkg);
  const documentXml = documentXmlOf(exported);
  const bodySection = documentXml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/i)?.[0] ?? "";
  assert.equal(twipsOf(pgMarOf(bodySection), "w:top"), 720);
  assert.equal(twipsOf(pgMarOf(bodySection), "w:left"), 1260);
  assert.equal(twipsOf(pgSzOf(bodySection), "w:w"), 16838);
  assert.match(pgSzOf(bodySection), /w:orient="landscape"/);

  const reimported = (await core.buildDocModelFromBytes(exported)).model;
  const reloadedXml = sectionPropertiesOf(reimported, 0);
  assert.equal(twipsOf(pgMarOf(reloadedXml), "w:top"), 720);
  assert.equal(twipsOf(pgMarOf(reloadedXml), "w:bottom"), 900);
  assert.match(pgSzOf(reloadedXml), /w:orient="landscape"/);
  const layout = core.parseSectionLayout(reloadedXml);
  assert.ok(layout.pageWidthPx > layout.pageHeightPx);
  assert.deepEqual(layout.marginsPx, { top: 48, right: 72, bottom: 60, left: 84 });
});

test("a mid-document section's page setup survives export and reload", async () => {
  const { package: pkg, model } = await loadSample("sections-columns.docx");
  const next = core.setPageSize(model, 0, {
    widthTwips: 15840,
    heightTwips: 12240,
    orientation: "landscape",
  });

  const exported = await core.serializeDocx(next, pkg);
  const documentXml = documentXmlOf(exported);
  // Two mid-document breaks plus the body-level sectPr.
  assert.equal((documentXml.match(/<w:sectPr\b/gi) ?? []).length, 3);
  assert.match(documentXml, /<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"\/>/);

  const reimported = (await core.buildDocModelFromBytes(exported)).model;
  assert.equal(reimported.metadata.sections.length, 3);
  assert.match(pgSzOf(sectionPropertiesOf(reimported, 0)), /w:orient="landscape"/);
  // The untouched landscape section stays landscape; portrait stays portrait.
  assert.match(pgSzOf(sectionPropertiesOf(reimported, 1)), /w:h="12240"/);
  assert.match(pgSzOf(sectionPropertiesOf(reimported, 2)), /w:h="15840"/);
});

test("a created header round-trips with its relationship and content type", async () => {
  const { package: pkg, model } = await loadSample("demo.docx");
  let next = core.addHeaderFooter(model, 0, "header", { referenceType: "first" });
  const partName = next.metadata.sections[0].headerSections.at(-1).partName;
  next = core.updateSectionParagraphTextAtLocation(
    next,
    { region: "header", partName, nodeIndex: 0 },
    "First page header",
  );

  const exported = await core.serializeDocx(next, pkg);
  const archive = unzipSync(new Uint8Array(exported));
  const fileName = partName.split("/").pop();

  assert.ok(archive[partName], "the part is written into the package");
  const relationships = strFromU8(archive["word/_rels/document.xml.rels"]);
  const relationshipId = sectionPropertiesOf(next, 0).match(
    /<w:headerReference\b[^>]*w:type="first"[^>]*r:id="([^"]+)"[^>]*>/,
  )?.[1];
  assert.ok(relationshipId);
  assert.match(
    relationships,
    new RegExp(`Id="${relationshipId}"[^>]*relationships/header"[^>]*Target="${fileName}"`),
  );
  const contentTypes = strFromU8(archive["[Content_Types].xml"]);
  assert.match(
    contentTypes,
    new RegExp(`PartName="/${partName}"[^>]*wordprocessingml\\.header\\+xml`),
  );
  assert.match(strFromU8(archive[partName]), /<w:t>First page header<\/w:t>/);

  const reimported = (await core.buildDocModelFromBytes(exported)).model;
  const reloaded = reimported.metadata.sections[0].headerSections.find(
    (entry) => entry.referenceType === "first",
  );
  assert.ok(reloaded, "the header is still referenced after reload");
  assert.equal(core.paragraphText(reloaded.nodes[0]), "First page header");
});

test("comments and a created header share one export wrapper", async () => {
  // Regression for the #96/#97 merge: comment parts and header parts are both
  // written by the same serialization pass, and each needs its own relationship
  // and content-type entry in the shared package wrapper.
  const { package: pkg, model } = await loadSample("demo.docx");
  const commented = core.addDocumentComment(model, {
    range: {
      start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 0 },
      end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 5 },
    },
    text: "Please review",
    author: "Reviewer",
  }).model;
  const withHeader = core.addHeaderFooter(commented, 0, "header", {
    referenceType: "first",
  });
  const partName = withHeader.metadata.sections[0].headerSections.at(-1).partName;
  const edited = core.updateSectionParagraphTextAtLocation(
    withHeader,
    { region: "header", partName, nodeIndex: 0 },
    "Header with comment support",
  );

  const exported = await core.serializeDocx(edited, pkg);
  const archive = unzipSync(new Uint8Array(exported));
  const relationships = strFromU8(archive["word/_rels/document.xml.rels"]);
  const contentTypes = strFromU8(archive["[Content_Types].xml"]);

  assert.ok(archive["word/comments.xml"], "the comment part is written");
  assert.ok(archive[partName], "the header part is written");
  assert.match(relationships, /relationships\/comments"/);
  assert.match(
    relationships,
    new RegExp(`relationships/header"[^>]*Target="${partName.split("/").pop()}"`),
  );
  assert.match(contentTypes, /wordprocessingml\.comments\+xml/);
  assert.match(
    contentTypes,
    new RegExp(`PartName="/${partName}"[^>]*wordprocessingml\\.header\\+xml`),
  );
  const documentXml = documentXmlOf(exported);
  assert.match(documentXml, /<w:commentRangeStart\b/);
  assert.match(documentXml, /<w:commentReference\b/);
  assert.match(strFromU8(archive[partName]), /Header with comment support/);

  const reimported = (await core.buildDocModelFromBytes(exported)).model;
  assert.equal(reimported.metadata.comments?.[0]?.text, "Please review");
  const reloadedHeader = reimported.metadata.sections[0].headerSections.find(
    (entry) => entry.referenceType === "first",
  );
  assert.ok(reloadedHeader, "the header survives reload next to the comment");
  assert.equal(
    core.paragraphText(reloadedHeader.nodes[0]),
    "Header with comment support",
  );
});

test("export leaves an untouched document's package wrapper alone", async () => {
  const { package: pkg, model } = await loadSample("demo.docx");
  const exported = await core.serializeDocx(model, pkg);
  const archive = unzipSync(new Uint8Array(exported));

  const relationshipKeys = (xml) =>
    [...xml.matchAll(/<Relationship\b[^>]*>/g)]
      .map((match) => {
        const id = match[0].match(/\bId="([^"]+)"/)?.[1];
        const type = match[0].match(/\bType="([^"]+)"/)?.[1];
        const target = match[0].match(/\bTarget="([^"]+)"/)?.[1];
        return `${id}|${type}|${target}`;
      })
      .sort();
  const overrideKeys = (xml) =>
    [...xml.matchAll(/<Override\b[^>]*>/g)].map((match) => match[0]).sort();

  // The serializer re-serializes the wrapper, but this patch must not add a
  // relationship or content type to a document that already declares its parts.
  assert.deepEqual(
    relationshipKeys(strFromU8(archive["word/_rels/document.xml.rels"])),
    relationshipKeys(pkg.parts.get("word/_rels/document.xml.rels").content),
  );
  assert.deepEqual(
    overrideKeys(strFromU8(archive["[Content_Types].xml"])),
    overrideKeys(pkg.parts.get("[Content_Types].xml").content),
  );
});

// ── 5. header edits reach every page of the section ──────────────────

test("a header edit is visible to every page of its section", async () => {
  const { model } = await loadSample("demo.docx");
  const partName = model.metadata.sections[0].headerSections[0].partName;
  const edited = core.updateSectionParagraphTextAtLocation(
    model,
    { region: "header", partName, nodeIndex: 0 },
    "Every page",
  );

  // The viewer resolves each page against the same section metadata, so every
  // page of section 0 renders the one edited part.
  const sections = core.resolveDocumentSectionsFromMetadata(edited.metadata);
  for (let pageIndex = 0; pageIndex < 3; pageIndex += 1) {
    const variant = core.selectSectionVariantForPage(
      sections[0].headerSections,
      sections[0].sectionPropertiesXml,
      pageIndex,
      { evenAndOddHeaders: true, documentPageIndex: pageIndex },
    );
    assert.equal(variant?.partName, partName);
    assert.equal(core.paragraphText(variant.nodes[0]), "Every page");
  }
  // A second section that inherits the part sees the same text, too.
  const inherited = core.resolveDocumentSectionsFromMetadata({
    ...edited.metadata,
    sections: [
      edited.metadata.sections[0],
      { startNodeIndex: 1, headerSections: [], footerSections: [] },
    ],
  });
  assert.equal(core.paragraphText(inherited[1].headerSections[0].nodes[0]), "Every page");
});

test("the page's header variant is the one an authoring command creates", () => {
  const withTitlePage =
    '<w:sectPr><w:titlePg/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
  const plain = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';

  assert.equal(
    core.resolveSectionVariantReferenceTypeForPage({
      sectionPropertiesXml: withTitlePage,
      sectionPageIndex: 0,
      documentPageIndex: 4,
      evenAndOddHeaders: true,
    }),
    "first",
  );
  assert.equal(
    core.resolveSectionVariantReferenceTypeForPage({
      sectionPropertiesXml: withTitlePage,
      sectionPageIndex: 1,
      documentPageIndex: 5,
      evenAndOddHeaders: true,
    }),
    "even",
  );
  assert.equal(
    core.resolveSectionVariantReferenceTypeForPage({
      sectionPropertiesXml: plain,
      sectionPageIndex: 2,
      documentPageIndex: 6,
      evenAndOddHeaders: true,
    }),
    "default",
  );
  // Even/odd variants are ignored when the document does not enable them.
  assert.equal(
    core.resolveSectionVariantReferenceTypeForPage({
      sectionPropertiesXml: plain,
      sectionPageIndex: 3,
      documentPageIndex: 7,
      evenAndOddHeaders: false,
    }),
    "default",
  );
});

test("resolvePageSetup reports the geometry the commands wrote", async () => {
  const { model } = await loadSample("sections-columns.docx");
  const next = core.setPageSize(
    core.setPageMargins(model, 1, { topTwips: 600 }),
    1,
    { orientation: "portrait", widthTwips: 11906, heightTwips: 16838 },
  );

  const setups = core.resolvePageSetup(next);
  assert.equal(setups.length, 3);
  assert.deepEqual(setups.map((entry) => entry.index), [0, 1, 2]);
  assert.equal(setups[1].margins.topTwips, 600);
  assert.equal(setups[1].pageSize.orientation, "portrait");
  assert.equal(setups[1].pageSize.widthTwips, 11906);
});

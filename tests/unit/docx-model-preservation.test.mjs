// Regression coverage for issue #45 — model-level data loss on edit.
//
// The issue has five bullets; every block below names the bullet it covers so
// the mapping stays checkable:
//
//   1. `cloneDocModel` dropped note references ............. clone round-trip
//   2. global replace invalidated the whole document's `sourceXml` ... replaceText
//   3. duplicated/pasted nodes copied document-scoped ids .. duplicate/paste
//   4. the paragraph-level image clone dropped crop/XML .... image runs
//   5. range style/link edits dropped note references ...... range mutators

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
const samples = new URL("../../apps/demo/public/samples/", import.meta.url);

const BOOKMARK_PARA_XML =
  '<w:p w14:paraId="1A2B3C4D" w14:textId="5E6F7A8B">' +
  '<w:bookmarkStart w:id="0" w:name="Section_1"/><w:r><w:t>Alpha</w:t></w:r>' +
  '<w:bookmarkEnd w:id="0"/></w:p>';
const DRAWING_PARA_XML =
  '<w:p w14:paraId="20000000"><w:r><w:drawing><wp:inline>' +
  '<wp:docPr id="7" name="Picture 1"/></wp:inline></w:drawing></w:r>' +
  '<w:commentRangeStart w:id="3"/><w:commentRangeEnd w:id="3"/>' +
  '<w:r><w:commentReference w:id="3"/></w:r>' +
  '<w:ins w:id="4" w:author="Reviewer"><w:r><w:t>inserted</w:t></w:r></w:ins></w:p>';
const PLAIN_PARA_XML = "<w:p><w:r><w:t>Beta stable</w:t></w:r></w:p>";
const TABLE_XML = "<w:tbl><w:tr><w:tc/></w:tr></w:tbl>";
const CELL_PARA_XML = "<w:p><w:r><w:t>stable cell</w:t></w:r></w:p>";
const UNTOUCHED_CELL_PARA_XML = "<w:p><w:r><w:t>untouched cell</w:t></w:r></w:p>";

function arrayBufferFromFile(name) {
  const bytes = readFileSync(new URL(name, samples));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function documentXmlFromExport(exported) {
  const archive = unzipSync(new Uint8Array(exported));
  return strFromU8(archive["word/document.xml"]);
}

function collectAttributeValues(xml, pattern) {
  return [...xml.matchAll(pattern)].map((match) => match[1]);
}

function countMatches(xml, pattern) {
  return [...xml.matchAll(pattern)].length;
}

// `cloneDocModel` materializes absent optional fields as explicit `undefined`, so
// JSON (which drops them) is the comparison that shows field-level fidelity.
// Keys are sorted because a clone may build the same values in another order.
function jsonOf(value) {
  return JSON.stringify(sortKeys(value));
}

function sortKeys(value) {
  if (Array.isArray(value)) {
    return value.map(sortKeys);
  }

  if (value && typeof value === "object" && !(value instanceof Uint8Array)) {
    return Object.fromEntries(
      Object.keys(value)
        .sort()
        .map((key) => [key, sortKeys(value[key])]),
    );
  }

  return value;
}

function noteReferencesOf(paragraph) {
  const seen = [];
  for (const child of paragraph.children) {
    if (child.type === "text" && child.noteReference) {
      seen.push(child.noteReference);
    }
  }
  return seen;
}

function paragraphImage({ alt = "Cropped" } = {}) {
  return {
    type: "image",
    src: "data:image/png;base64,AAAA",
    alt,
    widthPx: 100,
    heightPx: 60,
    partName: "word/media/image1.png",
    contentType: "image/png",
    data: Uint8Array.from([9, 8, 7]),
    sourceXml:
      '<w:drawing><wp:inline><wp:docPr id="2" name="Cropped"/></wp:inline></w:drawing>',
    crop: { leftFraction: 0.25, topFraction: 0.1, rightFraction: 0.15, bottomFraction: 0.05 },
    cssFilter: "saturate(1.4)",
    cssOpacity: 0.6,
    floating: {
      xPx: 12,
      yPx: 34,
      horizontalAlign: "right",
      verticalAlign: "bottom",
      wrapType: "none",
    },
  };
}

// A model built through `buildDocModelFromBytes`, i.e. the normalized import
// path the editor uses. It is the only path that populates the model-level
// `commentRanges` anchors that serialization re-emits.
async function sampleModelWithComments() {
  const imported = await core.buildDocModelFromBytes(
    arrayBufferFromFile("review-comments.docx"),
  );
  const commentedIndex = imported.model.nodes.findIndex(
    (node) => node.type === "paragraph" && (node.commentRanges?.length ?? 0) > 0,
  );
  assert.ok(commentedIndex >= 0, "the fixture carries a comment anchor");
  return { ...imported, commentedIndex };
}

function modelWithSourceXml() {
  const model = core.createBlankDocumentModel();
  model.nodes = [
    { type: "paragraph", sourceXml: BOOKMARK_PARA_XML, children: [{ type: "text", text: "Alpha" }] },
    { type: "paragraph", sourceXml: PLAIN_PARA_XML, children: [{ type: "text", text: "Beta stable" }] },
    {
      type: "table",
      sourceXml: TABLE_XML,
      rows: [
        {
          type: "table-row",
          cells: [
            {
              type: "table-cell",
              nodes: [
                { type: "paragraph", sourceXml: CELL_PARA_XML, children: [{ type: "text", text: "stable cell" }] },
              ],
            },
            {
              type: "table-cell",
              nodes: [
                {
                  type: "paragraph",
                  sourceXml: UNTOUCHED_CELL_PARA_XML,
                  children: [{ type: "text", text: "untouched cell" }],
                },
              ],
            },
          ],
        },
      ],
    },
  ];
  return model;
}

function buildFixtureModel() {
  const model = core.createBlankDocumentModel();
  model.nodes = [
    {
      type: "paragraph",
      style: {
        align: "justify",
        headingLevel: 2,
        styleId: "Heading2",
        styleName: "Heading 2",
        numbering: { numId: 2, ilvl: 1 },
        spacing: { beforeTwips: 120, afterTwips: 240, lineTwips: 300, lineRule: "auto" },
        indent: { leftTwips: 200, rightTwips: 100, firstLineTwips: 40, hangingTwips: 20 },
        backgroundColor: "#FFF2CC",
        borders: {
          top: { type: "single", color: "#FF0000", sizeEighthPt: 8, spacePt: 2, frame: true, shadow: false },
          bottom: { type: "double", color: "#00FF00", sizeEighthPt: 4, spacePt: 1, frame: false, shadow: true },
        },
        tabStops: [{ alignment: "center", leader: "dot", positionTwips: 1440 }],
        contextualSpacing: true,
        keepNext: true,
        keepLines: false,
        widowControl: true,
        pageBreakBefore: false,
        dropCap: {
          type: "drop",
          lines: 2,
          wrap: "around",
          horizontalAnchor: "margin",
          verticalAnchor: "text",
          xTwips: 1,
          yTwips: 2,
          horizontalSpaceTwips: 3,
          verticalSpaceTwips: 4,
        },
      },
      paragraphMarkDeleted: false,
      sourceXml: BOOKMARK_PARA_XML,
      // Model-level anchors that only the normalized import path produces.
      sectionBreak: {
        sectionPropertiesXml: '<w:sectPr><w:type w:val="continuous"/></w:sectPr>',
      },
      commentRanges: [
        { commentId: 5, start: 0, end: 5, anchorText: "Hello" },
      ],
      children: [
        {
          type: "text",
          text: "Hello",
          style: {
            bold: true,
            italic: false,
            underline: true,
            strike: false,
            color: "#112233",
            highlight: "#FFFF00",
            backgroundColor: "#EEEEEE",
            fontSizePt: 12,
            fontFamily: "Calibri",
            characterSpacingTwips: 20,
            verticalAlign: "superscript",
            runBorder: { type: "single", color: "#000000", sizeEighthPt: 4, spacePt: 1, frame: false, shadow: true },
          },
          link: "https://example.com/intro",
          noteReference: { kind: "footnote", id: 3 },
        },
        {
          type: "image",
          src: "data:image/png;base64,AAAA",
          alt: "Chart",
          widthPx: 120,
          heightPx: 80,
          partName: "word/media/image1.png",
          contentType: "image/png",
          data: Uint8Array.from([1, 2, 3, 4]),
          sourceXml: '<w:drawing><wp:inline><wp:docPr id="4" name="Chart"/></wp:inline></w:drawing>',
          crop: { leftFraction: 0.1, topFraction: 0.2, rightFraction: 0.3, bottomFraction: 0.4 },
          cssFilter: "contrast(1.2)",
          cssOpacity: 0.8,
          floating: {
            xPx: 10,
            yPx: 20,
            horizontalAlign: "center",
            verticalAlign: "top",
            horizontalRelativeTo: "paragraph",
            verticalRelativeTo: "paragraph",
            distLPx: 5,
            distRPx: 6,
            distTPx: 7,
            distBPx: 8,
            wrapType: "square",
            wrapText: "bothSides",
            behindDocument: false,
            zIndex: 2,
          },
          syntheticTextBox: false,
        },
        {
          type: "form-field",
          fieldType: "dropdown",
          sourceKind: "legacy",
          id: 5,
          tag: "pick",
          title: "Pick",
          placeholder: "Choose",
          checked: false,
          value: "a",
          options: [
            { displayText: "A", value: "a" },
            { displayText: "B", value: "b" },
          ],
          widget: {
            name: "pick",
            enabled: true,
            calcOnExit: false,
            text: { inputType: "regular", defaultText: "x", maxLength: 10, textFormat: "alpha" },
            checkbox: { defaultChecked: true, sizeMode: "exact", sizePt: 9 },
            dropdown: { defaultValue: "a" },
          },
          checkedSymbol: "☒",
          uncheckedSymbol: "☐",
          style: { bold: false },
          link: "https://example.com/field",
          sourceXml: "<w:sdt/>",
        },
        { type: "text", text: "", noteReference: { kind: "endnote", id: 7 } },
      ],
    },
    {
      type: "table",
      sourceXml: TABLE_XML,
      style: {
        widthTwips: 8000,
        indentTwips: 100,
        layout: "fixed",
        cellSpacingTwips: 20,
        cellMarginTwips: { topTwips: 1, rightTwips: 2, bottomTwips: 3, leftTwips: 4 },
        columnWidthsTwips: [3000, 5000],
        borders: {
          top: { type: "single", color: "#000000", sizeEighthPt: 4 },
          insideH: { type: "dashed", color: "#333333", sizeEighthPt: 2 },
        },
        floating: {
          xTwips: 1,
          yTwips: 2,
          leftFromTextTwips: 3,
          rightFromTextTwips: 4,
          topFromTextTwips: 5,
          bottomFromTextTwips: 6,
          horizontalAnchor: "margin",
          verticalAnchor: "page",
          horizontalAlign: "center",
          verticalAlign: "top",
        },
      },
      rows: [
        {
          type: "table-row",
          style: { backgroundColor: "#FFFFFF", heightTwips: 300, heightRule: "exact", cantSplit: true, isHeader: true },
          cells: [
            {
              type: "table-cell",
              style: {
                backgroundColor: "#F2F2F2",
                gridSpan: 2,
                rowSpan: 1,
                vMergeContinuation: false,
                widthTwips: 1000,
                marginTwips: { topTwips: 1, rightTwips: 2, bottomTwips: 3, leftTwips: 4 },
                verticalAlign: "center",
                borders: { left: { type: "single", color: "#000000", sizeEighthPt: 2 } },
              },
              nodes: [
                { type: "paragraph", sourceXml: CELL_PARA_XML, children: [{ type: "text", text: "cell" }] },
                { type: "table", style: {}, rows: [] },
              ],
            },
          ],
        },
      ],
    },
  ];
  model.metadata = {
    ...model.metadata,
    sourceParts: 7,
    warnings: ["a warning"],
    documentPageCount: 3,
    documentOpenTag: '<w:document test="1">',
    documentBackgroundColor: "#FFFFFF",
    sectionPropertiesXml: "<w:sectPr/>",
    sections: [
      {
        startNodeIndex: 0,
        sectionPropertiesXml: "<w:sectPr/>",
        headerSections: [
          {
            partName: "word/header1.xml",
            referenceType: "default",
            nodes: [{ type: "paragraph", sourceXml: "<w:p/>", children: [{ type: "text", text: "head" }] }],
          },
        ],
        footerSections: [],
      },
    ],
    headerSections: [
      {
        partName: "word/header1.xml",
        referenceType: "default",
        nodes: [{ type: "paragraph", children: [{ type: "text", text: "head" }] }],
      },
    ],
    footerSections: [
      {
        partName: "word/footer1.xml",
        referenceType: "first",
        nodes: [{ type: "table", sourceXml: TABLE_XML, rows: [] }],
      },
    ],
    defaultParagraphStyleId: "Normal",
    numberingDefinitions: {
      abstracts: [
        {
          abstractNumId: 1,
          levels: [
            {
              ilvl: 0,
              start: 1,
              format: "decimal",
              text: "%1.",
              suffix: "tab",
              indent: { leftTwips: 720 },
              runStyle: { bold: true },
              bulletFontFamily: "Symbol",
              bulletColor: "#000000",
              pictureBulletId: 1,
              pictureBullet: {
                numPicBulletId: 1,
                src: "data:image/png;base64,AAAA",
                widthPx: 10,
                heightPx: 10,
                partName: "word/media/bullet.png",
                contentType: "image/png",
              },
            },
          ],
        },
      ],
      instances: [
        {
          numId: 2,
          abstractNumId: 1,
          levelStartOverrides: { 0: 3 },
          levelOverrides: [
            { ilvl: 0, start: 5, format: "lowerLetter", text: "%1)", suffix: "space", runStyle: { italic: true } },
          ],
        },
      ],
    },
    compatibility: {
      suppressSpacingBeforeAfterPageBreak: true,
      usePrinterMetrics: false,
      doNotBreakWrappedTables: true,
    },
    footnotes: [
      {
        id: 3,
        text: "note text",
        nodes: [{ type: "paragraph", sourceXml: "<w:p/>", children: [{ type: "text", text: "note" }] }],
      },
    ],
    endnotes: [{ id: 7, text: "endnote text" }],
    comments: [
      { id: 1, author: "A", initials: "A", date: "2026-01-01T00:00:00Z", text: "comment", parentId: 0, resolved: true },
    ],
  };
  return model;
}

test("cloneDocModel round-trips every model field", () => {
  const model = buildFixtureModel();
  const snapshot = jsonOf(model);

  const clone = core.cloneDocModel(model);

  assert.equal(jsonOf(clone), snapshot, "every field of the fixture survives the clone");
  assert.equal(jsonOf(model), snapshot, "cloning must not mutate the source model");

  // No shared mutable state between the model and its clone.
  assert.notEqual(clone.nodes, model.nodes);
  assert.notEqual(clone.nodes[0], model.nodes[0]);
  assert.notEqual(clone.nodes[0].style, model.nodes[0].style);
  assert.notEqual(clone.nodes[0].style.borders.top, model.nodes[0].style.borders.top);
  assert.notEqual(clone.nodes[0].style.tabStops, model.nodes[0].style.tabStops);
  assert.notEqual(clone.nodes[0].style.dropCap, model.nodes[0].style.dropCap);
  assert.notEqual(clone.nodes[0].children[1].crop, model.nodes[0].children[1].crop);
  assert.notEqual(clone.nodes[0].children[1].data, model.nodes[0].children[1].data);
  assert.notEqual(clone.nodes[0].children[1].floating, model.nodes[0].children[1].floating);
  assert.notEqual(clone.nodes[0].children[2].widget, model.nodes[0].children[2].widget);
  assert.notEqual(clone.nodes[0].children[3].noteReference, model.nodes[0].children[3].noteReference);
  assert.notEqual(clone.nodes[1].style.floating, model.nodes[1].style.floating);
  assert.notEqual(clone.metadata.footnotes, model.metadata.footnotes);
  assert.notEqual(clone.metadata.warnings, model.metadata.warnings);
  // Model-level anchors are value objects on the paragraph, so they must be
  // copied (not aliased) as well.
  assert.notEqual(clone.nodes[0].sectionBreak, model.nodes[0].sectionBreak);
  assert.notEqual(clone.nodes[0].commentRanges, model.nodes[0].commentRanges);
  assert.notEqual(clone.nodes[0].commentRanges[0], model.nodes[0].commentRanges[0]);
  assert.deepEqual(clone.nodes[0].sectionBreak, model.nodes[0].sectionBreak);
  assert.deepEqual(clone.nodes[0].commentRanges, model.nodes[0].commentRanges);

  clone.nodes[0].style.tabStops[0].positionTwips = 99;
  clone.nodes[0].children[1].crop.leftFraction = 0.9;
  clone.nodes[0].children[2].widget.text.defaultText = "changed";
  clone.nodes[0].children[3].noteReference.id = 99;
  clone.nodes[0].commentRanges[0].end = 1;
  clone.nodes[0].sectionBreak.sectionPropertiesXml = "<w:sectPr/>";
  clone.nodes[1].rows[0].cells[0].nodes[0].children[0].text = "changed";

  assert.equal(jsonOf(model), snapshot, "mutating the clone must not touch the source model");
});

test("cloneDocModel keeps footnote and endnote references on text runs", () => {
  const clone = core.cloneDocModel(buildFixtureModel());
  const [textRun, , , noteRun] = clone.nodes[0].children;

  assert.deepEqual(textRun.noteReference, { kind: "footnote", id: 3 });
  assert.deepEqual(noteRun.noteReference, { kind: "endnote", id: 7 });
});

test("edit commands keep note references, crops and source XML through the shared clone path", () => {
  const model = core.createBlankDocumentModel();
  model.nodes = [
    {
      type: "paragraph",
      sourceXml: DRAWING_PARA_XML,
      children: [
        { type: "text", text: "Before " },
        {
          type: "image",
          src: "data:image/png;base64,AAAA",
          alt: "Cropped",
          widthPx: 100,
          heightPx: 60,
          partName: "word/media/image1.png",
          contentType: "image/png",
          data: Uint8Array.from([9, 8, 7]),
          sourceXml: '<w:drawing><wp:inline><wp:docPr id="2" name="Cropped"/></wp:inline></w:drawing>',
          crop: { leftFraction: 0.25, topFraction: 0.1, rightFraction: 0.15, bottomFraction: 0.05 },
          cssFilter: "saturate(1.4)",
          cssOpacity: 0.6,
          floating: { xPx: 12, yPx: 34, horizontalAlign: "right", verticalAlign: "bottom", wrapType: "none" },
        },
        { type: "text", text: " after", noteReference: { kind: "footnote", id: 3 } },
        { type: "text", text: "", noteReference: { kind: "endnote", id: 7 } },
      ],
    },
  ];
  const sourceImage = model.nodes[0].children.find((child) => child.type === "image");

  for (const [label, next] of [
    ["updateParagraphText", core.updateParagraphText(model, 0, "Before inserted after")],
    ["insertParagraph", core.insertParagraph(model, "inserted", 1)],
    ["applyRunStyle", core.applyRunStyle(model, 0, 0, { italic: true })],
    ["setParagraphHeading", core.setParagraphHeading(model, 0, 3)],
    ["duplicateParagraph", core.duplicateParagraph(model, 0)],
    ["replaceText", core.replaceText(model, "Before", "Before")],
  ]) {
    const children = next.nodes[0].children;
    assert.deepEqual(
      noteReferencesOf(next.nodes[0]).sort((left, right) => left.id - right.id),
      [{ kind: "footnote", id: 3 }, { kind: "endnote", id: 7 }],
      `${label} kept both note references`,
    );

    const image = children.find((child) => child.type === "image");
    assert.ok(image, `${label} kept the image run`);
    assert.deepEqual(image.crop, sourceImage.crop, `${label} kept the image crop`);
    assert.equal(image.cssFilter, sourceImage.cssFilter, `${label} kept the image filter`);
    assert.equal(image.cssOpacity, sourceImage.cssOpacity, `${label} kept the image opacity`);
    assert.equal(image.sourceXml, sourceImage.sourceXml, `${label} kept the drawing XML`);
  }
});

test("replaceText invalidates sourceXml only for the nodes it changed", () => {
  const model = modelWithSourceXml();
  const snapshot = JSON.stringify(model);

  const next = core.replaceText(model, "stable", "changed");
  const [firstParagraph, secondParagraph, table] = next.nodes;

  assert.equal(firstParagraph.sourceXml, BOOKMARK_PARA_XML, "unmatched paragraph keeps its XML");
  assert.equal(firstParagraph.children[0].text, "Alpha");
  assert.equal(secondParagraph.sourceXml, undefined, "changed paragraph loses its XML");
  assert.equal(secondParagraph.children[0].text, "Beta changed");
  assert.equal(table.sourceXml, undefined, "table XML is stale once a cell changed");
  assert.equal(table.rows[0].cells[0].nodes[0].sourceXml, undefined);
  assert.equal(table.rows[0].cells[0].nodes[0].children[0].text, "changed cell");
  assert.equal(
    table.rows[0].cells[1].nodes[0].sourceXml,
    UNTOUCHED_CELL_PARA_XML,
    "cell paragraphs that did not match keep their XML",
  );

  assert.equal(JSON.stringify(model), snapshot, "replaceText must not mutate the input model");
});

test("replaceText without a match preserves every sourceXml", () => {
  const model = modelWithSourceXml();
  const before = jsonOf(model);

  const next = core.replaceText(model, "no-such-term", "replacement");

  assert.equal(jsonOf(next), before, "a replace that matches nothing changes nothing");
  assert.equal(next.nodes[0].sourceXml, BOOKMARK_PARA_XML);
  assert.equal(next.nodes[1].sourceXml, PLAIN_PARA_XML);
  assert.equal(next.nodes[2].sourceXml, TABLE_XML);
});

test("replaceText with a global regex and a no-op replacement keeps sourceXml", () => {
  const model = modelWithSourceXml();
  const before = jsonOf(model);

  const next = core.replaceText(model, /alpha/gi, "Alpha");

  assert.equal(jsonOf(next), before);
  const changed = core.replaceText(model, /ALPHA/gi, "Gamma");
  assert.equal(changed.nodes[0].sourceXml, undefined);
  assert.equal(changed.nodes[0].children[0].text, "Gamma");
  assert.equal(changed.nodes[1].sourceXml, PLAIN_PARA_XML);
});

test("duplicateParagraph gives the copy fresh XML identifiers", () => {
  const model = core.createBlankDocumentModel();
  model.nodes = [
    {
      type: "paragraph",
      sourceXml: DRAWING_PARA_XML,
      children: [{ type: "text", text: "Alpha" }],
    },
    {
      type: "paragraph",
      // Occupies the first id the allocator would hand out, so a fresh id has
      // to skip it.
      sourceXml: '<w:p w14:paraId="10000000"><w:r><w:t>Beta</w:t></w:r></w:p>',
      children: [{ type: "text", text: "Beta" }],
    },
  ];

  const next = core.duplicateParagraph(model, 0);
  const original = next.nodes[0];
  const copy = next.nodes[1];

  assert.equal(next.nodes.length, 3);
  assert.equal(original.sourceXml, DRAWING_PARA_XML, "the source paragraph keeps its own XML");
  assert.equal(original.children[0].text, "Alpha");

  // Duplicated identifiers must not survive in the copy.
  assert.ok(copy.sourceXml, "the copy keeps its (rewritten) XML");
  assert.ok(!copy.sourceXml.includes("bookmarkStart"), "bookmarks are not duplicated");
  assert.ok(!copy.sourceXml.includes("bookmarkEnd"), "bookmark ends are not duplicated");
  assert.ok(!copy.sourceXml.includes("commentRangeStart"), "comment ranges are not duplicated");
  assert.ok(!copy.sourceXml.includes("commentRangeEnd"), "comment range ends are not duplicated");
  assert.ok(!copy.sourceXml.includes("commentReference"), "comment references are not duplicated");
  assert.ok(!copy.sourceXml.includes("1A2B3C4D"), "w14:paraId is renumbered");
  assert.ok(!copy.sourceXml.includes('id="7"'), "wp:docPr id is renumbered");
  assert.ok(!copy.sourceXml.includes('w:id="4"'), "tracked-change id is renumbered");

  const paraIds = next.nodes.flatMap((node) =>
    collectAttributeValues(node.sourceXml ?? "", /\bw14:paraId="([^"]*)"/g),
  );
  assert.equal(new Set(paraIds).size, paraIds.length, "paraIds stay unique across the model");

  const drawingIds = next.nodes.flatMap((node) =>
    collectAttributeValues(node.sourceXml ?? "", /<wp:docPr\b[^>]*?\bid="(\d+)"/g),
  );
  assert.equal(new Set(drawingIds).size, drawingIds.length, "drawing ids stay unique across the model");
});

test("exported duplicates carry no duplicated XML identifiers", async () => {
  const model = core.createBlankDocumentModel();
  model.nodes = [
    { type: "paragraph", sourceXml: BOOKMARK_PARA_XML, children: [{ type: "text", text: "Alpha" }] },
    { type: "paragraph", sourceXml: DRAWING_PARA_XML, children: [{ type: "text", text: "Beta" }] },
  ];

  const duplicated = core.duplicateParagraph(core.duplicateParagraph(model, 0), 2);
  const documentXml = documentXmlFromExport(await core.wasmSerializeDocx(duplicated));

  const bookmarkIds = collectAttributeValues(documentXml, /<w:bookmarkStart\b[^>]*?w:id="(\d+)"/g);
  assert.equal(bookmarkIds.length, 1, "only the original bookmark is exported");
  assert.equal(countMatches(documentXml, /<w:bookmarkEnd\b/g), 1);

  const paraIds = collectAttributeValues(documentXml, /\bw14:paraId="([^"]*)"/g);
  assert.equal(paraIds.length, 4);
  assert.equal(new Set(paraIds).size, paraIds.length, "exported paraIds are unique");

  const drawingIds = collectAttributeValues(documentXml, /<wp:docPr\b[^>]*?\bid="(\d+)"/g);
  assert.equal(drawingIds.length, 2);
  assert.equal(new Set(drawingIds).size, drawingIds.length, "exported drawing ids are unique");

  const revisionIds = collectAttributeValues(documentXml, /<w:ins\b[^>]*?w:id="(\d+)"/g);
  assert.equal(revisionIds.length, 2);
  assert.equal(new Set(revisionIds).size, revisionIds.length, "exported tracked-change ids are unique");
  assert.equal(countMatches(documentXml, /<w:commentReference\b/g), 1, "only the original comment reference is exported");
});

test("pasteParagraphs sanitizes pasted identifiers and keeps the target intact", () => {
  const model = modelWithSourceXml();
  const pasted = core.parseParagraphsFromClipboard(
    core.serializeParagraphsForClipboard([model.nodes[0]]),
  );

  assert.ok(pasted, "clipboard payload round-trips");
  const next = core.pasteParagraphs(model, model.nodes.length, pasted);
  const copy = next.nodes[next.nodes.length - 1];

  assert.equal(next.nodes[0].sourceXml, BOOKMARK_PARA_XML, "the source node keeps its identifiers");
  assert.ok(copy.sourceXml);
  assert.ok(!copy.sourceXml.includes("bookmarkStart"));
  assert.ok(!copy.sourceXml.includes("1A2B3C4D"));
  assert.equal(copy.children[0].text, "Alpha");
});

test("clipboard round-trip keeps note references, crops and source XML", () => {
  const paragraph = {
    type: "paragraph",
    sourceXml: BOOKMARK_PARA_XML,
    children: [
      { type: "text", text: "Hello", link: "https://example.com", noteReference: { kind: "footnote", id: 3 } },
      {
        type: "image",
        partName: "word/media/image1.png",
        contentType: "image/png",
        data: Uint8Array.from([1, 2]),
        sourceXml: "<w:drawing/>",
        crop: { leftFraction: 0.1, topFraction: 0.2, rightFraction: 0.3, bottomFraction: 0.4 },
        cssFilter: "blur(2px)",
        cssOpacity: 0.5,
      },
    ],
  };

  const parsed = core.parseParagraphsFromClipboard(core.serializeParagraphsForClipboard([paragraph]));

  assert.ok(parsed);
  assert.equal(parsed[0].sourceXml, BOOKMARK_PARA_XML);
  assert.deepEqual(parsed[0].children[0].noteReference, { kind: "footnote", id: 3 });
  assert.deepEqual(parsed[0].children[1].crop, paragraph.children[1].crop);
  assert.equal(parsed[0].children[1].cssFilter, "blur(2px)");
  assert.equal(parsed[0].children[1].cssOpacity, 0.5);
  assert.equal(parsed[0].children[1].sourceXml, "<w:drawing/>");
  assert.deepEqual([...parsed[0].children[1].data], [1, 2]);
});

test("updateParagraphText keeps crop and drawing XML of paragraph images", () => {
  const paragraph = {
    type: "paragraph",
    children: [
      { type: "text", text: "Before " },
      {
        type: "image",
        src: "data:image/png;base64,AAAA",
        alt: "Cropped",
        widthPx: 100,
        heightPx: 60,
        partName: "word/media/image1.png",
        contentType: "image/png",
        data: Uint8Array.from([9, 8, 7]),
        sourceXml: '<w:drawing><wp:inline><wp:docPr id="2" name="Cropped"/></wp:inline></w:drawing>',
        crop: { leftFraction: 0.25, topFraction: 0.1, rightFraction: 0.15, bottomFraction: 0.05 },
        cssFilter: "saturate(1.4)",
        cssOpacity: 0.6,
        floating: { xPx: 12, yPx: 34, horizontalAlign: "right", verticalAlign: "bottom", wrapType: "none" },
      },
      { type: "text", text: " after" },
    ],
  };
  const model = core.createBlankDocumentModel();
  model.nodes = [paragraph];

  const next = core.updateParagraphText(model, 0, "Before inserted after");
  const image = next.nodes[0].children.find((child) => child.type === "image");

  assert.ok(image, "the image survives the text edit");
  assert.deepEqual(image.crop, paragraph.children[1].crop);
  assert.equal(image.cssFilter, "saturate(1.4)");
  assert.equal(image.cssOpacity, 0.6);
  assert.equal(image.sourceXml, paragraph.children[1].sourceXml);
  assert.deepEqual(image.floating, paragraph.children[1].floating);
  assert.deepEqual([...image.data], [9, 8, 7]);
});

test("range style and link mutations keep note references", () => {
  const buildParagraph = () => ({
    type: "paragraph",
    children: [
      { type: "text", text: "Total: " },
      { type: "text", text: "", noteReference: { kind: "footnote", id: 1 } },
      { type: "text", text: "amount" },
    ],
  });

  const styled = buildParagraph();
  assert.equal(
    core.mutateParagraphTextStyleInRange(styled, 0, 6, (style) => ({ ...style, bold: true })),
    true,
  );
  const styledNoteRuns = styled.children.filter((child) => child.noteReference);
  assert.equal(styledNoteRuns.length, 1, "the footnote reference is still present once");
  assert.deepEqual(styledNoteRuns[0].noteReference, { kind: "footnote", id: 1 });

  const linked = buildParagraph();
  assert.equal(core.mutateParagraphLinkInRange(linked, 0, 5, "https://example.com"), true);
  const linkedNoteRuns = linked.children.filter((child) => child.noteReference);
  assert.equal(linkedNoteRuns.length, 1);
  assert.deepEqual(linkedNoteRuns[0].noteReference, { kind: "footnote", id: 1 });

  const untouchedRange = buildParagraph();
  core.mutateParagraphTextStyleInRange(untouchedRange, 7, 13, (style) => ({ ...style, italic: true }));
  const amountRun = untouchedRange.children.find((child) => child.text === "amount");
  assert.deepEqual(untouchedRange.children.filter((child) => child.noteReference).length, 1);
  assert.ok(amountRun);
});

test("splitting a note-bearing run by a range edit keeps one reference, on the last piece", () => {
  const paragraph = {
    type: "paragraph",
    children: [{ type: "text", text: "abcd", noteReference: { kind: "endnote", id: 5 } }],
  };

  core.mutateParagraphTextStyleInRange(paragraph, 1, 3, (style) => ({ ...style, bold: true }));

  const noteRuns = paragraph.children.filter((child) => child.noteReference);
  assert.equal(noteRuns.length, 1, "the reference is not duplicated across the split pieces");
  assert.equal(noteRuns[0].text, "d", "the reference stays at the end of the original run");
  assert.equal(paragraph.children.map((child) => child.text).join(""), "abcd");
});

test("mergeAdjacentTextRuns does not merge runs with different note references", () => {
  const merged = core.mergeAdjacentTextRuns([
    { type: "text", text: "a", style: { bold: true } },
    { type: "text", text: "b", style: { bold: true }, noteReference: { kind: "footnote", id: 1 } },
    { type: "text", text: "c", style: { bold: true } },
  ]);

  assert.equal(merged.length, 3);
  assert.deepEqual(merged[1].noteReference, { kind: "footnote", id: 1 });

  const sameReference = core.mergeAdjacentTextRuns([
    { type: "text", text: "a", style: { bold: true }, noteReference: { kind: "footnote", id: 1 } },
    { type: "text", text: "b", style: { bold: true }, noteReference: { kind: "footnote", id: 1 } },
  ]);
  assert.equal(sameReference.length, 1);
  assert.equal(sameReference[0].text, "ab");
  assert.deepEqual(sameReference[0].noteReference, { kind: "footnote", id: 1 });
});

test("range style and link mutations keep note references, crops and comment anchors", () => {
  const buildParagraph = () => ({
    type: "paragraph",
    // Comment anchors are character offsets; a style/link edit keeps the text, so
    // the anchors must come through untouched.
    commentRanges: [{ commentId: 2, start: 0, end: 7, anchorText: "Total: " }],
    children: [
      { type: "text", text: "Total: " },
      paragraphImage({ alt: "Range image" }),
      { type: "text", text: "", noteReference: { kind: "footnote", id: 1 } },
      { type: "text", text: "amount" },
    ],
  });

  for (const [label, mutate] of [
    [
      "style",
      (paragraph) =>
        core.mutateParagraphTextStyleInRange(paragraph, 0, 7, (style) => ({
          ...style,
          bold: true,
        })),
    ],
    [
      "link",
      (paragraph) =>
        core.mutateParagraphLinkInRange(paragraph, 0, 5, "https://example.com"),
    ],
  ]) {
    const paragraph = buildParagraph();
    assert.equal(mutate(paragraph), true, `${label} edit reports a change`);

    const noteRuns = paragraph.children.filter((child) => child.noteReference);
    assert.equal(noteRuns.length, 1, `${label} edit keeps the footnote reference once`);
    assert.deepEqual(noteRuns[0].noteReference, { kind: "footnote", id: 1 });

    const image = paragraph.children.find((child) => child.type === "image");
    assert.ok(image, `${label} edit keeps the image run`);
    assert.deepEqual(image.crop, paragraphImage().crop, `${label} edit keeps the crop`);
    assert.equal(image.sourceXml, paragraphImage().sourceXml, `${label} edit keeps the drawing XML`);
    assert.deepEqual([...image.data], [9, 8, 7]);

    assert.deepEqual(
      paragraph.commentRanges,
      [{ commentId: 2, start: 0, end: 7, anchorText: "Total: " }],
      `${label} edit keeps the comment anchor`,
    );
  }
});

test("splitting a paragraph keeps its image runs intact", () => {
  const paragraph = {
    type: "paragraph",
    children: [
      { type: "text", text: "Before " },
      paragraphImage({ alt: "Split image" }),
      { type: "text", text: " after" },
      { type: "text", text: "", noteReference: { kind: "endnote", id: 2 } },
    ],
  };

  const { beforeChildren, afterChildren } = core.splitParagraphChildrenAtTextOffsets(
    paragraph,
    "Before split after",
    6,
    6,
  );
  const children = [...beforeChildren, ...afterChildren];
  const image = children.find((child) => child.type === "image");

  assert.ok(image, "the image survives the split");
  assert.deepEqual(image.crop, paragraphImage().crop);
  assert.equal(image.sourceXml, paragraphImage().sourceXml);
  assert.deepEqual([...image.data], [9, 8, 7]);
  assert.equal(
    children.filter((child) => child.noteReference).length,
    1,
    "the note reference survives on exactly one side of the split",
  );
});

test("copy and paste do not inherit the source paragraph's comment anchors", async () => {
  const { model, commentedIndex } = await sampleModelWithComments();
  const originalAnchors = model.nodes[commentedIndex].commentRanges;

  const duplicated = core.duplicateParagraph(model, commentedIndex);
  const duplicate = duplicated.nodes[commentedIndex + 1];

  assert.equal(duplicate.commentRanges, undefined, "the copy carries no comment anchor");
  assert.equal(
    core.resolveParagraphCommentMarkup(duplicate),
    undefined,
    "the copy does not render the original comment",
  );
  assert.deepEqual(
    duplicated.nodes[commentedIndex].commentRanges,
    originalAnchors,
    "the original keeps its anchor",
  );

  const pasted = core.pasteParagraphs(model, model.nodes.length, [model.nodes[commentedIndex]]);
  assert.equal(
    pasted.nodes[pasted.nodes.length - 1].commentRanges,
    undefined,
    "the pasted copy carries no comment anchor",
  );
});

test("duplicating or pasting a commented paragraph exports one comment anchor, not two", async () => {
  const { model, package: pkg, commentedIndex } = await sampleModelWithComments();
  const originalText = core.paragraphText(model.nodes[commentedIndex]);

  for (const [label, edited] of [
    [
      "duplicate",
      core.updateParagraphText(
        core.duplicateParagraph(model, commentedIndex),
        commentedIndex + 1,
        `${originalText} (copy)`,
      ),
    ],
    [
      "paste",
      core.updateParagraphText(
        core.pasteParagraphs(model, model.nodes.length, [model.nodes[commentedIndex]]),
        model.nodes.length,
        `${originalText} (pasted)`,
      ),
    ],
  ]) {
    const documentXml = documentXmlFromExport(await core.serializeDocx(edited, pkg));
    const commentIds = collectAttributeValues(
      documentXml,
      /<w:commentRangeStart w:id="(\d+)"/g,
    );
    assert.deepEqual(commentIds, ["0"], `${label} exports the comment anchor once`);
    assert.deepEqual(
      collectAttributeValues(documentXml, /<w:commentRangeEnd w:id="(\d+)"/g),
      ["0"],
      `${label} exports the range end once`,
    );
    assert.deepEqual(
      collectAttributeValues(documentXml, /<w:commentReference w:id="(\d+)"/g),
      ["0"],
      `${label} exports the reference once`,
    );

    // Re-importing keeps the anchor on the original paragraph only.
    const reimported = await core.buildDocModelFromBytes(
      await core.serializeDocx(edited, pkg),
    );
    const rangesByParagraph = reimported.model.nodes
      .filter((node) => node.type === "paragraph")
      .map((node) => node.commentRanges?.length ?? 0);
    assert.deepEqual(
      rangesByParagraph,
      [0, 1, 0, 0, 0],
      `${label} leaves exactly one anchored paragraph`,
    );
  }
});

test("a rebuilt paragraph keeps its comment range and note marker in the export", async () => {
  const { model, package: pkg } = await sampleModelWithComments();
  const anchoredIndex = model.nodes.findIndex(
    (node) =>
      node.type === "paragraph" &&
      node.children.some((child) => child.type === "text" && child.noteReference),
  );
  assert.ok(anchoredIndex >= 0, "the fixture carries a note reference");

  // Append instead of replacing the commented words: the range is remapped onto
  // the surviving text, which is what keeps it exportable.
  const edited = core.updateParagraphText(
    model,
    anchoredIndex,
    "Monthly operational report (edited)",
  );
  assert.equal(edited.nodes[anchoredIndex].sourceXml, undefined);
  assert.deepEqual(
    noteReferencesOf(edited.nodes[anchoredIndex]),
    [{ kind: "footnote", id: 1 }],
  );

  const bytes = await core.serializeDocx(edited, pkg);
  const documentXml = documentXmlFromExport(bytes);
  assert.match(documentXml, /<w:commentRangeStart w:id="0"\/>/, "the comment range is exported");
  assert.match(documentXml, /<w:commentRangeEnd w:id="0"\/>/);
  assert.match(documentXml, /Monthly operational report \(edited\)/, "the edited text is exported");
  assert.match(
    documentXml,
    /<w:footnoteReference w:id="1"\/>/,
    "the note marker is re-emitted into the rebuilt paragraph",
  );

  // The reference and the note body both come back on re-import.
  const reimported = await core.buildDocModelFromBytes(bytes);
  assert.deepEqual(
    noteReferencesOf(reimported.model.nodes[anchoredIndex]),
    [{ kind: "footnote", id: 1 }],
  );
  assert.deepEqual(
    reimported.model.metadata.footnotes?.map((note) => note.text),
    model.metadata.footnotes?.map((note) => note.text),
  );
});

// ── Issue proposal: import → change one paragraph → export → re-import ────
test("editing one paragraph leaves every other node field-identical after a round trip", async () => {
  for (const sample of ["review-comments.docx", "report-with-image.docx", "legal-contract.docx"]) {
    const { model, package: pkg } = await core.buildDocModelFromBytes(
      arrayBufferFromFile(sample),
    );
    const editedText = `${core.paragraphText(model.nodes[1])} (edited)`;
    const edited = core.updateParagraphText(model, 1, editedText);
    const reimported = await core.buildDocModelFromBytes(
      await core.serializeDocx(edited, pkg),
    );

    assert.equal(reimported.model.nodes.length, model.nodes.length, sample);
    for (let index = 0; index < model.nodes.length; index += 1) {
      if (index === 1) {
        continue;
      }
      assert.equal(
        jsonOf(reimported.model.nodes[index]),
        jsonOf(model.nodes[index]),
        `${sample}: node ${index} must round-trip field-identical (XML, note ids, comment ranges, crops)`,
      );
    }

    // The edited paragraph keeps the anchors that live on the model, and the
    // serialization patch re-emits them into the exported file.
    assert.match(
      core.paragraphText(reimported.model.nodes[1]),
      /\(edited\)$/,
      `${sample}: the edit reaches the exported file`,
    );
    assert.deepEqual(
      noteReferencesOf(edited.nodes[1]),
      noteReferencesOf(model.nodes[1]),
      `${sample}: the edit keeps the paragraph's note references in the model`,
    );
    assert.deepEqual(
      noteReferencesOf(reimported.model.nodes[1]),
      noteReferencesOf(model.nodes[1]),
      `${sample}: the re-imported paragraph still carries its note references`,
    );
  }
});

test("a global replace leaves comment ranges, revisions and note anchors in the export", async () => {
  const imported = await core.importDocxBuffer(arrayBufferFromFile("review-comments.docx"), {
    useWorker: false,
  });
  const model = imported.model;
  const edited = core.replaceText(model, "term-that-does-not-exist", "replacement");

  assert.equal(jsonOf(edited), jsonOf(model), "a replace that matches nothing changes nothing");

  const documentXml = documentXmlFromExport(await core.wasmSerializeDocx(edited, imported.package));
  assert.match(documentXml, /<w:commentRangeStart\b/, "comment ranges survive the replace");
  assert.match(documentXml, /<w:commentReference\b/, "comment references survive the replace");
  assert.match(documentXml, /<w:footnoteReference\b/, "footnote anchors survive the replace");
  assert.match(documentXml, /<w:endnoteReference\b/, "endnote anchors survive the replace");
  assert.match(documentXml, /<w:ins\b/, "insertions survive the replace");
  assert.match(documentXml, /<w:del\b/, "deletions survive the replace");

  const reimported = await core.wasmBuildDocModelFromBytes(await core.wasmSerializeDocx(edited, imported.package));
  const noteReferences = reimported.model.nodes.flatMap((node) =>
    node.type === "paragraph"
      ? node.children.filter((child) => child.type === "text" && child.noteReference).map((child) => child.noteReference)
      : [],
  );
  assert.deepEqual(noteReferences, [
    { kind: "footnote", id: 1 },
    { kind: "endnote", id: 1 },
  ]);
});

test("editing a paragraph with a footnote does not drop its reference", async () => {
  const imported = await core.buildDocModelFromBytes(
    arrayBufferFromFile("review-comments.docx"),
  );

  const edited = core.updateParagraphText(imported.model, 1, "Monthly operational report edited");
  const editedParagraph = edited.nodes[1];
  const noteRun = editedParagraph.children.find(
    (child) => child.type === "text" && child.noteReference,
  );

  assert.ok(noteRun, "the edited paragraph keeps its footnote run");
  assert.deepEqual(noteRun.noteReference, { kind: "footnote", id: 1 });
  assert.equal(noteRun.text, "");
  assert.equal(editedParagraph.sourceXml, undefined, "the edited paragraph is re-serialized from the model");

  // Untouched neighbours keep their original XML, so their anchors survive export.
  assert.ok(edited.nodes[2].sourceXml, "the next paragraph keeps its source XML");
  const bytes = await core.serializeDocx(edited, imported.package);
  const documentXml = documentXmlFromExport(bytes);
  assert.match(documentXml, /<w:footnoteReference w:id="1"\/>/, "the edited paragraph's anchor is exported");
  assert.match(documentXml, /<w:endnoteReference\b/);
  assert.match(documentXml, /<w:ins\b/);

  const reimported = await core.buildDocModelFromBytes(bytes);
  assert.deepEqual(
    core.paragraphNoteReferences(reimported.model.nodes[1]),
    [{ kind: "footnote", id: 1 }],
    "the anchor survives export + re-import",
  );
  assert.ok(
    reimported.model.metadata.footnotes?.length > 0,
    "the note body survives export + re-import",
  );
});

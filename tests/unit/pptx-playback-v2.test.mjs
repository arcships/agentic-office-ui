import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { createTsModuleLoader } from "./ts-source-loader.mjs";

const rootDir = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const {
  DOMParser: XmlDomParser,
  Document: XmlDocument,
  Node: XmlNode,
} = requireFromXlsxCore("@xmldom/xmldom");
const { strFromU8, unzipSync } = requireFromXlsxCore("fflate");

globalThis.DOMParser ??= XmlDomParser;
globalThis.Document ??= XmlDocument;
globalThis.Node ??= XmlNode;
if (!XmlDocument.prototype.querySelector) {
  XmlDocument.prototype.querySelector = function querySelector() {
    return null;
  };
}

const typescript = requireFromXlsxCore("typescript");
const loadSource = createTsModuleLoader(
  new URL("../../packages/pptx-core/src/", import.meta.url),
  typescript,
);
const { parsePptxPlaybackDocument } = await import(loadSource("browser/playback-parser.ts"));
const {
  compilePptxSlideTracks,
  evaluatePptxTriggerGroup,
} = await import(loadSource("playback/track-compiler.ts"));
const {
  pptxEffectTargetKey,
} = await import(loadSource("playback/targets.ts"));
const {
  isPptxTransitionExecutable,
  pptxTransitionSoundMediaItem,
} = await import(loadSource("playback/transitions.ts"));

const NS = {
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
  dgm: "http://schemas.openxmlformats.org/drawingml/2006/diagram",
};

function slideXml({ shapes = "", timing = "", transition = "" }) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="${NS.p}" xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:c="${NS.c}" xmlns:dgm="${NS.dgm}">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr><p:grpSpPr/>
    ${shapes}
  </p:spTree></p:cSld>${transition}${timing ? `<p:timing><p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot"><p:childTnLst>${timing}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>` : ""}
</p:sld>`;
}

function textShape(id, name, paragraphs) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs.map((text) => `<a:p><a:r><a:t>${text}</a:t></a:r></a:p>`).join("")}</p:txBody></p:sp>`;
}

function tableShape(id) {
  const cell = (text) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`;
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblGrid><a:gridCol w="50"/><a:gridCol w="50"/></a:tblGrid><a:tr h="50">${cell("R0C0")}${cell("R0C1")}</a:tr><a:tr h="50">${cell("R1C0")}${cell("R1C1")}</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

function chartShape(id) {
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame>`;
}

function smartArtShape(id) {
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="dgm"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4"/></a:graphicData></a:graphic></p:graphicFrame>`;
}

let nextId = 2;
function effectXml({ targetId, inner = "", presetClass = "entr", filter = "fade", transition = "in", nodeType = "clickEffect", duration = 300 }) {
  const id = nextId;
  nextId += 3;
  return `<p:par><p:cTn id="${id}" presetID="10" presetClass="${presetClass}" nodeType="${nodeType}"><p:childTnLst><p:par><p:cTn id="${id + 1}" fill="hold"><p:childTnLst><p:animEffect transition="${transition}" filter="${filter}"><p:cBhvr><p:cTn id="${id + 2}" dur="${duration}"/><p:tgtEl><p:spTgt spid="${targetId}">${inner}</p:spTgt></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;
}

function parseSlides(slides, options = {}) {
  const files = {
    slides: new Map(slides.map((xml, index) => [`ppt/slides/slide${index + 1}.xml`, xml])),
    slideRels: new Map(slides.map((_, index) => [
      `ppt/slides/_rels/slide${index + 1}.xml.rels`,
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
    ])),
  };
  const presentation = {
    width: 9144000,
    height: 5143500,
    slides: slides.map((_, index) => ({
      index,
      slidePath: `ppt/slides/slide${index + 1}.xml`,
      hidden: false,
      rels: options.slideRels?.[index] ?? new Map(),
    })),
  };
  return parsePptxPlaybackDocument(files, presentation, options);
}

function features(report) {
  return report.features.map((feature) => `${feature.feature}=${feature.disposition}`);
}

test("PPTX by-word and by-character ranges compile to text targets", () => {
  nextId = 2;
  const timing =
    effectXml({ targetId: 2, inner: '<p:txEl><p:charRg st="0" end="4"/></p:txEl>' })
    + effectXml({ targetId: 2, inner: '<p:txEl><p:charRg st="6" end="8"/></p:txEl>' });
  const document = parseSlides([
    slideXml({
      shapes: textShape(2, "title", ["Hello world", "Second line"]),
      timing,
    }),
  ]);
  const slide = document.slides[0];
  const effects = Object.values(slide.nodes).flatMap((node) => node.effect ? [node.effect] : []);
  assert.deepEqual(
    effects.map((effect) => [effect.characterRange, effect.characterGranularity]),
    [
      [{ start: 0, end: 4 }, "word"],
      [{ start: 6, end: 8 }, "character"],
    ],
  );
  assert.ok(features(slide.capability).includes("text:by-word=strict"));
  assert.ok(features(slide.capability).includes("text:by-character=strict"));
  assert.deepEqual(
    features(slide.capability).filter((feature) => feature.startsWith("text:by")),
    ["text:by-word=strict", "text:by-character=strict"],
  );

  const compiled = compilePptxSlideTracks(slide);
  const tracks = compiled.groups.flatMap((group) => group.tracks);
  assert.deepEqual(
    tracks.map((track) => [track.target, track.property]),
    [
      [`${slide.objects[0].key}#text:0-4`, "display"],
      [`${slide.objects[0].key}#text:0-4`, "opacity"],
      [`${slide.objects[0].key}#text:6-8`, "display"],
      [`${slide.objects[0].key}#text:6-8`, "opacity"],
    ],
  );
  assert.equal(compiled.initialState[`${slide.objects[0].key}#text:0-4`].display, false);
  const firstGroup = compiled.groups.find((group) => group.key === "click:1");
  assert.ok(firstGroup);
  const start = evaluatePptxTriggerGroup(firstGroup, 0);
  assert.deepEqual(start[`${slide.objects[0].key}#text:0-4`], { display: true, opacity: 0 });
  const finished = evaluatePptxTriggerGroup(firstGroup, firstGroup.durationMs);
  assert.deepEqual(finished[`${slide.objects[0].key}#text:0-4`], { display: true, opacity: 1 });
});

test("PPTX character ranges without a text body stay reported and static", () => {
  nextId = 2;
  const timing = effectXml({ targetId: 4, inner: '<p:txEl><p:charRg st="0" end="1"/></p:txEl>' });
  const document = parseSlides([
    slideXml({
      shapes: '<p:pic><p:nvPicPr><p:cNvPr id="4" name="image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic>',
      timing,
    }),
  ]);
  const slide = document.slides[0];
  const effect = Object.values(slide.nodes).find((node) => node.effect)?.effect;
  assert.equal(effect.characterRangeUnavailable, true);
  assert.ok(features(slide.capability).includes("text:character-range-unavailable=static"));
  const compiled = compilePptxSlideTracks(slide);
  assert.equal(compiled.groups.flatMap((group) => group.tracks).length, 0);
});

test("PPTX table paragraph ranges address row-major cells", () => {
  nextId = 2;
  const timing =
    effectXml({ targetId: 9, inner: '<p:txEl><p:pRg st="1" end="1"/></p:txEl>' })
    + effectXml({ targetId: 9, inner: '<p:txEl><p:pRg st="3" end="3"/></p:txEl>' });
  const bldList = `<p:bldGraphic spid="9" grpId="0"><p:bldSub bld="cell" grpId="0"/></p:bldGraphic>`;
  const document = parseSlides([
    slideXml({
      shapes: tableShape(9),
      timing: timing.replace("</p:childTnLst></p:cTn></p:par></p:tnLst>", "</p:childTnLst></p:cTn></p:par></p:tnLst>"),
    }).replace("<p:timing>", `<p:timing><p:bldLst>${bldList}</p:bldLst>`),
  ]);
  const slide = document.slides[0];
  const effects = Object.values(slide.nodes).flatMap((node) => node.effect ? [node.effect] : []);
  assert.deepEqual(
    effects.map((effect) => effect.tableCell),
    [{ row: 0, column: 1 }, { row: 1, column: 1 }],
  );
  assert.equal(features(slide.capability).filter((feature) => feature === "table:cell=strict").length, 2);
  assert.ok(features(slide.capability).includes("table-build:cell=static"));
  const compiled = compilePptxSlideTracks(slide);
  assert.deepEqual(
    compiled.groups.flatMap((group) => group.tracks.map((track) => track.target)),
    [
      `${slide.objects[0].key}#cell:0:1`,
      `${slide.objects[0].key}#cell:0:1`,
      `${slide.objects[0].key}#cell:1:1`,
      `${slide.objects[0].key}#cell:1:1`,
    ],
  );
});

test("PPTX chart element targets separate whole-chart and series effects", () => {
  nextId = 2;
  const timing =
    effectXml({ targetId: 7, inner: '<p:oleChartEl type="chart"/>' })
    + effectXml({ targetId: 7, inner: '<p:oleChartEl type="ser" lvl="1" idx="1"/>' })
    + effectXml({ targetId: 7, inner: '<p:oleChartEl type="cat" idx="0"/>' })
    + effectXml({ targetId: 7, inner: '<p:oleChartEl type="ptInSer" idx="2"/>' });
  const document = parseSlides([slideXml({ shapes: chartShape(7), timing })]);
  const slide = document.slides[0];
  const effects = Object.values(slide.nodes).flatMap((node) => node.effect ? [node.effect] : []);
  assert.deepEqual(
    effects.map((effect) => [effect.chartElement.type, effect.chartElement.index, effect.chartElement.level]),
    [
      ["chart", undefined, undefined],
      ["series", 1, 1],
      ["category", 0, undefined],
      ["point-in-series", 2, undefined],
    ],
  );
  assert.ok(features(slide.capability).includes("chart:element=strict"));
  assert.ok(features(slide.capability).includes("chart:series=static"));
  assert.ok(features(slide.capability).includes("chart:category=static"));
  assert.ok(features(slide.capability).includes("chart:point=static"));
  const wholeSlot = effects[0];
  const seriesSlot = effects[1];
  assert.equal(pptxEffectTargetKey(wholeSlot), `${slide.objects[0].key}#chart:chart`);
  assert.equal(pptxEffectTargetKey(seriesSlot), `${slide.objects[0].key}#chart:series:1@1`);
  const compiled = compilePptxSlideTracks(slide);
  assert.deepEqual(
    [...new Set(compiled.groups.flatMap((group) => group.tracks.map((track) => track.target)))],
    [`${slide.objects[0].key}#chart:chart`],
  );
});

test("PPTX SmartArt node identities resolve while level-only targets are reported", () => {
  nextId = 2;
  const timing =
    effectXml({ targetId: 8, inner: '<p:graphicEl type="dgm" id="12"/>' })
    + effectXml({ targetId: 8, inner: '<p:graphicEl type="dgm" lvl="2"/>' })
    + effectXml({ targetId: 8, inner: '<p:subSp spid="13"/>' });
  const document = parseSlides([slideXml({ shapes: smartArtShape(8), timing })]);
  const slide = document.slides[0];
  const effects = Object.values(slide.nodes).flatMap((node) => node.effect ? [node.effect] : []);
  assert.equal(slide.objects[0].nodeType, "smartart");
  assert.equal(
    effects[0].graphicElement.objectKey,
    `${slide.objects[0].slidePath}|slide|group:8/shape:12`,
  );
  assert.equal(effects[1].graphicElement.objectKey, undefined);
  assert.equal(
    effects[2].graphicElement.objectKey,
    `${slide.objects[0].slidePath}|slide|group:8/shape:13`,
  );
  assert.ok(features(slide.capability).includes("smartart:node=strict"));
  assert.ok(features(slide.capability).includes("smartart:level=static"));
  const compiled = compilePptxSlideTracks(slide);
  assert.deepEqual(
    [...new Set(compiled.groups.flatMap((group) => group.tracks.map((track) => track.target)))].sort(),
    [
      `${slide.objects[0].slidePath}|slide|group:8/shape:12`,
      `${slide.objects[0].slidePath}|slide|group:8/shape:13`,
    ].sort(),
  );
});

test("PPTX transition sounds reuse media gating and 3D transitions stay reported", () => {
  const externalRelations = new Map([
    ["rId9", { target: "https://example.com/applause.wav", targetMode: "External", type: "audio" }],
  ]);
  const embeddedRelations = new Map([
    ["rId9", { target: "../media/applause.wav", targetMode: undefined, type: "audio" }],
  ]);
  const soundTransition = `<p:transition spd="fast"><p:fade/><p:sndAc><p:stSnd><p:snd r:embed="rId9"/></p:stSnd></p:sndAc></p:transition>`;
  const embedded = parseSlides([slideXml({ transition: soundTransition })], {
    slideRels: { 0: embeddedRelations },
  }).slides[0];
  assert.deepEqual(embedded.transition.sound, {
    relationId: "rId9",
    sourcePath: "ppt/media/applause.wav",
    contentType: "audio/wav",
    embedded: true,
    loop: false,
    volume: 1,
  });
  assert.ok(features(embedded.capability).includes("transition:sound=strict"));
  const soundItem = pptxTransitionSoundMediaItem(embedded.index, embedded.transition.sound);
  assert.deepEqual(
    [soundItem.id, soundItem.kind, soundItem.embedded, soundItem.autoplay],
    ["transition-sound:0", "audio", true, true],
  );

  const external = parseSlides([slideXml({ transition: soundTransition })], {
    slideRels: { 0: externalRelations },
  }).slides[0];
  assert.equal(external.transition.sound.embedded, false);
  assert.ok(features(external.capability).includes("transition:sound=static"));

  const externalAllowed = parseSlides([slideXml({ transition: soundTransition })], {
    slideRels: { 0: externalRelations },
    externalMedia: "allowed",
  }).slides[0];
  assert.ok(features(externalAllowed.capability).includes("transition:sound=strict"));

  const cube = parseSlides([
    slideXml({ transition: '<p:transition spd="slow"><p:cube dir="l"/></p:transition>' }),
  ]).slides[0];
  assert.ok(features(cube.capability).includes("transition:3d:cube=static"));
  assert.equal(isPptxTransitionExecutable("cube", "safe"), false);
  assert.equal(isPptxTransitionExecutable("fade", "off"), true);
  assert.equal(isPptxTransitionExecutable("morph", "safe"), true);
  assert.equal(isPptxTransitionExecutable("morph", "off"), false);
});

const FIXTURES = resolve(rootDir, "tests/fixtures/pptx/playback");

function xmlEntries(bytes) {
  const entries = unzipSync(bytes);
  return new Map(Object.entries(entries).map(([name, data]) => [name, strFromU8(data)]));
}

function relationships(xml) {
  const document = new XmlDomParser().parseFromString(xml, "application/xml");
  const map = new Map();
  for (const relation of [...document.getElementsByTagNameNS("*", "Relationship")]) {
    map.set(relation.getAttribute("Id"), {
      target: relation.getAttribute("Target"),
      targetMode: relation.getAttribute("TargetMode") || undefined,
      type: relation.getAttribute("Type") || undefined,
    });
  }
  return map;
}

function loadDeck(file) {
  const entries = xmlEntries(readFileSync(file));
  const presentationXml = entries.get("ppt/presentation.xml");
  const presentationDocument = new XmlDomParser().parseFromString(presentationXml, "application/xml");
  const relationshipsById = relationships(entries.get("ppt/_rels/presentation.xml.rels"));
  const size = [...presentationDocument.getElementsByTagNameNS("*", "sldSz")][0];
  const slides = [...presentationDocument.getElementsByTagNameNS("*", "sldId")].map((item, index) => {
    const relationId = item.getAttributeNS(NS.r, "id") ?? item.getAttribute("r:id");
    const relation = relationshipsById.get(relationId);
    if (!relation) throw new Error(`missing presentation relationship ${relationId}`);
    const slidePath = `ppt/${relation.target.replace(/^\.\//u, "")}`;
    const relsPath = slidePath.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
    return {
      index,
      slidePath,
      hidden: false,
      rels: entries.has(relsPath) ? relationships(entries.get(relsPath)) : new Map(),
    };
  });
  const files = {
    slides: new Map(slides.map((slide) => [slide.slidePath, entries.get(slide.slidePath)])),
    slideRels: new Map(slides.map((slide) => [
      slide.slidePath.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels",
      entries.get(slide.slidePath.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels") ?? "",
    ])),
  };
  return {
    entries,
    files,
    presentation: {
      width: Number(size.getAttribute("cx")),
      height: Number(size.getAttribute("cy")),
      slides,
    },
  };
}

test("PPTX controlled playback-v2 fixtures keep their XML branches and hashes", () => {
  const manifest = JSON.parse(readFileSync(resolve(FIXTURES, "manifest.json"), "utf8"));
  const controlled = manifest.entries.filter((entry) => entry.status === "controlled");
  assert.deepEqual(
    controlled.map((entry) => entry.file).sort(),
    ["chart-series.pptx", "table-cells.pptx", "text-by-word.pptx", "transition-sound.pptx"],
  );
  for (const entry of controlled) {
    const bytes = readFileSync(resolve(FIXTURES, entry.file));
    assert.equal(
      createHash("sha256").update(bytes).digest("hex"),
      entry.sha256,
      `${entry.file} hash`,
    );
    const deck = loadDeck(resolve(FIXTURES, entry.file));
    const document = parsePptxPlaybackDocument(deck.files, deck.presentation);
    assert.equal(document.slides.length, entry.slides, `${entry.file} slide count`);
    const focus = document.slides[entry.focusSlideIndex];
    assert.ok(focus, `${entry.file} focus slide`);
    for (const expectation of entry.expectations) {
      const count = focus.capability.features.filter((feature) =>
        feature.feature === expectation.feature
        && feature.disposition === expectation.disposition
      ).length;
      assert.equal(count, expectation.count, `${entry.file} ${expectation.feature}`);
    }
    for (const check of entry.otherSlideExpectations ?? []) {
      const slide = document.slides[check.slideIndex];
      assert.ok(slide, `${entry.file} slide ${check.slideIndex}`);
      for (const expectation of check.expectations) {
        const count = slide.capability.features.filter((feature) =>
          feature.feature === expectation.feature
          && feature.disposition === expectation.disposition
        ).length;
        assert.equal(count, expectation.count, `${entry.file} slide ${check.slideIndex} ${expectation.feature}`);
      }
    }
    const xml = [...deck.entries.entries()]
      .filter(([name]) => name.endsWith(".xml") || name.endsWith(".rels"))
      .map(([, content]) => content)
      .join("");
    for (const marker of entry.xmlMarkers) {
      assert.ok(xml.includes(marker), `${entry.file} missing ${marker}`);
    }
  }
});

#!/usr/bin/env node
/**
 * Builds the controlled playback-v2 decks from the PowerPoint-saved
 * `text-by-paragraph.pptx` fixture.
 *
 * These decks are not PowerPoint output. They are deterministic derivatives
 * that exercise XML branches the formal fixture set does not cover yet:
 * `p:charRg` word/character ranges, chart element targets, transition sounds
 * and a 3D transition. Re-save them in desktop PowerPoint before promoting any
 * of them from `controlled` to `ready` in `manifest.json`.
 */
import { createHash } from "node:crypto";
import { readFileSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const requireFromXlsxCore = createRequire(
  path.join(root, "packages", "xlsx-core", "package.json"),
);
const { unzipSync, zipSync, strToU8, strFromU8 } = requireFromXlsxCore("fflate");

const FIXTURES = path.join(root, "tests/fixtures/pptx/playback");
const BASE = path.join(FIXTURES, "text-by-paragraph.pptx");
const ZIP_TIME = new Date("2026-11-18T00:00:00Z");
const CHART_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

function readZip(file) {
  const entries = unzipSync(readFileSync(file));
  return new Map(Object.entries(entries).map(([name, data]) => [name, Buffer.from(data)]));
}

function text(entries, name) {
  const data = entries.get(name);
  if (!data) throw new Error(`missing archive entry ${name}`);
  return strFromU8(data);
}

function setText(entries, name, value) {
  entries.set(name, Buffer.from(strToU8(value)));
}

function writeZip(entries, file) {
  const archive = {};
  for (const name of [...entries.keys()].sort()) {
    archive[name] = [entries.get(name), { level: 9, mtime: ZIP_TIME }];
  }
  const bytes = zipSync(archive, { level: 9, mtime: ZIP_TIME });
  writeFileSync(file, bytes);
  return createHash("sha256").update(bytes).digest("hex");
}

function addContentType(entries, partName) {
  const current = text(entries, "[Content_Types].xml");
  if (current.includes(`PartName="${partName}"`)) return current;
  return current.replace(
    "</Types>",
    `<Override PartName="${partName}" ContentType="${CHART_CONTENT_TYPE}"/></Types>`,
  );
}

function addRelationship(xml, relationship) {
  return xml.replace("</Relationships>", `${relationship}</Relationships>`);
}

const byWordSlide = `<p:timing><p:tnLst><p:par>
  <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
    <p:par><p:cTn id="2" nodeType="clickEffect"><p:childTnLst><p:par>
      <p:cTn id="3" presetID="10" presetClass="entr" fill="hold" nodeType="withEffect"><p:childTnLst>
        <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="4" dur="300"/>
          <p:tgtEl><p:spTgt spid="3"><p:txEl><p:charRg st="0" end="2"/></p:txEl></p:spTgt></p:tgtEl>
        </p:cBhvr></p:animEffect>
      </p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>
    <p:par><p:cTn id="5" nodeType="clickEffect"><p:childTnLst><p:par>
      <p:cTn id="6" presetID="10" presetClass="entr" fill="hold" nodeType="withEffect"><p:childTnLst>
        <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="7" dur="300"/>
          <p:tgtEl><p:spTgt spid="3"><p:txEl><p:charRg st="4" end="6"/></p:txEl></p:spTgt></p:tgtEl>
        </p:cBhvr></p:animEffect>
      </p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>
    <p:par><p:cTn id="8" nodeType="clickEffect"><p:childTnLst><p:par>
      <p:cTn id="9" presetID="10" presetClass="entr" fill="hold" nodeType="withEffect"><p:childTnLst>
        <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="10" dur="300"/>
          <p:tgtEl><p:spTgt spid="3"><p:txEl><p:charRg st="8" end="10"/></p:txEl></p:spTgt></p:tgtEl>
        </p:cBhvr></p:animEffect>
      </p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>
  </p:childTnLst></p:cTn>
</p:par></p:tnLst><p:bldLst><p:bldP spid="3" grpId="0" build="p" bldLvl="1"/></p:bldLst></p:timing>`;

const chartTiming = `<p:timing><p:tnLst><p:par>
  <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
    <p:par><p:cTn id="2" nodeType="clickEffect"><p:childTnLst><p:par>
      <p:cTn id="3" presetID="10" presetClass="entr" fill="hold" nodeType="withEffect"><p:childTnLst>
        <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="4" dur="400"/>
          <p:tgtEl><p:spTgt spid="5"><p:oleChartEl type="chart"/></p:spTgt></p:tgtEl>
        </p:cBhvr></p:animEffect>
      </p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>
    <p:par><p:cTn id="5" nodeType="clickEffect"><p:childTnLst><p:par>
      <p:cTn id="6" presetID="10" presetClass="entr" fill="hold" nodeType="withEffect"><p:childTnLst>
        <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="7" dur="400"/>
          <p:tgtEl><p:spTgt spid="5"><p:oleChartEl type="ser" lvl="1" idx="1"/></p:spTgt></p:tgtEl>
        </p:cBhvr></p:animEffect>
      </p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>
  </p:childTnLst></p:cTn>
</p:par></p:tnLst><p:bldLst><p:bldGraphic spid="5" grpId="0"><p:bldSub bld="series" grpId="0"/></p:bldGraphic></p:bldLst></p:timing>`;

const chartFrame = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="chart-series"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="4572000" y="1371600"/><a:ext cx="3657600" cy="2286000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart1"/></a:graphicData></a:graphic></p:graphicFrame>`;

const tableFrame = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="table-cells"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="4572000" y="1371600"/><a:ext cx="3657600" cy="2286000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1"/><a:tblGrid><a:gridCol w="1828800"/><a:gridCol w="1828800"/></a:tblGrid><a:tr h="1143000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>R0C0</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>R0C1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr><a:tr h="1143000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>R1C0</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>R1C1</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;

const tableTiming = `<p:timing><p:tnLst><p:par>
  <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
    <p:par><p:cTn id="2" nodeType="clickEffect"><p:childTnLst><p:par>
      <p:cTn id="3" presetID="10" presetClass="entr" fill="hold" nodeType="withEffect"><p:childTnLst>
        <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="4" dur="300"/>
          <p:tgtEl><p:spTgt spid="6"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt></p:tgtEl>
        </p:cBhvr></p:animEffect>
      </p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>
    <p:par><p:cTn id="5" nodeType="clickEffect"><p:childTnLst><p:par>
      <p:cTn id="6" presetID="10" presetClass="entr" fill="hold" nodeType="withEffect"><p:childTnLst>
        <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="7" dur="300"/>
          <p:tgtEl><p:spTgt spid="6"><p:txEl><p:pRg st="3" end="3"/></p:txEl></p:spTgt></p:tgtEl>
        </p:cBhvr></p:animEffect>
      </p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>
  </p:childTnLst></p:cTn>
</p:par></p:tnLst><p:bldLst><p:bldGraphic spid="6" grpId="0"><p:bldSub bld="cell" grpId="0"/></p:bldGraphic></p:bldLst></p:timing>`;

const chartPart = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
    <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>系列一</c:v></c:pt></c:strCache></c:strRef></c:tx><c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
    <c:ser><c:idx val="1"/><c:order val="1"/><c:tx><c:strRef><c:f>Sheet1!$C$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>系列二</c:v></c:pt></c:strCache></c:strRef></c:tx><c:val><c:numRef><c:f>Sheet1!$C$2:$C$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
    <c:axId val="111111111"/><c:axId val="222222222"/>
  </c:barChart>
  <c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>
  <c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>
  </c:plotArea></c:chart>
</c:chartSpace>`;

function byWord() {
  const entries = readZip(BASE);
  const slide = text(entries, "ppt/slides/slide12.xml")
    .replace(/<p:timing>[\s\S]*<\/p:timing>/u, byWordSlide);
  setText(entries, "ppt/slides/slide12.xml", slide);
  return entries;
}

function chartSeries() {
  const entries = readZip(BASE);
  const slide = text(entries, "ppt/slides/slide12.xml")
    .replace(/<p:timing>[\s\S]*<\/p:timing>/u, chartTiming)
    .replace("</p:spTree>", `${chartFrame}</p:spTree>`);
  setText(entries, "ppt/slides/slide12.xml", slide);
  const rels = addRelationship(
    text(entries, "ppt/slides/_rels/slide12.xml.rels"),
    '<Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>',
  );
  setText(entries, "ppt/slides/_rels/slide12.xml.rels", rels);
  setText(entries, "ppt/charts/chart1.xml", chartPart);
  setText(entries, "[Content_Types].xml", addContentType(entries, "/ppt/charts/chart1.xml"));
  return entries;
}

function tableCells() {
  const entries = readZip(BASE);
  const slide = text(entries, "ppt/slides/slide12.xml")
    .replace(/<p:timing>[\s\S]*<\/p:timing>/u, tableTiming)
    .replace("</p:spTree>", `${tableFrame}</p:spTree>`);
  setText(entries, "ppt/slides/slide12.xml", slide);
  return entries;
}

function transitionSound() {
  const entries = readZip(BASE);
  const first = text(entries, "ppt/slides/slide1.xml").replace(
    "</p:sld>",
    '<p:transition spd="fast"><p:fade/><p:sndAc><p:stSnd><p:snd r:embed="rIdSound1"/></p:stSnd></p:sndAc></p:transition></p:sld>',
  );
  setText(entries, "ppt/slides/slide1.xml", first);
  setText(
    entries,
    "ppt/slides/_rels/slide1.xml.rels",
    addRelationship(
      text(entries, "ppt/slides/_rels/slide1.xml.rels"),
      '<Relationship Id="rIdSound1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio" Target="../media/transition-sound.wav"/>',
    ),
  );
  const second = text(entries, "ppt/slides/slide2.xml").replace(
    /<p:transition>[\s\S]*?<\/p:transition>/u,
    '<p:transition spd="slow"><p:cube dir="l"/></p:transition>',
  );
  setText(entries, "ppt/slides/slide2.xml", second);
  entries.set("ppt/media/transition-sound.wav", Buffer.from(entries.get("ppt/media/media2.wav")));
  return entries;
}

const outputs = [
  ["text-by-word.pptx", byWord()],
  ["chart-series.pptx", chartSeries()],
  ["table-cells.pptx", tableCells()],
  ["transition-sound.pptx", transitionSound()],
];
for (const [name, entries] of outputs) {
  const file = path.join(FIXTURES, name);
  const sha256 = writeZip(entries, file);
  console.log(`${name} ${sha256}`);
}

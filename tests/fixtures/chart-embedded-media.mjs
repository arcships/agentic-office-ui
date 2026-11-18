/**
 * Fixture builder for chart-embedded media tests.
 *
 * The repository's `charts-images.xlsx` sample already carries a real worksheet
 * chart (`xl/charts/chart1.xml`) inside a real drawing part. Excel-authored
 * chart pictures and `chartUserShapes` overlays are ordinary OOXML parts, so the
 * builder adds them to that sample instead of checking a second binary blob in:
 *
 * - `xl/charts/_rels/chart1.xml.rels` links a media image and a
 *   `chartUserShapes` part (the relationship types Excel writes).
 * - `xl/charts/chartUserShapes1.xml` holds a "source: …" text box and a picture.
 * - `chart1.xml` gains `a:blipFill` picture fills on the chart area, the plot
 *   area, the first series and one data point.
 */
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";

const requireFromXlsxCore = createRequire(
  new URL("../../packages/xlsx-core/package.json", import.meta.url),
);
const { strFromU8, strToU8, unzipSync, zipSync } = requireFromXlsxCore("fflate");

export const CHART_MEDIA_SAMPLE_PATH = new URL(
  "../../apps/demo/public/samples/charts-images.xlsx",
  import.meta.url,
);

export const CHART_USER_SHAPES_PATH = "xl/charts/chartUserShapes1.xml";
export const CHART_SHAPES_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/chartUserShapes";
export const CHART_USER_SHAPES_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml";
export const CHART_SOURCE_NOTE_TEXT = "source: Internal revenue dashboard";

const CHART_SHAPES_XML = `<cdr:wsDr xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cdr:twoCellAnchor><cdr:from><cdr:x>340000</cdr:x><cdr:y>1900000</cdr:y></cdr:from><cdr:to><cdr:x>3200000</cdr:x><cdr:y>2320000</cdr:y></cdr:to><cdr:sp><cdr:nvSpPr><cdr:cNvPr id="1" name="Source note" descr="Chart source note"/><cdr:cNvSpPr txBox="1"/></cdr:nvSpPr><cdr:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="1F2937"/></a:solidFill></a:ln><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 12000"/></a:avLst></a:prstGeom></cdr:spPr><cdr:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="l"/><a:r><a:rPr lang="en-US" sz="1000" b="1"/><a:t>${CHART_SOURCE_NOTE_TEXT}</a:t></a:r></a:p></cdr:txBody></cdr:sp><cdr:clientData/></cdr:twoCellAnchor><cdr:oneCellAnchor><cdr:from><cdr:x>4300000</cdr:x><cdr:y>260000</cdr:y></cdr:from><cdr:ext cx="900000" cy="600000"/><cdr:pic><cdr:nvPicPr><cdr:cNvPr id="2" name="Chart badge" descr="Chart badge"/><cdr:cNvPicPr/></cdr:nvPicPr><cdr:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></cdr:blipFill><cdr:spPr><a:ln><a:noFill/></a:ln><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></cdr:spPr></cdr:pic><cdr:clientData/></cdr:oneCellAnchor></cdr:wsDr>`;

const CHART_SHAPES_RELS_XML = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/chart-logo.png"/></Relationships>`;

const CHART_RELS_XML = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/chart-logo.png"/><Relationship Id="rIdShapes" Type="${CHART_SHAPES_REL_TYPE}" Target="chartUserShapes1.xml"/></Relationships>`;

const SERIES_BLIP_FILL = '<a:blipFill><a:blip r:embed="rIdLogo"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>';
const CHART_AREA_BLIP_FILL = '<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><a:blipFill><a:blip r:embed="rIdLogo"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></c:spPr>';
const PLOT_AREA_BLIP_FILL = '<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><a:blipFill><a:blip r:embed="rIdLogo"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></c:spPr>';
const POINT_BLIP_FILL = '<c:dPt xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:idx val="1"/><c:spPr><a:blipFill><a:blip r:embed="rIdLogo"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></c:spPr></c:dPt>';

function patchChartXml(xml) {
  let patched = xml.replace(
    /<chartSpace xmlns="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/chart">/,
    '<chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
  );
  if (!patched.includes("xmlns:r=")) {
    throw new Error("chart fixture lost its chartSpace namespace declarations");
  }
  // Chart-area picture fill: `c:spPr` is the last `chartSpace` child in the schema.
  patched = patched.replace("</chartSpace>", `${CHART_AREA_BLIP_FILL}</chartSpace>`);
  // Plot-area picture fill.
  patched = patched.replace(
    "</plotArea>",
    `${PLOT_AREA_BLIP_FILL}</plotArea>`,
  );
  // Series picture fill replaces the solid fill of the first series.
  patched = patched.replace(
    /<spPr><a:ln xmlns:a="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/main"><a:prstDash val="solid"\/><\/a:ln><\/spPr>/,
    `<spPr>${SERIES_BLIP_FILL}<a:ln><a:prstDash val="solid"/></a:ln></spPr>`,
  );
  if (!patched.includes("rIdLogo")) {
    throw new Error("chart fixture lost its series picture fill");
  }
  // Point picture fill on the second category.
  patched = patched.replace("</spPr><cat>", `</spPr>${POINT_BLIP_FILL}<cat>`);
  return patched;
}

/**
 * Returns bytes of the sample workbook with chart pictures and user shapes.
 *
 * @param {{ chartImageBytes?: Uint8Array, skipShapesPart?: boolean }} [options]
 */
export function buildChartMediaWorkbookBytes(options = {}) {
  const archive = unzipSync(new Uint8Array(readFileSync(CHART_MEDIA_SAMPLE_PATH)));
  const logoBytes = options.chartImageBytes ?? archive["xl/media/image1.png"];
  archive["xl/media/chart-logo.png"] = logoBytes;
  archive["xl/charts/chart1.xml"] = strToU8(patchChartXml(strFromU8(archive["xl/charts/chart1.xml"])));
  archive["xl/charts/_rels/chart1.xml.rels"] = strToU8(CHART_RELS_XML);
  if (!options.skipShapesPart) {
    archive[CHART_USER_SHAPES_PATH] = strToU8(CHART_SHAPES_XML);
    archive["xl/charts/_rels/chartUserShapes1.xml.rels"] = strToU8(CHART_SHAPES_RELS_XML);
  }
  const contentTypes = strFromU8(archive["[Content_Types].xml"]);
  archive["[Content_Types].xml"] = strToU8(
    options.skipShapesPart
      ? contentTypes
      : contentTypes.replace(
        "</Types>",
        `<Override PartName="/${CHART_USER_SHAPES_PATH}" ContentType="${CHART_USER_SHAPES_CONTENT_TYPE}"/></Types>`,
      ),
  );
  return zipSync(archive, { level: 6 });
}

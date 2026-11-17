import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const core = await import(
  pathToFileURL(requireFromDocxCore.resolve("@arcships/docx-core")).href
);

const WORD_NS =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

test("run styles keep double underline and double strike variants", () => {
  const underline = core.runStyleToCss({
    underline: true,
    underlineStyle: "double",
  });
  assert.equal(underline.textDecoration, "underline");
  assert.equal(underline.textDecorationStyle, "double");

  const strike = core.runStyleToCss({
    strike: true,
    strikeStyle: "double",
  });
  assert.equal(strike.textDecoration, "line-through");
  assert.equal(strike.textDecorationStyle, "double");

  const heavy = core.runStyleToCss({
    underline: true,
    underlineStyle: "dashedHeavy",
  });
  assert.equal(heavy.textDecorationStyle, "dashed");
  assert.equal(heavy.textDecorationThickness, "3px");
});

test("run borders map OOXML line styles to valid CSS", () => {
  const wave = core.runStyleToCss({
    runBorder: { type: "wave", sizeEighthPt: 8, color: "#FF0000" },
  });
  assert.equal(wave.borderStyle, "solid");
  assert.equal(wave.borderWidth, "2px");
  assert.equal(wave.borderColor, "#FF0000");

  const thick = core.runStyleToCss({
    runBorder: { type: "thick" },
  });
  assert.equal(thick.borderStyle, "solid");
  assert.equal(thick.borderWidth, "2px");

  const parsed = core.parseRunStyleFromRunXml(
    '<w:r><w:rPr><w:bdr w:val="dotDash" w:sz="8" w:space="1" w:color="FF0000"/></w:rPr><w:t>x</w:t></w:r>',
  );
  assert.equal(parsed.runBorder.type, "dotDash");
  assert.equal(parsed.runBorder.color, "#ff0000");
  assert.equal(parsed.runBorder.sizeEighthPt, 8);
  assert.equal(parsed.runBorder.spacePt, 1);

  const parsedCss = core.runStyleToCss(parsed);
  assert.equal(parsedCss.borderStyle, "dashed");
  assert.equal(parsedCss.borderWidth, "1.33px");
});

test("caps, small caps, character spacing and scale are modelled", () => {
  const caps = core.parseRunStyleFromRunXml(
    '<w:r><w:rPr><w:caps/><w:spacing w:val="20"/></w:rPr><w:t>Hello</w:t></w:r>',
  );
  assert.equal(caps.caps, true);
  assert.equal(caps.characterSpacingTwips, 20);

  const capsCss = core.runStyleToCss(caps);
  assert.equal(capsCss.textTransform, "uppercase");
  assert.equal(capsCss.letterSpacing, "1pt");

  const smallCaps = core.parseRunStyleFromRunXml(
    '<w:r><w:rPr><w:smallCaps/></w:rPr><w:t>Hello</w:t></w:r>',
  );
  assert.equal(core.runStyleToCss(smallCaps).fontVariantCaps, "small-caps");

  const scaled = core.parseRunStyleFromRunXml(
    '<w:r><w:rPr><w:w w:val="150"/></w:rPr><w:t>Wide</w:t></w:r>',
  );
  assert.equal(scaled.characterScalePercent, 150);
  const scaledCss = core.runStyleToCss(scaled);
  assert.equal(scaledCss.transform, "scaleX(1.5)");
  assert.equal(scaledCss.display, "inline-block");
});

test("RTL and extended justification survive parsing and CSS mapping", () => {
  const pPr = '<w:pPr><w:bidi/><w:jc w:val="both"/></w:pPr>';
  assert.equal(core.parseParagraphAlignmentFromXml(pPr), "both");
  assert.equal(core.parseParagraphDirectionFromXml(pPr), "rtl");

  assert.equal(core.paragraphTypographyStyle({ align: "both" }).textAlign, "justify");
  assert.equal(
    core.paragraphTypographyStyle({ align: "distribute" }).textAlign,
    "justify",
  );
  assert.equal(core.paragraphTypographyStyle({ align: "low" }).textAlign, "left");

  assert.deepEqual(
    core.paragraphTypographyStyle({ align: "distribute", direction: "rtl" }),
    { textAlign: "justify", direction: "rtl", unicodeBidi: "embed" },
  );

  const rtlRun = core.parseRunStyleFromRunXml(
    '<w:r><w:rPr><w:rtl/></w:rPr><w:t>مرحبا</w:t></w:r>',
  );
  assert.equal(rtlRun.direction, "rtl");
  assert.equal(core.runStyleToCss(rtlRun).direction, "rtl");
});

test("eastAsia and complex-script fonts build a font stack instead of being dropped", () => {
  const parsed = core.parseRunStyleFromRunXml(
    '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="SimSun" w:cs="Arial"/></w:rPr><w:t>中文 Text</w:t></w:r>',
  );
  assert.equal(parsed.fontFamily, "Calibri");
  assert.equal(parsed.fontFamilyEastAsia, "SimSun");
  assert.equal(parsed.fontFamilyComplexScript, "Arial");

  const stack = core.cssFontFamily(parsed.fontFamily, {
    eastAsia: parsed.fontFamilyEastAsia,
    complexScript: parsed.fontFamilyComplexScript,
  });
  assert.match(stack, /Calibri/);
  assert.match(stack, /SimSun/);
  assert.match(stack, /Arial/);


});

test("numbering formats cover decimalZero, Chinese and ideographic counters", () => {
  assert.equal(core.formatNumberingCounter("decimalZero", 1), "01");
  assert.equal(core.formatNumberingCounter("decimalZero", 12), "12");
  assert.equal(core.formatNumberingCounter("chineseCounting", 12), "十二");
  assert.equal(core.formatNumberingCounter("chineseCounting", 105), "一百零五");
  assert.equal(core.formatNumberingCounter("chineseLegalSimplified", 12), "壹拾贰");
  assert.equal(core.formatNumberingCounter("ideographDigital", 3), "三");
  assert.equal(core.formatNumberingCounter("ideographTraditional", 1), "甲");
  assert.equal(core.formatNumberingCounter("ordinal", 2), "2nd");
  assert.equal(core.formatNumberingCounter("decimalFullWidth", 12), "１２");
});

test("numbering level alignment and suffix are read from numbering.xml", async () => {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="${WORD_NS}"><w:body>` +
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
    "<w:r><w:t>Item</w:t></w:r></w:p>" +
    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
    "<w:r><w:t>Next</w:t></w:r></w:p>" +
    "</w:body></w:document>";
  const numberingXml =
    `<w:numbering xmlns:w="${WORD_NS}">` +
    '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0">' +
    '<w:start w:val="1"/><w:numFmt w:val="decimalZero"/>' +
    '<w:lvlText w:val="%1."/><w:suff w:val="space"/><w:lvlJc w:val="center"/>' +
    "</w:lvl></w:abstractNum>" +
    '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>' +
    "</w:numbering>";

  const pkg = core.createMinimalDocxPackage(documentXml);
  pkg.parts.set("word/numbering.xml", {
    name: "word/numbering.xml",
    content: numberingXml,
  });
  const model = await core.buildDocModel(pkg);
  const definitions = model.metadata.numberingDefinitions;
  const level = definitions.abstracts[0].levels[0];
  assert.equal(level.suffix, "space");
  assert.equal(level.justification, "center");

  const labels = core.buildParagraphNumberingLabels(model);
  const label = labels.get("p:0");
  assert.equal(label.text, "01. ");
  assert.equal(label.alignment, "center");
  assert.equal(labels.get("p:1").text, "02. ");

  const centered = core.numberingMarkerStyle(
    model.nodes[0],
    definitions,
    label,
    undefined,
    "light",
  );
  assert.equal(centered.justifyContent, "center");

  const left = core.numberingMarkerStyle(
    model.nodes[0],
    definitions,
    { ...label, alignment: "left" },
    undefined,
    "light",
  );
  assert.equal(left.justifyContent, "flex-start");
});

test("imported DOCX models are enriched from paragraph XML", async () => {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="${WORD_NS}"><w:body>` +
    "<w:p><w:pPr><w:bidi/><w:jc w:val=\"both\"/></w:pPr>" +
    "<w:r><w:rPr>" +
    '<w:rFonts w:ascii="Calibri" w:eastAsia="SimSun"/>' +
    '<w:u w:val="double"/><w:dstrike/><w:caps/><w:spacing w:val="20"/>' +
    '<w:w w:val="150"/><w:bdr w:val="wave" w:sz="8" w:color="FF0000"/>' +
    "<w:rtl/>" +
    "</w:rPr><w:t>中文 Text</w:t></w:r></w:p>" +
    "<w:p><w:r><w:t>plain</w:t></w:r></w:p>" +
    "</w:body></w:document>";

  const model = await core.buildDocModel(core.createMinimalDocxPackage(documentXml));
  const first = model.nodes[0];
  assert.equal(first.style.direction, "rtl");
  assert.ok(first.style.align === "both" || first.style.align === "justify");

  const run = first.children[0];
  assert.equal(run.style.underline, true);
  assert.equal(run.style.underlineStyle, "double");
  assert.equal(run.style.strike, true);
  assert.equal(run.style.strikeStyle, "double");
  assert.equal(run.style.caps, true);
  assert.equal(run.style.characterSpacingTwips, 20);
  assert.equal(run.style.characterScalePercent, 150);
  assert.equal(run.style.direction, "rtl");
  assert.equal(run.style.fontFamilyEastAsia, "SimSun");
  assert.equal(run.style.runBorder.type, "wave");

  const css = core.runStyleToCss(run.style);
  assert.equal(css.textDecorationStyle, "double");
  assert.equal(css.textTransform, "uppercase");
  assert.equal(css.letterSpacing, "1pt");
  assert.equal(css.borderStyle, "solid");
  assert.match(css.fontFamily, /SimSun/);
});

test("theme fonts from the imported package are attached to the model", async () => {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="${WORD_NS}"><w:body>` +
    "<w:p><w:r><w:rPr>" +
    '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia"/>' +
    "</w:rPr><w:t>中文</w:t></w:r></w:p>" +
    "</w:body></w:document>";
  const themeXml =
    '<a:theme><a:themeElements><a:fontScheme>' +
    '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="Microsoft YaHei"/></a:minorFont>' +
    "</a:fontScheme></a:themeElements></a:theme>";

  const pkg = core.createMinimalDocxPackage(documentXml);
  pkg.parts.set("word/theme/theme1.xml", {
    name: "word/theme/theme1.xml",
    content: themeXml,
  });
  const model = await core.buildDocModel(pkg);
  assert.equal(model.metadata.themeFonts.minor.ascii, "Calibri");
  assert.equal(model.metadata.themeFonts.minor.eastAsia, "Microsoft YaHei");

  const run = model.nodes[0].children[0];
  assert.equal(run.style.fontFamily, "Calibri");
  assert.equal(run.style.fontFamilyEastAsia, "Microsoft YaHei");
});

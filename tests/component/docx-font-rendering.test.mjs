// DOCX script-font rendering (issue #123).
//
// The layout measures one run per OOXML script slot; the DOM must paint the
// same split, with the same family, kerning and OpenType feature settings.
// DocxParagraphHost's run renderer emits static HTML, so the assertions read
// the spans out of its `innerHTML`.

import assert from "node:assert/strict";
import test from "node:test";

import { importFromDemo, mount, walk } from "./vue-test-renderer.mjs";

const { DocxParagraphHost } = await importFromDemo("@arcships/vue-docx");
const { DISABLED_DOCX_FONT_FEATURE_SETTINGS } = await importFromDemo(
  "@arcships/docx-core"
);

const SCRIPT_STYLE = {
  fontFamily: "TestLatin",
  fontFamilyAscii: "TestLatin",
  fontFamilyHAnsi: "TestLatin",
  fontFamilyEastAsia: "TestCJK",
  fontFamilyComplexScript: "TestCS",
  sourceFontFamily: "TestLatin",
  fontSizePt: 12,
  kerningMinFontSizePt: 12,
  ligatures: "standard",
  contextualAlternates: true,
};

function textRun(text, style) {
  return { type: "text", text, style };
}

function decodeHtml(value) {
  return value
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function parseStyleAttribute(attributes) {
  const styleMatch = attributes.match(/\bstyle="([^"]*)"/);
  const styles = {};
  if (!styleMatch) return styles;
  for (const declaration of decodeHtml(styleMatch[1]).split(";")) {
    const separator = declaration.indexOf(":");
    if (separator < 0) continue;
    styles[declaration.slice(0, separator).trim()] = declaration
      .slice(separator + 1)
      .trim();
  }
  return styles;
}

function spansFromHtml(html) {
  return [...String(html).matchAll(/<span\b([^>]*)>([\s\S]*?)<\/span>/g)].map(
    (match) => ({
      attributes: match[1],
      text: decodeHtml(match[2].replace(/<[^>]+>/g, "")),
      style: parseStyleAttribute(match[1]),
    })
  );
}

async function paragraphHtml(paragraph) {
  const mounted = await mount(DocxParagraphHost, {
    paragraph,
    paragraphIndex: 0,
  });
  try {
    const host = walk(mounted.root).find(
      (node) => node.props?.["data-testid"] === "editor-paragraph"
    );
    assert.ok(host, "paragraph host should render");
    return String(host.props.innerHTML);
  } finally {
    mounted.app.unmount();
  }
}

test("a mixed Latin/CJK/RTL run renders one span per font slot", async () => {
  const html = await paragraphHtml({
    type: "paragraph",
    children: [textRun("Hello 中文 مرحبا", SCRIPT_STYLE)],
  });
  const spans = spansFromHtml(html);
  const byText = new Map(spans.map((span) => [span.text, span.style]));

  const latin = byText.get("Hello ");
  const cjk = byText.get("中文");
  const arabic = byText.get("مرحبا");
  assert.ok(latin, "Latin segment span");
  assert.ok(cjk, "CJK segment span");
  assert.ok(arabic, "complex-script segment span");

  assert.match(latin["font-family"], /TestLatin/);
  assert.match(cjk["font-family"], /TestCJK/);
  assert.match(arabic["font-family"], /TestCS/);

  for (const style of [latin, cjk, arabic]) {
    assert.equal(style["font-kerning"], "normal");
    assert.match(style["font-feature-settings"], /"liga" 1/);
    assert.match(style["font-feature-settings"], /"calt" 1/);
  }

  // The rendered text is unchanged and in order.
  assert.equal(
    spans.map((span) => span.text).join(""),
    "Hello 中文 مرحبا"
  );
});

test("a script-neutral run keeps one span and disables Word-default ligatures", async () => {
  const html = await paragraphHtml({
    type: "paragraph",
    children: [
      textRun("Plain text", { fontFamily: "TestLatin", fontSizePt: 12 }),
    ],
  });
  const spans = spansFromHtml(html).filter(
    (span) => span.text === "Plain text"
  );
  assert.equal(spans.length, 1);
  assert.equal(spans[0].style["font-kerning"], "none");
  assert.equal(
    spans[0].style["font-feature-settings"],
    DISABLED_DOCX_FONT_FEATURE_SETTINGS
  );
});

test("a mixed-script hyperlink carries the script font on each child span", async () => {
  const html = await paragraphHtml({
    type: "paragraph",
    children: [
      {
        type: "text",
        text: "Latin 中文",
        link: "https://example.com/docs",
        style: SCRIPT_STYLE,
      },
    ],
  });

  assert.match(html, /<a\b[^>]*href="https:\/\/example\.com\/docs"/);
  const spans = spansFromHtml(html);
  const latin = spans.find((span) => span.text === "Latin ");
  const cjk = spans.find((span) => span.text === "中文");
  assert.ok(latin);
  assert.ok(cjk);
  assert.match(latin.style["font-family"], /TestLatin/);
  assert.match(cjk.style["font-family"], /TestCJK/);
  // The link colour still applies to every child.
  assert.equal(latin.style.color, "inherit");
  assert.equal(cjk.style.color, "inherit");
});

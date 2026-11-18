import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

import { importFromDemo, mount, walk } from "./vue-test-renderer.mjs";

const { DocxParagraphHost } = await importFromDemo("@arcships/vue-docx");
const { paragraphBaseFontSizePx, resolveMeasureFont } = await importFromDemo(
  "@arcships/docx-core",
);

function textRun(text, style) {
  return { type: "text", text, style };
}

function paragraph(style, children = [textRun("Body text")]) {
  return { type: "paragraph", style, children };
}

async function withParagraph(documentParagraph, assertions) {
  const mounted = await mount(DocxParagraphHost, {
    paragraph: documentParagraph,
    paragraphIndex: 0,
  });
  const host = walk(mounted.root).find(
    (node) => node.props?.["data-testid"] === "editor-paragraph",
  );
  assert.ok(host, "paragraph host should render");
  try {
    return await assertions(host.props.style, host, mounted);
  } finally {
    mounted.app.unmount();
  }
}

test("paragraph host renders spacing and first-line indentation", () =>
  withParagraph(
    paragraph({
      spacing: { beforeTwips: 240, afterTwips: 120 },
      indent: { firstLineTwips: 480 },
    }),
    (style) => {
      // 12pt before, 6pt after, 24pt first-line indent (15 twips per pixel).
      assert.equal(style.marginTop, "16px");
      assert.equal(style.marginBottom, "8px");
      assert.equal(style.textIndent, "32px");
      assert.equal(style.paddingLeft, undefined);
      assert.equal(style.paddingRight, undefined);
    },
  ));

test("paragraph host maps left and hanging indents to padding", () =>
  withParagraph(
    paragraph({
      indent: { leftTwips: 720, rightTwips: 360, hangingTwips: 360 },
    }),
    (style) => {
      assert.equal(style.paddingLeft, "48px");
      assert.equal(style.paddingRight, "24px");
      assert.equal(style.textIndent, "-24px");
    },
  ));

test("explicit first-line indent wins over a stale hanging indent", () =>
  withParagraph(
    paragraph({ indent: { firstLineTwips: 240, hangingTwips: 360 } }),
    (style) => {
      assert.equal(style.textIndent, "16px");
    },
  ));

test("paragraph host applies exact and atLeast line spacing", async () => {
  await withParagraph(
    paragraph({ spacing: { lineTwips: 360, lineRule: "exact" } }),
    (style) => assert.equal(style.lineHeight, "24px"),
  );
  await withParagraph(
    paragraph({ spacing: { lineTwips: 360, lineRule: "atLeast" } }),
    (style) => assert.equal(style.lineHeight, "24px"),
  );
});

test("paragraph host applies auto line spacing as a true multiple", async () => {
  await withParagraph(
    paragraph({ spacing: { lineTwips: 240, lineRule: "auto" } }),
    (style) => assert.equal(style.lineHeight, 1),
  );
  await withParagraph(
    paragraph({ spacing: { lineTwips: 360, lineRule: "auto" } }),
    (style) => assert.equal(style.lineHeight, 1.5),
  );
  await withParagraph(
    paragraph({ spacing: { lineTwips: 480, lineRule: "auto" } }),
    (style) => assert.equal(style.lineHeight, 2),
  );
  await withParagraph(paragraph(undefined), (style) => {
    assert.equal(style.lineHeight, undefined);
  });
});

test("paragraph host renders borders, border space and shading", () =>
  withParagraph(
    paragraph({
      backgroundColor: "#fff2cc",
      indent: { leftTwips: 720 },
      borders: {
        top: { type: "single", sizeEighthPt: 8, color: "#ff0000", spacePt: 1 },
        bottom: { type: "double", sizeEighthPt: 12, color: "#00ff00" },
        left: { type: "single", sizeEighthPt: 6, color: "#0000ff", spacePt: 2 },
      },
    }),
    (style) => {
      assert.equal(style.borderTop, "1.33px solid #ff0000");
      assert.equal(style.borderBottom, "2px double #00ff00");
      assert.equal(style.borderLeft, "1px solid #0000ff");
      assert.equal(style.backgroundColor, "#fff2cc");
      // Border stroke + space stay outside the text box.
      assert.equal(style.paddingLeft, "50.67px");
      assert.equal(style.paddingTop, "1.33px");
      assert.equal(style.paddingBottom, undefined);
    },
  ));

test("a paragraph bar renders as the left rule for quotes", () =>
  withParagraph(
    paragraph({
      borders: { bar: { type: "single", sizeEighthPt: 12, color: "#4472c4" } },
    }),
    (style) => {
      assert.equal(style.borderLeft, "2px solid #4472c4");
      assert.equal(style.borderTop, undefined);
      assert.equal(style.paddingLeft, undefined);
    },
  ));

test("explicit run fonts drive the paragraph base font", () =>
  withParagraph(
    paragraph(undefined, [
      textRun("Times", { fontFamily: "Times New Roman", fontSizePt: 12 }),
    ]),
    (style) => {
      assert.equal(style.fontFamily, '"Times New Roman", serif');
      assert.equal(style.fontSize, "16px");
    },
  ));

test("heading fallback size comes from the shared measurement base", () => {
  const model = paragraph({ headingLevel: 1 }, [textRun("Heading")]);
  return withParagraph(model, (style) => {
    assert.equal(style.fontSize, `${paragraphBaseFontSizePx(model)}px`);
    assert.equal(style.fontSize, "21px");
    assert.equal(style.fontWeight, 700);
  });
});

test("paragraph alignment is preserved", () =>
  withParagraph(paragraph({ align: "center" }), (style) => {
    assert.equal(style.textAlign, "center");
  }));

test("paragraph host default font matches the measurement font", () => {
  const model = paragraph(undefined, [textRun("No explicit font")]);
  return withParagraph(model, (style) => {
    const baseFontPx = paragraphBaseFontSizePx(model);
    assert.equal(style.fontSize, `${baseFontPx}px`);
    assert.ok(
      resolveMeasureFont(undefined, baseFontPx).endsWith(style.fontFamily),
      "measured font should use the same family as the paragraph host",
    );
  });
});

test("the paragraph host CSS keeps layout defaults but no typography", () => {
  const sfc = readFileSync(
    new URL(
      "../../packages/vue-docx/src/components/DocxParagraphHost.vue",
      import.meta.url,
    ),
    "utf8",
  );
  const styleBlock = sfc.match(/<style scoped>([\s\S]*?)<\/style>/)?.[1];
  assert.ok(styleBlock, "paragraph host should have a scoped style block");
  const hostRule = styleBlock.match(/\.docx-paragraph-host\s*\{([^}]*)\}/)?.[1];
  assert.ok(hostRule, "paragraph host base rule should exist");
  assert.match(hostRule, /min-height:\s*1em/);
  assert.match(hostRule, /white-space:\s*pre-wrap/);
  assert.doesNotMatch(hostRule, /(^|\s)line-height\s*:/);
  assert.doesNotMatch(hostRule, /(^|\s)font-family\s*:/);
  assert.doesNotMatch(hostRule, /(^|\s)font-size\s*:/);
});

// ── Automatic (HTML) spacing and the document tab grid (issue #129) ──

test("paragraph host adds Word's automatic spacing to explicit margins", async () => {
  // 14 pt at 96 dpi: 280 twips / 15 twips-per-pixel.
  const autoSpacingPx = 280 / 15;
  const mounted = await mount(DocxParagraphHost, {
    paragraph: paragraph({ spacing: { beforeTwips: 240, afterTwips: 120 } }),
    paragraphIndex: 0,
    autoSpacingBeforePx: autoSpacingPx,
    autoSpacingAfterPx: autoSpacingPx,
  });
  const host = walk(mounted.root).find(
    (node) => node.props?.["data-testid"] === "editor-paragraph",
  );
  try {
    assert.ok(host, "paragraph host should render");
    // 16 px + 18.67 px before, 8 px + 18.67 px after.
    assert.equal(host.props.style.marginTop, "34.67px");
    assert.equal(host.props.style.marginBottom, "26.67px");
  } finally {
    mounted.app.unmount();
  }
});

test("rendered tabs advance on the document default tab stop", async () => {
  const withTab = paragraph({}, [
    textRun("AA"),
    textRun("\t"),
    textRun("ZZ"),
  ]);

  const tabWidthPxFor = async (defaultTabStopPx) => {
    const mounted = await mount(DocxParagraphHost, {
      paragraph: withTab,
      paragraphIndex: 0,
      defaultTabStopPx,
    });
    try {
      const host = walk(mounted.root).find(
        (node) => node.props?.["data-testid"] === "editor-paragraph",
      );
      assert.ok(host, "paragraph host should render");
      // The run renderer emits its vnodes as static HTML.
      const spacerWidths = [
        ...String(host.props.innerHTML).matchAll(
          /display:inline-block;[^"]*?width:(\d+(?:\.\d+)?)px/g,
        ),
      ].map((match) => Number(match[1]));
      assert.equal(spacerWidths.length, 1, "one tab spacer");
      return spacerWidths[0];
    } finally {
      mounted.app.unmount();
    }
  };

  const narrow = await tabWidthPxFor(28);
  const wide = await tabWidthPxFor(48);
  // 420 twips vs 720 twips: the gap to the same first stop differs by 20 px.
  assert.equal(wide - narrow, 48 - 28);
  assert.ok(narrow > 0 && wide > 0);
});

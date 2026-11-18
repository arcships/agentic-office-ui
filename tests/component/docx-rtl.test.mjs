import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";

import { findByTestId, importFromDemo, mount } from "./vue-test-renderer.mjs";

const { DocxParagraphHost } = await importFromDemo("@arcships/vue-docx");
const core = await importFromDemo("@arcships/docx-core");

const ROOT = path.resolve(import.meta.dirname, "../..");

function paragraph(style, text = "مرحبا بالعالم") {
  return {
    type: "paragraph",
    style,
    children: [{ type: "text", text }],
  };
}

test("an RTL DOCX paragraph renders with dir=rtl and right alignment", async () => {
  const mounted = await mount(DocxParagraphHost, {
    paragraph: paragraph({ align: "right", direction: "rtl" }),
    paragraphIndex: 0,
    documentTheme: "light",
  });
  try {
    assert.equal(findByTestId(mounted.root, "editor-paragraph").props.dir, "rtl", "the paragraph host exposes dir=rtl");
    assert.equal(findByTestId(mounted.root, "editor-paragraph").props.style.direction, "rtl", "the bidi direction reaches CSS");
    assert.equal(findByTestId(mounted.root, "editor-paragraph").props.style.textAlign, "right", "w:jc=right stays right in RTL");
    assert.equal(findByTestId(mounted.root, "editor-paragraph").props.style.unicodeBidi, "embed", "punctuation follows the bidi embedding");
  } finally {
    mounted.app.unmount();
  }
});

test("a start-aligned RTL paragraph resolves to the right edge", async () => {
  const mounted = await mount(DocxParagraphHost, {
    paragraph: paragraph({ align: "start", direction: "rtl" }),
    paragraphIndex: 0,
    documentTheme: "light",
  });
  try {
    assert.equal(findByTestId(mounted.root, "editor-paragraph").props.dir, "rtl");
    assert.equal(findByTestId(mounted.root, "editor-paragraph").props.style.textAlign, "start", "start follows the paragraph direction");
    assert.equal(findByTestId(mounted.root, "editor-paragraph").props.style.direction, "rtl");
    assert.deepEqual(
      core.paragraphTypographyStyle({ align: "start", direction: "rtl" }),
      { textAlign: "start", direction: "rtl", unicodeBidi: "embed" },
    );
  } finally {
    mounted.app.unmount();
  }
});

test("an LTR paragraph keeps its direction and does not set dir", async () => {
  const mounted = await mount(DocxParagraphHost, {
    paragraph: paragraph({ align: "left", direction: "ltr" }, "Hello"),
    paragraphIndex: 0,
    documentTheme: "light",
  });
  try {
    assert.equal(findByTestId(mounted.root, "editor-paragraph").props.dir, "ltr", "an explicit w:bidi=0 is forwarded too");
    assert.equal(findByTestId(mounted.root, "editor-paragraph").props.style.direction, "ltr");
    const plain = await mount(DocxParagraphHost, {
      paragraph: paragraph({ align: "left" }, "Hello"),
      paragraphIndex: 0,
      documentTheme: "light",
    });
    try {
      assert.equal(findByTestId(plain.root, "editor-paragraph").props.dir, undefined, "no bidi mark leaves the paragraph LTR");
    } finally {
      plain.app.unmount();
    }
  } finally {
    mounted.app.unmount();
  }
});

test("table-cell paragraphs forward their own bidi direction", () => {
  const source = readFileSync(
    path.join(ROOT, "packages/vue-docx/src/components/DocxTableHost.vue"),
    "utf8",
  );
  assert.match(source, /:dir="node\.style\?\.direction"/);
});

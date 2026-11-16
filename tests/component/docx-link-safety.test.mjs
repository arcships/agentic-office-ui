import assert from "node:assert/strict";
import test from "node:test";

import {
  importFromDemo,
  mount,
  vue,
} from "./vue-test-renderer.mjs";

const {
  renderParagraphRuns,
  renderStaticHtml,
  useDocxEditor,
} = await importFromDemo("@arcships/vue-docx");
const { createBlankDocumentModel } = await importFromDemo("@arcships/docx-core");

const SAFE_LINKS = [
  "https://example.com/path?q=1#frag",
  "http://example.com/path",
  "mailto:person@example.com",
  "#bookmark",
  "/relative/path",
  "relative/path",
];

const UNSAFE_LINKS = [
  "javascript:alert(1)",
  "JaVaScRiPt:alert(1)",
  " javascript:alert(1)",
  "java\nscript:alert(1)",
  "data:text/html,<script>alert(1)</script>",
  "vbscript:msgbox(1)",
  "file:///etc/passwd",
];

function starterModel(nodes) {
  return {
    ...createBlankDocumentModel(),
    nodes,
  };
}

function textRun(model) {
  const children = model.nodes[0]?.type === "paragraph" ? model.nodes[0].children : [];
  return children.find((child) => child.type === "text");
}

async function mountController(model) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () => vue.h("div", { "data-testid": "controller-harness" });
    },
  });
  const mounted = await mount(Harness);
  return { ...mounted, get controller() { return controller; } };
}

function selectExampleText(controller) {
  controller.selectParagraph(0);
  controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 0 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 7 },
  });
}

function renderParagraph(paragraph) {
  return renderStaticHtml(renderParagraphRuns(paragraph, "link-test"));
}

test("setLink rejects executable schemes and stores safe targets", async () => {
  const mounted = await mountController(starterModel([
    { type: "paragraph", children: [{ type: "text", text: "Example" }] },
  ]));
  const { controller } = mounted;
  selectExampleText(controller);
  await vue.nextTick();

  for (const unsafe of UNSAFE_LINKS) {
    controller.setLink(unsafe);
    await vue.nextTick();
    assert.equal(textRun(controller.model)?.link, undefined, `rejected ${JSON.stringify(unsafe)}`);
  }

  for (const safe of SAFE_LINKS) {
    controller.setLink(safe);
    await vue.nextTick();
    assert.equal(textRun(controller.model)?.link, safe, `accepted ${JSON.stringify(safe)}`);
  }

  controller.setLink("");
  await vue.nextTick();
  assert.equal(textRun(controller.model)?.link, undefined, "empty input removes the link");

  controller.setLink(undefined);
  await vue.nextTick();
  assert.equal(textRun(controller.model)?.link, undefined);
  assert.deepEqual(mounted.warnings, []);
  mounted.app.unmount();
});

test("renderer neutralises unsafe text and form-field links", () => {
  for (const unsafe of UNSAFE_LINKS) {
    const textHtml = renderParagraph({
      type: "paragraph",
      children: [{ type: "text", text: "Example", link: unsafe }],
    });
    assert.equal(textHtml.includes("<a"), false, `plain text for ${JSON.stringify(unsafe)}`);
    assert.equal(textHtml.includes(unsafe), false);
    assert.match(textHtml, /Example/);

    const fieldHtml = renderParagraph({
      type: "paragraph",
      children: [{
        type: "form-field",
        fieldType: "text",
        value: "Field",
        link: unsafe,
      }],
    });
    assert.equal(fieldHtml.includes("<a"), false, `plain field for ${JSON.stringify(unsafe)}`);
    assert.equal(fieldHtml.includes(unsafe), false);
    assert.match(fieldHtml, /Field/);
  }
});

test("renderer keeps safe hrefs and internal-link click handling", () => {
  const external = renderParagraph({
    type: "paragraph",
    children: [{ type: "text", text: "Example", link: "https://example.com" }],
  });
  assert.match(external, /<a[^>]*href="https:\/\/example\.com"/);
  assert.match(external, /target="_blank"/);
  assert.match(external, /rel="noreferrer noopener"/);

  const internal = renderParagraph({
    type: "paragraph",
    children: [{ type: "text", text: "Jump", link: "#bookmark" }],
  });
  assert.match(internal, /<a[^>]*href="#bookmark"/);
  assert.equal(internal.includes('target="_blank"'), false);
});

test("static HTML drops an unsafe anchor href defensively", () => {
  const html = renderStaticHtml(
    vue.h("a", { href: "javascript:alert(1)", target: "_blank" }, "Unsafe"),
  );
  assert.equal(html.includes("javascript:"), false);
  assert.equal(html.includes("href="), false);
  assert.match(html, /Unsafe/);
});

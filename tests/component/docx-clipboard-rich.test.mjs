import assert from "node:assert/strict";
import test from "node:test";

import { importFromDemo, mount, vue, walk } from "./vue-test-renderer.mjs";

const { DocxParagraphHost, DocxViewerRoot, useDocxEditor } = await importFromDemo("@arcships/vue-docx");
const { DOCX_CLIPBOARD_MIME, createBlankDocumentModel } = await importFromDemo("@arcships/docx-core");

function paragraphText(node) {
  if (!node || node.type !== "paragraph") return "";
  return node.children.map((child) => (child.type === "text" ? child.text : "")).join("");
}

function textRuns(node) {
  return node.children.filter((child) => child.type === "text");
}

function starterModel(nodes) {
  return {
    ...createBlankDocumentModel(),
    nodes,
  };
}

async function mountController(model) {
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () => vue.h("div", { "data-testid": "controller-harness" }, controller.status);
    },
  });
  const mounted = await mount(Harness);
  return {
    ...mounted,
    get controller() {
      return controller;
    },
  };
}

class FakeClipboardItem {
  constructor(items) {
    this.items = items;
  }
  get types() {
    return Object.keys(this.items);
  }
  getType(type) {
    const blob = this.items[type];
    return blob ? Promise.resolve(blob) : Promise.reject(new Error(`missing ${type}`));
  }
}

/**
 * Installs a clipboard that records `write()` items and lets a test choose what
 * `read()` returns. Node has `Blob` but no `ClipboardItem`, so both the
 * constructor and `navigator` are restored afterwards.
 */
function installClipboard(t, { rejectCompositeWrite = false, rejectCustomMime = false } = {}) {
  const written = [];
  const textWrites = [];
  let readItems = [];
  const originalNavigator = Object.getOwnPropertyDescriptor(globalThis, "navigator");
  const originalClipboardItem = Object.getOwnPropertyDescriptor(globalThis, "ClipboardItem");

  Object.defineProperty(globalThis, "ClipboardItem", {
    configurable: true,
    value: FakeClipboardItem,
  });
  Object.defineProperty(globalThis, "navigator", {
    configurable: true,
    value: {
      clipboard: {
        async write(items) {
          if (rejectCompositeWrite) throw new Error("clipboard write denied");
          if (rejectCustomMime && items[0]?.types?.includes(DOCX_CLIPBOARD_MIME)) {
            throw new Error("custom MIME types are not supported");
          }
          written.push(items[0]);
        },
        async writeText(text) {
          textWrites.push(text);
        },
        async read() {
          return readItems;
        },
        async readText() {
          return readItems
            .map((item) => item.items["text/plain"])
            .filter(Boolean)
            .join("");
        },
      },
    },
  });

  t.after(() => {
    if (originalNavigator) Object.defineProperty(globalThis, "navigator", originalNavigator);
    else delete globalThis.navigator;
    if (originalClipboardItem) {
      Object.defineProperty(globalThis, "ClipboardItem", originalClipboardItem);
    } else {
      delete globalThis.ClipboardItem;
    }
  });

  return {
    written,
    textWrites,
    /** Make the next read return the last composite write. */
    readLastWrite() {
      readItems = written.slice(-1).map((item) => ({
        types: item.types,
        getType: (type) => item.getType(type),
      }));
    },
    readTypes(types) {
      readItems = [
        {
          types: Object.keys(types),
          getType: (type) => Promise.resolve(new Blob([types[type]], { type })),
        },
      ];
    },
  };
}

function selectBoldLink(controller) {
  controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 0 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 9 },
  });
}

function boldLinkModel() {
  return starterModel([
    {
      type: "paragraph",
      children: [
        { type: "text", text: "bold ", style: { bold: true, color: "#ff0000" } },
        { type: "text", text: "link", link: "https://example.com/a" },
      ],
    },
  ]);
}

test("copy writes the internal payload, text/html and text/plain flavours", async (t) => {
  const clipboard = installClipboard(t);
  const mounted = await mountController(boldLinkModel());
  const { controller } = mounted;

  selectBoldLink(controller);
  await controller.copy();

  assert.equal(controller.status, "Copied to clipboard");
  assert.equal(clipboard.written.length, 1, "one composite clipboard item is written");
  const item = clipboard.written[0];
  assert.deepEqual(
    item.types.sort(),
    [DOCX_CLIPBOARD_MIME, "text/html", "text/plain", "text/rtf"].sort(),
  );

  const internal = JSON.parse(await (await item.getType(DOCX_CLIPBOARD_MIME)).text());
  assert.equal(internal.version, 1);
  assert.equal(internal.kind, "docx-fragment");
  assert.deepEqual(internal.range, {
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 0 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 9 },
  });
  assert.deepEqual(internal.paragraphs[0].runs, [
    { text: "bold ", style: { bold: true, color: "#ff0000" } },
    { text: "link", link: "https://example.com/a" },
  ]);
  assert.equal(internal.text, "bold link");

  const html = await (await item.getType("text/html")).text();
  assert.match(html, /font-weight:bold/);
  assert.match(html, /color:#ff0000/);
  assert.match(html, /<a href="https:\/\/example.com\/a">link<\/a>/);
  const rtf = await (await item.getType("text/rtf")).text();
  assert.match(rtf, /^\{\\rtf1\\ansi/);
  assert.match(rtf, /\{\\b[^{}]*bold/);
  assert.match(rtf, /HYPERLINK "https:\/\/example.com\/a"/);
  assert.equal(await (await item.getType("text/plain")).text(), "bold link");

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("copy falls back to writeText when the composite clipboard write is denied", async (t) => {
  const clipboard = installClipboard(t, { rejectCompositeWrite: true });
  const mounted = await mountController(boldLinkModel());
  const { controller } = mounted;

  selectBoldLink(controller);
  await controller.copy();

  assert.deepEqual(clipboard.textWrites, ["bold link"]);
  assert.equal(clipboard.written.length, 0);
  assert.equal(controller.status, "Copied to clipboard");

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("copy retries with the standard flavours when a custom MIME write is rejected", async (t) => {
  const clipboard = installClipboard(t, { rejectCustomMime: true });
  const mounted = await mountController(boldLinkModel());
  const { controller } = mounted;

  selectBoldLink(controller);
  await controller.copy();

  assert.equal(clipboard.written.length, 1, "the retry reaches the clipboard");
  assert.deepEqual(clipboard.written[0].types.sort(), ["text/html", "text/plain", "text/rtf"]);
  const html = await (await clipboard.written[0].getType("text/html")).text();
  assert.match(html, /font-weight:bold/);
  assert.match(html, /<a href="https:\/\/example.com\/a">link<\/a>/);
  assert.equal(clipboard.textWrites.length, 0, "the HTML flavour is kept");

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("copy without a selection writes nothing and keeps the status", async (t) => {
  const clipboard = installClipboard(t);
  const mounted = await mountController(boldLinkModel());
  const { controller } = mounted;
  controller.setStatus("Ready");
  controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 2 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 2 },
  });

  await controller.copy();

  assert.equal(clipboard.written.length, 0);
  assert.equal(clipboard.textWrites.length, 0);
  assert.equal(controller.status, "Ready");

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("pasting our own copy round-trips bold, colour and links", async (t) => {
  const clipboard = installClipboard(t);
  const mounted = await mountController(boldLinkModel());
  const { controller } = mounted;

  selectBoldLink(controller);
  await controller.copy();
  clipboard.readLastWrite();

  controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 9 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 9 },
  });
  await controller.paste();
  await vue.nextTick();

  assert.equal(controller.status, "Pasted from clipboard");
  const node = controller.model.nodes[0];
  assert.equal(paragraphText(node), "bold linkbold link");
  const runs = textRuns(node);
  assert.equal(runs.length, 4);
  assert.deepEqual(runs[2].style, { bold: true, color: "#ff0000" });
  assert.equal(runs[3].link, "https://example.com/a");

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("pasting Word-style HTML keeps bold, links and lists", async (t) => {
  const clipboard = installClipboard(t);
  const mounted = await mountController(
    starterModel([{ type: "paragraph", children: [{ type: "text", text: "" }] }]),
  );
  const { controller } = mounted;

  clipboard.readTypes({
    "text/html":
      '<p><b>Bold</b> and <a href="https://example.com/b">link</a></p>' +
      '<ul><li>bullet one</li><li>bullet two</li></ul>' +
      '<p><a href="javascript:alert(1)">unsafe</a></p>',
    "text/plain": "Bold and link\nbullet one\nbullet two\nunsafe",
  });
  controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 0 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 0 },
  });
  await controller.paste();
  await vue.nextTick();

  const nodes = controller.model.nodes;
  assert.equal(nodes.length, 4);
  assert.deepEqual(textRuns(nodes[0])[0].style, { bold: true });
  assert.equal(textRuns(nodes[0])[1].text, " and ");
  assert.equal(textRuns(nodes[0])[2].link, "https://example.com/b");

  assert.equal(paragraphText(nodes[1]), "bullet one");
  assert.ok(nodes[1].style.numbering.numId > 0, "the pasted list keeps its numbering");
  assert.equal(nodes[2].style.numbering.numId, nodes[1].style.numbering.numId);
  assert.equal(paragraphText(nodes[3]), "unsafe");
  assert.equal(textRuns(nodes[3])[0].link, undefined, "an unsafe URL is stripped by the link policy");

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("internal payload wins over the HTML flavour on paste", async (t) => {
  const clipboard = installClipboard(t);
  const mounted = await mountController(
    starterModel([{ type: "paragraph", children: [{ type: "text", text: "" }] }]),
  );
  const { controller } = mounted;

  clipboard.readTypes({
    [DOCX_CLIPBOARD_MIME]: JSON.stringify({
      version: 1,
      kind: "docx-fragment",
      text: "internal",
      paragraphs: [{ runs: [{ text: "internal", style: { italic: true } }] }],
    }),
    "text/html": "<p>html</p>",
    "text/plain": "plain",
  });
  controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 0 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 0 },
  });
  await controller.paste();
  await vue.nextTick();

  assert.equal(paragraphText(controller.model.nodes[0]), "internal");
  assert.deepEqual(textRuns(controller.model.nodes[0])[0].style, { italic: true });

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("plain-text clipboard still pastes when no rich flavour parses", async (t) => {
  const clipboard = installClipboard(t);
  const mounted = await mountController(
    starterModel([{ type: "paragraph", children: [{ type: "text", text: "ab" }] }]),
  );
  const { controller } = mounted;

  clipboard.readTypes({ "text/html": "<meta charset='utf-8'>", "text/plain": "XY" });
  controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 1 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 1 },
  });
  await controller.paste();
  await vue.nextTick();

  assert.equal(paragraphText(controller.model.nodes[0]), "aXYb");
  assert.equal(controller.status, "Pasted from clipboard");

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("rich paste across paragraphs is one undo step and keeps formatting", async (t) => {
  const clipboard = installClipboard(t);
  const mounted = await mountController(
    starterModel([
      { type: "paragraph", children: [{ type: "text", text: "first" }] },
      { type: "paragraph", children: [{ type: "text", text: "second" }] },
    ]),
  );
  const { controller } = mounted;
  const before = JSON.stringify(controller.model);

  clipboard.readTypes({
    "text/html": "<p><b>new</b></p>",
    "text/plain": "new",
  });
  controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 2 },
    end: { location: { kind: "paragraph", nodeIndex: 1 }, offset: 3 },
  });
  await controller.paste();
  await vue.nextTick();

  assert.equal(controller.model.nodes.length, 1);
  assert.equal(paragraphText(controller.model.nodes[0]), "finewond");
  assert.deepEqual(textRuns(controller.model.nodes[0])[1].style, { bold: true });
  assert.equal(controller.canUndo, true);

  controller.undo();
  await vue.nextTick();
  assert.equal(
    JSON.stringify(controller.model),
    before,
    "a single undo restores the collapsed selection",
  );

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("writeSelectionToClipboardData fills a copy event sink synchronously", async (t) => {
  installClipboard(t);
  const mounted = await mountController(boldLinkModel());
  const { controller } = mounted;

  controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 2 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 2 },
  });
  const collapsed = [];
  assert.equal(
    controller.writeSelectionToClipboardData({ setData: (type, value) => collapsed.push([type, value]) }),
    false,
    "a collapsed caret leaves the native copy alone",
  );
  assert.deepEqual(collapsed, []);

  selectBoldLink(controller);
  const copied = [];
  assert.equal(
    controller.writeSelectionToClipboardData({ setData: (type, value) => copied.push([type, value]) }),
    true,
  );
  assert.deepEqual(copied.map(([type]) => type), [
    DOCX_CLIPBOARD_MIME,
    "text/html",
    "text/rtf",
    "text/plain",
  ]);
  assert.equal(copied[3][1], "bold link");

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("the paragraph host paste handler forwards every flavour to the controller", async (t) => {
  installClipboard(t);
  const model = starterModel([{ type: "paragraph", children: [{ type: "text", text: "ab" }] }]);
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: model });
      return () =>
        vue.h(DocxParagraphHost, {
          paragraph: controller.model.nodes[0],
          paragraphIndex: 0,
          editable: true,
          controller,
        });
    },
  });
  const mounted = await mount(Harness);
  const host = walk(mounted.root).find((node) => node.props?.["data-docx-paragraph-host"] === true);
  assert.ok(host, "the paragraph host renders");

  const data = {
    [DOCX_CLIPBOARD_MIME]: "",
    "text/html": '<p><b>bold</b></p>',
    "text/plain": "bold",
  };
  let defaultPrevented = false;
  controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 2 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 2 },
  });
  host.props.onPaste({
    clipboardData: { getData: (type) => data[type] ?? "" },
    preventDefault() {
      defaultPrevented = true;
    },
  });
  await vue.nextTick();
  await vue.nextTick();

  assert.equal(defaultPrevented, true);
  assert.equal(paragraphText(controller.model.nodes[0]), "abbold");
  assert.deepEqual(textRuns(controller.model.nodes[0]).at(-1).style, { bold: true });
  assert.equal(controller.status, "Pasted from clipboard");

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("pasting RTF keeps bold, links and lists when no HTML flavour is available", async (t) => {
  const clipboard = installClipboard(t);
  const mounted = await mountController(
    starterModel([{ type: "paragraph", children: [{ type: "text", text: "" }] }]),
  );
  const { controller } = mounted;

  clipboard.readTypes({
    "text/rtf":
      "{\\rtf1\\ansi\\ansicpg1252\\deff0 " +
      "{\\colortbl;\\red0\\green0\\blue255;}" +
      "\\pard {\\b bold} and " +
      '{\\field{\\*\\fldinst{HYPERLINK "https://example.com/r"}}{\\fldrslt{\\ul\\cf1 link}}} ' +
      "\\par " +
      "\\pard{\\pntext\\f0 \\u8226?\\tab}bullet item\\par " +
      "\\pard{\\pntext\\f0 1.\\tab}numbered item\\par}",
    "text/plain": "bold and link\nbullet item\nnumbered item",
  });
  controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 0 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 0 },
  });
  await controller.paste();
  await vue.nextTick();

  const nodes = controller.model.nodes;
  assert.equal(nodes.length, 3);
  assert.equal(textRuns(nodes[0])[0].text, "bold");
  assert.equal(textRuns(nodes[0])[0].style.bold, true);
  const linkRun = textRuns(nodes[0]).find((run) => run.text === "link");
  assert.equal(linkRun.link, "https://example.com/r");
  assert.equal(linkRun.style.underline, true);
  assert.equal(paragraphText(nodes[1]), "bullet item");
  assert.equal(nodes[1].style.numbering.numId > 0, true);
  assert.equal(paragraphText(nodes[2]), "numbered item");

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("the HTML flavour wins over RTF on paste", async (t) => {
  const clipboard = installClipboard(t);
  const mounted = await mountController(
    starterModel([{ type: "paragraph", children: [{ type: "text", text: "" }] }]),
  );
  const { controller } = mounted;

  clipboard.readTypes({
    "text/html": "<p>from html</p>",
    "text/rtf": "{\\rtf1\\ansi from rtf\\par}",
    "text/plain": "plain",
  });
  controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 0 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 0 },
  });
  await controller.paste();
  await vue.nextTick();

  assert.equal(paragraphText(controller.model.nodes[0]), "from html");

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("copy caps drop the rich flavours but keep the plain-text copy", async (t) => {
  const clipboard = installClipboard(t);
  const mounted = await mountController(
    starterModel([{ type: "paragraph", children: [{ type: "text", text: "a".repeat(600 * 1024) }] }]),
  );
  const { controller } = mounted;

  controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 0 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 600 * 1024 },
  });
  await controller.copy();

  assert.equal(clipboard.written.length, 0, "no oversized composite flavour is written");
  assert.equal(clipboard.textWrites.length, 1, "the plain-text copy still reaches the clipboard");
  assert.equal(clipboard.textWrites[0].length, 600 * 1024);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

test("viewer root copy and cut events use the rich flavours", async (t) => {
  installClipboard(t);
  let controller;
  const Harness = vue.defineComponent({
    setup() {
      controller = useDocxEditor({ starterModel: boldLinkModel() });
      return () => vue.h(DocxViewerRoot, { model: controller.model, controller, editable: true });
    },
  });
  const mounted = await mount(Harness);
  const root = walk(mounted.root).find((node) =>
    String(node.props?.class || "").split(/\s+/).includes("docx-viewer-root")
  );
  assert.ok(root, "the viewer root renders");

  selectBoldLink(controller);
  const copied = [];
  let prevented = false;
  root.props.onCopy({
    clipboardData: { setData: (type, value) => copied.push([type, value]) },
    preventDefault() {
      prevented = true;
    },
  });
  assert.equal(prevented, true);
  assert.deepEqual(copied.map(([type]) => type), [
    DOCX_CLIPBOARD_MIME,
    "text/html",
    "text/rtf",
    "text/plain",
  ]);
  assert.equal(copied[3][1], "bold link");

  controller.setActiveTextRange({
    start: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 3 },
    end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: 3 },
  });
  let collapsedCopyPrevented = false;
  root.props.onCopy({
    clipboardData: { setData: () => assert.fail("a collapsed caret must not write") },
    preventDefault() {
      collapsedCopyPrevented = true;
    },
  });
  assert.equal(collapsedCopyPrevented, false, "the native copy keeps running without a selection");

  selectBoldLink(controller);
  const cut = [];
  let cutPrevented = false;
  root.props.onCut({
    clipboardData: { setData: (type, value) => cut.push([type, value]) },
    preventDefault() {
      cutPrevented = true;
    },
  });
  await vue.nextTick();
  assert.equal(cutPrevented, true);
  assert.deepEqual(cut.map(([type]) => type), [
    DOCX_CLIPBOARD_MIME,
    "text/html",
    "text/rtf",
    "text/plain",
  ]);
  assert.equal(paragraphText(controller.model.nodes[0]), "");
  assert.equal(controller.canUndo, true);

  mounted.app.unmount();
  assert.deepEqual(mounted.warnings, []);
});

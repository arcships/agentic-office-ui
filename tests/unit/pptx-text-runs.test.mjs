import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { dirname, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";

const rootDir = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const sourceUrl = (relative) => pathToFileURL(resolve(rootDir, relative)).href;
const resolveHookUrl = `data:text/javascript,${encodeURIComponent(`
  export async function resolve(specifier, context, nextResolve) {
    try {
      return await nextResolve(specifier, context);
    } catch (error) {
      if (specifier.startsWith(".") && !specifier.endsWith(".ts")) {
        return nextResolve(specifier + ".ts", context);
      }
      throw error;
    }
  }
`)}`;
const registerHookUrl = `data:text/javascript,${encodeURIComponent(`
  import { register } from "node:module";
  register(${JSON.stringify(resolveHookUrl)}, import.meta.url);
`)}`;

const modules = {
  trackPlayer: sourceUrl("packages/pptx-core/src/browser/track-player.ts"),
  textRuns: sourceUrl("packages/pptx-core/src/browser/text-runs.ts"),
  tableTargets: sourceUrl("packages/pptx-core/src/browser/table-targets.ts"),
  targets: sourceUrl("packages/pptx-core/src/playback/targets.ts"),
};

const script = `
  import assert from "node:assert/strict";

  const [track, textRuns, tables, targets] = await Promise.all(
    ${JSON.stringify(Object.values(modules))}.map((url) => import(url)),
  );

  const TEXT_NODE = 3;

  class FakeText {
    constructor(document, data) {
      this.ownerDocument = document;
      this.nodeType = TEXT_NODE;
      this.data = data;
      this.parentNode = null;
    }
    get textContent() {
      return this.data;
    }
  }

  class FakeNode {
    constructor(document) {
      this.ownerDocument = document;
      this.parentNode = null;
      this.childNodes = [];
    }
    get children() {
      return this.childNodes.filter((node) => node.nodeType !== TEXT_NODE);
    }
    get parentElement() {
      return this.parentNode && this.parentNode.nodeType !== TEXT_NODE ? this.parentNode : null;
    }
    appendChild(node) {
      if (node.parentNode) node.parentNode.removeChild(node);
      node.parentNode = this;
      this.childNodes.push(node);
      return node;
    }
    insertBefore(node, reference) {
      if (node.parentNode) node.parentNode.removeChild(node);
      const index = reference ? this.childNodes.indexOf(reference) : this.childNodes.length;
      node.parentNode = this;
      this.childNodes.splice(index < 0 ? this.childNodes.length : index, 0, node);
      return node;
    }
    append(...nodes) {
      for (const node of nodes) this.appendChild(node);
    }
    removeChild(node) {
      const index = this.childNodes.indexOf(node);
      if (index >= 0) this.childNodes.splice(index, 1);
      node.parentNode = null;
      return node;
    }
    replaceChild(node, reference) {
      const index = this.childNodes.indexOf(reference);
      if (index >= 0) {
        this.childNodes[index] = node;
        node.parentNode = this;
        reference.parentNode = null;
      }
      return reference;
    }
    get textContent() {
      return this.childNodes.map((node) => node.textContent ?? "").join("");
    }
    set textContent(value) {
      for (const node of this.childNodes) node.parentNode = null;
      this.childNodes = [];
      if (value !== "") this.appendChild(new FakeText(this.ownerDocument, String(value)));
    }
  }

  class FakeElement extends FakeNode {
    constructor(document, localName) {
      super(document);
      this.localName = localName;
      this.nodeType = 1;
      this.dataset = {};
      this.attributes = new Map();
      this.style = {
        visibility: "", display: "", opacity: "", pointerEvents: "", translate: "", scale: "", rotate: "",
        clipPath: "", filter: "", backgroundColor: "", borderColor: "", color: "", whiteSpace: "",
      };
      this.colSpan = 1;
      this.rowSpan = 1;
    }
    getAttribute(name) {
      return this.attributes.has(name) ? this.attributes.get(name) : null;
    }
    setAttribute(name, value) {
      this.attributes.set(name, String(value));
    }
    matchesSelector(selector) {
      if (selector === "table") return this.localName === "table";
      if (selector === "tr") return this.localName === "tr";
      if (selector === "td" || selector === "th") return this.localName === "td" || this.localName === "th";
      const attribute = /^\\[([a-z0-9-]+)\\]$/u.exec(selector);
      if (!attribute) return false;
      const camel = attribute[1].replace(/^data-/u, "").replace(/-([a-z])/gu, (_, letter) => letter.toUpperCase());
      return this.dataset[camel] !== undefined;
    }
    querySelectorAll(selector) {
      const selectorList = selector.split(",").map((item) => item.trim());
      const result = [];
      const visit = (node) => {
        for (const child of node.children) {
          if (selectorList.some((item) => child.matchesSelector(item))) result.push(child);
          visit(child);
        }
      };
      visit(this);
      return result;
    }
    querySelector(selector) {
      return this.querySelectorAll(selector)[0] ?? null;
    }
    closest(selector) {
      let current = this;
      while (current) {
        if (current.matchesSelector?.(selector)) return current;
        current = current.parentNode;
      }
      return null;
    }
  }

  class FakeDocument {
    createElement(localName) {
      return new FakeElement(this, localName);
    }
    createTextNode(data) {
      return new FakeText(this, data);
    }
    createDocumentFragment() {
      return new FakeNode(this);
    }
  }

  function element(document, localName, values = {}) {
    const node = document.createElement(localName);
    Object.assign(node.dataset, values.dataset ?? {});
    for (const [name, value] of Object.entries(values.attributes ?? {})) node.setAttribute(name, value);
    return node;
  }

  function paragraph(document, index, text) {
    const node = element(document, "div", { dataset: { pptxParagraphIndex: String(index) } });
    node.appendChild(new FakeText(document, text));
    return node;
  }

  function textObject(document, objectKey, paragraphs) {
    const object = element(document, "div", { dataset: { pptxObjectKey: objectKey } });
    for (const item of paragraphs) object.appendChild(paragraph(document, item.index, item.text));
    return object;
  }

  function tableObject(document, objectKey) {
    const object = element(document, "div", { dataset: { pptxObjectKey: objectKey } });
    const table = element(document, "table");
    object.appendChild(table);
    const rows = [
      ["R0C0", "R0C1"],
      ["R1C0", "R1C1"],
    ];
    rows.forEach((cells, rowIndex) => {
      const row = element(document, "tr");
      table.appendChild(row);
      cells.forEach((text, columnIndex) => {
        const cell = element(document, "td");
        cell.appendChild(paragraph(document, columnIndex, text));
        row.appendChild(cell);
      });
    });
    return { object, table };
  }

  const document = new FakeDocument();
  globalThis.document = document;
  globalThis.getComputedStyle = (node) => ({
    visibility: node.style.visibility || "visible",
    display: node.style.display || "block",
    opacity: node.style.opacity || "1",
    clipPath: node.style.clipPath || "",
    filter: node.style.filter || "",
    transform: "",
    translate: "",
    scale: "",
    rotate: "",
  });

  const objectKey = "ppt/slides/slide1.xml|slide|shape:3";

  // 1. Word and character boxes carry exact paragraph-local offsets.
  {
    const object = textObject(document, objectKey, [
      { index: 0, text: "Hello world" },
      { index: 1, text: "Second line" },
    ]);
    const layout = textRuns.layoutPptxTextBody(object);
    assert.equal(layout.text, "Hello world\\rSecond line");
    assert.deepEqual([...layout.starts], [0, 12]);
    assert.deepEqual(
      layout.runs[0].words.map((word) => [word.element.textContent, word.start, word.end]),
      [["Hello", 0, 4], ["world", 6, 10]],
    );
    assert.deepEqual(
      layout.runs[0].characters.map((item) => [item.element.textContent, item.start, item.end]),
      [..."Hello world"].flatMap((character, index) => character === " " ? [] : [[character, index, index]]),
    );
    assert.equal(layout.runs[0].paragraph.textContent, "Hello world");
    // Splitting is idempotent.
    const again = textRuns.layoutPptxTextBody(object);
    assert.equal(again.runs[0].characters.length, layout.runs[0].characters.length);
    assert.equal(object.querySelectorAll("[data-pptx-text-run]").length, 24);
  }

  // 2. Whole-word ranges target the word box; partial ranges target characters.
  {
    const object = textObject(document, objectKey, [{ index: 0, text: "Hello world" }]);
    const word = textRuns.resolvePptxTextRange(object, { start: 0, end: 4 });
    assert.equal(word.length, 1);
    assert.equal(word[0].dataset.pptxTextRun, "word");
    assert.equal(word[0].textContent, "Hello");
    const partial = textRuns.resolvePptxTextRange(object, { start: 1, end: 3 });
    assert.deepEqual(partial.map((item) => item.textContent), ["e", "l", "l"]);
    const range = textRuns.resolvePptxTextRange(object, { start: 6, end: 10 });
    assert.equal(range[0].textContent, "world");
    const mixed = textRuns.resolvePptxTextRange(object, { start: 0, end: 8 });
    assert.deepEqual(mixed.map((item) => item.textContent), ["Hello", "w", "o", "r"]);
  }

  // 3. Character targets are resolved by the track index and apply per word.
  {
    const object = textObject(document, objectKey, [{ index: 0, text: "Hello world" }]);
    const root = element(document, "div");
    root.appendChild(object);
    const index = track.createPptxTrackTargets(root, [
      objectKey + "#text:0-4",
      objectKey + "#text:6-10",
    ]);
    assert.equal(index.elements.get(objectKey + "#text:0-4").length, 1);
    assert.equal(index.elements.get(objectKey + "#text:0-4")[0].textContent, "Hello");
    track.applyPptxObjectState(root, {
      [objectKey + "#text:0-4"]: { display: false, opacity: 0 },
    }, index);
    assert.equal(index.elements.get(objectKey + "#text:0-4")[0].style.visibility, "hidden");
    assert.equal(index.elements.get(objectKey + "#text:6-10")[0].style.visibility, "");
  }

  // 4. Table cells expose visual grid positions and row-major paragraph order.
  {
    const { object, table } = tableObject(document, objectKey);
    const cells = tables.pptxRenderedTableCells(object);
    assert.deepEqual(
      cells.map((cell) => [cell.row, cell.column, cell.paragraphs.length]),
      [[0, 0, 1], [0, 1, 1], [1, 0, 1], [1, 1, 1]],
    );
    assert.deepEqual(
      tables.pptxTableParagraphs(object).map((item) => item.textContent),
      ["R0C0", "R0C1", "R1C0", "R1C1"],
    );
    const root = element(document, "div");
    root.appendChild(object);
    const index = track.createPptxTrackTargets(root);
    assert.equal(index.elements.get(objectKey + "#cell:0:1")[0], cells[1].element);
    assert.equal(index.elements.get(objectKey + "#paragraph:2-2")[0].textContent, "R1C0");
    track.applyPptxObjectState(root, {
      [objectKey + "#cell:1:0"]: { display: false },
    }, index);
    assert.equal(cells[2].element.style.visibility, "hidden");
  }

  // 5. Row-span and col-span cells expand into the visual grid.
  {
    const object = element(document, "div", { dataset: { pptxObjectKey: objectKey } });
    const table = element(document, "table");
    const firstRow = element(document, "tr");
    const spanning = element(document, "td");
    spanning.colSpan = 2;
    spanning.appendChild(paragraph(document, 0, "span"));
    firstRow.appendChild(spanning);
    const firstTail = element(document, "td");
    firstTail.appendChild(paragraph(document, 0, "tail"));
    firstRow.appendChild(firstTail);
    const secondRow = element(document, "tr");
    const secondLead = element(document, "td");
    secondLead.rowSpan = 2;
    secondLead.appendChild(paragraph(document, 0, "lead"));
    secondRow.appendChild(secondLead);
    const secondCell = element(document, "td");
    secondCell.appendChild(paragraph(document, 0, "second"));
    secondRow.appendChild(secondCell);
    const thirdRow = element(document, "tr");
    const thirdCell = element(document, "td");
    thirdCell.appendChild(paragraph(document, 0, "third"));
    thirdRow.appendChild(thirdCell);
    table.append(firstRow, secondRow, thirdRow);
    object.appendChild(table);
    const cells = tables.pptxRenderedTableCells(object);
    assert.deepEqual(
      cells.map((cell) => [cell.row, cell.column, cell.element.textContent]),
      [
        [0, 0, "span"],
        [0, 2, "tail"],
        [1, 0, "lead"],
        [1, 1, "second"],
        [2, 1, "third"],
      ],
    );
  }

  // 6. SmartArt child and whole-chart element targets resolve from identity.
  {
    const object = element(document, "div", { dataset: { pptxObjectKey: objectKey } });
    const child = element(document, "div", { dataset: { pptxNodeId: "12" } });
    object.appendChild(child);
    const chartObject = element(document, "div", {
      dataset: { pptxObjectKey: "ppt/slides/slide1.xml|slide|shape:7" },
    });
    const root = element(document, "div");
    root.append(object, chartObject);
    const index = track.createPptxTrackTargets(root, [
      objectKey + "#graphic:dgm:12",
      "ppt/slides/slide1.xml|slide|shape:7#chart:chart",
      "ppt/slides/slide1.xml|slide|shape:7#chart:series:1",
    ]);
    assert.equal(index.elements.get(objectKey + "#graphic:dgm:12")[0], child);
    assert.equal(index.elements.get("ppt/slides/slide1.xml|slide|shape:7#chart:chart")[0], chartObject);
    assert.equal(index.elements.has("ppt/slides/slide1.xml|slide|shape:7#chart:series:1"), false);
  }

  // 7. Target helpers keep stable selectors for every element target.
  {
    assert.equal(
      targets.pptxEffectTargetKey({ id: "1", kind: "fade-in", targetObjectKey: objectKey, characterRange: { start: 2, end: 4 }, values: {} }),
      objectKey + "#text:2-4",
    );
    assert.equal(
      targets.pptxEffectTargetKey({ id: "2", kind: "fade-in", targetObjectKey: objectKey, tableCell: { row: 1, column: 2 }, values: {} }),
      objectKey + "#cell:1:2",
    );
    assert.equal(
      targets.pptxEffectTargetKey({ id: "3", kind: "fade-in", targetObjectKey: objectKey, chartElement: { type: "series", index: 1 }, values: {} }),
      objectKey + "#chart:series:1",
    );
    assert.equal(targets.pptxTargetObjectKey(objectKey + "#chart:series:1"), objectKey);
  }

  process.stdout.write(JSON.stringify({ status: "PASS" }));
`;

test("PPTX text runs, table cells and element targets resolve on the rendered DOM", () => {
  const result = spawnSync(process.execPath, [
    "--experimental-strip-types",
    "--import",
    registerHookUrl,
    "--input-type=module",
    "--eval",
    script,
  ], {
    cwd: rootDir,
    encoding: "utf8",
    env: { ...process.env, NODE_NO_WARNINGS: "1" },
  });
  assert.equal(result.status, 0, result.stderr || result.stdout);
  assert.equal(JSON.parse(result.stdout).status, "PASS");
});

import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { dirname, resolve } from "node:path";
import test from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";

const rootDir = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const sourceUrl = (relative) => pathToFileURL(resolve(rootDir, relative)).href;
// Node cannot resolve the source package's extensionless relative imports, so
// the child process registers the same resolver hook the other PPTX unit tests
// use and strips the TypeScript types itself.
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
const pptxCorePackageUrl = pathToFileURL(
  resolve(rootDir, "packages/pptx-core/package.json"),
).href;
const xlsxCorePackageUrl = pathToFileURL(
  resolve(rootDir, "packages/xlsx-core/package.json"),
).href;

function runNode(args) {
  return spawnSync(process.execPath, args, {
    cwd: rootDir,
    encoding: "utf8",
    env: { ...process.env, NODE_NO_WARNINGS: "1" },
  });
}

const notesScript = `
  import assert from "node:assert/strict";
  import { createRequire } from "node:module";

  const requireFromPptxCore = createRequire(${JSON.stringify(pptxCorePackageUrl)});
  const { strToU8, zipSync } = requireFromPptxCore("fflate");
  const { DOMParser } = createRequire(${JSON.stringify(xlsxCorePackageUrl)})("@xmldom/xmldom");
  globalThis.DOMParser = DOMParser;
  const domParser = new DOMParser();

  const notes = await import(${JSON.stringify(sourceUrl("packages/pptx-core/src/notes.ts"))});

  const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
  const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
  const LIMITS = { maxSingleXmlBytes: 1048576, maxXmlDepth: 64 };

  function notesXml(paragraphs) {
    const body = paragraphs
      .map((paragraph) => \`<a:p><a:r><a:rPr lang="en-US"/><a:t>\${paragraph}</a:t></a:r></a:p>\`)
      .join("");
    return \`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <p:notes xmlns:p="\${NS_P}" xmlns:a="\${NS_A}"><p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr/><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>
    <p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>\${body}</p:txBody></p:sp>
    <p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Number Placeholder 3"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum" sz="quarter" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{2}" type="slidenum"><a:rPr/><a:t>7</a:t></a:fld></a:p></p:txBody></p:sp>
    </p:spTree></p:cSld></p:notes>\`;
  }

  const brokenXml = \`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <p:notes xmlns:p="\${NS_P}" xmlns:a="\${NS_A}"><p:cSld><p:spTree><p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
  <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>
  <a:p><a:r><a:t>Slide </a:t></a:r><a:fld id="{1}" type="slidenum"><a:t>3</a:t></a:fld><a:br/><a:r><a:t>is next</a:t></a:r></a:p>
  </p:txBody></p:sp></p:spTree></p:cSld></p:notes>\`;

  function relsXml(target, { type = "notesSlide", targetMode } = {}) {
    const mode = targetMode ? \` TargetMode="\${targetMode}"\` : "";
    return \`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="\${NS_REL}"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/\${type}" Target="\${target}"\${mode}/></Relationships>\`;
  }

  function archive(parts) {
    const entries = {};
    for (const [name, text] of Object.entries(parts)) entries[name] = strToU8(text);
    const buffer = zipSync(entries);
    return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
  }

  // 1. Notes text: content paragraphs survive, image/number placeholders do not.
  assert.equal(
    notes.parsePptxNotesXml(notesXml(["Open with the quarter", "Then pause."]), domParser),
    "Open with the quarter\\nThen pause.",
  );
  assert.equal(notes.parsePptxNotesXml(notesXml([]), domParser), "");
  assert.equal(notes.parsePptxNotesXml(brokenXml, domParser), "Slide 3\\nis next");

  // 2. Notes part paths come from the slide relationships and stay inside ppt/.
  const rels = new Map([
    ["ppt/slides/_rels/slide1.xml.rels", relsXml("../notesSlides/notesSlide1.xml")],
    ["ppt/slides/_rels/slide10.xml.rels", relsXml("/ppt/notesSlides/notesSlide10.xml")],
    ["ppt/slides/_rels/slide2.xml.rels", relsXml("../../../escape.xml")],
    ["ppt/slides/_rels/slide3.xml.rels", relsXml("https://example.test/notes.xml", { targetMode: "External" })],
    ["ppt/slides/_rels/slide4.xml.rels", relsXml("../notesSlides/notesSlide4.xml", { type: "slideLayout" })],
  ]);
  assert.equal(
    notes.pptxNotesPartPath("ppt/slides/slide1.xml", rels, domParser),
    "ppt/notesSlides/notesSlide1.xml",
  );
  assert.equal(notes.pptxNotesPartPath("ppt/slides/slide2.xml", rels, domParser), undefined);
  assert.equal(notes.pptxNotesPartPath("ppt/slides/slide3.xml", rels, domParser), undefined);
  assert.equal(notes.pptxNotesPartPath("ppt/slides/slide4.xml", rels, domParser), undefined);
  assert.equal(notes.pptxNotesPartPath("ppt/slides/slide9.xml", rels, domParser), undefined);
  assert.equal(
    notes.pptxNotesPartPath("ppt/slides/slide10.xml", rels, domParser),
    "ppt/notesSlides/notesSlide10.xml",
    "a package-absolute target resolves too",
  );

  // 3. Only slides with a notes relationship get notes.
  {
    const buffer = archive({
      "ppt/notesSlides/notesSlide1.xml": notesXml(["First slide notes"]),
      "ppt/notesSlides/notesSlide2.xml": notesXml(["Second slide notes"]),
    });
    const slideRels = new Map([
      ["ppt/slides/_rels/slide1.xml.rels", relsXml("../notesSlides/notesSlide1.xml")],
      ["ppt/slides/_rels/slide2.xml.rels", relsXml("../slideLayouts/slideLayout2.xml", { type: "slideLayout" })],
      ["ppt/slides/_rels/slide3.xml.rels", relsXml("../notesSlides/notesSlide2.xml")],
    ]);
    const read = await notes.readPptxSpeakerNotes(
      buffer,
      [
        { slideIndex: 0, slidePath: "ppt/slides/slide1.xml" },
        { slideIndex: 1, slidePath: "ppt/slides/slide2.xml" },
        { slideIndex: 2, slidePath: "ppt/slides/slide3.xml" },
      ],
      { slideRels, limits: LIMITS },
    );
    assert.deepEqual([...read.entries()], [[0, "First slide notes"], [2, "Second slide notes"]]);
  }

  // 4. A deck without notes relationships never re-opens the container.
  {
    let reads = 0;
    const read = await notes.readPptxSpeakerNotes(
      new ArrayBuffer(8),
      [{ slideIndex: 0, slidePath: "ppt/slides/slide1.xml" }],
      {
        slideRels: new Map([
          ["ppt/slides/_rels/slide1.xml.rels", relsXml("../slideLayouts/slideLayout1.xml", { type: "slideLayout" })],
        ]),
        limits: LIMITS,
        reader: {
          async read() {
            reads += 1;
            return new Map();
          },
        },
      },
    );
    assert.equal(reads, 0);
    assert.equal(read.size, 0);
  }

  // 5. An unreadable notes part only drops that slide's notes.
  {
    const buffer = archive({
      "ppt/notesSlides/notesSlide1.xml": "<p:notes>truncated",
      "ppt/notesSlides/notesSlide2.xml": notesXml(["Still readable"]),
    });
    const slideRels = new Map([
      ["ppt/slides/_rels/slide1.xml.rels", relsXml("../notesSlides/notesSlide1.xml")],
      ["ppt/slides/_rels/slide2.xml.rels", relsXml("../notesSlides/notesSlide2.xml")],
    ]);
    const read = await notes.readPptxSpeakerNotes(
      buffer,
      [
        { slideIndex: 0, slidePath: "ppt/slides/slide1.xml" },
        { slideIndex: 1, slidePath: "ppt/slides/slide2.xml" },
      ],
      { slideRels, limits: LIMITS },
    );
    assert.deepEqual([...read.entries()], [[1, "Still readable"]]);
  }

  // 6. Budgets and cancellations still fail the open request.
  {
    const buffer = archive({ "ppt/notesSlides/notesSlide1.xml": notesXml(["Budget me"]) });
    const slides = [{ slideIndex: 0, slidePath: "ppt/slides/slide1.xml" }];
    const slideRels = new Map([
      ["ppt/slides/_rels/slide1.xml.rels", relsXml("../notesSlides/notesSlide1.xml")],
    ]);
    await assert.rejects(
      notes.readPptxSpeakerNotes(buffer, slides, { slideRels, limits: { maxSingleXmlBytes: 16 } }),
      (error) => error.code === "LIMIT_EXCEEDED",
    );
    const controller = new AbortController();
    controller.abort();
    await assert.rejects(
      notes.readPptxSpeakerNotes(buffer, slides, { slideRels, limits: LIMITS, signal: controller.signal }),
      (error) => error.code === "ABORTED",
    );
  }

  // 7. A container that cannot be re-read degrades to no notes.
  {
    const read = await notes.readPptxSpeakerNotes(
      archive({}),
      [{ slideIndex: 0, slidePath: "ppt/slides/slide1.xml" }],
      {
        slideRels: new Map([
          ["ppt/slides/_rels/slide1.xml.rels", relsXml("../notesSlides/notesSlide1.xml")],
        ]),
        limits: LIMITS,
        reader: {
          async read() {
            throw new Error("zip64 archives are unsupported");
          },
        },
      },
    );
    assert.equal(read.size, 0);
  }

  process.stdout.write(JSON.stringify({ status: "PASS" }));
`;

test("PPTX speaker notes are parsed from the notes parts with archive budgets intact", () => {
  const result = runNode([
    "--experimental-strip-types",
    "--import",
    registerHookUrl,
    "--input-type=module",
    "--eval",
    notesScript,
  ]);
  assert.equal(result.status, 0, result.stderr || result.stdout);
  assert.equal(JSON.parse(result.stdout).status, "PASS");
});

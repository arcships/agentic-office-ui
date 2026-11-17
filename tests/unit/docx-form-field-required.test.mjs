import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { test } from "node:test";
import { pathToFileURL } from "node:url";

const requireFromVueDocx = createRequire(
  new URL("../../packages/vue-docx/package.json", import.meta.url),
);
const requireFromDocxCore = createRequire(
  new URL("../../packages/docx-core/package.json", import.meta.url),
);
const core = await import(
  pathToFileURL(requireFromVueDocx.resolve("@arcships/docx-core")).href
);
const { strFromU8, unzipSync } = await import(
  pathToFileURL(requireFromDocxCore.resolve("fflate")).href
);

function legacyTextFieldDocumentXml({ required }) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Employee"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput><w:maxLength w:val="20"/></w:textInput>${required ? "<w:required/>" : ""}</w:ffData></w:fldChar></w:r>
      <w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>Alice</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>`;
}

async function requiredModel(required) {
  const pkg = core.createMinimalDocxPackage(legacyTextFieldDocumentXml({ required }));
  const model = await core.buildDocModel(pkg);
  const field = model.nodes[0].children[0];
  assert.equal(field.type, "form-field");
  return { pkg, model, field };
}

test("legacy form fields parse w:ffData/w:required into FormFieldRunNode", async () => {
  const required = await requiredModel(true);
  assert.equal(required.field.required, true);
  assert.equal(required.field.widget?.name, "Employee");

  const optional = await requiredModel(false);
  assert.equal(optional.field.required, undefined);
});

test("serialized DOCX preserves required on rebuilt form fields", async () => {
  const { model, field } = await requiredModel(true);
  field.value = "Ada";
  // Editing clears sourceXml so the serializer rebuilds the field from model
  // properties; the required flag must still reach the output.
  field.sourceXml = undefined;
  model.nodes[0].sourceXml = undefined;

  const documentXml = await core.modelToDocumentXml(model);
  assert.match(documentXml, /<w:ffData>[\s\S]*<w:required\/>[\s\S]*<\/w:ffData>/);
  assert.match(documentXml, /<w:t>Ada<\/w:t>/);

  const pkg = await core.serializeDocModel(model);
  assert.match(
    pkg.parts.get("word/document.xml").content,
    /<w:ffData>[\s\S]*<w:required\/>[\s\S]*<\/w:ffData>/,
  );

  const bytes = await core.serializeDocx(model);
  const archive = unzipSync(new Uint8Array(bytes));
  assert.match(
    strFromU8(archive["word/document.xml"]),
    /<w:ffData>[\s\S]*<w:required\/>[\s\S]*<\/w:ffData>/,
  );
});

test("serialized DOCX removes w:required when the field is no longer required", async () => {
  const { model, field } = await requiredModel(true);
  field.required = false;
  field.value = "Ada";
  field.sourceXml = undefined;
  model.nodes[0].sourceXml = undefined;

  const documentXml = await core.modelToDocumentXml(model);
  assert.doesNotMatch(documentXml, /<w:required\b/);
});

test("serializer maps required flags to the matching ffData blocks", async () => {
  const field = (name, required, text) =>
    `<w:p><w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="${name}"/><w:enabled/><w:textInput/>${required ? "<w:required/>" : ""}</w:ffData></w:fldChar></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t>${text}</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>${field("Employee", true, "Alice")}${field("Nickname", false, "Al")}</w:body>
</w:document>`;

  const model = await core.buildDocModel(core.createMinimalDocxPackage(documentXml));
  assert.equal(model.nodes[0].children[0].required, true);
  assert.equal(model.nodes[1].children[0].required, undefined);
  for (const node of model.nodes) {
    node.children[0].sourceXml = undefined;
    node.sourceXml = undefined;
  }

  const xml = await core.modelToDocumentXml(model);
  const blocks = [...xml.matchAll(/<w:ffData\b[^>]*>[\s\S]*?<\/w:ffData>/g)].map((match) => match[0]);
  assert.equal(blocks.length, 2);
  assert.match(blocks[0], /<w:required\/>/);
  assert.doesNotMatch(blocks[1], /<w:required\/>/);
});

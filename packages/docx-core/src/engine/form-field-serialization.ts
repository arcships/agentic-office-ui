// Form-field serialization support.
//
// The WASM serializer rebuilds legacy form fields from the common model
// properties but does not know about `w:ffData/w:required`. Patch the emitted
// document XML in model order so required flags survive edits (which clear the
// original `sourceXml`, forcing the serializer to rebuild the field).

import type { DocModel, FormFieldRunNode, ParagraphNode, TableNode } from "./types";

const FF_DATA_BLOCK_PATTERN = /<w:ffData\b[^>]*>[\s\S]*?<\/w:ffData>/g;
const REQUIRED_SELF_CLOSING_PATTERN = /<w:required\b[^>]*\/>/g;
const REQUIRED_PAIRED_PATTERN = /<w:required\b[^>]*>[\s\S]*?<\/w:required>/g;

function collectParagraphFields(
  paragraph: ParagraphNode,
  collected: FormFieldRunNode[],
): void {
  for (const child of paragraph.children) {
    if (child.type === "form-field" && child.sourceKind !== "sdt") {
      collected.push(child);
    }
  }
}

function collectTableFields(
  table: TableNode,
  collected: FormFieldRunNode[],
): void {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      for (const node of cell.nodes) {
        if (node.type === "paragraph") {
          collectParagraphFields(node, collected);
        } else {
          collectTableFields(node, collected);
        }
      }
    }
  }
}

export function collectLegacyFormFieldsInDocumentOrder(
  model: DocModel,
): FormFieldRunNode[] {
  const collected: FormFieldRunNode[] = [];
  for (const node of model.nodes) {
    if (node.type === "paragraph") {
      collectParagraphFields(node, collected);
    } else {
      collectTableFields(node, collected);
    }
  }
  return collected;
}

function patchRequiredFlag(block: string, required: boolean): string {
  const hasRequired = REQUIRED_SELF_CLOSING_PATTERN.test(block) ||
    REQUIRED_PAIRED_PATTERN.test(block);
  REQUIRED_SELF_CLOSING_PATTERN.lastIndex = 0;
  REQUIRED_PAIRED_PATTERN.lastIndex = 0;

  if (required) {
    if (!hasRequired) {
      return block.replace(/<\/w:ffData>$/, "<w:required/></w:ffData>");
    }
    // Normalise an explicitly disabled flag (`w:val="0"`) to an enabled one.
    return block
      .replace(REQUIRED_SELF_CLOSING_PATTERN, "<w:required/>")
      .replace(REQUIRED_PAIRED_PATTERN, "<w:required/>");
  }

  if (!hasRequired) return block;
  return block
    .replace(REQUIRED_SELF_CLOSING_PATTERN, "")
    .replace(REQUIRED_PAIRED_PATTERN, "");
}

export function applyFormFieldRequiredToDocumentXml(
  xml: string,
  model: DocModel,
): string {
  const fields = collectLegacyFormFieldsInDocumentOrder(model);
  if (fields.length === 0) return xml;

  let index = 0;
  return xml.replace(FF_DATA_BLOCK_PATTERN, (block) => {
    if (index >= fields.length) return block;
    const required = fields[index++]?.required === true;
    return patchRequiredFlag(block, required);
  });
}

export function modelNeedsFormFieldRequiredPatch(model: DocModel): boolean {
  return collectLegacyFormFieldsInDocumentOrder(model).some((field) =>
    field.required !== undefined || /<w:required\b/i.test(field.sourceXml ?? ""),
  );
}

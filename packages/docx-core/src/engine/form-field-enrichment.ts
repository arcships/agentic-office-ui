// Post-import form-field enrichment.
//
// The WASM model builder exposes the common widget settings but drops the
// legacy `w:ffData/w:required` flag. The original field XML is still available
// on imported form fields, so recover the flag once after the model is built.
// This keeps `required` on FormFieldRunNode without changing the WASM ABI.

import type { DocModel, FormFieldRunNode, ParagraphNode, TableNode } from "./types";
import { extractBalancedTagRanges, xmlAttribute } from "./text-typography";

function firstTagBlock(xml: string, tagName: string): string | undefined {
  const range = extractBalancedTagRanges(xml, tagName)[0];
  return range ? xml.slice(range.start, range.end) : undefined;
}

function onOffValue(tagXml: string | undefined): boolean {
  if (!tagXml) return false;
  const value = xmlAttribute(tagXml, "w:val")?.trim().toLowerCase();
  return value !== "0" && value !== "false" && value !== "off";
}

function applyRequiredFlag(field: FormFieldRunNode): void {
  if (field.sourceKind === "sdt" || !field.sourceXml) return;
  const ffData = firstTagBlock(field.sourceXml, "w:ffData");
  if (!ffData) return;
  const requiredTag = firstTagBlock(ffData, "w:required");
  if (requiredTag) {
    field.required = onOffValue(requiredTag);
  }
}

function enrichParagraph(paragraph: ParagraphNode): void {
  for (const child of paragraph.children) {
    if (child.type === "form-field") {
      applyRequiredFlag(child);
    }
  }
}

function walkTable(table: TableNode): void {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      for (const node of cell.nodes) {
        if (node.type === "paragraph") {
          enrichParagraph(node);
        } else {
          walkTable(node);
        }
      }
    }
  }
}

export function enrichDocModelFormFields(model: DocModel): void {
  for (const node of model.nodes) {
    if (node.type === "paragraph") {
      enrichParagraph(node);
    } else {
      walkTable(node);
    }
  }
}

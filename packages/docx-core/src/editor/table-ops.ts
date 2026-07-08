// Table cell text editing operations.
// Split from editor-ops.ts to stay within the ≤1000-line constraint.

import type {
  DocModel,
  ParagraphNode,
  TableCellContentNode,
  TableNode,
} from "../engine/types";
import { cloneDocModel } from "../engine/clone";
import type { UpdateTextOptions } from "./paragraph-ops";
import {
  distributeTextAcrossParagraphChildren,
  getParagraph,
  paragraphFromText,
} from "./paragraph-ops";
import { cloneTextStyle } from "./helpers/text-mutation";

export function updateTableCellText(
  model: DocModel,
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  text: string,
  options?: UpdateTextOptions
): DocModel {
  const next = cloneDocModel(model);
  const tableNode = next.nodes[tableIndex];
  if (!tableNode || tableNode.type !== "table") {
    return next;
  }

  const row = tableNode.rows[rowIndex];
  const cell = row?.cells[cellIndex];
  if (!cell) {
    return next;
  }

  const paragraphs = cell.nodes.filter((node): node is ParagraphNode => node.type === "paragraph");
  const incomingParagraphTexts = text
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n");

  if (incomingParagraphTexts.length === 0) {
    incomingParagraphTexts.push("");
  }

  if (paragraphs.length === 0) {
    cell.nodes.push(paragraphFromText(incomingParagraphTexts[0] ?? "", {
      runStyle: cloneTextStyle(options?.insertedStyle)
    }));
    tableNode.sourceXml = undefined;
    return next;
  }

  paragraphs.forEach((paragraph, paragraphIndex) => {
    const paragraphText = incomingParagraphTexts[paragraphIndex] ?? "";
    paragraph.children = distributeTextAcrossParagraphChildren(paragraph, paragraphText, options);
    paragraph.sourceXml = undefined;
  });

  if (incomingParagraphTexts.length > paragraphs.length) {
    for (let paragraphIndex = paragraphs.length; paragraphIndex < incomingParagraphTexts.length; paragraphIndex += 1) {
      cell.nodes.push(paragraphFromText(incomingParagraphTexts[paragraphIndex] ?? "", {
        runStyle: cloneTextStyle(options?.insertedStyle)
      }));
    }
  }

  tableNode.sourceXml = undefined;
  return next;
}

export function updateTableCellParagraphText(
  model: DocModel,
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  paragraphIndex: number,
  text: string,
  options?: UpdateTextOptions
): DocModel {
  const next = cloneDocModel(model);
  const tableNode = next.nodes[tableIndex];
  if (!tableNode || tableNode.type !== "table") {
    return next;
  }

  const row = tableNode.rows[rowIndex];
  const cell = row?.cells[cellIndex];
  const paragraph = cell?.nodes.filter((node): node is ParagraphNode => node.type === "paragraph")[paragraphIndex];

  if (!cell || !paragraph) {
    return next;
  }

  paragraph.children = distributeTextAcrossParagraphChildren(paragraph, text, options);
  paragraph.sourceXml = undefined;
  tableNode.sourceXml = undefined;

  return next;
}

export function updateTableCellParagraphTextRecursive(
  model: DocModel,
  tableIndex: number,
  rowIndex: number,
  cellIndex: number,
  paragraphIndex: number,
  text: string,
  options?: UpdateTextOptions
): DocModel {
  const next = cloneDocModel(model);
  const tableNode = next.nodes[tableIndex];
  if (!tableNode || tableNode.type !== "table") {
    return next;
  }

  const row = tableNode.rows[rowIndex];
  const cell = row?.cells[cellIndex];
  if (!cell) {
    return next;
  }

  const targetParagraphIndex = Math.max(0, Math.round(paragraphIndex));
  let paragraphCursor = 0;

  const updateInNodes = (
    nodes: TableCellContentNode[],
    ancestorTables: TableNode[]
  ): boolean => {
    for (const node of nodes) {
      if (node.type === "paragraph") {
        if (paragraphCursor !== targetParagraphIndex) {
          paragraphCursor += 1;
          continue;
        }

        node.children = distributeTextAcrossParagraphChildren(node, text, options);
        node.sourceXml = undefined;
        ancestorTables.forEach((ancestorTable) => {
          ancestorTable.sourceXml = undefined;
        });
        return true;
      }

      ancestorTables.push(node);
      for (const nestedRow of node.rows) {
        for (const nestedCell of nestedRow.cells) {
          if (updateInNodes(nestedCell.nodes, ancestorTables)) {
            ancestorTables.pop();
            return true;
          }
        }
      }
      ancestorTables.pop();
    }

    return false;
  };

  updateInNodes(cell.nodes, [tableNode]);
  return next;
}

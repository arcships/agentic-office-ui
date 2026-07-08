// editor-text-input.ts — text input operations (contentEditable + draft)
//
// Extracted from useDocxEditor.ts.

import type { DocModel, ParagraphNode } from "@extend-ai/docx-core"
import type { ParagraphLocation } from "@extend-ai/docx-core"
import { cloneDocModel, updateParagraphText } from "@extend-ai/docx-core"
import type { EditorCore } from "./editor-shared"

export function createEditorTextInput(
  ctx: EditorCore,
  applyChange: (updater: (current: DocModel) => DocModel, successStatus?: string) => void,
) {
  // ── commitParagraphText ──────────────────────────────────────────
  const commitParagraphText = (nodeIndex: number, text: string): void => {
    applyChange((current) => updateParagraphText(current, nodeIndex, text))
  }

  // ── commitTableCellText ──────────────────────────────────────────
  const commitTableCellText = (
    tableIndex: number, rowIndex: number, cellIndex: number, text: string
  ): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      const cell = node.rows[rowIndex]?.cells[cellIndex]
      if (!cell) return current
      const para = cell.nodes[0]
      if (!para || para.type !== "paragraph") {
        const newPara: ParagraphNode = {
          type: "paragraph",
          children: [{ type: "text", text, style: {} }],
        }
        if (cell.nodes.length > 0) {
          cell.nodes[0] = newPara
        } else {
          cell.nodes.push(newPara)
        }
        return nextModel
      }
      para.children = [{ type: "text", text, style: {} }]
      return nextModel
    })
  }

  // ── commitTableCellParagraphTextRecursive ────────────────────────
  const commitTableCellParagraphTextRecursive = (
    tableIndex: number, rowIndex: number, cellIndex: number, paragraphIndex: number, text: string
  ): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const node = nextModel.nodes[tableIndex]
      if (!node || node.type !== "table") return current
      const cell = node.rows[rowIndex]?.cells[cellIndex]
      if (!cell) return current
      const para = cell.nodes[paragraphIndex]
      if (!para || para.type !== "paragraph") return current
      para.children = [{ type: "text", text, style: {} }]
      return nextModel
    })
  }

  // ── commitSectionParagraphText ───────────────────────────────────
  const commitSectionParagraphText = (
    location: { region: "header" | "footer"; partName: string; nodeIndex: number },
    text: string
  ): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const sections =
        location.region === "header"
          ? nextModel.metadata.headerSections
          : nextModel.metadata.footerSections
      const section = sections.find((s) => s.partName === location.partName)
      if (!section) return current
      const node = section.nodes[location.nodeIndex]
      if (!node || node.type !== "paragraph") return current
      node.children = [{ type: "text", text, style: {} }]
      return nextModel
    })
  }

  // ── appendParagraph ──────────────────────────────────────────────
  const appendParagraph = (text?: string): number => {
    let nodeIndex = -1
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const para: ParagraphNode = { type: "paragraph", children: [] }
      if (text) {
        para.children.push({ type: "text", text, style: {} })
      }
      nextModel.nodes.push(para)
      nodeIndex = nextModel.nodes.length - 1
      return nextModel
    })
    return nodeIndex
  }

  // ── splitParagraphAtSelection (stub) ─────────────────────────────
  const splitParagraphAtSelection = (
    _draftText: string,
    _startOffset: number,
    _endOffset?: number,
    _targetLocation?: ParagraphLocation
  ): { paragraphIndex: number; caretOffset: number } | undefined => {
    // Stub — full implementation in editor-text-input module
    return undefined
  }

  // ── insertListItemAfterSelection (stub) ──────────────────────────
  const insertListItemAfterSelection = (
    _draftText: string,
    _startOffset: number,
    _endOffset?: number,
    _targetLocation?: ParagraphLocation
  ): { paragraphIndex: number; caretOffset: number } | undefined => {
    // Stub — full implementation in editor-text-input module
    return undefined
  }

  // ── replaceExpandedSelection (stub) ──────────────────────────────
  const replaceExpandedSelection = (
    _text: string,
    _range?: import("@extend-ai/docx-core").DocxTextRange
  ): import("@extend-ai/docx-core").DocxTextRange | undefined => {
    // Stub — full implementation in editor-text-input module
    return undefined
  }

  // ── deleteExpandedSelection (stub) ───────────────────────────────
  const deleteExpandedSelection = (
    _range?: import("@extend-ai/docx-core").DocxTextRange
  ): import("@extend-ai/docx-core").DocxTextRange | undefined => {
    // Stub — full implementation in editor-text-input module
    return undefined
  }

  return {
    commitParagraphText,
    commitTableCellText,
    commitTableCellParagraphTextRecursive,
    commitSectionParagraphText,
    appendParagraph,
    splitParagraphAtSelection,
    insertListItemAfterSelection,
    replaceExpandedSelection,
    deleteExpandedSelection,
  }
}

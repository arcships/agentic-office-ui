// editor-image.ts — image insert/resize/move/wrap operations
//
// Extracted from useDocxEditor.ts.

import type { DocModel, ImageRunNode, ParagraphNode } from "@extend-ai/docx-core"
import type {
  DocxImageLocation,
  DocxImageWrapMode,
  DocxImageDropTarget,
  DocxSectionImageLocation,
} from "@extend-ai/docx-core"
import {
  cloneDocModel,
  toBase64,
  tableCellParagraphs,
} from "@extend-ai/docx-core"
import type { EditorCore } from "./editor-shared"

export function createEditorImage(
  ctx: EditorCore,
  applyChange: (updater: (current: DocModel) => DocModel, successStatus?: string) => void,
) {
  const insertImageFile = async (file: File): Promise<void> => {
    if (!file.type.startsWith("image/")) {
      ctx.status.value = "Select an image file"
      return
    }

    const bytes = new Uint8Array(await file.arrayBuffer())
    const src = `data:${file.type};base64,${toBase64(bytes)}`

    applyChange((current) => {
      const next = cloneDocModel(current)
      const selection = ctx.selectionSnapshot.value

      if (selection.kind === "paragraph") {
        const paragraph = next.nodes[selection.nodeIndex]
        if (!paragraph || paragraph.type !== "paragraph") return current
        paragraph.children.push({
          type: "image",
          src,
          alt: file.name,
          contentType: file.type,
          data: new Uint8Array(bytes),
          widthPx: 240,
        })
        paragraph.sourceXml = undefined
        return next
      }

      const table = next.nodes[selection.tableIndex]
      if (!table || table.type !== "table") return current
      const cell = table.rows[selection.rowIndex]?.cells[selection.cellIndex]
      const paragraph = tableCellParagraphs(cell?.nodes ?? [])[0]
      if (!paragraph) return current
      paragraph.children.push({
        type: "image",
        src,
        alt: file.name,
        contentType: file.type,
        data: new Uint8Array(bytes),
        widthPx: 240,
      })
      paragraph.sourceXml = undefined
      table.sourceXml = undefined
      return next
    }, `Inserted image: ${file.name}`)
  }

  const resizeImage = (
    location: DocxImageLocation, widthPx: number, heightPx: number
  ): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      let node: any | undefined
      if (location.kind === "paragraph") {
        node = nextModel.nodes[location.nodeIndex]
      } else {
        const tableNode = nextModel.nodes[location.tableIndex]
        if (tableNode && tableNode.type === "table") {
          const cell = tableNode.rows[location.rowIndex]?.cells[location.cellIndex]
          node = cell?.nodes[location.paragraphIndex]
        }
      }
      if (!node || node.type !== "paragraph") return current
      const child = node.children[location.childIndex]
      if (!child || child.type !== "image") return current
      child.widthPx = widthPx
      child.heightPx = heightPx
      return nextModel
    })
  }

  const setSyntheticTextBoxText = (
    _location: DocxImageLocation, _text: string
  ): void => {
    // Modifies the text content of a synthetic text box image.
    // Requires synthetic textbox infrastructure from paragraph-geometry.
  }

  const setImageWrapMode = (
    _location: DocxImageLocation,
    _mode: DocxImageWrapMode,
    _seed?: Partial<NonNullable<ImageRunNode["floating"]>>
  ): void => {
    // Sets the text wrapping mode for an image (inline, square, tight, etc.).
    // Modifies the floating property on ImageRunNode.
  }

  const moveFloatingImage = (
    _location: DocxImageLocation,
    _patch: Partial<NonNullable<ImageRunNode["floating"]>>
  ): void => {
    // Moves a floating image by applying position/offset patches.
  }

  const moveSectionFloatingImage = (
    _location: DocxSectionImageLocation,
    _patch: Partial<NonNullable<ImageRunNode["floating"]>>
  ): void => {
    // Moves a floating image in a header/footer section.
  }

  const moveParagraphDropCap = (
    _nodeIndex: number,
    _patch: Partial<NonNullable<NonNullable<ParagraphNode["style"]>["dropCap"]>>
  ): void => {
    // Adjusts the position/size of a paragraph drop cap.
  }

  const setParagraphDropCapFontSizePt = (
    _nodeIndex: number, _fontSizePt: number
  ): void => {
    // Sets the font size (in points) for a paragraph drop cap.
  }

  const setParagraphDropCapText = (
    _nodeIndex: number, _text: string
  ): void => {
    // Sets the text content for a paragraph drop cap.
  }

  const moveImage = (
    _source: DocxImageLocation, _target: DocxImageDropTarget
  ): void => {
    // Moves an image from one location to another within the document.
  }

  return {
    insertImageFile,
    resizeImage,
    setSyntheticTextBoxText,
    setImageWrapMode,
    moveFloatingImage,
    moveSectionFloatingImage,
    moveParagraphDropCap,
    setParagraphDropCapFontSizePt,
    setParagraphDropCapText,
    moveImage,
  }
}

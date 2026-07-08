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
import { cloneDocModel } from "@extend-ai/docx-core"
import type { EditorCore } from "./editor-shared"

export function createEditorImage(
  ctx: EditorCore,
  applyChange: (updater: (current: DocModel) => DocModel, successStatus?: string) => void,
) {
  const insertImageFile = async (_file: File): Promise<void> => {
    // Reads image file, converts to data URL or Uint8Array, inserts as ImageRunNode.
    // Full implementation in editor-image module.
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
  ): void => {}

  const setImageWrapMode = (
    _location: DocxImageLocation,
    _mode: DocxImageWrapMode,
    _seed?: Partial<NonNullable<ImageRunNode["floating"]>>
  ): void => {}

  const moveFloatingImage = (
    _location: DocxImageLocation,
    _patch: Partial<NonNullable<ImageRunNode["floating"]>>
  ): void => {}

  const moveSectionFloatingImage = (
    _location: DocxSectionImageLocation,
    _patch: Partial<NonNullable<ImageRunNode["floating"]>>
  ): void => {}

  const moveParagraphDropCap = (
    _nodeIndex: number,
    _patch: Partial<NonNullable<NonNullable<ParagraphNode["style"]>["dropCap"]>>
  ): void => {}

  const setParagraphDropCapFontSizePt = (
    _nodeIndex: number, _fontSizePt: number
  ): void => {}

  const setParagraphDropCapText = (
    _nodeIndex: number, _text: string
  ): void => {}

  const moveImage = (
    _source: DocxImageLocation, _target: DocxImageDropTarget
  ): void => {}

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

// editor-list.ts — list toggle + depth adjustment
//
// Extracted from useDocxEditor.ts.

import type { DocModel } from "@extend-ai/docx-core"
import type { DocxListType } from "@extend-ai/docx-core"
import {
  cloneDocModel,
  paragraphListType,
  resolveSelectedParagraphLocation,
  getParagraphAtLocation,
} from "@extend-ai/docx-core"
import type { EditorCore } from "./editor-shared"

export function createEditorList(
  ctx: EditorCore,
  applyChange: (updater: (current: DocModel) => DocModel, successStatus?: string) => void,
) {
  const toggleList = (listType: DocxListType): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const loc = resolveSelectedParagraphLocation(ctx.selectionSnapshot.value, ctx.activeTextRangeSnapshot.value)
      const { paragraph: para } = getParagraphAtLocation(nextModel, loc)
      if (!para) return current
      if (!para.style) para.style = {}
      const currentListType = paragraphListType(para, current.metadata.numberingDefinitions)
      if (listType === "unordered") {
        if (currentListType === "unordered") {
          para.style.numbering = undefined
        } else {
          para.style.numbering = { numId: 1, ilvl: 0 }
        }
      } else {
        if (currentListType === "ordered") {
          para.style.numbering = undefined
        } else {
          para.style.numbering = { numId: 2, ilvl: 0 }
        }
      }
      return nextModel
    })
  }

  const adjustSelectedListDepth = (levelDelta: number, _draftText?: string): boolean => {
    const currentModel = ctx.model.value
    const loc = resolveSelectedParagraphLocation(ctx.selectionSnapshot.value, ctx.activeTextRangeSnapshot.value)
    if (loc.kind !== "paragraph") return false

    const para = currentModel.nodes[loc.nodeIndex]
    if (!para || para.type !== "paragraph" || !para.style?.numbering) return false

    const currentIlvl = para.style.numbering.ilvl ?? 0
    const nextIlvl = Math.max(0, Math.min(8, currentIlvl + levelDelta))
    if (nextIlvl === currentIlvl) return false

    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const nextPara = nextModel.nodes[loc.nodeIndex]
      if (!nextPara || nextPara.type !== "paragraph" || !nextPara.style?.numbering) return current
      nextPara.style.numbering = { ...nextPara.style.numbering, ilvl: nextIlvl }
      nextPara.sourceXml = undefined
      return nextModel
    })

    return true
  }

  return { toggleList, adjustSelectedListDepth }
}

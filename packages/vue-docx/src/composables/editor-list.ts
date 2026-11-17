// editor-list.ts — list toggle + depth adjustment
//
// Extracted from useDocxEditor.ts.

import type { DocModel, ParagraphNode } from "@arcships/docx-core"
import type { DocxListType } from "@arcships/docx-core"
import {
  cloneDocModel,
  ensureListNumberingId,
  getParagraphAtLocation,
  paragraphListType,
  resolveMaxNumberingLevel,
  resolveSelectedParagraphLocation,
} from "@arcships/docx-core"
import type { ApplyModelChangeOptions, EditorCore } from "./editor-shared"
import { selectedParagraphLocations } from "./editor-paragraph-range"

export function createEditorList(
  ctx: EditorCore,
  applyChange: (
    updater: (current: DocModel) => DocModel,
    successStatus?: string,
    options?: ApplyModelChangeOptions
  ) => void,
) {
  const toggleList = (listType: DocxListType): void => {
    applyChange((current) => {
      const locations = selectedParagraphLocations(
        current,
        ctx.selectionSnapshot.value,
        ctx.activeTextRangeSnapshot.value
      )
      const nextModel = cloneDocModel(current)
      const paragraphs: ParagraphNode[] = []
      for (const location of locations) {
        const { paragraph } = getParagraphAtLocation(nextModel, location)
        if (paragraph) paragraphs.push(paragraph)
      }
      if (paragraphs.length === 0) {
        ctx.status.value = "Cannot change the list: paragraph not found"
        return current
      }

      // Toggling a list type applies to the whole selection: it is removed only
      // when every selected paragraph already carries that list type.
      const removeList = paragraphs.every(
        (para) => paragraphListType(para, current.metadata.numberingDefinitions) === listType
      )
      if (removeList) {
        for (const para of paragraphs) {
          if (!para.style) para.style = {}
          para.style.numbering = undefined
          para.sourceXml = undefined
        }
        return nextModel
      }

      const numId = ensureListNumberingId(nextModel, listType)
      for (const para of paragraphs) {
        if (!para.style) para.style = {}
        para.style.numbering = {
          numId,
          ilvl: Math.max(0, para.style.numbering?.ilvl ?? 0),
        }
        para.sourceXml = undefined
      }
      return nextModel
    }, listType === "unordered" ? "Applied bullet list" : "Applied numbered list")
  }

  const adjustSelectedListDepth = (levelDelta: number, _draftText?: string): boolean => {
    const currentModel = ctx.model.value
    const loc = resolveSelectedParagraphLocation(ctx.selectionSnapshot.value, ctx.activeTextRangeSnapshot.value)
    const para = getParagraphAtLocation(currentModel, loc).paragraph
    if (!para?.style?.numbering) return false

    const currentIlvl = Math.max(0, para.style.numbering.ilvl ?? 0)
    const maxLevel = resolveMaxNumberingLevel(
      currentModel.metadata.numberingDefinitions,
      para.style.numbering.numId
    )
    const nextIlvl = Math.max(0, Math.min(maxLevel, currentIlvl + levelDelta))
    if (nextIlvl === currentIlvl) {
      if (levelDelta > 0 && currentIlvl >= maxLevel) {
        ctx.status.value = `Maximum list level is ${maxLevel + 1}`
      }
      return false
    }

    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const nextPara = getParagraphAtLocation(nextModel, loc).paragraph
      if (!nextPara?.style?.numbering) return current
      nextPara.style.numbering = { ...nextPara.style.numbering, ilvl: nextIlvl }
      nextPara.sourceXml = undefined
      return nextModel
    })

    return true
  }

  return { toggleList, adjustSelectedListDepth }
}

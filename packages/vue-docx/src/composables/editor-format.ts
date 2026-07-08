// editor-format.ts — character & paragraph formatting operations
//
// Extracted from useDocxEditor.ts.

import type {
  DocModel,
  ParagraphNode,
  TextRunNode,
  HeadingLevel,
  ParagraphAlignment,
} from "@extend-ai/docx-core"
import type {
  DocxBorderPreset,
  DocxEditorTransactionContext,
  DocxEditorTransactionPatch,
  DocxListType,
} from "@extend-ai/docx-core"
import { cloneDocModel, mutateParagraphTextStyleInRange } from "@extend-ai/docx-core"
import {
  resolveSelectedParagraphLocation,
  getParagraphAtLocation,
} from "@extend-ai/docx-core"
import type { EditorCore } from "./editor-shared"

function typedKeys<T extends object>(obj: T): (keyof T)[] {
  return Object.keys(obj) as (keyof T)[]
}

export function createEditorFormat(
  ctx: EditorCore,
  dispatch: (resolver: (txCtx: DocxEditorTransactionContext) => DocxEditorTransactionPatch | undefined) => boolean,
  applyChange: (updater: (current: DocModel) => DocModel, successStatus?: string) => void,
) {
  // ── applySelectedStyleChange ─────────────────────────────────────
  function applySelectedStyleChange(
    patch: Partial<TextRunNode["style"]>,
    replace?: Partial<TextRunNode["style"]>
  ): void {
    const modelToUse = ctx.modelSnapshot.value
    const rng = ctx.activeTextRangeSnapshot.value
    const sp = getParagraphAtLocation(
      modelToUse,
      resolveSelectedParagraphLocation(ctx.selectionSnapshot.value, rng)
    ).paragraph

    if (!rng || !sp) {
      // collapsed selection: mutate pendingRunStyle
      dispatch((txCtx) => {
        const nextPending = {
          ...(txCtx.pendingRunStyle ?? {}),
          ...patch,
        }
        if (replace) {
          typedKeys(replace).forEach((k) => {
            if (replace[k] === undefined) {
              delete nextPending[k]
            } else {
              (nextPending as Record<string, unknown>)[k as string] = replace[k]
            }
          })
        }
        return { pendingRunStyle: nextPending }
      })
      return
    }

    dispatch((txCtx) => {
      const nextModel = cloneDocModel(txCtx.model)
      const { paragraph: para } = getParagraphAtLocation(
        nextModel,
        resolveSelectedParagraphLocation(ctx.selectionSnapshot.value, rng)
      )
      if (!para) return undefined

      mutateParagraphTextStyleInRange(para, rng.start.offset, rng.end.offset, (currentStyle) => {
        const next = { ...currentStyle, ...patch }
        if (replace) {
          typedKeys(replace).forEach((k) => {
            if (replace[k] === undefined) {
              delete (next as Record<string, unknown>)[k as string]
            } else {
              (next as Record<string, unknown>)[k as string] = replace[k]
            }
          })
        }
        return next
      })

      return { model: nextModel }
    })
  }

  // ── character formatting ─────────────────────────────────────────
  const toggleBold = (): void =>
    applySelectedStyleChange({ bold: !(ctx.selectedRunStyle.value?.bold) })
  const toggleItalic = (): void =>
    applySelectedStyleChange({ italic: !(ctx.selectedRunStyle.value?.italic) })
  const toggleUnderline = (): void =>
    applySelectedStyleChange({ underline: !(ctx.selectedRunStyle.value?.underline) })
  const toggleStrike = (): void =>
    applySelectedStyleChange({ strike: !(ctx.selectedRunStyle.value?.strike) })
  const toggleSuperscript = (): void =>
    applySelectedStyleChange({
      verticalAlign:
        ctx.selectedRunStyle.value?.verticalAlign === "superscript" ? undefined : "superscript",
    })
  const toggleSubscript = (): void =>
    applySelectedStyleChange({
      verticalAlign:
        ctx.selectedRunStyle.value?.verticalAlign === "subscript" ? undefined : "subscript",
    })

  const setFontFamily = (fontFamily: string): void => applySelectedStyleChange({ fontFamily })
  const setFontSize = (fontSizePt: number): void => applySelectedStyleChange({ fontSizePt })
  const setHighlight = (highlight?: string): void =>
    applySelectedStyleChange({}, { highlight })
  const setTextColor = (color?: string): void =>
    applySelectedStyleChange({}, { color })

  // ── hyperlink ────────────────────────────────────────────────────
  const setLink = (link?: string): void => {
    dispatch((txCtx) => {
      const rng = ctx.activeTextRangeSnapshot.value
      if (!rng) return undefined
      const nextModel = cloneDocModel(txCtx.model)
      const { paragraph: para } = getParagraphAtLocation(
        nextModel,
        resolveSelectedParagraphLocation(ctx.selectionSnapshot.value, rng)
      )
      if (!para) return undefined
      const safeStart = Math.max(0, Math.min(rng.start.offset, rng.end.offset))
      const safeEnd = Math.max(safeStart, Math.max(rng.start.offset, rng.end.offset))
      if (safeStart === safeEnd) return undefined
      let cursor = 0
      for (const child of para.children) {
        if (child.type === "text") {
          const childLen = child.text.length
          const childStart = cursor
          const childEnd = cursor + childLen
          const overlapStart = Math.max(safeStart, childStart)
          const overlapEnd = Math.min(safeEnd, childEnd)
          if (overlapStart < overlapEnd) {
            if (link) {
              child.link = link
            } else {
              delete child.link
            }
          }
          cursor = childEnd
        }
      }
      return { model: nextModel }
    })
  }

  // ── paragraph formatting ─────────────────────────────────────────
  const setHeading = (heading?: HeadingLevel): void => {
    dispatch((txCtx) => {
      const nextModel = cloneDocModel(txCtx.model)
      const loc = resolveSelectedParagraphLocation(ctx.selectionSnapshot.value, ctx.activeTextRangeSnapshot.value)
      const { paragraph: para } = getParagraphAtLocation(nextModel, loc)
      if (!para) return undefined
      if (!para.style) para.style = {}
      para.style.headingLevel = heading
      return { model: nextModel }
    })
  }

  const setParagraphStyle = (styleId?: string): void => {
    dispatch((txCtx) => {
      const nextModel = cloneDocModel(txCtx.model)
      const loc = resolveSelectedParagraphLocation(ctx.selectionSnapshot.value, ctx.activeTextRangeSnapshot.value)
      const { paragraph: para } = getParagraphAtLocation(nextModel, loc)
      if (!para) return undefined
      if (!para.style) para.style = {}
      para.style.styleId = styleId
      delete para.style.headingLevel
      return { model: nextModel }
    })
  }

  const setLineSpacing = (lineMultiple: number): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const loc = resolveSelectedParagraphLocation(ctx.selectionSnapshot.value, ctx.activeTextRangeSnapshot.value)
      const { paragraph: para } = getParagraphAtLocation(nextModel, loc)
      if (!para) return current
      if (!para.style) para.style = {}
      if (!para.style.spacing) para.style.spacing = {}
      const lineTwips = Math.max(1, Math.round(lineMultiple * 240))
      para.style.spacing.lineTwips = lineTwips
      para.style.spacing.lineRule = "auto"
      return nextModel
    })
  }

  const setAlignment = (align?: ParagraphAlignment): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const loc = resolveSelectedParagraphLocation(ctx.selectionSnapshot.value, ctx.activeTextRangeSnapshot.value)
      const { paragraph: para } = getParagraphAtLocation(nextModel, loc)
      if (!para) return current
      if (!para.style) para.style = {}
      para.style.align = align
      return nextModel
    })
  }

  const applyBorderPreset = (_preset: DocxBorderPreset): void => {
    // Schema stub — actual border application is handled by the toolbar
    // with paragraph-level or table-level border style mutations.
  }

  return {
    applySelectedStyleChange,
    toggleBold,
    toggleItalic,
    toggleUnderline,
    toggleStrike,
    toggleSuperscript,
    toggleSubscript,
    setFontFamily,
    setFontSize,
    setHighlight,
    setTextColor,
    setLink,
    setHeading,
    setParagraphStyle,
    setLineSpacing,
    setAlignment,
    applyBorderPreset,
  }
}

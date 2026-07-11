// editor-format.ts — character & paragraph formatting operations
//
// Extracted from useDocxEditor.ts.

import type {
  DocModel,
  ParagraphNode,
  TextRunNode,
  HeadingLevel,
  ParagraphAlignment,
} from "@arcships/docx-core"
import type {
  DocxBorderPreset,
  DocxEditorTransactionContext,
  DocxEditorTransactionPatch,
  DocxListType,
} from "@arcships/docx-core"
import {
  cloneDocModel,
  mutateParagraphTextStyleInRange,
  resolveSelectedParagraphLocation,
  getParagraphAtLocation,
  normalizeTextRange,
  compareTextRangeBoundaries,
  cloneParagraphStyle,
  tableBorderPresetState,
  applyTableBorderPreset,
  applyParagraphBorderPresetForRangeEntry,
  paragraphRangePresetActive,
  paragraphRangeForMutate,
} from "@arcships/docx-core"
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

  // ── applyBorderPreset ────────────────────────────────────────────
  const applyBorderPreset = (preset: DocxBorderPreset): void => {
    const targetLocation = ctx.selectedParagraphLocation.value
    const activeRange = ctx.activeTextRangeSnapshot.value
    const normalizedRange = activeRange ? normalizeTextRange(activeRange) : undefined
    const hasExpandedRange = Boolean(
      normalizedRange && compareTextRangeBoundaries(normalizedRange.start, normalizedRange.end) < 0
    )

    applyChange((current) => {
      if (targetLocation.kind === "table-cell") {
        const next = cloneDocModel(current)
        const tableNode = next.nodes[targetLocation.tableIndex]
        if (!tableNode || tableNode.type !== "table") return current

        const targetCell = tableNode.rows[targetLocation.rowIndex]?.cells[targetLocation.cellIndex]
        if (!targetCell) return current

        const presetState = tableBorderPresetState(tableNode.style?.borders, targetCell.style?.borders)
        const removePreset = preset !== "none" && presetState[preset]

        // Diagonal borders are deferred to component integration
        if (preset === "diagonal-down" || preset === "diagonal-up") return current

        const nextBorders = applyTableBorderPreset(tableNode.style?.borders, preset, removePreset)
        if (!nextBorders) return current

        tableNode.style = { ...(tableNode.style ?? {}), borders: nextBorders }
        tableNode.sourceXml = undefined
        return next
      }

      // Paragraph border preset
      if (hasExpandedRange && normalizedRange) {
        const rangeLocations = paragraphRangeForMutate(
          current,
          normalizedRange.start.location,
          normalizedRange.end.location
        ).map((entry) => entry.location)
        if (rangeLocations.length === 0) return current

        const paragraphsWithLocation = rangeLocations
          .map((location) => {
            const lookup = getParagraphAtLocation(current, location)
            if (!lookup.paragraph) return undefined
            return { location, borders: lookup.paragraph.style?.borders as any }
          })
          .filter((entry): entry is { location: import("@arcships/docx-core").ParagraphLocation; borders: any } => Boolean(entry))
        if (paragraphsWithLocation.length === 0) return current

        const removePreset = preset !== "none" &&
          paragraphRangePresetActive(preset, paragraphsWithLocation.map((e) => e.borders))
        const next = cloneDocModel(current)

        for (let index = 0; index < paragraphsWithLocation.length; index += 1) {
          const target = paragraphsWithLocation[index]
          const { paragraph, tableNode } = getParagraphAtLocation(next, target.location)
          if (!paragraph) continue

          const clonedStyle = cloneParagraphStyle(paragraph.style) ?? {}
          const nextBorders = applyParagraphBorderPresetForRangeEntry(
            clonedStyle.borders,
            preset,
            removePreset,
            index,
            paragraphsWithLocation.length
          )
          paragraph.style = { ...clonedStyle, borders: nextBorders }
          paragraph.sourceXml = undefined
          if (tableNode) tableNode.sourceXml = undefined
        }
        return next
      }

      // Single paragraph — apply border directly
      const next = cloneDocModel(current)
      const loc = resolveSelectedParagraphLocation(ctx.selectionSnapshot.value, ctx.activeTextRangeSnapshot.value)
      const { paragraph, tableNode } = getParagraphAtLocation(next, loc)
      if (!paragraph) return current

      const clonedStyle = cloneParagraphStyle(paragraph.style) ?? {}
      const borderActive = !!(
        clonedStyle.borders?.top && clonedStyle.borders?.bottom &&
        clonedStyle.borders?.left && clonedStyle.borders?.right
      )
      const removePreset: boolean = preset !== "none" && borderActive
      const nextBorders = applyParagraphBorderPresetForRangeEntry(
        clonedStyle.borders,
        preset,
        removePreset,
        0,
        1
      )
      paragraph.style = { ...clonedStyle, borders: nextBorders }
      paragraph.sourceXml = undefined
      if (tableNode) tableNode.sourceXml = undefined
      return next
    })
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

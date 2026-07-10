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
  updateSectionImageFloatingAtLocation,
} from "@extend-ai/docx-core"
import type { EditorCore } from "./editor-shared"

type LocatedImage = {
  image: ImageRunNode
  paragraph: ParagraphNode
  table?: Extract<DocModel["nodes"][number], { type: "table" }>
}

function locateImage(
  model: DocModel,
  location: DocxImageLocation,
): LocatedImage | undefined {
  let paragraph: ParagraphNode | undefined
  let table: Extract<DocModel["nodes"][number], { type: "table" }> | undefined

  if (location.kind === "paragraph") {
    const node = model.nodes[location.nodeIndex]
    if (node?.type === "paragraph") paragraph = node
  } else {
    const node = model.nodes[location.tableIndex]
    if (node?.type === "table") {
      table = node
      const candidate = node.rows[location.rowIndex]?.cells[location.cellIndex]
        ?.nodes[location.paragraphIndex]
      if (candidate?.type === "paragraph") paragraph = candidate
    }
  }

  const image = paragraph?.children[location.childIndex]
  if (!paragraph || image?.type !== "image") return undefined
  return { image, paragraph, table }
}

function invalidateImageSource(location: LocatedImage): void {
  location.image.sourceXml = undefined
  location.paragraph.sourceXml = undefined
  if (location.table) location.table.sourceXml = undefined
}

export function createEditorImage(
  ctx: EditorCore,
  applyChange: (updater: (current: DocModel) => DocModel, successStatus?: string) => void,
) {
  const unsupported = (feature: string): void => {
    ctx.status.value = `Unsupported: ${feature}`
  }

  const applyImageChange = (
    location: DocxImageLocation,
    update: (located: LocatedImage) => void,
    status: string,
  ): void => {
    if (!locateImage(ctx.modelSnapshot.value, location)) {
      unsupported(`${status} target is unavailable`)
      return
    }
    applyChange((current) => {
      const next = cloneDocModel(current)
      const located = locateImage(next, location)
      if (!located) return current
      update(located)
      invalidateImageSource(located)
      return next
    }, status)
  }

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
    if (!Number.isFinite(widthPx) || !Number.isFinite(heightPx) || widthPx <= 0 || heightPx <= 0) {
      unsupported("image size must be positive")
      return
    }
    const width = Math.max(1, Math.round(widthPx))
    const height = Math.max(1, Math.round(heightPx))
    applyImageChange(location, ({ image }) => {
      image.widthPx = width
      image.heightPx = height
    }, `Resized image to ${width}×${height}`)
  }

  const setSyntheticTextBoxText = (
    location: DocxImageLocation, text: string
  ): void => {
    const located = locateImage(ctx.modelSnapshot.value, location)
    if (!located?.image.syntheticTextBox) {
      unsupported("synthetic text box editing is unavailable for this image")
      return
    }
    applyImageChange(location, ({ image }) => {
      image.textBoxText = text
    }, "Updated text box")
  }

  const setImageWrapMode = (
    location: DocxImageLocation,
    mode: DocxImageWrapMode,
    seed: Partial<NonNullable<ImageRunNode["floating"]>> = {}
  ): void => {
    applyImageChange(location, ({ image }) => {
      if (mode === "inline") {
        image.floating = undefined
        return
      }

      const floating: NonNullable<ImageRunNode["floating"]> = {
        xPx: 0,
        yPx: 0,
        ...(image.floating ?? {}),
        ...seed,
      }
      if (mode === "behindText") {
        floating.wrapType = "none"
        floating.behindDocument = true
        floating.zIndex = Math.min(floating.zIndex ?? -1, -1)
      } else if (mode === "inFrontOfText") {
        floating.wrapType = "none"
        floating.behindDocument = false
        floating.zIndex = Math.max(floating.zIndex ?? 1, 1)
      } else {
        floating.wrapType = mode
        floating.behindDocument = false
      }
      image.floating = floating
    }, `Set image wrap: ${mode}`)
  }

  const moveFloatingImage = (
    location: DocxImageLocation,
    patch: Partial<NonNullable<ImageRunNode["floating"]>>
  ): void => {
    applyImageChange(location, ({ image }) => {
      const floating: NonNullable<ImageRunNode["floating"]> = {
        xPx: 0,
        yPx: 0,
        wrapType: "square",
        ...(image.floating ?? {}),
        ...patch,
      }
      if (patch.xPx !== undefined) delete floating.horizontalAlign
      if (patch.yPx !== undefined) delete floating.verticalAlign
      image.floating = floating
    }, "Moved image")
  }

  const moveSectionFloatingImage = (
    location: DocxSectionImageLocation,
    patch: Partial<NonNullable<ImageRunNode["floating"]>>
  ): void => {
    const next = updateSectionImageFloatingAtLocation(
      ctx.modelSnapshot.value,
      location,
      patch,
    )
    if (next === ctx.modelSnapshot.value) {
      unsupported("section image target is unavailable")
      return
    }
    applyChange(() => next, "Moved section image")
  }

  const moveParagraphDropCap = (
    nodeIndex: number,
    patch: Partial<NonNullable<NonNullable<ParagraphNode["style"]>["dropCap"]>>
  ): void => {
    const paragraph = ctx.modelSnapshot.value.nodes[nodeIndex]
    if (paragraph?.type !== "paragraph" || !paragraph.style?.dropCap) {
      unsupported("drop cap target is unavailable")
      return
    }
    applyChange((current) => {
      const next = cloneDocModel(current)
      const target = next.nodes[nodeIndex]
      if (target?.type !== "paragraph" || !target.style?.dropCap) return current
      target.style.dropCap = { ...target.style.dropCap, ...patch }
      target.sourceXml = undefined
      return next
    }, "Moved drop cap")
  }

  const setParagraphDropCapFontSizePt = (
    nodeIndex: number, fontSizePt: number
  ): void => {
    const paragraph = ctx.modelSnapshot.value.nodes[nodeIndex]
    const textRun = paragraph?.type === "paragraph"
      ? paragraph.children.find((child) => child.type === "text" && child.text.length > 0)
      : undefined
    if (!textRun || !Number.isFinite(fontSizePt) || fontSizePt <= 0) {
      unsupported("drop cap font size target is unavailable")
      return
    }
    applyChange((current) => {
      const next = cloneDocModel(current)
      const target = next.nodes[nodeIndex]
      if (target?.type !== "paragraph") return current
      const firstText = target.children.find((child) => child.type === "text" && child.text.length > 0)
      if (!firstText || firstText.type !== "text") return current
      firstText.style = { ...(firstText.style ?? {}), fontSizePt }
      target.sourceXml = undefined
      return next
    }, `Set drop cap font size to ${fontSizePt}pt`)
  }

  const setParagraphDropCapText = (
    nodeIndex: number, text: string
  ): void => {
    const paragraph = ctx.modelSnapshot.value.nodes[nodeIndex]
    const textRun = paragraph?.type === "paragraph"
      ? paragraph.children.find((child) => child.type === "text" && child.text.length > 0)
      : undefined
    if (!textRun || text.length === 0) {
      unsupported("drop cap text target is unavailable")
      return
    }
    applyChange((current) => {
      const next = cloneDocModel(current)
      const target = next.nodes[nodeIndex]
      if (target?.type !== "paragraph") return current
      const firstText = target.children.find((child) => child.type === "text" && child.text.length > 0)
      if (!firstText || firstText.type !== "text") return current
      const characters = Array.from(firstText.text)
      firstText.text = text + characters.slice(1).join("")
      target.sourceXml = undefined
      return next
    }, "Updated drop cap text")
  }

  const moveImage = (
    source: DocxImageLocation, target: DocxImageDropTarget
  ): void => {
    if (!locateImage(ctx.modelSnapshot.value, source)) {
      unsupported("image move source is unavailable")
      return
    }
    applyChange((current) => {
      const next = cloneDocModel(current)
      const from = locateImage(next, source)
      if (!from) return current

      let targetParagraph: ParagraphNode | undefined
      let targetTable: Extract<DocModel["nodes"][number], { type: "table" }> | undefined
      if (target.kind === "paragraph") {
        const node = next.nodes[target.nodeIndex]
        if (node?.type === "paragraph") targetParagraph = node
      } else {
        const node = next.nodes[target.tableIndex]
        if (node?.type === "table") {
          targetTable = node
          const candidate = node.rows[target.rowIndex]?.cells[target.cellIndex]
            ?.nodes[target.paragraphIndex]
          if (candidate?.type === "paragraph") targetParagraph = candidate
        }
      }
      if (!targetParagraph) {
        unsupported("image move target is unavailable")
        return current
      }

      const [image] = from.paragraph.children.splice(source.childIndex, 1)
      if (!image || image.type !== "image") return current
      const sameParagraph = from.paragraph === targetParagraph
      const insertionIndex = Math.max(
        0,
        Math.min(
          targetParagraph.children.length,
          target.childIndex - (sameParagraph && source.childIndex < target.childIndex ? 1 : 0),
        ),
      )
      targetParagraph.children.splice(insertionIndex, 0, image)
      from.paragraph.sourceXml = undefined
      targetParagraph.sourceXml = undefined
      if (from.table) from.table.sourceXml = undefined
      if (targetTable) targetTable.sourceXml = undefined
      return next
    }, "Moved image")
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

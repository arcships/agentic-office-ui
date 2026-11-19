import type { OfficeObjectDescriptor, ResolveReferenceResult } from "./types"
import type { OfficeMessages } from "./i18n"
import { formatOfficeMessage } from "./i18n"

export type OfficeDescriptorMessages = Pick<
  OfficeMessages["common"],
  "descriptors" | "selectedText"
>

function numberAfter(label: string, prefix: string): number | undefined {
  const match = new RegExp(`^${prefix}\\s+(\\d+)`, "u").exec(label.trim())
  return match ? Number(match[1]) : undefined
}

/** Localize descriptor labels that core adapters emit in English. */
export function localizeOfficeDescriptorLabel(
  label: string,
  messages: OfficeDescriptorMessages,
): string {
  const descriptors = messages.descriptors
  const text = label.trim()
  if (!text) return label
  if (text === "Document") return descriptors.document
  if (text === "Workbook") return descriptors.workbook
  if (text === "Presentation") return descriptors.presentation
  if (text === "Image") return descriptors.image
  if (text === "Chart") return descriptors.chart
  if (text === "Comment") return descriptors.comment
  if (text === "Group") return descriptors.group
  if (text === "Object") return descriptors.object
  if (text === "Shape") return descriptors.shape
  if (text === "Text box") return descriptors.textBox
  if (text === "PDF") return descriptors.pdf
  if (text.startsWith("Comment by ")) {
    return formatOfficeMessage(descriptors.commentBy, { author: text.slice(11) })
  }
  if (text.startsWith("Comment: ")) {
    return formatOfficeMessage(descriptors.commentText, { text: text.slice(9) })
  }
  if (text.startsWith("Image: ")) {
    return formatOfficeMessage(descriptors.imageText, { text: text.slice(7) })
  }
  if (text.startsWith("Chart ")) {
    return formatOfficeMessage(descriptors.chartName, { name: text.slice(6) })
  }
  const comment = numberAfter(text, "Comment")
  if (comment !== undefined) return formatOfficeMessage(descriptors.commentById, { id: comment })
  const group = numberAfter(text, "Group")
  if (group !== undefined) return formatOfficeMessage(descriptors.group, { number: group })
  const shape = /^Shape\s+(.+)$/u.exec(text)
  if (shape) return formatOfficeMessage(descriptors.shape, { id: shape[1] })
  const object = /^Object\s+(.+)$/u.exec(text)
  if (object) return formatOfficeMessage(descriptors.object, { id: object[1] })
  const textBox = /^Text box\s+(.+)$/u.exec(text)
  if (textBox) return formatOfficeMessage(descriptors.textBox, { id: textBox[1] })
  const lowerShape = /^shape\s+(.+)$/u.exec(text)
  if (lowerShape) return formatOfficeMessage(descriptors.shape, { id: lowerShape[1] })
  const lowerGroup = /^group\s+(.+)$/u.exec(text)
  if (lowerGroup) return formatOfficeMessage(descriptors.group, { number: lowerGroup[1] })
  const lowerTable = /^table\s+(.+)$/u.exec(text)
  if (lowerTable) return formatOfficeMessage(descriptors.table, { number: lowerTable[1] })
  const lowerChart = /^chart\s+(.+)$/u.exec(text)
  if (lowerChart) return formatOfficeMessage(descriptors.chartName, { name: lowerChart[1] })
  const lowerImage = /^image\s+(.+)$/u.exec(text)
  if (lowerImage) return formatOfficeMessage(descriptors.imageId, { id: lowerImage[1] })
  const lowerTextBox = /^text-box\s+(.+)$/u.exec(text)
  if (lowerTextBox) return formatOfficeMessage(descriptors.textBox, { id: lowerTextBox[1] })
  if (text === "Heading") return descriptors.heading
  if (text === "List item") return descriptors.listItem
  if (text === "Text block") return descriptors.textBlock
  if (text === "Link") return descriptors.link
  if (text === "Form field") return descriptors.formField
  if (text.startsWith("Form field: ")) {
    return formatOfficeMessage(descriptors.formFieldName, { name: text.slice(12) })
  }
  if (text === "Annotation") return descriptors.annotation
  if (text.startsWith("Annotation: ")) {
    return formatOfficeMessage(descriptors.annotationText, { text: text.slice(12) })
  }
  if (text === "Sheet protection") return descriptors.protection
  if (text === "Slide transition") return descriptors.transition
  if (text === "Action") return descriptors.action
  if (text === "Slide master") return descriptors.master
  if (text === "Slide layout") return descriptors.layout
  if (text === "Speaker notes") return descriptors.note
  if (text === "Permissions") return descriptors.permission
  if (text.startsWith("Table ")) {
    const name = text.slice(6)
    if (name && !/^\d+$/u.test(name)) return formatOfficeMessage(descriptors.tableByName, { name })
  }
  if (text.startsWith("Series: ")) {
    return formatOfficeMessage(descriptors.chartSeries, { name: text.slice(8) })
  }
  if (text.startsWith("Point ")) {
    const point = text.slice(6)
    if (/^\d+$/u.test(point)) return formatOfficeMessage(descriptors.chartPoint, { number: point })
  }
  if (text.startsWith("Legend entry ")) {
    const legend = text.slice(13)
    if (/^\d+$/u.test(legend)) return formatOfficeMessage(descriptors.legendEntry, { number: legend })
  }
  if (text.startsWith("Style: ")) return formatOfficeMessage(descriptors.styleName, { name: text.slice(7) })
  if (text.startsWith("Bookmark: ")) return formatOfficeMessage(descriptors.bookmarkName, { name: text.slice(10) })
  if (text.startsWith("Field: ")) return formatOfficeMessage(descriptors.fieldName, { name: text.slice(7) })
  if (text.startsWith("Revision by ")) return formatOfficeMessage(descriptors.trackedChange, { author: text.slice(12) })
  if (text.startsWith("Comment state: ")) return formatOfficeMessage(descriptors.commentState, { state: text.slice(15) })
  if (text.startsWith("Cross-reference: ")) return formatOfficeMessage(descriptors.crossReference, { name: text.slice(17) })
  if (text.startsWith("Formula: ")) return formatOfficeMessage(descriptors.formulaLabel, { formula: text.slice(9) })
  if (text.startsWith("Conditional format ")) {
    const format = text.slice(19)
    if (/^\d+$/u.test(format)) return formatOfficeMessage(descriptors.conditionalFormat, { number: format })
  }
  if (text.startsWith("Data validation ")) {
    const validation = text.slice(16)
    if (/^\d+$/u.test(validation)) return formatOfficeMessage(descriptors.dataValidation, { number: validation })
  }
  if (text.startsWith("Named range: ")) return formatOfficeMessage(descriptors.namedRange, { name: text.slice(13) })
  if (text.startsWith("Animation ")) {
    const animation = text.slice(10)
    if (/^\d+$/u.test(animation)) return formatOfficeMessage(descriptors.animation, { number: animation })
  }
  if (text.startsWith("Attachment: ")) return formatOfficeMessage(descriptors.attachment, { name: text.slice(12) })
  if (text.startsWith("Optional content: ")) return formatOfficeMessage(descriptors.layer, { name: text.slice(18) })
  if (text.startsWith("Signature: ")) return formatOfficeMessage(descriptors.signature, { name: text.slice(11) })
  if (text === "Selected text") return messages.selectedText
  if (text === "Selected region") return descriptors.selectedRegion
  if (text === "Text range") return descriptors.textRange
  if (text.startsWith("Text: ")) {
    return formatOfficeMessage(descriptors.textLabel, { text: text.slice(6) })
  }
  const page = numberAfter(text, "Page")
  if (page !== undefined) return formatOfficeMessage(descriptors.page, { number: page })
  const slide = numberAfter(text, "Slide")
  if (slide !== undefined) return formatOfficeMessage(descriptors.slide, { number: slide })
  const table = numberAfter(text, "Table")
  if (table !== undefined) return formatOfficeMessage(descriptors.table, { number: table })
  const tableAtNode = numberAfter(text, "Table at node")
  if (tableAtNode !== undefined) return formatOfficeMessage(descriptors.table, { number: tableAtNode })
  const row = numberAfter(text, "Row")
  if (row !== undefined) return formatOfficeMessage(descriptors.row, { number: row })
  const cell = numberAfter(text, "Cell")
  if (cell !== undefined) return formatOfficeMessage(descriptors.cell, { number: cell })
  const section = numberAfter(text, "Section")
  if (section !== undefined) return formatOfficeMessage(descriptors.section, { number: section })
  const paragraph = numberAfter(text, "Paragraph")
  if (paragraph !== undefined) return formatOfficeMessage(descriptors.paragraph, { number: paragraph })
  const column = /^Column\s+(.+)$/u.exec(text)
  if (column) return formatOfficeMessage(descriptors.column, { column: column[1] })
  const regionPage = numberAfter(text, "Region on page")
  if (regionPage !== undefined) return formatOfficeMessage(descriptors.regionOnPage, { number: regionPage })
  const regionSlide = numberAfter(text, "Region on slide")
  if (regionSlide !== undefined) return formatOfficeMessage(descriptors.regionOnSlide, { number: regionSlide })
  if (text.startsWith("Region on ")) {
    return formatOfficeMessage(descriptors.regionOnSheet, { name: text.slice(10) })
  }
  return label
}

export function localizeOfficeObjectDescriptor(
  descriptor: OfficeObjectDescriptor,
  messages: OfficeDescriptorMessages,
): OfficeObjectDescriptor {
  return {
    ...descriptor,
    label: localizeOfficeDescriptorLabel(descriptor.label, messages),
    path: descriptor.path.map((entry) => ({
      ...entry,
      label: localizeOfficeDescriptorLabel(entry.label, messages),
    })),
  }
}

/** Localize every descriptor carried by a reference resolution result. */
export function localizeOfficeResolveResult(
  result: ResolveReferenceResult,
  messages: OfficeDescriptorMessages,
): ResolveReferenceResult {
  if (result.status === "exact" || result.status === "relocated") {
    return { ...result, descriptor: localizeOfficeObjectDescriptor(result.descriptor, messages) }
  }
  if (result.status === "ambiguous") {
    return {
      ...result,
      candidates: result.candidates.map((descriptor) =>
        localizeOfficeObjectDescriptor(descriptor, messages),
      ),
    }
  }
  return result
}

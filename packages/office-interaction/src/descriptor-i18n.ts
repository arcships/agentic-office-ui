import type { OfficeObjectDescriptor } from "./types"
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
  if (text === "PDF") return descriptors.pdf
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
  const paragraph = numberAfter(text, "Paragraph")
  if (paragraph !== undefined) return formatOfficeMessage(descriptors.paragraph, { number: paragraph })
  const column = /^Column\s+(.+)$/u.exec(text)
  if (column) return formatOfficeMessage(descriptors.column, { column: column[1] })
  const regionPage = numberAfter(text, "Region on page")
  if (regionPage !== undefined) return formatOfficeMessage(descriptors.regionOnPage, { number: regionPage })
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

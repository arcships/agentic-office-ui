// docx-form-fields.ts — page-scoped form field collection
//
// The form field layer and the page surface both need to agree on which form
// fields belong to a page. Both walk the model through the core collector so
// table-cell fields are found too, then filter by the page's node segments.
// A table split across pages only contributes fields for the rows it renders.

import type {
  DocModel,
  DocxFormFieldLocation,
  DocumentPageNodeSegment,
  FormFieldRunNode,
} from "@arcships/docx-core"
import { collectFormFieldsFromModel, getParagraphAtLocation } from "@arcships/docx-core"

export interface DocxPageFormField {
  field: FormFieldRunNode
  location: DocxFormFieldLocation
  /** Stable key for `v-for` and DOM attribute lookups. */
  locationKey: string
}

/** Serializes a form field location into a stable, comparable key. */
export function docxFormFieldLocationKey(
  location: DocxFormFieldLocation
): string {
  if (location.kind === "paragraph") {
    return `p:${location.nodeIndex}:${location.childIndex}`
  }
  return `t:${location.tableIndex}:${location.rowIndex}:${location.cellIndex}:${location.paragraphIndex}:${location.childIndex}`
}

/**
 * Parses the key written to `data-docx-form-field-location` back into a
 * location. Returns undefined for values that are not shaped like a key, so a
 * host-provided attribute cannot address an unrelated node.
 */
export function parseDocxFormFieldLocationKey(
  key: string | null | undefined
): DocxFormFieldLocation | undefined {
  if (!key) return undefined
  const parts = key.split(":")
  const numbers = parts.slice(1).map((part) => Number(part))
  if (numbers.some((value) => !Number.isInteger(value) || value < 0)) {
    return undefined
  }
  if (parts[0] === "p" && numbers.length === 2) {
    return { kind: "paragraph", nodeIndex: numbers[0], childIndex: numbers[1] }
  }
  if (parts[0] === "t" && numbers.length === 5) {
    return {
      kind: "table-cell",
      tableIndex: numbers[0],
      rowIndex: numbers[1],
      cellIndex: numbers[2],
      paragraphIndex: numbers[3],
      childIndex: numbers[4],
    }
  }
  return undefined
}

/** Resolves the form-field node a location points at, if it still exists. */
export function resolveFormFieldAt(
  model: DocModel,
  location: DocxFormFieldLocation | undefined
): FormFieldRunNode | undefined {
  if (!location) return undefined
  const { paragraph } = getParagraphAtLocation(model, location)
  const field = paragraph?.children[location.childIndex]
  return field && field.type === "form-field" ? field : undefined
}

function topLevelNodeIndex(location: DocxFormFieldLocation): number {
  return location.kind === "paragraph" ? location.nodeIndex : location.tableIndex
}

function segmentShowsLocation(
  segment: DocumentPageNodeSegment,
  location: DocxFormFieldLocation
): boolean {
  if (location.kind === "paragraph" || !segment.tableRowRange) return true
  return (
    location.rowIndex >= segment.tableRowRange.startRowIndex &&
    location.rowIndex < segment.tableRowRange.endRowIndex
  )
}

/** Collects the form fields rendered on the page described by `pageNodeSegments`. */
export function collectPageFormFields(
  model: DocModel,
  pageNodeSegments: DocumentPageNodeSegment[]
): DocxPageFormField[] {
  const segmentsByNodeIndex = new Map<number, DocumentPageNodeSegment[]>()
  for (const segment of pageNodeSegments) {
    const segments = segmentsByNodeIndex.get(segment.nodeIndex)
    if (segments) segments.push(segment)
    else segmentsByNodeIndex.set(segment.nodeIndex, [segment])
  }

  const fields: DocxPageFormField[] = []
  for (const collected of collectFormFieldsFromModel(model)) {
    const location = collected.location
    const segments = segmentsByNodeIndex.get(topLevelNodeIndex(location))
    if (!segments) continue
    if (!segments.some((segment) => segmentShowsLocation(segment, location))) continue
    fields.push({
      field: collected.field,
      location,
      locationKey: docxFormFieldLocationKey(location),
    })
  }
  return fields
}

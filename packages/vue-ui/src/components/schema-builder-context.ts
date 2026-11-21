// Shared row context for the Vue SchemaBuilder block (#179).
//
// Rows are recursive: an object field renders its own children. Passing five
// mutation callbacks down every level would make each nested row re-emit its
// children's events, so the root provides one context and every row injects it.
import type { InjectionKey } from "vue"
import type { OfficeSelectEntry } from "./ui/types"
import type { SchemaBuilderField, SchemaBuilderFieldType } from "../types"
import type { SchemaBuilderIssue } from "../schema-builder"

export interface SchemaBuilderRowContext {
  readOnly: () => boolean
  /** Issues attached to one field id (may be empty). */
  issuesFor: (fieldId: string) => SchemaBuilderIssue[]
  updateField: (fieldId: string, patch: Partial<SchemaBuilderField>) => void
  removeField: (fieldId: string) => void
  /** `-1` moves up, `1` moves down inside the sibling list. */
  moveField: (fieldId: string, delta: -1 | 1) => void
  addChild: (fieldId: string) => void
  /** Localized labels for the five builder types. */
  typeLabel: (type: SchemaBuilderFieldType) => string
  /** Select options for the field type control. */
  typeOptions: () => OfficeSelectEntry[]
  /** Select options for the array element type control. */
  itemTypeOptions: () => OfficeSelectEntry[]
}

export const SCHEMA_BUILDER_CONTEXT: InjectionKey<SchemaBuilderRowContext> = Symbol(
  "arcships.schema-builder",
)

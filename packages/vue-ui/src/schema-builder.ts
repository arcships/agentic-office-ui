// Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn.
// Extraction-schema model for the Vue SchemaBuilder block (#179).
//
// The design audit (#185 §10) called out a missing persistence contract: the
// upstream block only kept a React reducer, so a host had no way to save or
// reload a schema. This module is that contract — a plain JSON-safe tree, a
// deterministic JSON Schema serializer, a tolerant importer and structural
// validation. Messages stay out of it on purpose: issues carry codes and the
// component formats them through the shared i18n catalog.
import type {
  SchemaBuilderArrayItemType,
  SchemaBuilderField,
  SchemaBuilderFieldType,
  SchemaBuilderSchema,
} from "./types"

export const SCHEMA_BUILDER_FIELD_TYPES: readonly SchemaBuilderFieldType[] = [
  "string",
  "number",
  "boolean",
  "object",
  "array",
]

export const SCHEMA_BUILDER_ARRAY_ITEM_TYPES: readonly SchemaBuilderArrayItemType[] = [
  "string",
  "number",
  "boolean",
  "object",
]

export function isSchemaBuilderFieldType(value: unknown): value is SchemaBuilderFieldType {
  return typeof value === "string" && (SCHEMA_BUILDER_FIELD_TYPES as readonly string[]).includes(value)
}

export function isSchemaBuilderArrayItemType(
  value: unknown,
): value is SchemaBuilderArrayItemType {
  return typeof value === "string" && (SCHEMA_BUILDER_ARRAY_ITEM_TYPES as readonly string[]).includes(value)
}

export function createEmptySchemaBuilderSchema(): SchemaBuilderSchema {
  return { properties: [] }
}

/**
 * Create a field with a caller-supplied id. The component keeps a counter; the
 * importer uses a deterministic one so a round trip is reproducible.
 */
export function createSchemaBuilderField(
  id: string,
  type: SchemaBuilderFieldType = "string",
): SchemaBuilderField {
  return { id, name: "", type, required: false, description: "" }
}

/** `field_1`, `field_2`, … skipping names already used by siblings. */
export function nextSchemaBuilderFieldName(siblings: readonly SchemaBuilderField[]): string {
  const used = new Set(siblings.map((field) => field.name.trim()))
  let index = 1
  while (used.has(`field_${index}`)) index += 1
  return `field_${index}`
}

export interface SchemaBuilderFieldLocation {
  field: SchemaBuilderField
  /** Sibling list that owns the field; `null` for the root properties array. */
  owner: SchemaBuilderField | null
  siblings: SchemaBuilderField[]
  /** Dot path of names, e.g. `line_items.sku`. */
  path: string
}

/** Depth-first lookup by id. */
export function findSchemaBuilderField(
  schema: SchemaBuilderSchema,
  fieldId: string,
): SchemaBuilderFieldLocation | null {
  let found: SchemaBuilderFieldLocation | null = null
  const visit = (siblings: SchemaBuilderField[], owner: SchemaBuilderField | null, prefix: string) => {
    for (const field of siblings) {
      if (found) return
      const path = prefix ? `${prefix}.${field.name || "?"}` : field.name || "?"
      if (field.id === fieldId) {
        found = { field, owner, siblings, path }
        return
      }
      if (field.properties?.length) visit(field.properties, field, path)
    }
  }
  visit(schema.properties, null, "")
  return found
}

export interface SchemaBuilderIssue {
  code: "empty-name" | "duplicate-name" | "invalid-type"
  fieldId: string
  /** Dotted display path with `?` for unnamed ancestors. */
  path: string
  name: string
  /** Offending type, present for `invalid-type`. */
  type?: string
}

/**
 * Structural validation. Duplicate detection is scoped to siblings, which is
 * how JSON Schema `properties` work; two different branches may reuse `id`.
 */
export function validateSchemaBuilderSchema(schema: SchemaBuilderSchema): SchemaBuilderIssue[] {
  const issues: SchemaBuilderIssue[] = []

  const visit = (siblings: SchemaBuilderField[], prefix: string) => {
    const seen = new Set<string>()
    for (const field of siblings) {
      const name = typeof field.name === "string" ? field.name.trim() : ""
      const path = prefix ? `${prefix}.${name || "?"}` : name || "?"
      if (!name) {
        issues.push({ code: "empty-name", fieldId: field.id, path, name: "" })
      } else if (seen.has(name)) {
        issues.push({ code: "duplicate-name", fieldId: field.id, path, name })
      } else {
        seen.add(name)
      }
      if (!isSchemaBuilderFieldType(field.type)) {
        issues.push({
          code: "invalid-type",
          fieldId: field.id,
          path,
          name,
          type: String(field.type),
        })
      } else if (field.type === "array" && !isSchemaBuilderArrayItemType(field.itemType ?? "string")) {
        issues.push({
          code: "invalid-type",
          fieldId: field.id,
          path,
          name,
          type: String(field.itemType),
        })
      }
      if (field.properties?.length) visit(field.properties, path)
    }
  }

  visit(schema.properties ?? [], "")
  return issues
}

export interface SerializedSchemaField {
  type: string
  description?: string
  properties?: Record<string, SerializedSchemaField>
  required?: string[]
  items?: SerializedSchemaField
}

export interface SerializedSchema {
  type: "object"
  properties: Record<string, SerializedSchemaField>
  required?: string[]
}

function requiredNames(fields: readonly SchemaBuilderField[]): string[] {
  return fields
    .filter((field) => field.required && field.name.trim().length > 0)
    .map((field) => field.name.trim())
}

function serializeField(field: SchemaBuilderField): SerializedSchemaField {
  const result: SerializedSchemaField = { type: field.type }
  const description = field.description?.trim()
  if (description) result.description = description

  if (field.type === "object") {
    result.properties = serializeProperties(field.properties ?? [])
    const required = requiredNames(field.properties ?? [])
    if (required.length) result.required = required
    return result
  }

  if (field.type === "array") {
    const itemType = isSchemaBuilderArrayItemType(field.itemType ?? "string")
      ? field.itemType ?? "string"
      : "string"
    const items: SerializedSchemaField = { type: itemType }
    if (itemType === "object") {
      items.properties = serializeProperties(field.properties ?? [])
      const required = requiredNames(field.properties ?? [])
      if (required.length) items.required = required
    }
    result.items = items
    return result
  }

  return result
}

function serializeProperties(fields: readonly SchemaBuilderField[]): Record<string, SerializedSchemaField> {
  const properties: Record<string, SerializedSchemaField> = {}
  for (const field of fields) {
    const name = field.name.trim()
    if (!name) continue
    properties[name] = serializeField(field)
  }
  return properties
}

/**
 * JSON Schema (draft 2020-12 subset) for a builder schema. Unnamed fields are
 * skipped so an in-progress tree still produces a valid document.
 */
export function serializeSchemaBuilderSchema(schema: SchemaBuilderSchema): SerializedSchema {
  const result: SerializedSchema = {
    type: "object",
    properties: serializeProperties(schema.properties ?? []),
  }
  const required = requiredNames(schema.properties ?? [])
  if (required.length) result.required = required
  return result
}

export interface SchemaBuilderImportWarning {
  code: "unsupported-type" | "unsupported-item-type" | "unsupported-item-shape" | "missing-name"
  path: string
  type?: string
}

export interface SchemaBuilderParseResult {
  schema: SchemaBuilderSchema
  warnings: SchemaBuilderImportWarning[]
}

export type SchemaBuilderParseErrorCode = "invalid-root" | "invalid-properties" | "invalid-property"

/** Thrown by the importer; the component maps `code` to a catalog message. */
export class SchemaBuilderParseError extends Error {
  readonly code: SchemaBuilderParseErrorCode
  readonly path: string

  constructor(code: SchemaBuilderParseErrorCode, path = "") {
    super(code)
    this.name = "SchemaBuilderParseError"
    this.code = code
    this.path = path
  }
}

type JsonRecord = Record<string, unknown>

function isRecord(value: unknown): value is JsonRecord {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value)
}

const SCALAR_FALLBACK: Record<string, SchemaBuilderFieldType> = {
  integer: "number",
  null: "string",
}

function mapFieldType(
  input: unknown,
  fallback: SchemaBuilderFieldType,
  path: string,
  warnings: SchemaBuilderImportWarning[],
): SchemaBuilderFieldType {
  if (isSchemaBuilderFieldType(input)) return input
  if (typeof input === "string" && SCALAR_FALLBACK[input]) {
    warnings.push({ code: "unsupported-type", path, type: input })
    return SCALAR_FALLBACK[input]!
  }
  if (input !== undefined) {
    warnings.push({ code: "unsupported-type", path, type: String(input) })
  }
  return fallback
}

/**
 * Import a JSON Schema document. Supported: `object` with `properties` +
 * `required`, `array` with scalar or object `items`, `description`. `integer`
 * and `null` collapse to `number`/`string`; `enum`, `anyOf`, `$ref` and other
 * keywords are reported as warnings instead of being silently dropped.
 */
export function parseJsonSchema(input: unknown): SchemaBuilderParseResult {
  const warnings: SchemaBuilderImportWarning[] = []
  let counter = 0
  const nextId = () => {
    counter += 1
    return `field-${counter}`
  }

  const readProperties = (
    container: JsonRecord,
    path: string,
  ): SchemaBuilderField[] => {
    const value = container.properties
    if (value === undefined) return []
    if (!isRecord(value)) throw new SchemaBuilderParseError("invalid-properties", path)
    const required = new Set(
      Array.isArray(container.required)
        ? container.required.filter((name): name is string => typeof name === "string")
        : [],
    )
    const fields: SchemaBuilderField[] = []
    for (const [name, raw] of Object.entries(value)) {
      const childPath = path ? `${path}.${name}` : name
      if (!isRecord(raw)) throw new SchemaBuilderParseError("invalid-property", childPath)
      if (!name.trim()) warnings.push({ code: "missing-name", path: childPath })
      const type = mapFieldType(raw.type, "string", childPath, warnings)
      const field: SchemaBuilderField = {
        id: nextId(),
        name,
        type,
        required: required.has(name),
        description: typeof raw.description === "string" ? raw.description : "",
      }
      if (type === "object") {
        field.properties = readProperties(raw, childPath)
      }
      if (type === "array") {
        const items = raw.items
        if (isRecord(items)) {
          const itemTypeInput = items.type
          if (isSchemaBuilderFieldType(itemTypeInput) && itemTypeInput !== "array") {
            field.itemType = itemTypeInput
            if (itemTypeInput === "object") {
              field.properties = readProperties(items, `${childPath}[]`)
            }
          } else {
            warnings.push({
              code: "unsupported-item-type",
              path: `${childPath}[]`,
              type: String(itemTypeInput),
            })
            field.itemType = "string"
          }
        } else {
          if (items !== undefined) {
            warnings.push({ code: "unsupported-item-shape", path: `${childPath}[]` })
          }
          field.itemType = "string"
        }
      }
      fields.push(field)
    }
    return fields
  }

  if (!isRecord(input)) throw new SchemaBuilderParseError("invalid-root")
  if (input.type !== undefined && input.type !== "object") {
    throw new SchemaBuilderParseError("invalid-root")
  }
  const properties = readProperties(input, "")

  return { schema: { properties }, warnings }
}

/** `JSON.parse` + {@link parseJsonSchema}; a syntax error surfaces as-is. */
export function parseSchemaBuilderJsonText(text: string): SchemaBuilderParseResult {
  return parseJsonSchema(JSON.parse(text) as unknown)
}

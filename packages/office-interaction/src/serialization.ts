import type {
  JsonValue,
  OfficeObjectReference,
  OfficeReferenceConfirmEvent,
} from "./types"
import {
  assertOfficeJsonValue,
  OfficeInteractionValidationError,
  parseOfficeObjectReference,
  parseOfficeReferenceConfirmEvent,
} from "./validation"

function sortJson(value: JsonValue): JsonValue {
  if (Array.isArray(value)) return value.map(sortJson)
  if (value && typeof value === "object") {
    const object = value as { readonly [key: string]: JsonValue }
    return Object.fromEntries(
      Object.keys(object)
        .sort()
        .map((key) => [key, sortJson(object[key])]),
    )
  }
  return value
}

export function canonicalStringifyOfficeJson(value: unknown): string {
  assertOfficeJsonValue(value)
  return JSON.stringify(sortJson(value))
}

function parseJson(text: string, label: string): unknown {
  try {
    return JSON.parse(text)
  } catch (error) {
    throw new OfficeInteractionValidationError({
      code: "INVALID_JSON_VALUE",
      path: label,
      message: error instanceof Error ? error.message : "invalid JSON",
    })
  }
}

export function serializeOfficeObjectReference(reference: OfficeObjectReference): string {
  return canonicalStringifyOfficeJson(parseOfficeObjectReference(reference))
}

export function deserializeOfficeObjectReference(text: string): OfficeObjectReference {
  return parseOfficeObjectReference(parseJson(text, "$reference"))
}

export function serializeOfficeReferenceConfirmEvent(event: OfficeReferenceConfirmEvent): string {
  return canonicalStringifyOfficeJson(parseOfficeReferenceConfirmEvent(event))
}

export function deserializeOfficeReferenceConfirmEvent(text: string): OfficeReferenceConfirmEvent {
  return parseOfficeReferenceConfirmEvent(parseJson(text, "$confirm"))
}

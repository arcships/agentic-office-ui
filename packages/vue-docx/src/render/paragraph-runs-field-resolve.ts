// Field resolution helpers for renderParagraphRuns.
// Extracted from the main function to keep paragraph-runs.ts under 1000 lines.
// Upstream editor.tsx: 18175-18183, 18430-18516.

import type { ParagraphNode } from "@arcships/docx-core"

export interface FieldResolutionContext {
  hasPageField: boolean
  hasStyleRefField: boolean
  pageFieldSequence: unknown[]
  pageFieldValueSequence: unknown[]
  styleRefFieldValueSequence: unknown[]
  pageNumber?: number
  totalPages?: number
  pageNumberFormat?: string
  resolveStyleRefFieldValue?: (target: string) => string | undefined
}

export interface FieldResolutionState {
  consumedPageFieldValues: number
  consumedStyleRefFieldValues: number
}

/**
 * Normalize text for field comparison (replace nbsp, collapse whitespace).
 */
export function normalizeFieldComparableText(input: string): string {
  return input.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim()
}

/**
 * Resolve PAGE/NUMPAGES field values in text.
 * Returns the resolved text and mutates state.consumedPageFieldValues.
 */
export function resolvePageFieldText(
  value: string,
  ctx: FieldResolutionContext,
  state: FieldResolutionState,
  preferredZone?: number
): string {
  if (!ctx.hasPageField || value.trim().length === 0) return value

  let fieldKind = ctx.pageFieldSequence[state.consumedPageFieldValues] as string | undefined
  const valueToken = ctx.pageFieldValueSequence[state.consumedPageFieldValues] as { rawText: string; kind: string } | undefined
  if (valueToken) {
    if (
      normalizeFieldComparableText(value) !==
      normalizeFieldComparableText(valueToken.rawText)
    ) {
      return value
    }
    fieldKind = valueToken.kind
  } else {
    if (!fieldKind) return value
    const normalized = value.trim()
    const likelyFieldResult =
      /^\d+$/.test(normalized) || /^[ivxlcdm]+$/i.test(normalized)
    if (!likelyFieldResult && preferredZone !== 1) return value
  }

  const resolvedFieldValue =
    fieldKind === "NUMPAGES"
      ? Number.isFinite(ctx.totalPages) && (ctx.totalPages as number) > 0
        ? Math.max(1, Math.round(ctx.totalPages as number))
        : undefined
      : Number.isFinite(ctx.pageNumber) && (ctx.pageNumber as number) > 0
      ? Math.max(1, Math.round(ctx.pageNumber as number))
      : undefined
  if (!Number.isFinite(resolvedFieldValue)) return value

  state.consumedPageFieldValues += 1
  const leadingWhitespace = value.match(/^\s*/)?.[0] ?? ""
  const trailingWhitespace = value.match(/\s*$/)?.[0] ?? ""
  const normalizedValue = Math.max(1, Math.round(resolvedFieldValue as number))
  const formattedFieldValue =
    fieldKind === "PAGE"
      ? formatPageFieldValue(normalizedValue, ctx.pageNumberFormat)
      : String(normalizedValue)
  return `${leadingWhitespace}${formattedFieldValue}${trailingWhitespace}`
}

/**
 * Resolve StyleRef field values in text.
 * Returns the resolved text and mutates state.consumedStyleRefFieldValues.
 */
export function resolveStyleRefFieldText(
  value: string,
  ctx: FieldResolutionContext,
  state: FieldResolutionState
): string {
  if (!ctx.hasStyleRefField || value.trim().length === 0) return value
  const valueToken = ctx.styleRefFieldValueSequence[state.consumedStyleRefFieldValues] as { rawText: string; target: string } | undefined
  if (!valueToken) return value
  if (
    normalizeFieldComparableText(value) !==
    normalizeFieldComparableText(valueToken.rawText)
  ) {
    return value
  }
  const resolvedValue = ctx
    .resolveStyleRefFieldValue?.(valueToken.target)
    ?.trim()
  if (!resolvedValue) return value
  state.consumedStyleRefFieldValues += 1
  const leadingWhitespace = value.match(/^\s*/)?.[0] ?? ""
  const trailingWhitespace = value.match(/\s*$/)?.[0] ?? ""
  return `${leadingWhitespace}${resolvedValue}${trailingWhitespace}`
}

/**
 * Resolve both PAGE/NUMPAGES and StyleRef fields.
 */
export function resolveFieldText(
  value: string,
  ctx: FieldResolutionContext,
  state: FieldResolutionState,
  preferredZone?: number
): string {
  return resolvePageFieldText(
    resolveStyleRefFieldText(value, ctx, state),
    ctx,
    state,
    preferredZone
  )
}

// ---- Helpers ----

function formatPageFieldValue(value: number, format?: string): string {
  if (!format) return String(value)
  const rawFormat = format.trim()
  if (!rawFormat) return String(value)
  const normalized = rawFormat.toLowerCase()

  // OOXML ST_NumberFormat values (w:pgNumType/@w:fmt) are case-insensitive.
  switch (normalized) {
    case "lowerroman":
    case "romanlower":
      return numberToRoman(value).toLowerCase()
    case "upperroman":
    case "romanupper":
    case "roman":
      return numberToRoman(value)
    case "lowerletter":
    case "loweralpha":
      return numberToLetters(value).toLowerCase()
    case "upperletter":
    case "upperalpha":
      return numberToLetters(value)
    case "none":
      return ""
    case "decimalzero":
      return String(value).padStart(2, "0")
    default:
      break
  }

  // Legacy single-letter aliases: "i"/"I" select lower/upper Roman, and
  // "a".."z"/"A".."Z" select letter numbering with that letter as offset.
  if (normalized === "i") {
    return rawFormat === "I"
      ? numberToRoman(value)
      : numberToRoman(value).toLowerCase()
  }

  if (/^[a-z]$/.test(normalized)) {
    const offset = normalized.charCodeAt(0) - 97
    const uppercase = rawFormat === rawFormat.toUpperCase()
    const base = uppercase ? 65 : 97
    return String.fromCharCode(base + ((value - 1 + offset) % 26))
  }

  return String(value)
}

export function numberToLetters(value: number): string {
  if (!Number.isFinite(value) || value <= 0) return ""
  let remaining = Math.floor(value)
  let output = ""
  while (remaining > 0) {
    const offset = (remaining - 1) % 26
    output = String.fromCharCode(65 + offset) + output
    remaining = Math.floor((remaining - 1) / 26)
  }
  return output
}

export function numberToRoman(value: number): string {
  if (!Number.isFinite(value) || value <= 0) return ""
  const numerals: Array<[number, string]> = [
    [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"],
    [100, "C"], [90, "XC"], [50, "L"], [40, "XL"],
    [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
  ]
  let remaining = Math.floor(value)
  let output = ""
  for (const [base, numeral] of numerals) {
    while (remaining >= base) {
      output += numeral
      remaining -= base
    }
  }
  return output
}

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
  sectionPages?: number
  pageNumberFormat?: string
  resolveStyleRefFieldValue?: (target: string) => string | undefined
  resolveBookmarkFieldValue?: (target: string) => string | undefined
  resolveBookmarkPageNumber?: (target: string) => number | undefined
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
 * Resolve dynamic field values (PAGE, NUMPAGES, SECTIONPAGES, DATE, TIME, REF,
 * PAGEREF, IF) in text. Returns the resolved text and mutates
 * state.consumedPageFieldValues.
 */
export function resolvePageFieldText(
  value: string,
  ctx: FieldResolutionContext,
  state: FieldResolutionState,
  preferredZone?: number
): string {
  if (!ctx.hasPageField || value.trim().length === 0) return value

  let fieldKind = ctx.pageFieldSequence[state.consumedPageFieldValues] as string | undefined
  const valueToken = ctx.pageFieldValueSequence[state.consumedPageFieldValues] as
    | { rawText: string; kind: string; instruction?: string }
    | undefined
  let instruction = ""
  if (valueToken) {
    if (
      normalizeFieldComparableText(value) !==
      normalizeFieldComparableText(valueToken.rawText)
    ) {
      return value
    }
    fieldKind = valueToken.kind
    instruction = valueToken.instruction ?? ""
  } else {
    if (!fieldKind) return value
    const normalized = value.trim()
    const likelyFieldResult =
      /^\d+$/.test(normalized) || /^[ivxlcdm]+$/i.test(normalized)
    if (!likelyFieldResult && preferredZone !== 1) return value
  }

  const resolvedFieldValue = resolveDynamicFieldValue(fieldKind, instruction, ctx)
  if (resolvedFieldValue === undefined) return value

  state.consumedPageFieldValues += 1
  const leadingWhitespace = value.match(/^\s*/)?.[0] ?? ""
  const trailingWhitespace = value.match(/\s*$/)?.[0] ?? ""
  return `${leadingWhitespace}${resolvedFieldValue}${trailingWhitespace}`
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

function positiveNumber(value: number | undefined): number | undefined {
  return Number.isFinite(value) && (value as number) > 0
    ? Math.max(1, Math.round(value as number))
    : undefined
}

function resolveDynamicFieldValue(
  kind: string | undefined,
  instruction: string,
  ctx: FieldResolutionContext
): string | undefined {
  switch (kind) {
    case "PAGE": {
      const pageNumber = positiveNumber(ctx.pageNumber)
      return pageNumber === undefined
        ? undefined
        : formatPageFieldValue(pageNumber, ctx.pageNumberFormat)
    }
    case "NUMPAGES": {
      const totalPages = positiveNumber(ctx.totalPages)
      return totalPages === undefined ? undefined : String(totalPages)
    }
    case "SECTIONPAGES": {
      const sectionPages = positiveNumber(ctx.sectionPages ?? ctx.totalPages)
      return sectionPages === undefined ? undefined : String(sectionPages)
    }
    case "DATE":
      return formatDateFieldValue(new Date(), instruction)
    case "TIME":
      return formatTimeFieldValue(new Date(), instruction)
    case "REF": {
      const target = fieldInstructionTarget(instruction, "REF")
      return target ? ctx.resolveBookmarkFieldValue?.(target) : undefined
    }
    case "PAGEREF": {
      const target = fieldInstructionTarget(instruction, "PAGEREF")
      const pageNumber = target
        ? positiveNumber(ctx.resolveBookmarkPageNumber?.(target))
        : undefined
      return pageNumber === undefined ? undefined : String(pageNumber)
    }
    case "IF":
      return evaluateIfField(instruction, ctx)
    case "UNSUPPORTED":
      return unsupportedFieldMarker(instruction)
    default:
      return undefined
  }
}

function unsupportedFieldMarker(instruction: string): string {
  const keyword =
    instruction
      .replace(/^\s*=+\s*/, "")
      .replace(/\s+/g, " ")
      .trim()
      .toUpperCase()
      .split(/\s+/, 1)[0] ?? "FIELD"
  return `[${keyword}]`
}

function fieldInstructionTarget(
  instruction: string,
  keyword: string
): string | undefined {
  const rest = instruction
    .replace(new RegExp(`^\\s*${keyword}\\b`, "i"), "")
    .trim()
  const match = rest.match(/^"([^"]+)"|^([^\s\\]+)/)
  const target = (match?.[1] ?? match?.[2] ?? "").trim()
  return target.length > 0 ? target : undefined
}

function extractFieldFormatSwitch(instruction: string): string | undefined {
  const match = instruction.match(/\\@\s*(?:"([^"]*)"|(\S+))/i)
  const format = match?.[1] ?? match?.[2]
  return format && format.trim().length > 0 ? format : undefined
}

const MONTH_NAMES_LONG = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
]
const MONTH_NAMES_SHORT = MONTH_NAMES_LONG.map((name) => name.slice(0, 3))
const WEEKDAY_NAMES_LONG = [
  "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday",
]
const WEEKDAY_NAMES_SHORT = WEEKDAY_NAMES_LONG.map((name) => name.slice(0, 3))

function pad2(value: number): string {
  return String(value).padStart(2, "0")
}

function formatDateFieldValue(date: Date, instruction: string): string {
  const format = extractFieldFormatSwitch(instruction)
  if (!format) return date.toLocaleDateString()
  return applyDatePattern(date, format)
}

function formatTimeFieldValue(date: Date, instruction: string): string {
  const format = extractFieldFormatSwitch(instruction)
  if (!format) return date.toLocaleTimeString()
  return applyDatePattern(date, format)
}

function applyDatePattern(date: Date, pattern: string): string {
  const hours24 = date.getHours()
  const hours12 = hours24 % 12 === 0 ? 12 : hours24 % 12
  return pattern.replace(
    /yyyy|yy|MMMM|MMM|MM|M|dddd|ddd|dd|d|HH|H|hh|h|mm|m|ss|s|AM\/PM|am\/pm/g,
    (token) => {
      switch (token) {
        case "yyyy": return String(date.getFullYear())
        case "yy": return pad2(date.getFullYear() % 100)
        case "MMMM": return MONTH_NAMES_LONG[date.getMonth()]
        case "MMM": return MONTH_NAMES_SHORT[date.getMonth()]
        case "MM": return pad2(date.getMonth() + 1)
        case "M": return String(date.getMonth() + 1)
        case "dddd": return WEEKDAY_NAMES_LONG[date.getDay()]
        case "ddd": return WEEKDAY_NAMES_SHORT[date.getDay()]
        case "dd": return pad2(date.getDate())
        case "d": return String(date.getDate())
        case "HH": return pad2(hours24)
        case "H": return String(hours24)
        case "hh": return pad2(hours12)
        case "h": return String(hours12)
        case "mm": return pad2(date.getMinutes())
        case "m": return String(date.getMinutes())
        case "ss": return pad2(date.getSeconds())
        case "s": return String(date.getSeconds())
        case "AM/PM": return hours24 < 12 ? "AM" : "PM"
        case "am/pm": return hours24 < 12 ? "am" : "pm"
        default: return token
      }
    }
  )
}

interface IfFieldToken {
  value: string
  quoted: boolean
}

function tokenizeIfInstruction(source: string): IfFieldToken[] {
  const tokens: IfFieldToken[] = []
  let current = ""
  let quoted = false
  let inQuotes = false

  const flush = (): void => {
    if (current.length === 0) return
    tokens.push({ value: current, quoted })
    current = ""
    quoted = false
  }

  for (let index = 0; index < source.length; index += 1) {
    const char = source[index]
    if (char === '"') {
      if (inQuotes) {
        tokens.push({ value: current, quoted: true })
        current = ""
        quoted = false
        inQuotes = false
      } else {
        flush()
        inQuotes = true
      }
      continue
    }
    if (inQuotes) {
      current += char
      continue
    }
    if (/\s/.test(char)) {
      flush()
      continue
    }
    if (/[<>=]/.test(char)) {
      flush()
      let operator = char
      while (index + 1 < source.length && /[<>=]/.test(source[index + 1])) {
        index += 1
        operator += source[index]
      }
      tokens.push({ value: operator, quoted: false })
      continue
    }
    quoted = false
    current += char
  }
  flush()
  return tokens
}

function resolveIfOperand(
  token: IfFieldToken | undefined,
  ctx: FieldResolutionContext
): string | number | undefined {
  if (!token) return undefined
  if (token.quoted) return token.value
  const value = token.value.trim()
  if (/^-?\d+(?:\.\d+)?$/.test(value)) return Number(value)
  if (/^PAGE$/i.test(value)) return positiveNumber(ctx.pageNumber)
  if (/^NUMPAGES$/i.test(value)) return positiveNumber(ctx.totalPages)
  if (/^SECTIONPAGES$/i.test(value)) {
    return positiveNumber(ctx.sectionPages ?? ctx.totalPages)
  }
  if (/^REF\s+/i.test(value)) {
    const target = value.replace(/^REF\s+/i, "").trim()
    return target ? ctx.resolveBookmarkFieldValue?.(target) : undefined
  }
  return value
}

function evaluateIfField(
  instruction: string,
  ctx: FieldResolutionContext
): string | undefined {
  const content = instruction
    .replace(/^\s*IF\b/i, "")
    .replace(/\s+\\[*#@!].*$/, "")
    .trim()
  if (!content) return undefined
  const tokens = tokenizeIfInstruction(content)
  if (tokens.length < 5) return undefined

  const trueText = tokens[tokens.length - 2]
  const falseText = tokens[tokens.length - 1]
  const condition = tokens.slice(0, -2)
  const operatorIndex = condition.findIndex((token) =>
    /^(?:=|==|<>|!=|>=|<=|>|<)$/.test(token.value)
  )
  if (operatorIndex <= 0 || operatorIndex >= condition.length - 1) {
    return undefined
  }

  const leftToken = condition[operatorIndex - 1]
  const rightToken = condition[operatorIndex + 1]
  const left = resolveIfOperand(leftToken, ctx)
  const right = resolveIfOperand(rightToken, ctx)
  if (left === undefined || right === undefined) return undefined

  const operator = condition[operatorIndex].value
  const bothNumbers = typeof left === "number" && typeof right === "number"
  const leftValue = bothNumbers ? (left as number) : String(left)
  const rightValue = bothNumbers ? (right as number) : String(right)
  let matched = false
  switch (operator) {
    case "=":
    case "==":
      matched = leftValue === rightValue
      break
    case "<>":
    case "!=":
      matched = leftValue !== rightValue
      break
    case ">":
      matched = leftValue > rightValue
      break
    case "<":
      matched = leftValue < rightValue
      break
    case ">=":
      matched = leftValue >= rightValue
      break
    case "<=":
      matched = leftValue <= rightValue
      break
    default:
      return undefined
  }
  return String(matched ? trueText.value : falseText.value)
}

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

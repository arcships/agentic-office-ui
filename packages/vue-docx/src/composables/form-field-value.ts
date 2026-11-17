// form-field-value.ts — validation, normalization and date formatting for
// DOCX form fields.
//
// The same helpers back the inline widget layer and the controller write path
// so a value that the UI accepts is exactly the value that reaches the model
// (and therefore the exported XML).

import type { FormFieldRunNode } from "@arcships/docx-core"
import { formatOfficeMessage, officeMessagesEn } from "@arcships/office-interaction"

export interface FormFieldValidationMessages {
  invalidValue: string
  required: string
  noOptions: string
  chooseValue: string
  invalidDate: string
  /** Template with a `{max}` placeholder. */
  maxLength: string
}

export interface FormFieldValidationResult {
  valid: boolean
  /** Normalized value to store when valid. */
  value: string
  error?: string
}

export interface FormFieldValidationOptions {
  /** Reject empty required fields. Used on blur and before export. */
  enforceRequired?: boolean
  /** Localized validation copy; defaults to the English catalog. */
  messages?: Partial<FormFieldValidationMessages>
}

export const FORM_FIELD_REQUIRED_ERROR = officeMessagesEn.docx.formField.errors.required

function validationMessages(
  options: FormFieldValidationOptions,
): FormFieldValidationMessages {
  return { ...officeMessagesEn.docx.formField.errors, ...options.messages }
}

/** Current model value, normalized for validation and export checks. */
export function formFieldCurrentValue(field: FormFieldRunNode): string {
  if (field.fieldType === "checkbox") {
    const explicit = field.value?.trim().toLowerCase()
    return field.checked === true || explicit === "1" || explicit === "true"
      ? "1"
      : "0"
  }
  return field.value ?? ""
}

/** True when a required field has no usable value yet. */
export function isRequiredFormFieldEmpty(field: FormFieldRunNode): boolean {
  if (field.required !== true) return false
  if (field.fieldType === "checkbox") {
    return formFieldCurrentValue(field) !== "1"
  }
  return formFieldCurrentValue(field).trim() === ""
}

interface DateParts {
  year: number
  month: number
  day: number
}

const MONTH_NAMES = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
]

const MONTH_ABBREVIATIONS = MONTH_NAMES.map((name) => name.slice(0, 3))

const DATE_TOKEN_PATTERN = /yyyy|yy|MMMM|MMM|MM|M|dd|d/g

function padNumber(value: number, length = 2): string {
  return String(value).padStart(length, "0")
}

function isValidDateParts(parts: DateParts): boolean {
  if (parts.month < 1 || parts.month > 12) return false
  if (parts.day < 1 || parts.day > 31) return false
  const date = new Date(Date.UTC(parts.year, parts.month - 1, parts.day))
  return (
    date.getUTCFullYear() === parts.year &&
    date.getUTCMonth() === parts.month - 1 &&
    date.getUTCDate() === parts.day
  )
}

function datePartsFromIso(value: string): DateParts | undefined {
  const match = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/.exec(value.trim())
  if (!match) return undefined
  const parts = {
    year: Number(match[1]),
    month: Number(match[2]),
    day: Number(match[3]),
  }
  return isValidDateParts(parts) ? parts : undefined
}

function monthIndexFromName(name: string): number | undefined {
  const normalized = name.trim().toLowerCase()
  const abbreviated = normalized.slice(0, 3)
  const index = MONTH_ABBREVIATIONS.findIndex(
    (month) => month.toLowerCase() === abbreviated
  )
  return index >= 0 ? index + 1 : undefined
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
}

function datePartsFromFormat(value: string, format: string): DateParts | undefined {
  const tokens: { kind: "literal" | "token"; value: string }[] = []
  let cursor = 0
  for (const match of format.matchAll(DATE_TOKEN_PATTERN)) {
    if (match.index !== undefined && match.index > cursor) {
      tokens.push({ kind: "literal", value: format.slice(cursor, match.index) })
    }
    tokens.push({ kind: "token", value: match[0] })
    cursor = (match.index ?? 0) + match[0].length
  }
  if (cursor < format.length) {
    tokens.push({ kind: "literal", value: format.slice(cursor) })
  }

  const captures: string[] = []
  let pattern = "^"
  for (const token of tokens) {
    if (token.kind === "literal") {
      pattern += escapeRegExp(token.value).replace(/\s+/g, "\\s*")
      continue
    }
    captures.push(token.value)
    switch (token.value) {
      case "yyyy":
        pattern += "(\\d{4})"
        break
      case "yy":
        pattern += "(\\d{2})"
        break
      case "MMMM":
      case "MMM":
        pattern += "([A-Za-z]+)"
        break
      default:
        pattern += "(\\d{1,2})"
        break
    }
  }
  pattern += "$"

  const match = new RegExp(pattern, "i").exec(value.trim())
  if (!match) return undefined

  let year: number | undefined
  let month: number | undefined
  let day: number | undefined
  captures.forEach((token, index) => {
    const raw = match[index + 1]
    if (raw === undefined) return
    switch (token) {
      case "yyyy":
        year = Number(raw)
        break
      case "yy":
        year = 2000 + Number(raw)
        break
      case "MMMM":
      case "MMM":
        month = monthIndexFromName(raw)
        break
      case "MM":
      case "M":
        month = Number(raw)
        break
      case "dd":
      case "d":
        day = Number(raw)
        break
    }
  })

  if (year === undefined || month === undefined || day === undefined) {
    return undefined
  }
  const parts = { year, month, day }
  return isValidDateParts(parts) ? parts : undefined
}

function datePartsFromLooseValue(value: string): DateParts | undefined {
  const timestamp = Date.parse(value.trim())
  if (Number.isNaN(timestamp)) return undefined
  const date = new Date(timestamp)
  return {
    year: date.getFullYear(),
    month: date.getMonth() + 1,
    day: date.getDate(),
  }
}

function parseFormFieldDate(
  value: string,
  format?: string
): DateParts | undefined {
  const trimmed = value.trim()
  if (!trimmed) return undefined
  const iso = datePartsFromIso(trimmed)
  if (iso) return iso
  // A value that looks like ISO but is not a real calendar date must be
  // rejected instead of being allowed to roll over in Date.parse.
  if (/^\d{4}-\d{1,2}-\d{1,2}(?:[T\s].*)?$/.test(trimmed)) return undefined
  if (format?.trim()) {
    const formatted = datePartsFromFormat(trimmed, format.trim())
    if (formatted) return formatted
  }
  return datePartsFromLooseValue(trimmed)
}

export function isDateFormField(field: FormFieldRunNode): boolean {
  if (field.fieldType === "date") return true
  const inputType = field.widget?.text?.inputType
  return field.fieldType === "text" && (inputType === "date" || inputType === "currentDate")
}

export function isoDateValueForField(
  field: FormFieldRunNode,
  value?: string
): string {
  const parts = parseFormFieldDate(
    value ?? field.value ?? "",
    field.widget?.text?.textFormat
  )
  if (!parts) return ""
  return `${padNumber(parts.year, 4)}-${padNumber(parts.month)}-${padNumber(parts.day)}`
}

function formatDateParts(parts: DateParts, format: string): string {
  return format.replace(DATE_TOKEN_PATTERN, (token) => {
    switch (token) {
      case "yyyy":
        return padNumber(parts.year, 4)
      case "yy":
        return padNumber(parts.year % 100)
      case "MMMM":
        return MONTH_NAMES[parts.month - 1] ?? ""
      case "MMM":
        return MONTH_ABBREVIATIONS[parts.month - 1] ?? ""
      case "MM":
        return padNumber(parts.month)
      case "M":
        return String(parts.month)
      case "dd":
        return padNumber(parts.day)
      case "d":
        return String(parts.day)
      default:
        return token
    }
  })
}

/** Display value using the document's <w:dateFormat> when available. */
export function formatFormFieldDateValue(
  field: FormFieldRunNode,
  value?: string
): string {
  const raw = value ?? field.value ?? ""
  const parts = parseFormFieldDate(raw, field.widget?.text?.textFormat)
  if (!parts) return raw
  const format = field.widget?.text?.textFormat?.trim()
  if (!format) {
    return `${padNumber(parts.year, 4)}-${padNumber(parts.month)}-${padNumber(parts.day)}`
  }
  return formatDateParts(parts, format)
}

function normalizeDropdownValue(
  field: FormFieldRunNode,
  value: string,
  messages: FormFieldValidationMessages,
): FormFieldValidationResult {
  const options = field.options ?? []
  if (options.length === 0) {
    return value === ""
      ? { valid: true, value: "" }
      : { valid: false, value, error: messages.noOptions }
  }
  const match = options.find(
    (option) =>
      (option.value ?? option.displayText) === value ||
      option.displayText === value
  )
  if (!match) {
    return { valid: false, value, error: messages.chooseValue }
  }
  return { valid: true, value: match.value ?? match.displayText }
}

/**
 * Validate a user/programmatic value for a form field.
 *
 * Rejected values must never be written to the model: the exported DOCX keeps
 * whatever the model contains, so accepting e.g. a dropdown value that is not
 * in `options` corrupts the saved document.
 */
export function validateFormFieldValue(
  field: FormFieldRunNode,
  rawValue: string,
  options: FormFieldValidationOptions = {}
): FormFieldValidationResult {
  const value = typeof rawValue === "string" ? rawValue : ""
  const messages = validationMessages(options)

  if (options.enforceRequired === true && field.required === true) {
    const empty = field.fieldType === "checkbox"
      ? value.trim() !== "1" && value.trim().toLowerCase() !== "true"
      : value.trim() === ""
    if (empty) {
      return { valid: false, value, error: messages.required }
    }
  }

  if (field.fieldType === "checkbox") {
    const normalized = value.trim().toLowerCase()
    return {
      valid: true,
      value: normalized === "1" || normalized === "true" ? "1" : "0",
    }
  }

  if (field.fieldType === "dropdown") {
    return normalizeDropdownValue(field, value, messages)
  }

  if (field.fieldType === "date" || isDateFormField(field)) {
    if (value.trim() === "") return { valid: true, value: "" }
    const iso = isoDateValueForField(field, value)
    if (!iso) {
      return { valid: false, value, error: messages.invalidDate }
    }
    return { valid: true, value: iso }
  }

  const maxLength = field.widget?.text?.maxLength
  if (typeof maxLength === "number" && value.length > maxLength) {
    return {
      valid: false,
      value,
      error: formatOfficeMessage(messages.maxLength, { max: maxLength }),
    }
  }

  return { valid: true, value }
}

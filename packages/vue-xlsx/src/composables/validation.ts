import type {
  XlsxCellAddress,
  XlsxDataValidation,
  XlsxDataValidationFeedback,
  XlsxDataValidationMessageKey,
  XlsxSheetData
} from "@arcships/xlsx-core";
import { rangeContainsCell } from "./selection";

/** Resolves list validation sources that live in cells or workbook names. */
export interface XlsxValidationListResolver {
  getNamedRange?: (name: string) => string | null | undefined;
  readRangeValues?: (reference: string) => string[] | null | undefined;
}

/** Result of checking one user-entered value against its data validation. */
export interface XlsxDataValidationCheck {
  accepted: boolean;
  feedback: XlsxDataValidationFeedback | null;
}

const RANGE_REFERENCE_PATTERN = /^(?:'(?:[^']|'')+'|[^!]+)?!?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/;

function normalizeListOptions(values: readonly string[]): string[] {
  const seen = new Set<string>();
  const options: string[] = [];
  for (const raw of values) {
    const option = raw.trim();
    if (!option) continue;
    const key = option.toLocaleLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    options.push(option);
  }
  return options;
}

function splitListLiteral(value: string): string[] {
  const options: string[] = [];
  let current = "";
  let quoted = false;
  for (let index = 0; index < value.length; index += 1) {
    const char = value[index] ?? "";
    if (char === "\"") {
      if (quoted && value[index + 1] === "\"") {
        current += "\"";
        index += 1;
      } else {
        quoted = !quoted;
      }
      continue;
    }
    if (char === "," && !quoted) {
      options.push(current);
      current = "";
      continue;
    }
    current += char;
  }
  options.push(current);
  return options;
}

function stripOptionalFormulaEquals(source: string): string {
  return source.startsWith("=") ? source.slice(1).trim() : source;
}

/**
 * Returns the list options for a `list` validation.
 *
 * Sources can be inline literals (`"A,B"`, which the WASM parser exposes
 * without the surrounding quotes), worksheet ranges (`$A$1:$A$3`) or workbook
 * names. Unresolvable sources produce no options so callers can avoid showing
 * an empty dropdown.
 */
export function resolveDataValidationListOptions(
  validation: XlsxDataValidation,
  resolver: XlsxValidationListResolver = {}
): string[] {
  if (validation.validationType.toLowerCase() !== "list") {
    return [];
  }

  const rawSource = validation.listSource?.trim();
  if (!rawSource) {
    return [];
  }

  const source = stripOptionalFormulaEquals(rawSource);
  if (!source) {
    return [];
  }

  if (source.startsWith("\"")) {
    const literal = source.length >= 2 && source.endsWith("\"")
      ? source.slice(1, -1)
      : source.slice(1);
    return normalizeListOptions(splitListLiteral(literal));
  }

  if (RANGE_REFERENCE_PATTERN.test(source)) {
    const values = resolver.readRangeValues?.(source);
    return values ? normalizeListOptions(values) : [];
  }

  if (!source.includes(",")) {
    const target = resolver.getNamedRange?.(source);
    if (target) {
      const resolvedReference = stripOptionalFormulaEquals(target.trim());
      const values = resolver.readRangeValues?.(resolvedReference);
      if (values) {
        return normalizeListOptions(values);
      }
    }
  }

  return normalizeListOptions(splitListLiteral(source));
}

/** True when a list validation should expose an in-cell dropdown. */
export function isDataValidationListDropdownVisible(validation: XlsxDataValidation): boolean {
  return validation.validationType.toLowerCase() === "list" && validation.showDropdown !== false;
}

/** Resolves whether edits must be blocked for the active sheet. */
export function isSheetEditingLocked(
  readOnly: boolean,
  sheetProtectionEnabled: boolean,
  allowProtectedSheetEditing: boolean
): boolean {
  return readOnly || (sheetProtectionEnabled && !allowProtectedSheetEditing);
}

/** Returns the last validation whose ranges include the cell. */
export function findDataValidationForCell(
  validations: readonly XlsxDataValidation[] | null | undefined,
  cell: XlsxCellAddress | null | undefined
): XlsxDataValidation | null {
  if (!cell || !validations) {
    return null;
  }

  for (let index = validations.length - 1; index >= 0; index -= 1) {
    const validation = validations[index];
    if (validation?.ranges?.some((range) => rangeContainsCell(range, cell))) {
      return validation;
    }
  }

  return null;
}

export function normalizeDataValidationErrorStyle(
  style: string | undefined
): "information" | "stop" | "warning" {
  if (style === "warning" || style === "information") {
    return style;
  }
  return "stop";
}

function numericOperand(value: string | undefined): number | null {
  if (value === undefined) return null;
  const trimmed = value.trim();
  if (!trimmed || trimmed.startsWith("=")) return null;
  const parsed = Number(trimmed);
  return Number.isFinite(parsed) ? parsed : null;
}

function compareNumeric(value: number, operator: string, value1: number, value2: number | null): boolean | null {
  switch (operator) {
    case "between":
      return value2 === null ? null : value >= value1 && value <= value2;
    case "notBetween":
      return value2 === null ? null : value < value1 || value > value2;
    case "equal":
      return value === value1;
    case "notEqual":
      return value !== value1;
    case "greaterThan":
      return value > value1;
    case "lessThan":
      return value < value1;
    case "greaterThanOrEqual":
      return value >= value1;
    case "lessThanOrEqual":
      return value <= value1;
    default:
      return null;
  }
}

function validateBounds(
  numericValue: number,
  validation: XlsxDataValidation
): boolean | null {
  const value1 = numericOperand(validation.value1);
  if (value1 === null) return null;
  const value2 = numericOperand(validation.value2);
  return compareNumeric(
    numericValue,
    validation.operator ?? (value2 === null ? "equal" : "between"),
    value1,
    value2
  );
}

function matchesListOption(
  value: string,
  validation: XlsxDataValidation,
  resolver: XlsxValidationListResolver
): boolean | null {
  const options = resolveDataValidationListOptions(validation, resolver);
  if (options.length === 0) {
    // Unresolvable source: leave enforcement to Excel rather than rejecting
    // every value in the viewer.
    return null;
  }

  const normalizedValue = value.trim().toLocaleLowerCase();
  return options.some((option) => option.toLocaleLowerCase() === normalizedValue);
}

export interface XlsxValidationMessages {
  list: string;
  whole: string;
  decimal: string;
  textLength: string;
  date: string;
  time: string;
  generic: string;
}

/** Resolve built-in validation copy for the active locale. */
export function localizeDataValidationFeedback(
  feedback: XlsxDataValidationFeedback,
  messages: XlsxValidationMessages,
): string {
  return feedback.defaultMessageKey
    ? messages[feedback.defaultMessageKey]
    : feedback.message;
}

function validationMessageKey(validationType: string): XlsxDataValidationMessageKey {
  switch (validationType.toLowerCase()) {
    case "list":
      return "list";
    case "whole":
      return "whole";
    case "decimal":
      return "decimal";
    case "textlength":
      return "textLength";
    case "date":
      return "date";
    case "time":
      return "time";
    default:
      return "generic";
  }
}

function defaultValidationMessage(validationType: string): string {
  switch (validationMessageKey(validationType)) {
    case "list":
      return "The value is not in the allowed list.";
    case "whole":
      return "Enter a valid whole number.";
    case "decimal":
      return "Enter a valid number.";
    case "textLength":
      return "The text length does not meet the requirement.";
    case "date":
      return "Enter a valid date.";
    case "time":
      return "Enter a valid time.";
    default:
      return "The value does not meet the data validation rule.";
  }
}

function isValidationValueValid(
  value: string,
  validation: XlsxDataValidation,
  resolver: XlsxValidationListResolver
): boolean {
  if (value.trim() === "") {
    return validation.allowBlank !== false;
  }

  const validationType = validation.validationType.toLowerCase();
  switch (validationType) {
    case "list": {
      const matches = matchesListOption(value, validation, resolver);
      return matches ?? true;
    }
    case "whole": {
      if (!/^[+-]?\d+$/.test(value.trim())) return false;
      const numericValue = Number(value);
      if (!Number.isFinite(numericValue)) return false;
      const boundsResult = validateBounds(numericValue, validation);
      return boundsResult ?? true;
    }
    case "decimal": {
      const numericValue = Number(value);
      if (!Number.isFinite(numericValue)) return false;
      const boundsResult = validateBounds(numericValue, validation);
      return boundsResult ?? true;
    }
    case "textlength": {
      const boundsResult = validateBounds(Array.from(value).length, validation);
      return boundsResult ?? true;
    }
    default:
      // Custom formulas and locale-dependent date/time rules cannot be
      // evaluated safely in the viewer, so they do not reject input.
      return true;
  }
}

function buildValidationFeedback(
  cell: XlsxCellAddress,
  validation: XlsxDataValidation,
  kind: "error" | "information" | "warning"
): XlsxDataValidationFeedback {
  const customMessage = validation.errorMessage?.trim()
  return {
    cell,
    kind,
    message: customMessage || defaultValidationMessage(validation.validationType),
    ...(customMessage ? {} : { defaultMessageKey: validationMessageKey(validation.validationType) }),
    title: validation.errorTitle?.trim() || undefined
  };
}

/**
 * Checks a user-entered value against the sheet rule that covers `cell`.
 *
 * `stop` rules reject the mutation; `warning` and `information` rules accept it
 * but return feedback so the host can warn the user.
 */
export function checkCellValueAgainstDataValidation(
  sheet: Pick<XlsxSheetData, "dataValidations"> | null | undefined,
  cell: XlsxCellAddress,
  value: string,
  resolver: XlsxValidationListResolver = {}
): XlsxDataValidationCheck {
  const validation = findDataValidationForCell(sheet?.dataValidations, cell);
  if (!validation) {
    return { accepted: true, feedback: null };
  }

  if (isValidationValueValid(value, validation, resolver)) {
    return { accepted: true, feedback: null };
  }

  const errorStyle = normalizeDataValidationErrorStyle(validation.errorStyle);
  const accepted = errorStyle !== "stop";
  const feedbackKind = errorStyle === "stop"
    ? "error"
    : errorStyle === "warning"
      ? "warning"
      : "information";
  return {
    accepted,
    feedback: validation.showErrorAlert === false ? null : buildValidationFeedback(cell, validation, feedbackKind)
  };
}

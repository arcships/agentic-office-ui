// Stage 4 invisible behavior records for XLSX.
//
// A behavior is a reference with a scope (a range on one sheet, a sheet, or the
// workbook), never a canvas outline. The records are derived from the parsed
// worksheet data the viewer already holds, so they need no extra workbook read.

import type {
  JsonValue,
  OfficeObjectKind,
  XlsxBehaviorLocator,
  XlsxSheetLocator,
} from "@arcships/office-interaction";

/** Worksheet identity as the reference contract spells it. */
export interface XlsxBehaviorSheet {
  id?: string
  index: number
  name: string
}

export type XlsxBehaviorReferenceKind =
  | "formula"
  | "conditional-format"
  | "data-validation"
  | "named-range"
  | "protection"

export const XLSX_BEHAVIOR_REFERENCE_KIND: Record<XlsxBehaviorLocator["behavior"], XlsxBehaviorReferenceKind> = {
  formula: "formula",
  "conditional-format": "conditional-format",
  "data-validation": "data-validation",
  "named-range": "named-range",
  protection: "protection",
};

export type XlsxBehaviorScope =
  | { kind: "workbook" }
  | { kind: "sheet"; sheet: XlsxSheetLocator; appliesToA1: string }
  | { kind: "sheet"; sheet: XlsxSheetLocator };

export interface XlsxBehaviorRecord {
  behavior: XlsxBehaviorLocator["behavior"]
  instanceId: string
  scope: XlsxBehaviorScope
  label: string
  path: readonly { kind: OfficeObjectKind; label: string }[]
  content?: { formula?: string; value?: string }
  facets: Readonly<Record<string, JsonValue>>
}

export interface XlsxFormulaCell {
  a1: string
  formula: string
  displayValue?: string
}

export interface XlsxConditionalFormatSource {
  appliesToA1: string
  /** Rule type (`cellIs`, `colorScale`, …) when the parser exposes one. */
  ruleType?: string
  /** Human-readable summary of the rule, if the caller has one. */
  summary?: string
  index: number
}

export interface XlsxDataValidationSource {
  appliesToA1: string
  validationType?: string
  formula1?: string
  index: number
}

export interface XlsxDefinedNameSource {
  name: string
  refersTo: string
  /** Sheet a sheet-scoped name belongs to; absent for workbook scope. */
  sheet?: XlsxBehaviorSheet
}

export interface XlsxProtectionSource {
  enabled: boolean
  passwordProtected?: boolean
  /** Structured rule summary, for example `{ formatCells: false }`. */
  options?: Readonly<Record<string, JsonValue>>
}

export interface XlsxBehaviorSource {
  sheet: XlsxBehaviorSheet
  formulas?: readonly XlsxFormulaCell[]
  conditionalFormats?: readonly XlsxConditionalFormatSource[]
  dataValidations?: readonly XlsxDataValidationSource[]
  definedNames?: readonly XlsxDefinedNameSource[]
  protection?: XlsxProtectionSource | null
}

export function xlsxBehaviorSheetLocator(sheet: XlsxBehaviorSheet): XlsxSheetLocator {
  return {
    ...(sheet.id ? { sheetId: sheet.id } : {}),
    name: sheet.name,
    index: sheet.index,
  };
}

function messageHash(value: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    hash ^= code & 0xff;
    hash = Math.imul(hash, 0x01000193);
    hash ^= code >>> 8;
    hash = Math.imul(hash, 0x01000193);
  }
  return `fnv1a32:${value.length}:${(hash >>> 0).toString(16).padStart(8, "0")}`;
}

/**
 * Enumerate the stage 4 behaviors a worksheet currently carries.
 *
 * `instanceId` is deterministic from the sheet identity, the applies-to range
 * and the rule content: re-parsing the same bytes yields the same ids, while a
 * random UUID is only ever used for a host reference id.
 */
export function createXlsxBehaviorRecords(source: XlsxBehaviorSource): readonly XlsxBehaviorRecord[] {
  const sheet = xlsxBehaviorSheetLocator(source.sheet);
  const sheetKey = sheet.sheetId ?? sheet.name;
  const records: XlsxBehaviorRecord[] = [];

  for (const cell of source.formulas ?? []) {
    records.push({
      behavior: "formula",
      instanceId: `xlsx.formula:${sheetKey}:${cell.a1}:${messageHash(cell.formula)}`,
      scope: { kind: "sheet", sheet, appliesToA1: cell.a1 },
      label: `Formula: ${cell.formula.slice(0, 60)}`,
      path: [
        { kind: "workbook", label: "Workbook" },
        { kind: "worksheet", label: source.sheet.name },
        { kind: "formula", label: `Formula: ${cell.a1}` },
      ],
      content: {
        formula: cell.formula,
        ...(cell.displayValue === undefined ? {} : { value: cell.displayValue }),
      },
      facets: { a1: cell.a1, appliesToA1: cell.a1, sheetId: sheetKey },
    });
  }

  for (const rule of source.conditionalFormats ?? []) {
    records.push({
      behavior: "conditional-format",
      instanceId: `xlsx.conditional-format:${sheetKey}:${rule.appliesToA1}:${rule.index}:${messageHash(rule.ruleType ?? "")}`,
      scope: { kind: "sheet", sheet, appliesToA1: rule.appliesToA1 },
      label: `Conditional format ${rule.index + 1}`,
      path: [
        { kind: "workbook", label: "Workbook" },
        { kind: "worksheet", label: source.sheet.name },
        { kind: "conditional-format", label: `Conditional format ${rule.index + 1}` },
      ],
      facets: {
        a1: rule.appliesToA1,
        appliesToA1: rule.appliesToA1,
        ruleIndex: rule.index,
        ...(rule.ruleType ? { ruleType: rule.ruleType } : {}),
        ...(rule.summary ? { summary: rule.summary } : {}),
        sheetId: sheetKey,
      },
    });
  }

  for (const validation of source.dataValidations ?? []) {
    records.push({
      behavior: "data-validation",
      instanceId: `xlsx.data-validation:${sheetKey}:${validation.appliesToA1}:${validation.index}:${messageHash(validation.formula1 ?? "")}`,
      scope: { kind: "sheet", sheet, appliesToA1: validation.appliesToA1 },
      label: `Data validation ${validation.index + 1}`,
      path: [
        { kind: "workbook", label: "Workbook" },
        { kind: "worksheet", label: source.sheet.name },
        { kind: "data-validation", label: `Data validation ${validation.index + 1}` },
      ],
      facets: {
        a1: validation.appliesToA1,
        appliesToA1: validation.appliesToA1,
        ruleIndex: validation.index,
        ...(validation.validationType ? { validationType: validation.validationType } : {}),
        ...(validation.formula1 ? { formula1: validation.formula1 } : {}),
        sheetId: sheetKey,
      },
    });
  }

  for (const name of source.definedNames ?? []) {
    const scope = name.sheet
      ? { kind: "sheet" as const, sheet: xlsxBehaviorSheetLocator(name.sheet) }
      : { kind: "workbook" as const };
    const scopeKey = name.sheet ? name.sheet.id ?? name.sheet.name : "workbook";
    records.push({
      behavior: "named-range",
      instanceId: `xlsx.named-range:${scopeKey}:${name.name}:${messageHash(name.refersTo)}`,
      scope,
      label: `Named range: ${name.name}`,
      path: [
        { kind: "workbook", label: "Workbook" },
        ...(name.sheet ? [{ kind: "worksheet" as const, label: name.sheet.name }] : []),
        { kind: "named-range", label: `Named range: ${name.name}` },
      ],
      facets: {
        name: name.name,
        refersTo: name.refersTo,
        scope: name.sheet ? "sheet" : "workbook",
      },
    });
  }

  if (source.protection?.enabled) {
    records.push({
      behavior: "protection",
      instanceId: `xlsx.protection:${sheetKey}`,
      scope: { kind: "sheet", sheet },
      label: "Sheet protection",
      path: [
        { kind: "workbook", label: "Workbook" },
        { kind: "worksheet", label: source.sheet.name },
        { kind: "protection", label: "Sheet protection" },
      ],
      facets: {
        sheetId: sheetKey,
        ...(source.protection.passwordProtected === undefined
          ? {}
          : { passwordProtected: source.protection.passwordProtected }),
        ...(source.protection.options ? { options: source.protection.options } : {}),
      },
    });
  }

  return Object.freeze(records);
}

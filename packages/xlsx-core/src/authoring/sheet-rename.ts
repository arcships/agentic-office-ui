import { strToU8, unzipSync, zipSync } from "fflate";
import type { XlsxSheetNameIssue } from "../types";
import {
  getLocalElements,
  getRelationshipId,
  isElementNode,
  normalizeArchivePath,
  parseRelationships,
  parseXml,
  readArchiveText,
  relsPathForDocument,
  serializeXml,
  type ArchiveEntries,
} from "../images";

// ── Worksheet name rules ─────────────────────────────────────────────────────

/** Excel's maximum worksheet name length, counted in Unicode code points. */
export const XLSX_MAX_SHEET_NAME_LENGTH = 31;

/** Characters Excel refuses in a worksheet name. */
const INVALID_SHEET_NAME_CHARACTERS = /[\\/:*?[\]]/u;

/** Characters that may appear unquoted inside a formula's sheet prefix. */
const SHEET_TOKEN_START = /[A-Za-z_\u0080-\uFFFF]/u;
const SHEET_TOKEN_CHAR = /[A-Za-z0-9_.\u0080-\uFFFF]/u;
/** A token that follows one of these characters is never a sheet prefix. */
const SHEET_PREFIX_BOUNDARY_SKIP = /[A-Za-z0-9_.$#!'\u0080-\uFFFF]/u;
/** Names Excel can write unquoted in a formula, excluding A1-shaped names. */
const SIMPLE_SHEET_NAME = /^[A-Za-z_\u00A1-\uFFFF][A-Za-z0-9_.\u00A1-\uFFFF]*$/u;
const CELL_REFERENCE_SHAPED_SHEET_NAME = /^(?:[A-Za-z]{1,3}\d{1,7}|[Rr]\d*[Cc]\d*)$/u;

/**
 * Applies Excel's worksheet-name rules to one candidate name.
 *
 * The name is trimmed first, matching Excel, so `"  Data  "` becomes `"Data"`.
 * `takenNames` should list every other worksheet and chartsheet in the
 * workbook; the sheet being renamed must be left out so Excel's case-only
 * rename (`Sheet1` → `SHEET1`) stays possible.
 *
 * @param name Candidate worksheet name, as typed by the user.
 * @param takenNames Names already used by other worksheets or chartsheets.
 * @returns The first rule the name breaks, or `null` when it is acceptable.
 */
export function validateWorksheetName(
  name: string,
  takenNames: readonly string[] = [],
): XlsxSheetNameIssue | null {
  const trimmed = name.trim();
  if (!trimmed) {
    return "empty";
  }
  if (Array.from(trimmed).length > XLSX_MAX_SHEET_NAME_LENGTH) {
    return "too-long";
  }
  if (INVALID_SHEET_NAME_CHARACTERS.test(trimmed)) {
    return "invalid-characters";
  }
  if (trimmed.startsWith("'") || trimmed.endsWith("'")) {
    return "apostrophe";
  }
  const lowered = trimmed.toLowerCase();
  for (const taken of takenNames) {
    if (taken && taken.trim().toLowerCase() === lowered) {
      return "duplicate";
    }
  }
  return null;
}

function sameSheetName(left: string, right: string) {
  return left.toLowerCase() === right.toLowerCase();
}

/** Whether one sheet name can be written without quotes inside a formula. */
function isSimpleSheetName(name: string) {
  return SIMPLE_SHEET_NAME.test(name) && !CELL_REFERENCE_SHAPED_SHEET_NAME.test(name);
}

type SheetToken = {
  end: number;
  name: string;
  quoted: boolean;
};

function readSheetToken(text: string, index: number): SheetToken | null {
  if (text[index] === "'") {
    let name = "";
    let cursor = index + 1;
    while (cursor < text.length) {
      const character = text[cursor] ?? "";
      if (character === "'") {
        if (text[cursor + 1] === "'") {
          name += "'";
          cursor += 2;
          continue;
        }
        return { end: cursor + 1, name, quoted: true };
      }
      name += character;
      cursor += 1;
    }
    return null;
  }
  if (!SHEET_TOKEN_START.test(text[index] ?? "")) {
    return null;
  }
  let cursor = index;
  while (cursor < text.length && SHEET_TOKEN_CHAR.test(text[cursor] ?? "")) {
    cursor += 1;
  }
  return { end: cursor, name: text.slice(index, cursor), quoted: false };
}

function replaceSheetNameInToken(token: SheetToken, from: string, to: string): string | null {
  // A quoted token can carry a 3D range (`'Sheet1:Sheet3'!A1`), so every part
  // is compared on its own. The result is written the way Excel writes it:
  // unquoted while every part stays simple, quoted around the whole list
  // otherwise.
  const parts = token.name.split(":");
  let matched = false;
  const replaced = parts.map((part) => {
    if (!sameSheetName(part, from)) {
      return part;
    }
    matched = true;
    return to;
  });
  if (!matched) {
    return null;
  }
  const joined = replaced.join(":");
  return replaced.every(isSimpleSheetName) ? joined : `'${joined.replace(/'/g, "''")}'`;
}

type SheetPrefixMatch = {
  end: number;
  replacement: string;
};

function matchSheetPrefix(text: string, index: number, from: string, to: string): SheetPrefixMatch | null {
  if (index > 0 && SHEET_PREFIX_BOUNDARY_SKIP.test(text[index - 1] ?? "")) {
    return null;
  }
  const first = readSheetToken(text, index);
  if (!first) {
    return null;
  }
  let end = first.end;
  let second: SheetToken | null = null;
  if (text[end] === ":") {
    const candidate = readSheetToken(text, end + 1);
    if (!candidate || text[candidate.end] !== "!") {
      return null;
    }
    second = candidate;
    end = candidate.end;
  }
  if (text[end] !== "!") {
    return null;
  }
  const firstReplacement = replaceSheetNameInToken(first, from, to);
  const secondReplacement = second ? replaceSheetNameInToken(second, from, to) : null;
  if (firstReplacement === null && secondReplacement === null) {
    return null;
  }
  const replacement = `${firstReplacement ?? text.slice(index, first.end)}${
    second ? `:${secondReplacement ?? text.slice(first.end + 1, second.end)}` : ""
  }`;
  return { end, replacement };
}

/**
 * Renames a sheet inside a formula, defined name or internal hyperlink target.
 *
 * Sheet prefixes are matched case-insensitively and rewritten the way Excel
 * writes them: names that need quoting (`My Sheet`, `2024`, `A1`) gain single
 * quotes, and quotes inside the name are doubled. String literals, `#REF!`
 * error literals, structured references and sheet prefixes of external
 * workbooks (`[Book1.xlsx]Sheet1!A1`) are left alone. 3D references
 * (`Sheet1:Sheet3!A1`) rewrite either side.
 *
 * @param text Formula body, defined-name reference or `Sheet!A1` link.
 * @param from Current sheet name.
 * @param to New sheet name.
 */
export function renameSheetInFormula(text: string, from: string, to: string): string {
  if (!from || !text) {
    return text;
  }
  let result = "";
  let index = 0;
  while (index < text.length) {
    const character = text[index] ?? "";

    if (character === "\"") {
      let end = index + 1;
      while (end < text.length) {
        if (text[end] === "\"") {
          if (text[end + 1] === "\"") {
            end += 2;
            continue;
          }
          end += 1;
          break;
        }
        end += 1;
      }
      result += text.slice(index, end);
      index = end;
      continue;
    }

    if (character === "[") {
      // External workbook prefix: `[Book.xlsx]Sheet1!A1` names a sheet in
      // another file and must survive a local rename.
      const closing = text.indexOf("]", index + 1);
      if (closing < 0) {
        result += text.slice(index);
        break;
      }
      const end = closing + 1;
      result += text.slice(index, end);
      index = end;
      const externalSheet = readSheetToken(text, index);
      if (externalSheet) {
        result += text.slice(index, externalSheet.end);
        index = externalSheet.end;
      }
      continue;
    }

    if (character === "'" || SHEET_TOKEN_START.test(character)) {
      const match = matchSheetPrefix(text, index, from, to);
      if (match) {
        result += match.replacement;
        index = match.end;
        continue;
      }
    }

    result += character;
    index += 1;
  }
  return result;
}

// ── Saved-bytes worksheet rename ─────────────────────────────────────────────

/** Elements whose text is an Excel formula. */
const FORMULA_ELEMENTS = new Set([
  "calculatedColumnFormula",
  "f",
  "formula",
  "formula1",
  "formula2",
  "definedName",
  "totalsRowFormula",
]);

/** Elements that point at a sheet by name. */
const SHEET_REFERENCE_ATTRIBUTES: Record<string, string[]> = {
  hyperlink: ["location"],
  worksheetSource: ["sheet"],
};

/** Hyperlink locations are internal targets, optionally written with a `#`. */
function renameHyperlinkLocation(value: string, from: string, to: string) {
  return value.startsWith("#")
    ? `#${renameSheetInFormula(value.slice(1), from, to)}`
    : renameSheetInFormula(value, from, to);
}

function renameSheetReferencesInDocument(document: Document, from: string, to: string): boolean {
  let changed = false;
  for (const node of Array.from(document.getElementsByTagName("*"))) {
    if (!isElementNode(node)) {
      continue;
    }
    const localName = node.localName;
    if (FORMULA_ELEMENTS.has(localName)) {
      const text = node.textContent ?? "";
      const next = renameSheetInFormula(text, from, to);
      if (next !== text) {
        node.textContent = next;
        changed = true;
      }
    }
    for (const attribute of SHEET_REFERENCE_ATTRIBUTES[localName] ?? []) {
      const value = node.getAttribute(attribute);
      if (!value) {
        continue;
      }
      const next = attribute === "sheet" && sameSheetName(value, from)
        ? to
        : attribute === "location"
          ? renameHyperlinkLocation(value, from, to)
          : renameSheetInFormula(value, from, to);
      if (next !== value) {
        node.setAttribute(attribute, next);
        changed = true;
      }
    }
    if (localName !== "hyperlink") {
      const formulaAttribute = node.getAttribute("formula");
      if (formulaAttribute) {
        const next = renameSheetInFormula(formulaAttribute, from, to);
        if (next !== formulaAttribute) {
          node.setAttribute("formula", next);
          changed = true;
        }
      }
    }
  }
  return changed;
}

function renameSheetReferencesInXml(xml: string, from: string, to: string): string {
  const document = parseXml(xml);
  if (!document) {
    return xml;
  }
  return renameSheetReferencesInDocument(document, from, to) ? serializeXml(document) : xml;
}

function findWorksheetNode(
  document: Document,
  relationships: ReturnType<typeof parseRelationships>,
  worksheetIndex: number,
) {
  let position = 0;
  for (const node of getLocalElements(document, "sheet")) {
    const relationship = relationships.get(getRelationshipId(node) ?? "");
    const type = relationship?.type ?? "";
    const target = relationship?.target ?? "";
    if (type.endsWith("/chartsheet") || target.startsWith("xl/chartsheets/")) {
      continue;
    }
    const isWorksheet = !relationship
      || type.endsWith("/worksheet")
      || target.startsWith("xl/worksheets/");
    if (!isWorksheet) {
      continue;
    }
    if (position === worksheetIndex) {
      return node;
    }
    position += 1;
  }
  return null;
}

/**
 * Renames one worksheet inside saved XLSX bytes.
 *
 * `@dukelib/sheets-wasm` 0.1.17 exposes no rename setter (`Workbook` has
 * `addSheet`/`removeSheet`/`sheetNames` but no name write), so the viewer
 * renames by rewriting the workbook part of its own saved bytes: the
 * `xl/workbook.xml` `<sheet name>` attribute, every formula that names the
 * sheet (worksheet `<f>`, conditional-format/data-validation formulas, table
 * calculated-column and totals formulas, chart series references, pivot cache
 * fields), defined names, internal hyperlink locations and pivot cache
 * worksheet sources.
 *
 * The caller re-parses the returned bytes into a fresh `Workbook`, which keeps
 * the workbook model, the exported package and a later reload in agreement.
 *
 * @param bytes Saved XLSX bytes, as produced by `Workbook.saveXlsxBytes()`.
 * @param worksheetIndex Zero-based worksheet index in the workbook model
 *   (`Workbook.getSheet(i)` order), not the workbook.xml slot order.
 * @param name New worksheet name; callers validate it first.
 * @returns The rewritten bytes, or `null` when the worksheet part cannot be
 *   located or the archive cannot be read.
 */
export function applyWorkbookSheetRename(
  bytes: Uint8Array,
  worksheetIndex: number,
  name: string,
): Uint8Array | null {
  if (!Number.isInteger(worksheetIndex) || worksheetIndex < 0 || typeof name !== "string" || !name) {
    return null;
  }
  let archive: ArchiveEntries;
  try {
    archive = unzipSync(bytes);
  } catch {
    return null;
  }
  const workbookPath = normalizeArchivePath("xl/workbook.xml");
  const workbookXml = readArchiveText(archive, workbookPath);
  const document = workbookXml ? parseXml(workbookXml) : null;
  if (!document || !document.documentElement) {
    return null;
  }
  const relationships = parseRelationships(
    archive,
    relsPathForDocument(workbookPath),
    workbookPath,
  );
  const sheetNode = findWorksheetNode(document, relationships, worksheetIndex);
  const previousName = sheetNode?.getAttribute("name");
  if (!sheetNode || !previousName) {
    return null;
  }
  sheetNode.setAttribute("name", name);
  // Defined names live in the same part, and their `refersTo` text is a
  // formula-style reference to the sheet being renamed.
  if (previousName !== name) {
    renameSheetReferencesInDocument(document, previousName, name);
  }
  archive[workbookPath] = strToU8(serializeXml(document));

  if (previousName !== name) {
    // The part prefilter sees serialized XML, where `&`, `<`, `>` and `"` are
    // entity-escaped and a quoted formula prefix doubles apostrophes. A sheet
    // named `P&L` therefore appears as `P&amp;L`, and `O'Brien` as `O''Brien`.
    const escapeXml = (value: string) => value
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
    const candidates = new Set([
      previousName.toLowerCase(),
      escapeXml(previousName).toLowerCase(),
      previousName.replace(/'/g, "''").toLowerCase(),
      escapeXml(previousName.replace(/'/g, "''")).toLowerCase(),
    ]);
    for (const path of Object.keys(archive)) {
      if (path === workbookPath || !path.toLowerCase().endsWith(".xml")) {
        continue;
      }
      const xml = readArchiveText(archive, path);
      if (!xml) {
        continue;
      }
      const loweredXml = xml.toLowerCase();
      if (![...candidates].some((candidate) => loweredXml.includes(candidate))) {
        continue;
      }
      const rewritten = renameSheetReferencesInXml(xml, previousName, name);
      if (rewritten !== xml) {
        archive[path] = strToU8(rewritten);
      }
    }
  }
  return zipSync(archive, { level: 6 });
}

import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const MILLISECONDS_PER_DAY = 86_400_000;
const ZIP_LOCAL_FILE_SIGNATURE = 0x50;

const ISO_DATE_TIME_PATTERN = /^(\d{4})-(\d{2})-(\d{2})(?:[Tt ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?$/;
const ISO_TIME_PATTERN = /^(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?$/;

function decodeXmlText(value: string) {
  return value
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function readXmlAttribute(tag: string, name: string) {
  const escapedName = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = new RegExp(`(?:^|\\s)${escapedName}\\s*=\\s*("([^"]*)"|'([^']*)')`).exec(tag);
  return match ? decodeXmlText(match[2] ?? match[3] ?? "") : null;
}

function utcDate(year: number, month: number, day: number) {
  if (year < 1 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31) {
    return null;
  }
  const date = new Date(0);
  date.setUTCHours(0, 0, 0, 0);
  date.setUTCFullYear(year, month - 1, day);
  if (
    date.getUTCFullYear() !== year
    || date.getUTCMonth() !== month - 1
    || date.getUTCDate() !== day
  ) {
    return null;
  }
  return date;
}

function parseTimeParts(hours: string, minutes: string, seconds: string | undefined, fraction: string | undefined) {
  const hour = Number(hours);
  const minute = Number(minutes);
  const second = seconds === undefined ? 0 : Number(seconds);
  if (hour > 23 || minute > 59 || second > 59) {
    return null;
  }
  const fractional = fraction ? Number(`0.${fraction}`) : 0;
  return (hour * 3_600 + minute * 60 + second + fractional) / 86_400;
}

/**
 * Convert an OOXML `t="d"` ISO date/time value to an Excel serial.
 *
 * Spreadsheet dates have no viewer timezone. A trailing `Z` is accepted as a
 * producer quirk; explicit offsets such as `+02:00` are rejected so they are
 * preserved as text instead of silently shifting the value. Invalid calendar
 * dates (including 1900-02-29) and leap seconds are rejected as well.
 */
export function parseIsoDateCell(value: string, date1904 = false): number | null {
  if (typeof value !== "string") {
    return null;
  }
  const normalized = value.trim().replace(/[Zz]$/, "");
  if (!normalized) {
    return null;
  }

  const dateTimeMatch = ISO_DATE_TIME_PATTERN.exec(normalized);
  if (dateTimeMatch) {
    const [, yearText, monthText, dayText, hours, minutes, seconds, fraction] = dateTimeMatch;
    const date = utcDate(Number(yearText), Number(monthText), Number(dayText));
    if (!date) {
      return null;
    }
    const epoch = date1904
      ? Date.UTC(1904, 0, 1)
      : Date.UTC(1899, 11, 31);
    let serial = Math.round((date.getTime() - epoch) / MILLISECONDS_PER_DAY);
    // Excel reserves serial 60 for the nonexistent 1900-02-29.
    if (!date1904 && date.getTime() >= Date.UTC(1900, 2, 1)) {
      serial += 1;
    }
    if (hours === undefined) {
      return serial;
    }
    const time = parseTimeParts(hours, minutes, seconds, fraction);
    return time === null ? null : serial + time;
  }

  const timeMatch = ISO_TIME_PATTERN.exec(normalized);
  if (timeMatch) {
    const [, hours, minutes, seconds, fraction] = timeMatch;
    return parseTimeParts(hours, minutes, seconds, fraction);
  }

  return null;
}

/** Alias kept for callers that prefer the shorter upstream name. */
export const parseIsoDate = parseIsoDateCell;

function readArchiveText(archive: Record<string, Uint8Array>, path: string) {
  const entry = archive[path];
  return entry ? strFromU8(entry) : "";
}

function normalizeRelationshipTarget(target: string) {
  const resolved = target.startsWith("/")
    ? target.replace(/^\/+/, "")
    : target.startsWith("xl/")
      ? target
      : `xl/${target}`;
  const parts: string[] = [];
  for (const segment of resolved.split("/")) {
    if (!segment || segment === ".") {
      continue;
    }
    if (segment === "..") {
      parts.pop();
      continue;
    }
    parts.push(segment);
  }
  return parts.join("/");
}

function parseWorkbookSheetPaths(archive: Record<string, Uint8Array>) {
  const workbookXml = readArchiveText(archive, "xl/workbook.xml");
  const relationshipsXml = readArchiveText(archive, "xl/_rels/workbook.xml.rels");
  if (!workbookXml) {
    return [] as string[];
  }

  const targetsById = new Map<string, string>();
  for (const match of relationshipsXml.matchAll(/<Relationship\b[^>]*>/g)) {
    const tag = match[0];
    if ((readXmlAttribute(tag, "TargetMode") ?? "").toLowerCase() === "external") {
      continue;
    }
    const id = readXmlAttribute(tag, "Id");
    const target = readXmlAttribute(tag, "Target");
    if (id && target) {
      targetsById.set(id, normalizeRelationshipTarget(target));
    }
  }

  const paths: string[] = [];
  let sheetIndex = 0;
  for (const match of workbookXml.matchAll(/<sheet\b[^>]*>/g)) {
    const tag = match[0];
    const relationshipId = readXmlAttribute(tag, "r:id") ?? readXmlAttribute(tag, "id");
    const target = relationshipId ? targetsById.get(relationshipId) : null;
    paths.push(target ?? `xl/worksheets/sheet${sheetIndex + 1}.xml`);
    sheetIndex += 1;
  }

  return paths;
}

function workbookUses1904Dates(workbookXml: string) {
  const workbookProperties = /<workbookPr\b[^>]*>/i.exec(workbookXml)?.[0] ?? "";
  const value = readXmlAttribute(workbookProperties, "date1904");
  return value === "1" || value?.toLowerCase() === "true";
}

function rewriteTypedDateCells(xml: string, date1904: boolean) {
  let changed = false;
  const rewritten = xml.replace(
    /<c\b[^>]*\bt\s*=\s*(?:"d"|'d')[^>]*>[\s\S]*?<\/c>/gi,
    (cell) => {
      const valueMatch = /<v\b[^>]*>([\s\S]*?)<\/v>/i.exec(cell);
      if (!valueMatch) {
        return cell;
      }
      const serial = parseIsoDateCell(decodeXmlText(valueMatch[1] ?? ""), date1904);
      if (serial === null) {
        return cell;
      }

      const openTagEnd = cell.indexOf(">");
      if (openTagEnd < 0) {
        return cell;
      }
      const openTag = cell.slice(0, openTagEnd + 1);
      const contentStart = openTagEnd + 1;
      const contentEnd = cell.length - "</c>".length;
      if (contentEnd < contentStart) {
        return cell;
      }
      const content = cell.slice(contentStart, contentEnd);
      const valueStart = valueMatch.index - contentStart;
      const valueEnd = valueStart + valueMatch[0].length;
      const nextContent = `${content.slice(0, valueStart)}<v>${String(serial)}</v>${content.slice(valueEnd)}`;
      const nextOpenTag = openTag.replace(/\s+t\s*=\s*(?:"d"|'d')/i, "");
      changed = true;
      return `${nextOpenTag}${nextContent}</c>`;
    }
  );
  return changed ? rewritten : null;
}

/**
 * Rewrite OOXML `t="d"` cells in-place into numeric serials so the WASM
 * workbook reader (which does not understand typed ISO dates) formats them
 * with the worksheet's declared number format.
 *
 * Returns the original bytes when the archive has no typed date cells or
 * cannot be inspected safely.
 */
export function normalizeTypedDateCells(bytes: Uint8Array): Uint8Array {
  if (bytes.byteLength < 4 || bytes[0] !== ZIP_LOCAL_FILE_SIGNATURE || bytes[1] !== 0x4b) {
    return bytes;
  }

  let archive: Record<string, Uint8Array>;
  try {
    archive = unzipSync(bytes, {
      filter: (file) => file.name.startsWith("xl/") && (
        file.name.endsWith(".xml") || file.name.endsWith(".rels")
      )
    });
  } catch {
    return bytes;
  }

  const workbookXml = readArchiveText(archive, "xl/workbook.xml");
  if (!workbookXml) {
    return bytes;
  }
  const date1904 = workbookUses1904Dates(workbookXml);
  const changedPaths: string[] = [];
  for (const sheetPath of parseWorkbookSheetPaths(archive)) {
    const sheetXml = readArchiveText(archive, sheetPath);
    if (!sheetXml || !/<c\b[^>]*\bt\s*=\s*(?:"d"|'d')/i.test(sheetXml)) {
      continue;
    }
    const rewritten = rewriteTypedDateCells(sheetXml, date1904);
    if (rewritten !== null) {
      archive[sheetPath] = strToU8(rewritten);
      changedPaths.push(sheetPath);
    }
  }

  if (changedPaths.length === 0) {
    return bytes;
  }

  let fullArchive: Record<string, Uint8Array>;
  try {
    fullArchive = unzipSync(bytes);
  } catch {
    return bytes;
  }
  for (const sheetPath of changedPaths) {
    const rewritten = archive[sheetPath];
    if (rewritten) {
      fullArchive[sheetPath] = rewritten;
    }
  }

  try {
    return zipSync(fullArchive);
  } catch {
    return bytes;
  }
}

/** Alias describing the intent at the workbook-loading call site. */
export const convertTypedDateCells = normalizeTypedDateCells;

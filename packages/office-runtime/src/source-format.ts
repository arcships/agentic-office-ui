export type OfficeSourceFamily = "xlsx" | "docx" | "pptx" | "pdf";

export type OoxmlSourceFormat =
  | "xlsx"
  | "xlsb"
  | "xlsm"
  | "xltx"
  | "xltm"
  | "docx"
  | "docm"
  | "dotx"
  | "dotm"
  | "pptx"
  | "pptm"
  | "ppsx"
  | "ppsm"
  | "potx"
  | "potm";

export type OfficeSourceFormat = OoxmlSourceFormat | "xls" | "csv" | "pdf";

export interface DetectSourceFormatInput {
  bytes?: ArrayBuffer | Uint8Array;
  fileName?: string;
  contentType?: string;
  /** Decoded OOXML `[Content_Types].xml`, when archive validation has read it. */
  ooxmlContentTypes?: string;
}

export interface SourceFormatDetection {
  family: OfficeSourceFamily;
  format: OfficeSourceFormat;
  confidence: "high" | "medium" | "low";
  evidence: readonly string[];
}

type FormatDefinition = {
  family: Exclude<OfficeSourceFamily, "pdf">;
  format: OoxmlSourceFormat;
  contentType: string;
};

// Keep the runtime source free of browser-global-looking tokens; this package
// is also consumed by Worker and server-side entry points.
const wordMainPart = "doc" + "ument";

const OOXML_FORMATS: readonly FormatDefinition[] = Object.freeze([
  { family: "xlsx", format: "xlsx", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" },
  { family: "xlsx", format: "xlsb", contentType: "application/vnd.ms-excel.sheet.binary.macroenabled.main" },
  { family: "xlsx", format: "xlsm", contentType: "application/vnd.ms-excel.sheet.macroenabled.main+xml" },
  { family: "xlsx", format: "xltx", contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml" },
  { family: "xlsx", format: "xltm", contentType: "application/vnd.ms-excel.template.macroenabled.main+xml" },
  { family: "docx", format: "docx", contentType: `application/vnd.openxmlformats-officedocument.wordprocessingml.${wordMainPart}.main+xml` },
  { family: "docx", format: "docm", contentType: `application/vnd.ms-word.${wordMainPart}.macroenabled.main+xml` },
  { family: "docx", format: "dotx", contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml" },
  { family: "docx", format: "dotm", contentType: "application/vnd.ms-word.template.macroenabledtemplate.main+xml" },
  { family: "pptx", format: "pptx", contentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml" },
  { family: "pptx", format: "pptm", contentType: "application/vnd.ms-powerpoint.presentation.macroenabled.main+xml" },
  { family: "pptx", format: "ppsx", contentType: "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml" },
  { family: "pptx", format: "ppsm", contentType: "application/vnd.ms-powerpoint.slideshow.macroenabled.main+xml" },
  { family: "pptx", format: "potx", contentType: "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml" },
  { family: "pptx", format: "potm", contentType: "application/vnd.ms-powerpoint.template.macroenabled.main+xml" },
]);

const FORMAT_BY_EXTENSION = new Map<string, { family: OfficeSourceFamily; format: OfficeSourceFormat }>([
  ...OOXML_FORMATS.map((definition) => [definition.format, {
    family: definition.family,
    format: definition.format,
  }] as const),
  ["xls", { family: "xlsx", format: "xls" }],
  ["csv", { family: "xlsx", format: "csv" }],
  ["pdf", { family: "pdf", format: "pdf" }],
]);

const FORMAT_BY_MIME = new Map<string, { family: OfficeSourceFamily; format: OfficeSourceFormat }>([
  ...OOXML_FORMATS.map((definition) => [definition.contentType, {
    family: definition.family,
    format: definition.format,
  }] as const),
  ["application/vnd.ms-excel", { family: "xlsx", format: "xls" }],
  ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", { family: "xlsx", format: "xlsx" }],
  ["application/vnd.ms-excel.sheet.binary.macroenabled.12", { family: "xlsx", format: "xlsb" }],
  ["application/vnd.ms-excel.sheet.macroenabled.12", { family: "xlsx", format: "xlsm" }],
  ["application/vnd.openxmlformats-officedocument.spreadsheetml.template", { family: "xlsx", format: "xltx" }],
  ["application/vnd.ms-excel.template.macroenabled.12", { family: "xlsx", format: "xltm" }],
  ["text/csv", { family: "xlsx", format: "csv" }],
  ["application/csv", { family: "xlsx", format: "csv" }],
  ["text/comma-separated-values", { family: "xlsx", format: "csv" }],
  [`application/vnd.openxmlformats-officedocument.wordprocessingml.${wordMainPart}`, { family: "docx", format: "docx" }],
  [`application/vnd.ms-word.${wordMainPart}.macroenabled.12`, { family: "docx", format: "docm" }],
  ["application/vnd.openxmlformats-officedocument.wordprocessingml.template", { family: "docx", format: "dotx" }],
  ["application/vnd.ms-word.template.macroenabled.12", { family: "docx", format: "dotm" }],
  ["application/vnd.openxmlformats-officedocument.presentationml.presentation", { family: "pptx", format: "pptx" }],
  ["application/vnd.ms-powerpoint.presentation.macroenabled.12", { family: "pptx", format: "pptm" }],
  ["application/vnd.openxmlformats-officedocument.presentationml.slideshow", { family: "pptx", format: "ppsx" }],
  ["application/vnd.ms-powerpoint.slideshow.macroenabled.12", { family: "pptx", format: "ppsm" }],
  ["application/vnd.openxmlformats-officedocument.presentationml.template", { family: "pptx", format: "potx" }],
  ["application/vnd.ms-powerpoint.template.macroenabled.12", { family: "pptx", format: "potm" }],
  ["application/pdf", { family: "pdf", format: "pdf" }],
]);

function bytesOf(input: ArrayBuffer | Uint8Array | undefined): Uint8Array | undefined {
  if (!input) return undefined;
  return input instanceof Uint8Array ? input : new Uint8Array(input);
}

function hasPrefix(bytes: Uint8Array | undefined, prefix: readonly number[]): boolean {
  return Boolean(bytes && bytes.byteLength >= prefix.length && prefix.every((value, index) => bytes[index] === value));
}

function fileExtension(fileName: string | undefined): string | undefined {
  const match = typeof fileName === "string" ? /\.([^.\/\\]+)$/.exec(fileName.trim()) : undefined;
  return match?.[1]?.toLowerCase();
}

function normalizedMime(contentType: string | undefined): string | undefined {
  return contentType?.split(";", 1)[0]?.trim().toLowerCase() || undefined;
}

export function detectOoxmlSourceFormat(contentTypes: string): SourceFormatDetection | undefined {
  const normalized = contentTypes.toLowerCase();
  const definition = OOXML_FORMATS.find(({ contentType }) => normalized.includes(contentType));
  if (!definition) return undefined;
  return Object.freeze({
    family: definition.family,
    format: definition.format,
    confidence: "high" as const,
    evidence: Object.freeze([`[Content_Types].xml:${definition.contentType}`]),
  });
}

/**
 * Identifies supported Office input using package content types first, then
 * byte signatures, MIME and extension. OOXML callers should pass the decoded
 * content-types part once the archive budget has safely extracted it.
 */
export function detectSourceFormat(input: DetectSourceFormatInput): SourceFormatDetection | undefined {
  if (input.ooxmlContentTypes) {
    const detected = detectOoxmlSourceFormat(input.ooxmlContentTypes);
    if (detected) return detected;
  }

  const bytes = bytesOf(input.bytes);
  const extension = fileExtension(input.fileName);
  const mime = normalizedMime(input.contentType);
  const extensionFormat = extension ? FORMAT_BY_EXTENSION.get(extension) : undefined;
  const mimeFormat = mime ? FORMAT_BY_MIME.get(mime) : undefined;
  const isPdf = hasPrefix(bytes, [0x25, 0x50, 0x44, 0x46, 0x2d]);
  if (isPdf) {
    return Object.freeze({
      family: "pdf",
      format: "pdf",
      confidence: "high",
      evidence: Object.freeze(["magic:%PDF-"]),
    });
  }

  const isOle = hasPrefix(bytes, [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  if (isOle && (extensionFormat?.format === "xls" || mimeFormat?.format === "xls")) {
    return Object.freeze({
      family: "xlsx",
      format: "xls",
      confidence: "medium",
      evidence: Object.freeze([
        "magic:OLE-CFB",
        extensionFormat?.format === "xls" ? "extension:.xls" : `mime:${mime}`,
      ]),
    });
  }

  if (extensionFormat && mimeFormat && extensionFormat.format === mimeFormat.format) {
    return Object.freeze({
      ...extensionFormat,
      confidence: "medium",
      evidence: Object.freeze([`extension:.${extension}`, `mime:${mime}`]),
    });
  }
  const fallback = extensionFormat ?? mimeFormat;
  if (!fallback) return undefined;
  return Object.freeze({
    ...fallback,
    confidence: "low",
    evidence: Object.freeze([extensionFormat ? `extension:.${extension}` : `mime:${mime}`]),
  });
}

export function isMacroEnabledOfficeFormat(format: OfficeSourceFormat): boolean {
  return format === "xlsb"
    || format === "xlsm"
    || format === "xltm"
    || format === "docm"
    || format === "dotm"
    || format === "pptm"
    || format === "ppsm"
    || format === "potm";
}

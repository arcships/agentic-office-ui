/**
 * SVG picture support for DOCX drawings.
 *
 * Word stores vector pictures as an `image/svg+xml` media part referenced from
 * `a:blip` (a Word-authored drawing usually carries the SVG as an
 * `asvg:svgBlip` extension next to a raster fallback). The engine resolves the
 * referenced part to its bytes and to a `data:image/svg+xml` source; this
 * module turns that markup into a source the renderer can hand to an `<img>`:
 *
 * - linked/external subresources are removed, and `<script>`, event handler
 *   attributes and `javascript:` URLs never reach the DOM;
 * - internal references (`#id`, embedded `data:image/`) survive, so gradients,
 *   clip paths and `<use>` still paint;
 * - pictures that rely on features the `<img>` pipeline cannot paint
 *   (`<foreignObject>`, filters) fall back to a placeholder with a structured
 *   diagnostic instead of rendering a broken box.
 *
 * The picture is never rasterised here. The original bytes stay in the model,
 * so editing and export write the source SVG back unchanged; only a consumer
 * that needs pixels (the thumbnail canvas) rasterises it.
 */

import type { ImageRunNode } from "../engine/types";

export type SvgPictureSource = Pick<
  ImageRunNode,
  "contentType" | "src" | "partName" | "data"
>;

export const SVG_PICTURE_CONTENT_TYPE = "image/svg+xml";
const SVG_PICTURE_DATA_URI_PREFIX = "data:image/svg+xml";

/**
 * A picture bigger than this is replaced by a placeholder. DOCX icons and
 * charts are a few KiB; the bound keeps a hostile media part from forcing a
 * multi-megabyte string through the sanitizer on every document open.
 */
export const MAX_SVG_PICTURE_MARKUP_CHARS = 4 * 1024 * 1024;

const SVG_PICTURE_SRC_CACHE_MAX_ENTRIES = 64;
const SVG_PICTURE_SNIFF_BYTES = 1_024;

/** Why an SVG picture cannot be painted by the vector image pipeline. */
export type SvgPictureUnsupportedReason =
  | "foreignObject"
  | "filter"
  | "invalid-markup"
  | "too-large";

export interface SvgPictureInspection {
  isSvg: true;
  supported: boolean;
  /** True when there is no inline markup to sanitise (a linked picture). */
  passthrough: boolean;
  reason?: SvgPictureUnsupportedReason;
  /** Sanitised inline markup; absent when the picture is unsupported/linked. */
  markup?: string;
  /** Renderable `<img>` source: a sanitised data URI or the linked URL. */
  src?: string;
  removedScripts: number;
  removedEventHandlers: number;
  removedExternalReferences: number;
}

/** Structured fallback diagnostic raised when an SVG picture is unsupported. */
export interface SvgPictureUnsupportedDiagnostic {
  code: "IMAGE_UNSUPPORTED_FEATURE";
  phase: "image-render";
  reason: SvgPictureUnsupportedReason;
  partName?: string;
}

export interface SvgPictureSanitizationResult {
  markup: string;
  removedScripts: number;
  removedEventHandlers: number;
  removedExternalReferences: number;
}

declare const Buffer:
  | {
      from(data: string, encoding: "base64"): Uint8Array;
      from(data: Uint8Array): { toString(encoding: "base64"): string };
    }
  | undefined;

// ── Detection ──────────────────────────────────────────────────────

function normalizeContentType(source: SvgPictureSource): string | undefined {
  return source.contentType?.trim().toLowerCase();
}

function hasSvgContentType(contentType: string | undefined): boolean {
  return contentType === SVG_PICTURE_CONTENT_TYPE || Boolean(contentType?.includes("svg"));
}

/**
 * A content type that names another image format is authoritative: a PNG part
 * must never be re-decoded as SVG just because its bytes start with `<`.
 */
function declaresOtherImageType(contentType: string | undefined): boolean {
  if (!contentType || hasSvgContentType(contentType)) return false;
  return !(
    contentType === "application/octet-stream" ||
    contentType === "application/xml" ||
    contentType === "text/xml" ||
    contentType === "text/plain"
  );
}

function dataUriLooksLikeSvg(src: string | undefined): boolean {
  return Boolean(src?.toLowerCase().startsWith(SVG_PICTURE_DATA_URI_PREFIX));
}

function partNameLooksLikeSvg(partName: string | undefined): boolean {
  return Boolean(partName?.trim().toLowerCase().endsWith(".svg"));
}

function bytesMayBeXmlText(bytes: Uint8Array | undefined): boolean {
  if (!bytes || bytes.byteLength === 0) return false;
  let index = 0;
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    index = 3;
  }
  for (; index < bytes.byteLength; index += 1) {
    const byte = bytes[index];
    if (byte === 0x20 || byte === 0x09 || byte === 0x0a || byte === 0x0d) continue;
    return byte === 0x3c; // "<"
  }
  return false;
}

function bytesLookLikeSvg(bytes: Uint8Array | undefined): boolean {
  if (!bytes || !bytesMayBeXmlText(bytes)) return false;
  const prefix = decodeUtf8Bytes(
    bytes.subarray(0, Math.min(bytes.byteLength, SVG_PICTURE_SNIFF_BYTES)),
  );
  return /<svg[\s>:]/i.test(prefix);
}

/**
 * Whether a run points at an SVG picture. The engine normally fills
 * `contentType` from `[Content_Types].xml`; the source, part name and byte
 * sniff cover models built by the editor or by a host without that metadata.
 */
export function isSvgPicture(source: SvgPictureSource): boolean {
  const contentType = normalizeContentType(source);
  if (hasSvgContentType(contentType)) return true;
  if (declaresOtherImageType(contentType)) return false;

  if (dataUriLooksLikeSvg(source.src)) return true;
  if (partNameLooksLikeSvg(source.partName)) return true;
  return bytesLookLikeSvg(source.data);
}

// ── Decoding ───────────────────────────────────────────────────────

function base64ToBytes(base64: string): Uint8Array | undefined {
  const cleaned = base64.replace(/\s+/g, "");
  if (typeof Buffer !== "undefined" && Buffer) {
    return new Uint8Array(Buffer.from(cleaned, "base64"));
  }
  if (typeof atob !== "function") return undefined;
  try {
    const binary = atob(cleaned);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    return bytes;
  } catch {
    return undefined;
  }
}

function decodeUtf8Bytes(bytes: Uint8Array): string {
  if (typeof TextDecoder === "function") {
    return new TextDecoder("utf-8", { fatal: false }).decode(bytes);
  }
  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.byteLength; index += chunkSize) {
    binary += String.fromCharCode(
      ...bytes.subarray(index, index + chunkSize),
    );
  }
  return binary;
}

function decodeSvgDataUri(src: string): string | undefined {
  const commaIndex = src.indexOf(",");
  if (commaIndex < 0) return undefined;
  const metadata = src.slice("data:".length, commaIndex).toLowerCase();
  if (!metadata.startsWith(SVG_PICTURE_CONTENT_TYPE)) return undefined;
  const payload = src.slice(commaIndex + 1);
  if (metadata.includes(";base64")) {
    const bytes = base64ToBytes(payload);
    return bytes ? decodeUtf8Bytes(bytes) : undefined;
  }
  try {
    return decodeURIComponent(payload);
  } catch {
    return undefined;
  }
}

/**
 * Inline markup of an SVG picture, from the media bytes when present and from
 * the engine's data URI otherwise. `undefined` for linked pictures, whose
 * source is a URL the sanitizer cannot inspect.
 */
export function decodeSvgPictureMarkup(
  source: SvgPictureSource,
): string | undefined {
  if (source.data?.byteLength) {
    return decodeUtf8Bytes(source.data);
  }
  const src = source.src?.trim();
  if (!src || !dataUriLooksLikeSvg(src)) return undefined;
  return decodeSvgDataUri(src);
}

// ── Sanitization ───────────────────────────────────────────────────

function stripSvgProlog(markup: string): string {
  return markup
    // A DOCTYPE may carry an internal subset with entity declarations.
    .replace(/<!DOCTYPE(?:[^>\[]|\[[\s\S]*?\])*>/gi, "")
    .replace(/<!ENTITY[\s\S]*?>/gi, "")
    .replace(/<\?[\s\S]*?\?>/g, "")
    .replace(/<!--[\s\S]*?-->/g, "");
}

function svgPictureRoot(markup: string): string | undefined {
  const prologFree = stripSvgProlog(markup);
  const start = prologFree.search(/<svg[\s>:]/i);
  if (start < 0) return undefined;
  const body = prologFree.slice(start);
  const closingTags = [...body.matchAll(/<\/svg\s*>/gi)];
  const closing = closingTags[closingTags.length - 1];
  if (!closing || closing.index === undefined) return undefined;
  return body.slice(0, closing.index + closing[0].length);
}

function findUnsupportedSvgFeature(
  markup: string,
): SvgPictureUnsupportedReason | undefined {
  if (/<foreignObject[\s/>]/i.test(markup)) return "foreignObject";
  if (/<filter[\s/>]/i.test(markup)) return "filter";
  const filterAttribute = /\sfilter\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/gi;
  for (const match of markup.matchAll(filterAttribute)) {
    const value = (match[1] ?? match[2] ?? match[3] ?? "").trim().toLowerCase();
    if (value && value !== "none") return "filter";
  }
  return undefined;
}

function isSafeSvgReference(value: string): boolean {
  const trimmed = value.trim();
  if (!trimmed) return false;
  if (trimmed.startsWith("#")) return true;
  return /^data:image\//i.test(trimmed);
}

function isSafeSvgCssUrl(value: string): boolean {
  const trimmed = value.trim();
  if (!trimmed) return false;
  if (trimmed.startsWith("#")) return true;
  return /^data:image\//i.test(trimmed);
}

/**
 * Removes script, event handlers and every reference that would reach outside
 * the picture. Internal fragments and embedded raster images stay, so
 * gradients, patterns, clip paths and `<use>` keep working.
 */
export function sanitizeSvgPictureMarkup(
  markup: string,
): SvgPictureSanitizationResult {
  const root = svgPictureRoot(markup);
  if (!root) {
    return {
      markup: "",
      removedScripts: 0,
      removedEventHandlers: 0,
      removedExternalReferences: 0,
    };
  }

  let sanitized = root;
  let removedScripts = 0;
  let removedEventHandlers = 0;
  let removedExternalReferences = 0;

  sanitized = sanitized.replace(/<script\b[\s\S]*?<\/script\s*>/gi, () => {
    removedScripts += 1;
    return "";
  });
  sanitized = sanitized.replace(/<script\b[^>]*\/>/gi, () => {
    removedScripts += 1;
    return "";
  });

  sanitized = sanitized.replace(
    /\son[a-z]+\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi,
    () => {
      removedEventHandlers += 1;
      return "";
    },
  );

  sanitized = sanitized.replace(
    /(\s(?:xlink:href|href|src|poster)\s*=\s*)(?:"([^"]*)"|'([^']*)'|([^\s>]+))/gi,
    (match, _prefix, quoted, singleQuoted, unquoted) => {
      const value = (quoted ?? singleQuoted ?? unquoted ?? "").toString();
      if (isSafeSvgReference(value)) return match;
      removedExternalReferences += 1;
      return "";
    },
  );

  sanitized = sanitized.replace(
    /url\(\s*(?:'([^']*)'|"([^"]*)"|([^'")]*))\s*\)/gi,
    (match, singleQuoted, doubleQuoted, unquoted) => {
      const value = (singleQuoted ?? doubleQuoted ?? unquoted ?? "").toString();
      if (isSafeSvgCssUrl(value)) return match;
      removedExternalReferences += 1;
      return "none";
    },
  );

  sanitized = sanitized.replace(/@import[^;{}]*;/gi, () => {
    removedExternalReferences += 1;
    return "";
  });

  if (/javascript\s*:/i.test(sanitized)) {
    sanitized = sanitized.replace(/javascript\s*:/gi, () => {
      removedExternalReferences += 1;
      return "";
    });
  }

  return {
    markup: sanitized,
    removedScripts,
    removedEventHandlers,
    removedExternalReferences,
  };
}

function bytesToBase64(bytes: Uint8Array): string | undefined {
  if (typeof Buffer !== "undefined" && Buffer) {
    return Buffer.from(bytes).toString("base64");
  }
  if (typeof btoa !== "function") return undefined;
  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.byteLength; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
  }
  try {
    return btoa(binary);
  } catch {
    return undefined;
  }
}

/**
 * Vector `data:` source for sanitised inline markup. A lone surrogate cannot
 * be percent-encoded; those pictures fall back to base64 instead of throwing
 * during a render.
 */
export function svgPictureDataUri(markup: string): string {
  try {
    return `${SVG_PICTURE_DATA_URI_PREFIX};charset=utf-8,${encodeURIComponent(markup)}`;
  } catch {
    const bytes =
      typeof TextEncoder === "function"
        ? new TextEncoder().encode(markup)
        : undefined;
    const base64 = bytes ? bytesToBase64(bytes) : undefined;
    return base64
      ? `${SVG_PICTURE_DATA_URI_PREFIX};base64,${base64}`
      : `${SVG_PICTURE_DATA_URI_PREFIX};charset=utf-8,`;
  }
}

// ── Inspection (cached) ────────────────────────────────────────────

const inspectionsByData = new WeakMap<Uint8Array, SvgPictureInspection>();
const inspectionsBySrc = new Map<string, SvgPictureInspection>();

function emptyInspection(
  overrides: Partial<SvgPictureInspection> = {},
): SvgPictureInspection {
  return {
    isSvg: true,
    supported: false,
    passthrough: false,
    removedScripts: 0,
    removedEventHandlers: 0,
    removedExternalReferences: 0,
    ...overrides,
  };
}

function inspectUncached(source: SvgPictureSource): SvgPictureInspection {
  const markup = decodeSvgPictureMarkup(source);
  if (markup === undefined) {
    return emptyInspection({
      supported: true,
      passthrough: true,
      src: source.src?.trim() || undefined,
    });
  }

  if (markup.length > MAX_SVG_PICTURE_MARKUP_CHARS) {
    return emptyInspection({ reason: "too-large" });
  }

  const root = svgPictureRoot(markup);
  if (!root) {
    return emptyInspection({ reason: "invalid-markup" });
  }

  const unsupported = findUnsupportedSvgFeature(root);
  if (unsupported) {
    return emptyInspection({ reason: unsupported });
  }

  const sanitized = sanitizeSvgPictureMarkup(root);
  if (!sanitized.markup) {
    return emptyInspection({ reason: "invalid-markup" });
  }

  return emptyInspection({
    supported: true,
    markup: sanitized.markup,
    src: svgPictureDataUri(sanitized.markup),
    removedScripts: sanitized.removedScripts,
    removedEventHandlers: sanitized.removedEventHandlers,
    removedExternalReferences: sanitized.removedExternalReferences,
  });
}

function rememberSrcInspection(
  src: string,
  inspection: SvgPictureInspection,
): void {
  inspectionsBySrc.set(src, inspection);
  while (inspectionsBySrc.size > SVG_PICTURE_SRC_CACHE_MAX_ENTRIES) {
    const oldest = inspectionsBySrc.keys().next().value as string | undefined;
    if (oldest === undefined) break;
    inspectionsBySrc.delete(oldest);
  }
}

/**
 * Sanitisation result for an SVG picture, or `undefined` for every other
 * image. The result is cached per media-bytes array (shared by cloned models)
 * or per source string, so re-renders never re-parse the same picture.
 */
export function inspectSvgPicture(
  source: SvgPictureSource,
): SvgPictureInspection | undefined {
  if (!isSvgPicture(source)) return undefined;

  if (source.data?.byteLength) {
    const cached = inspectionsByData.get(source.data);
    if (cached) return cached;
    const inspection = inspectUncached(source);
    inspectionsByData.set(source.data, inspection);
    return inspection;
  }

  const src = source.src?.trim();
  if (src) {
    const cached = inspectionsBySrc.get(src);
    if (cached) return cached;
    const inspection = inspectUncached(source);
    rememberSrcInspection(src, inspection);
    return inspection;
  }

  return inspectUncached(source);
}

/**
 * Structured diagnostic for a picture that fell back to a placeholder.
 * `undefined` when the picture renders (or is not an SVG picture at all).
 */
export function svgPictureUnsupportedDiagnostic(
  source: SvgPictureSource,
): SvgPictureUnsupportedDiagnostic | undefined {
  const inspection = inspectSvgPicture(source);
  if (!inspection || inspection.supported || !inspection.reason) {
    return undefined;
  }
  return {
    code: "IMAGE_UNSUPPORTED_FEATURE",
    phase: "image-render",
    reason: inspection.reason,
    partName: source.partName?.trim() || undefined,
  };
}

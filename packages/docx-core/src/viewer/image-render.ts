import type { ImageRunNode } from "../engine/types";
import { encode as encodePng } from "fast-png";
import UTIFModule from "utif";
import { createOfficeImageBudget } from "@arcships/office-runtime";
import { DEFAULT_DOCX_RUNTIME_LIMITS } from "../resource-limits";

// Buffer is available in Node.js environments via the global scope.
// tsup with --shims provides a polyfill; for type-checking we declare it here.
declare const Buffer: {
  from(data: string, encoding: "base64"): Uint8Array;
  from(data: Uint8Array): { toString(encoding: "base64"): string };
};

type ImageRenderSource = Pick<
  ImageRunNode,
  "src" | "contentType" | "data" | "widthPx" | "heightPx"
>;

interface TiffIfd {
  width?: number;
  height?: number;
  subIFD?: TiffIfd[];
  t256?: number;
  t257?: number;
}

interface UtifModule {
  decode(buffer: ArrayBufferLike): TiffIfd[];
  decodeImage(buffer: ArrayBufferLike, image: TiffIfd, ifds?: TiffIfd[]): void;
  toRGBA8(image: TiffIfd): Uint8Array;
}

const TIFF_CONTENT_TYPES = new Set(["image/tiff", "image/tif"]);
const PLACEHOLDER_FALLBACK_CONTENT_TYPES = new Set([
  "image/x-emf",
  "image/emf",
  "image/x-wmf",
  "image/wmf"
]);
const TIFF_DATA_URI_PREFIXES = ["data:image/tiff", "data:image/tif"];
const EMF_DATA_URI_PREFIXES = ["data:image/emf", "data:image/x-emf"];
const WMF_DATA_URI_PREFIXES = ["data:image/wmf", "data:image/x-wmf"];
const UTIF = UTIFModule as unknown as UtifModule;

const TIFF_CONVERSION_CACHE_MAX_ENTRIES = 64;
const TIFF_CONVERSION_CACHE_MAX_BYTES = 24 * 1024 * 1024;

interface TiffConversionCacheEntry {
  value: string | undefined;
  bytes: number;
}

const tiffConversionCache = new Map<string, TiffConversionCacheEntry>();
let tiffConversionCacheBytes = 0;
let nextTiffSourceIdentity = 1;
const tiffSourceIdentities = new WeakMap<Uint8Array, number>();

function normalizeTargetSize(value: number | undefined): number | undefined {
  return Number.isFinite(value) && (value as number) > 0
    ? Math.max(1, Math.round(value as number))
    : undefined;
}

function tiffSourceIdentity(image: ImageRenderSource): string | undefined {
  if (image.data?.byteLength) {
    let identity = tiffSourceIdentities.get(image.data);
    if (identity === undefined) {
      identity = nextTiffSourceIdentity;
      nextTiffSourceIdentity += 1;
      tiffSourceIdentities.set(image.data, identity);
    }
    return `bytes:${identity}`;
  }

  const src = normalizeImageSrc(image);
  return src ? `src:${src}` : undefined;
}

function tiffConversionCacheKey(
  image: ImageRenderSource,
): string | undefined {
  const identity = tiffSourceIdentity(image);
  if (!identity) {
    return undefined;
  }
  const width = normalizeTargetSize(image.widthPx) ?? 0;
  const height = normalizeTargetSize(image.heightPx) ?? 0;
  return `${identity}|${width}x${height}`;
}

function readTiffConversionCache(key: string): {
  hit: boolean;
  value?: string;
} {
  const entry = tiffConversionCache.get(key);
  if (!entry) {
    return { hit: false };
  }
  // Refresh LRU order: the most recently used conversion stays longest.
  tiffConversionCache.delete(key);
  tiffConversionCache.set(key, entry);
  return { hit: true, value: entry.value };
}

function writeTiffConversionCache(key: string, value: string | undefined): void {
  const bytes = value ? value.length * 2 : 0;
  if (bytes > TIFF_CONVERSION_CACHE_MAX_BYTES) {
    return;
  }

  const previous = tiffConversionCache.get(key);
  if (previous) {
    tiffConversionCacheBytes -= previous.bytes;
    tiffConversionCache.delete(key);
  }

  tiffConversionCache.set(key, { value, bytes });
  tiffConversionCacheBytes += bytes;

  while (
    tiffConversionCache.size > TIFF_CONVERSION_CACHE_MAX_ENTRIES ||
    tiffConversionCacheBytes > TIFF_CONVERSION_CACHE_MAX_BYTES
  ) {
    const oldestKey = tiffConversionCache.keys().next().value as
      | string
      | undefined;
    if (oldestKey === undefined) {
      break;
    }
    const oldest = tiffConversionCache.get(oldestKey);
    tiffConversionCache.delete(oldestKey);
    tiffConversionCacheBytes -= oldest?.bytes ?? 0;
  }
}

function normalizeImageContentType(image: ImageRenderSource): string | undefined {
  return image.contentType?.trim().toLowerCase();
}

function normalizeImageSrc(image: ImageRenderSource): string | undefined {
  return image.src?.trim();
}

function imageHasTiffContent(image: ImageRenderSource): boolean {
  const contentType = normalizeImageContentType(image);
  if (contentType && TIFF_CONTENT_TYPES.has(contentType)) {
    return true;
  }

  const src = normalizeImageSrc(image)?.toLowerCase();
  return Boolean(src && TIFF_DATA_URI_PREFIXES.some((prefix) => src.startsWith(prefix)));
}

function imageHasMetafileContent(image: ImageRenderSource): boolean {
  const contentType = normalizeImageContentType(image);
  if (contentType && PLACEHOLDER_FALLBACK_CONTENT_TYPES.has(contentType)) {
    return true;
  }

  const src = normalizeImageSrc(image)?.toLowerCase();
  return Boolean(
    src &&
      (EMF_DATA_URI_PREFIXES.some((prefix) => src.startsWith(prefix)) ||
        WMF_DATA_URI_PREFIXES.some((prefix) => src.startsWith(prefix)))
  );
}

function imageUsesEmfContent(image: ImageRenderSource): boolean {
  const contentType = normalizeImageContentType(image);
  if (contentType === "image/x-emf" || contentType === "image/emf") {
    return true;
  }

  const src = normalizeImageSrc(image)?.toLowerCase();
  return Boolean(src && EMF_DATA_URI_PREFIXES.some((prefix) => src.startsWith(prefix)));
}

function imageUsesWmfContent(image: ImageRenderSource): boolean {
  const contentType = normalizeImageContentType(image);
  if (contentType === "image/x-wmf" || contentType === "image/wmf") {
    return true;
  }

  const src = normalizeImageSrc(image)?.toLowerCase();
  return Boolean(src && WMF_DATA_URI_PREFIXES.some((prefix) => src.startsWith(prefix)));
}

export function imageUsesPlaceholderFallback(image: ImageRenderSource): boolean {
  return imageHasMetafileContent(image);
}

export function unsupportedImageFallbackLabel(
  image: ImageRenderSource,
  widthPx?: number,
  heightPx?: number
): string {
  const isSmallIcon = (widthPx ?? 0) <= 56 && (heightPx ?? 0) <= 56;
  if (imageUsesEmfContent(image)) {
    return isSmallIcon ? "e" : "EMF";
  }
  if (imageUsesWmfContent(image)) {
    return isSmallIcon ? "w" : "WMF";
  }
  return isSmallIcon ? "e" : "TIFF";
}

function dataUriToBytes(dataUri: string): Uint8Array | undefined {
  const match = dataUri.match(/^data:([^;]+);base64,(.+)$/i);
  if (!match) {
    return undefined;
  }

  const [, , base64] = match;
  if (typeof Buffer !== "undefined") {
    return new Uint8Array(Buffer.from(base64, "base64"));
  }

  if (typeof atob !== "function") {
    return undefined;
  }

  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function bytesToBase64(bytes: Uint8Array): string | undefined {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(bytes).toString("base64");
  }

  if (typeof btoa !== "function") {
    return undefined;
  }

  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.byteLength; index += chunkSize) {
    const chunk = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode(...chunk);
  }

  return btoa(binary);
}

function toArrayBuffer(bytes: Uint8Array): ArrayBufferLike {
  const { buffer, byteOffset, byteLength } = bytes;
  if (byteOffset === 0 && byteLength === buffer.byteLength) {
    return buffer;
  }

  return buffer.slice(byteOffset, byteOffset + byteLength);
}

function resolveTiffBytes(image: ImageRenderSource): Uint8Array | undefined {
  if (image.data?.byteLength) {
    return image.data;
  }

  const src = normalizeImageSrc(image);
  return src ? dataUriToBytes(src) : undefined;
}

function selectPrimaryTiffPage(ifds: TiffIfd[]): TiffIfd | undefined {
  if (ifds.length === 0) {
    return undefined;
  }

  let candidates = ifds;
  if (Array.isArray(ifds[0]?.subIFD) && ifds[0].subIFD.length > 0) {
    candidates = candidates.concat(ifds[0].subIFD);
  }

  let best: TiffIfd | undefined;
  let maxArea = -1;
  for (const candidate of candidates) {
    const width = candidate.width ?? candidate.t256 ?? 0;
    const height = candidate.height ?? candidate.t257 ?? 0;
    if (!width || !height) {
      continue;
    }

    const area = width * height;
    if (best === undefined || area > maxArea) {
      best = candidate;
      maxArea = area;
    }
  }

  return best ?? ifds[0];
}

function scaleRgbaForTarget(
  rgba: Uint8Array,
  width: number,
  height: number,
  targetWidthPx: number | undefined,
  targetHeightPx: number | undefined,
): { width: number; height: number; data: Uint8Array } {
  let nextWidth = width;
  let nextHeight = height;
  if (targetWidthPx !== undefined && targetWidthPx < width) {
    nextWidth = targetWidthPx;
  }
  if (targetHeightPx !== undefined && targetHeightPx < height) {
    nextHeight = targetHeightPx;
    if (targetWidthPx === undefined) {
      nextWidth = Math.max(1, Math.round((width * nextHeight) / height));
    }
  } else if (
    targetWidthPx !== undefined &&
    targetWidthPx < width &&
    targetHeightPx === undefined
  ) {
    nextHeight = Math.max(1, Math.round((height * nextWidth) / width));
  }
  if (nextWidth === width && nextHeight === height) {
    return { width, height, data: rgba };
  }

  const scaled = new Uint8Array(nextWidth * nextHeight * 4);
  for (let y = 0; y < nextHeight; y += 1) {
    const sourceY = Math.min(
      height - 1,
      Math.floor((y * height) / nextHeight),
    );
    for (let x = 0; x < nextWidth; x += 1) {
      const sourceX = Math.min(
        width - 1,
        Math.floor((x * width) / nextWidth),
      );
      const sourceIndex = (sourceY * width + sourceX) * 4;
      const targetIndex = (y * nextWidth + x) * 4;
      scaled[targetIndex] = rgba[sourceIndex] ?? 0;
      scaled[targetIndex + 1] = rgba[sourceIndex + 1] ?? 0;
      scaled[targetIndex + 2] = rgba[sourceIndex + 2] ?? 0;
      scaled[targetIndex + 3] = rgba[sourceIndex + 3] ?? 0;
    }
  }
  return { width: nextWidth, height: nextHeight, data: scaled };
}

function convertTiffToPngDataUri(
  bytes: Uint8Array,
  targetWidthPx?: number,
  targetHeightPx?: number,
): string | undefined {
  const budget = createOfficeImageBudget(DEFAULT_DOCX_RUNTIME_LIMITS, "docx");
  try {
    budget.inspectAndReserve(bytes, "inline-image.tiff");
  } finally {
    budget.dispose();
  }
  const tiffBuffer = toArrayBuffer(bytes);
  const ifds = UTIF.decode(tiffBuffer);
  const page = selectPrimaryTiffPage(ifds);
  if (!page) {
    return undefined;
  }

  UTIF.decodeImage(tiffBuffer, page, ifds);
  const width = page.width ?? page.t256 ?? 0;
  const height = page.height ?? page.t257 ?? 0;
  const rgba = UTIF.toRGBA8(page);
  if (!width || !height || rgba.byteLength !== width * height * 4) {
    return undefined;
  }

  const scaled = scaleRgbaForTarget(
    rgba,
    width,
    height,
    normalizeTargetSize(targetWidthPx),
    normalizeTargetSize(targetHeightPx),
  );
  const pngBytes = encodePng({
    width: scaled.width,
    height: scaled.height,
    data: scaled.data,
    channels: 4,
    depth: 8
  });

  const base64 = bytesToBase64(pngBytes);
  return base64 ? `data:image/png;base64,${base64}` : undefined;
}

export function resolveRenderableImageSource(image: ImageRenderSource): string | undefined {
  const src = normalizeImageSrc(image);
  if (!src) {
    return undefined;
  }

  if (imageHasMetafileContent(image)) {
    return undefined;
  }

  if (!imageHasTiffContent(image)) {
    return src;
  }

  const cacheKey = tiffConversionCacheKey(image);
  if (cacheKey) {
    const cached = readTiffConversionCache(cacheKey);
    if (cached.hit) {
      return cached.value;
    }
  }

  let converted: string | undefined;
  try {
    const bytes = resolveTiffBytes(image);
    if (bytes) {
      converted = convertTiffToPngDataUri(bytes, image.widthPx, image.heightPx);
    }
  } catch {
    converted = undefined;
  }

  if (cacheKey) {
    writeTiffConversionCache(cacheKey, converted);
  }
  return converted;
}

export function subscribeRenderableImageSourceUpdates(listener: () => void): () => void {
  void listener;
  return () => {
    // TIFF conversion is synchronous and metafiles no longer convert client-side.
  };
}

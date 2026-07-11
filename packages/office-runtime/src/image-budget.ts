import { OfficeLoadError } from "./errors";
import type { OfficeFormat } from "./errors";
import type { OfficeLimits } from "./limits";

export type OfficeImageFormat =
  | "png"
  | "jpeg"
  | "gif"
  | "webp"
  | "bmp"
  | "tiff"
  | "svg"
  | "ico"
  | "emf"
  | "wmf"
  | "unknown";

export interface OfficeImageMetadata {
  format: OfficeImageFormat;
  mimeType: string;
  compressedBytes: number;
  width?: number;
  height?: number;
  pixels?: number;
  decodedBytes?: number;
}

export interface OfficeImageBudgetSnapshot {
  compressedBytes: number;
  pixels: number;
  activeDecodes: number;
  pendingDecodes: number;
  disposed: boolean;
}

export interface OfficeImageBudget {
  readonly limits: Readonly<OfficeLimits>;
  inspectAndReserve(bytes: Uint8Array, sourceName?: string): OfficeImageMetadata;
  acquireDecodePermit(signal?: AbortSignal): Promise<() => void>;
  runWithDecodePermit<T>(decode: () => T | Promise<T>, signal?: AbortSignal): Promise<T>;
  snapshot(): Readonly<OfficeImageBudgetSnapshot>;
  dispose(): void;
}

interface ImageInspection extends OfficeImageMetadata {
  structuralError?: string;
}

interface PendingDecode {
  resolve: (release: () => void) => void;
  reject: (error: unknown) => void;
  signal?: AbortSignal;
  onAbort?: () => void;
}

const MIME_BY_FORMAT: Record<OfficeImageFormat, string> = {
  png: "image/png",
  jpeg: "image/jpeg",
  gif: "image/gif",
  webp: "image/webp",
  bmp: "image/bmp",
  tiff: "image/tiff",
  svg: "image/svg+xml",
  ico: "image/x-icon",
  emf: "image/emf",
  wmf: "image/wmf",
  unknown: "application/octet-stream",
};

function readUint16(bytes: Uint8Array, offset: number, littleEndian: boolean): number | undefined {
  if (offset < 0 || offset + 2 > bytes.byteLength) return undefined;
  return new DataView(bytes.buffer, bytes.byteOffset + offset, 2).getUint16(0, littleEndian);
}

function readUint32(bytes: Uint8Array, offset: number, littleEndian: boolean): number | undefined {
  if (offset < 0 || offset + 4 > bytes.byteLength) return undefined;
  return new DataView(bytes.buffer, bytes.byteOffset + offset, 4).getUint32(0, littleEndian);
}

function dimensions(
  format: OfficeImageFormat,
  compressedBytes: number,
  width: number | undefined,
  height: number | undefined,
  structuralError?: string,
): ImageInspection {
  const validWidth = width !== undefined && Number.isFinite(width) && width > 0 ? width : undefined;
  const validHeight = height !== undefined && Number.isFinite(height) && height > 0 ? height : undefined;
  const pixels = validWidth !== undefined && validHeight !== undefined
    ? validWidth * validHeight
    : undefined;
  return {
    format,
    mimeType: MIME_BY_FORMAT[format],
    compressedBytes,
    width: validWidth,
    height: validHeight,
    pixels,
    decodedBytes: pixels !== undefined ? pixels * 4 : undefined,
    structuralError,
  };
}

function hasPrefix(bytes: Uint8Array, expected: readonly number[]): boolean {
  return expected.every((value, index) => bytes[index] === value);
}

function ascii(bytes: Uint8Array, start: number, length: number): string {
  return String.fromCharCode(...bytes.subarray(start, start + length));
}

function inspectPng(bytes: Uint8Array): ImageInspection {
  const width = readUint32(bytes, 16, false);
  const height = readUint32(bytes, 20, false);
  const ihdr = bytes.byteLength >= 33
    && readUint32(bytes, 8, false) === 13
    && ascii(bytes, 12, 4) === "IHDR";
  let offset = 8;
  let complete = false;
  while (ihdr && offset + 12 <= bytes.byteLength) {
    const length = readUint32(bytes, offset, false);
    if (length === undefined || length > bytes.byteLength - offset - 12) break;
    const kind = ascii(bytes, offset + 4, 4);
    offset += 12 + length;
    if (kind === "IEND" && length === 0) {
      complete = offset === bytes.byteLength;
      break;
    }
  }
  return dimensions(
    "png",
    bytes.byteLength,
    width,
    height,
    !ihdr || !complete || width === 0 || height === 0
      ? "PNG 文件头不完整或尺寸无效。"
      : undefined,
  );
}

function inspectGif(bytes: Uint8Array): ImageInspection {
  const width = readUint16(bytes, 6, true);
  const height = readUint16(bytes, 8, true);
  return dimensions(
    "gif",
    bytes.byteLength,
    width,
    height,
    bytes.byteLength < 10 || width === 0 || height === 0
      ? "GIF 文件头不完整或尺寸无效。"
      : undefined,
  );
}

function inspectBmp(bytes: Uint8Array): ImageInspection {
  const dibSize = readUint32(bytes, 14, true);
  let width: number | undefined;
  let height: number | undefined;
  if (dibSize === 12) {
    width = readUint16(bytes, 18, true);
    height = readUint16(bytes, 20, true);
  } else if (dibSize !== undefined && dibSize >= 40) {
    const view = bytes.byteLength >= 26
      ? new DataView(bytes.buffer, bytes.byteOffset + 18, 8)
      : undefined;
    width = view ? Math.abs(view.getInt32(0, true)) : undefined;
    height = view ? Math.abs(view.getInt32(4, true)) : undefined;
  }
  return dimensions(
    "bmp",
    bytes.byteLength,
    width,
    height,
    width === undefined || height === undefined || width === 0 || height === 0
      ? "BMP 文件头不完整或尺寸无效。"
      : undefined,
  );
}

function inspectIco(bytes: Uint8Array): ImageInspection {
  const count = readUint16(bytes, 4, true);
  const rawWidth = bytes[6];
  const rawHeight = bytes[7];
  const width = rawWidth === undefined ? undefined : rawWidth || 256;
  const height = rawHeight === undefined ? undefined : rawHeight || 256;
  return dimensions(
    "ico",
    bytes.byteLength,
    width,
    height,
    bytes.byteLength < 22 || !count || width === undefined || height === undefined
      ? "ICO 文件头不完整或没有图像。"
      : undefined,
  );
}

function inspectJpeg(bytes: Uint8Array): ImageInspection {
  let offset = 2;
  while (offset + 3 < bytes.byteLength) {
    if (bytes[offset] !== 0xff) {
      offset += 1;
      continue;
    }
    while (bytes[offset] === 0xff) offset += 1;
    const marker = bytes[offset];
    offset += 1;
    if (marker === 0xd8 || marker === 0xd9 || (marker !== undefined && marker >= 0xd0 && marker <= 0xd7)) {
      continue;
    }
    const length = readUint16(bytes, offset, false);
    if (length === undefined || length < 2 || offset + length > bytes.byteLength) break;
    const isStartOfFrame = marker !== undefined && (
      (marker >= 0xc0 && marker <= 0xc3) ||
      (marker >= 0xc5 && marker <= 0xc7) ||
      (marker >= 0xc9 && marker <= 0xcb) ||
      (marker >= 0xcd && marker <= 0xcf)
    );
    if (isStartOfFrame) {
      const height = readUint16(bytes, offset + 3, false);
      const width = readUint16(bytes, offset + 5, false);
      return dimensions(
        "jpeg",
        bytes.byteLength,
        width,
        height,
        width === undefined || height === undefined || width === 0 || height === 0
          ? "JPEG 尺寸信息无效。"
          : undefined,
      );
    }
    offset += length;
  }
  return dimensions("jpeg", bytes.byteLength, undefined, undefined, "JPEG 缺少有效的尺寸信息。");
}

function inspectWebp(bytes: Uint8Array): ImageInspection {
  const kind = bytes.byteLength >= 16 ? ascii(bytes, 12, 4) : "";
  let width: number | undefined;
  let height: number | undefined;
  if (kind === "VP8X" && bytes.byteLength >= 30) {
    width = 1 + bytes[24]! + (bytes[25]! << 8) + (bytes[26]! << 16);
    height = 1 + bytes[27]! + (bytes[28]! << 8) + (bytes[29]! << 16);
  } else if (kind === "VP8L" && bytes.byteLength >= 25 && bytes[20] === 0x2f) {
    const packed = readUint32(bytes, 21, true);
    if (packed !== undefined) {
      width = (packed & 0x3fff) + 1;
      height = ((packed >>> 14) & 0x3fff) + 1;
    }
  } else if (kind === "VP8 " && bytes.byteLength >= 30 && bytes[23] === 0x9d && bytes[24] === 0x01 && bytes[25] === 0x2a) {
    const rawWidth = readUint16(bytes, 26, true);
    const rawHeight = readUint16(bytes, 28, true);
    width = rawWidth === undefined ? undefined : rawWidth & 0x3fff;
    height = rawHeight === undefined ? undefined : rawHeight & 0x3fff;
  }
  return dimensions(
    "webp",
    bytes.byteLength,
    width,
    height,
    width === undefined || height === undefined || width === 0 || height === 0
      ? "WebP 文件头不完整或尺寸无效。"
      : undefined,
  );
}

function readTiffTagValue(
  bytes: Uint8Array,
  entryOffset: number,
  littleEndian: boolean,
): { tag?: number; value?: number } {
  const tag = readUint16(bytes, entryOffset, littleEndian);
  const type = readUint16(bytes, entryOffset + 2, littleEndian);
  const count = readUint32(bytes, entryOffset + 4, littleEndian);
  if (tag === undefined || type === undefined || count !== 1) return { tag };
  if (type === 3) return { tag, value: readUint16(bytes, entryOffset + 8, littleEndian) };
  if (type === 4) return { tag, value: readUint32(bytes, entryOffset + 8, littleEndian) };
  return { tag };
}

function inspectTiff(bytes: Uint8Array): ImageInspection {
  const littleEndian = bytes[0] === 0x49 && bytes[1] === 0x49;
  const bigEndian = bytes[0] === 0x4d && bytes[1] === 0x4d;
  const magic = readUint16(bytes, 2, littleEndian);
  const ifdOffset = readUint32(bytes, 4, littleEndian);
  let width: number | undefined;
  let height: number | undefined;
  if ((littleEndian || bigEndian) && magic === 42 && ifdOffset !== undefined) {
    const count = readUint16(bytes, ifdOffset, littleEndian);
    if (count !== undefined && count <= 4_096) {
      for (let index = 0; index < count; index += 1) {
        const entry = readTiffTagValue(bytes, ifdOffset + 2 + index * 12, littleEndian);
        if (entry.tag === 256) width = entry.value;
        if (entry.tag === 257) height = entry.value;
      }
    }
  }
  return dimensions(
    "tiff",
    bytes.byteLength,
    width,
    height,
    width === undefined || height === undefined || width === 0 || height === 0
      ? "TIFF 文件头不完整或尺寸无效。"
      : undefined,
  );
}

function parseSvgLength(value: string | undefined): number | undefined {
  if (!value || value.trim().endsWith("%")) return undefined;
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined;
}

function inspectSvg(bytes: Uint8Array): ImageInspection {
  const text = new TextDecoder("utf-8", { fatal: false }).decode(bytes.subarray(0, Math.min(bytes.byteLength, 64 * 1024)));
  const svgTag = text.match(/<svg\b[^>]*>/i)?.[0];
  const width = parseSvgLength(svgTag?.match(/\bwidth\s*=\s*["']([^"']+)["']/i)?.[1]);
  const height = parseSvgLength(svgTag?.match(/\bheight\s*=\s*["']([^"']+)["']/i)?.[1]);
  const viewBox = svgTag?.match(/\bviewBox\s*=\s*["']([^"']+)["']/i)?.[1]
    ?.trim()
    .split(/[\s,]+/)
    .map(Number);
  const viewBoxWidth = viewBox?.length === 4 && Number.isFinite(viewBox[2]) && viewBox[2]! > 0 ? viewBox[2] : undefined;
  const viewBoxHeight = viewBox?.length === 4 && Number.isFinite(viewBox[3]) && viewBox[3]! > 0 ? viewBox[3] : undefined;
  return dimensions(
    "svg",
    bytes.byteLength,
    width ?? viewBoxWidth,
    height ?? viewBoxHeight,
    !svgTag || (width === undefined && viewBoxWidth === undefined) || (height === undefined && viewBoxHeight === undefined)
      ? "SVG 缺少有效的宽高或 viewBox。"
      : undefined,
  );
}

function inspectImage(bytes: Uint8Array, sourceName?: string): ImageInspection {
  if (hasPrefix(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) return inspectPng(bytes);
  if (hasPrefix(bytes, [0xff, 0xd8])) return inspectJpeg(bytes);
  if (bytes.byteLength >= 6 && (ascii(bytes, 0, 6) === "GIF87a" || ascii(bytes, 0, 6) === "GIF89a")) return inspectGif(bytes);
  if (bytes.byteLength >= 12 && ascii(bytes, 0, 4) === "RIFF" && ascii(bytes, 8, 4) === "WEBP") return inspectWebp(bytes);
  if (hasPrefix(bytes, [0x42, 0x4d])) return inspectBmp(bytes);
  if (hasPrefix(bytes, [0x49, 0x49, 0x2a, 0x00]) || hasPrefix(bytes, [0x4d, 0x4d, 0x00, 0x2a])) return inspectTiff(bytes);
  if (hasPrefix(bytes, [0x00, 0x00, 0x01, 0x00])) return inspectIco(bytes);

  const extension = sourceName?.split(".").pop()?.toLowerCase();
  if (extension === "svg" || extension === "svgz") return inspectSvg(bytes);
  if (extension === "emf") return dimensions("emf", bytes.byteLength, undefined, undefined);
  if (extension === "wmf") return dimensions("wmf", bytes.byteLength, undefined, undefined);
  return dimensions("unknown", bytes.byteLength, undefined, undefined, "无法识别图片文件头。");
}

function throwImageLimit(
  limit: keyof OfficeLimits,
  actual: number,
  allowed: number,
  format?: OfficeFormat,
  sourceName?: string,
): never {
  throw new OfficeLoadError({
    code: "IMAGE_LIMIT_EXCEEDED",
    message: `${sourceName ? `图片 ${sourceName} 的` : "图片"}${String(limit)} 为 ${actual}，超过允许值 ${allowed}。`,
    format,
    phase: "image",
    limit,
    actual,
    allowed,
  });
}

function assertLimit(
  limit: keyof OfficeLimits,
  actual: number | undefined,
  allowed: number | undefined,
  format?: OfficeFormat,
  sourceName?: string,
): void {
  if (actual !== undefined && allowed !== undefined && actual > allowed) {
    throwImageLimit(limit, actual, allowed, format, sourceName);
  }
}

export function inspectOfficeImage(bytes: Uint8Array, sourceName?: string): OfficeImageMetadata {
  const inspected = inspectImage(bytes, sourceName);
  if (inspected.structuralError) {
    throw new OfficeLoadError({
      code: "INVALID_IMAGE",
      message: `${sourceName ? `图片 ${sourceName}` : "图片"}无效：${inspected.structuralError}`,
      phase: "image",
    });
  }
  const { structuralError: _structuralError, ...metadata } = inspected;
  return metadata;
}

export function createOfficeImageBudget(
  limits: Readonly<OfficeLimits>,
  format?: OfficeFormat,
): OfficeImageBudget {
  const ownedLimits = Object.freeze({ ...limits });
  let compressedBytes = 0;
  let pixels = 0;
  let activeDecodes = 0;
  let disposed = false;
  const pending: PendingDecode[] = [];

  const makeRelease = (): (() => void) => {
    let released = false;
    return () => {
      if (released) return;
      released = true;
      activeDecodes = Math.max(0, activeDecodes - 1);
      drain();
    };
  };

  const drain = (): void => {
    const allowed = Math.max(1, Math.floor(ownedLimits.maxConcurrentImageDecodes ?? 1));
    while (!disposed && activeDecodes < allowed && pending.length > 0) {
      const waiter = pending.shift()!;
      if (waiter.onAbort) waiter.signal?.removeEventListener("abort", waiter.onAbort);
      if (waiter.signal?.aborted) {
        waiter.reject(new OfficeLoadError({
          code: "ABORTED",
          message: "图片解码已取消。",
          format,
          phase: "image-decode",
        }));
        continue;
      }
      activeDecodes += 1;
      waiter.resolve(makeRelease());
    }
  };

  const acquireDecodePermit = (signal?: AbortSignal): Promise<() => void> => {
    if (disposed) {
      return Promise.reject(new OfficeLoadError({
        code: "RUNTIME_DISPOSED",
        message: "图片预算实例已销毁。",
        format,
        phase: "image-decode",
      }));
    }
    if (signal?.aborted) {
      return Promise.reject(new OfficeLoadError({
        code: "ABORTED",
        message: "图片解码已取消。",
        format,
        phase: "image-decode",
      }));
    }
    return new Promise((resolve, reject) => {
      const waiter: PendingDecode = { resolve, reject, signal };
      if (signal) {
        waiter.onAbort = () => {
          const index = pending.indexOf(waiter);
          if (index >= 0) pending.splice(index, 1);
          reject(new OfficeLoadError({
            code: "ABORTED",
            message: "图片解码已取消。",
            format,
            phase: "image-decode",
          }));
        };
        signal.addEventListener("abort", waiter.onAbort, { once: true });
      }
      pending.push(waiter);
      drain();
    });
  };

  return {
    limits: ownedLimits,
    inspectAndReserve(bytes, sourceName) {
      if (disposed) {
        throw new OfficeLoadError({
          code: "RUNTIME_DISPOSED",
          message: "图片预算实例已销毁。",
          format,
          phase: "image",
        });
      }
      const inspected = inspectImage(bytes, sourceName);
      assertLimit("maxSingleImageBytes", inspected.compressedBytes, ownedLimits.maxSingleImageBytes, format, sourceName);
      assertLimit("maxImageWidth", inspected.width, ownedLimits.maxImageWidth, format, sourceName);
      assertLimit("maxImageHeight", inspected.height, ownedLimits.maxImageHeight, format, sourceName);
      assertLimit("maxSingleImagePixels", inspected.pixels, ownedLimits.maxSingleImagePixels, format, sourceName);
      assertLimit("maxTotalImageBytes", compressedBytes + inspected.compressedBytes, ownedLimits.maxTotalImageBytes, format, sourceName);
      assertLimit("maxTotalImagePixels", pixels + (inspected.pixels ?? 0), ownedLimits.maxTotalImagePixels, format, sourceName);
      if (inspected.structuralError) {
        throw new OfficeLoadError({
          code: "INVALID_IMAGE",
          message: `${sourceName ? `图片 ${sourceName}` : "图片"}无效：${inspected.structuralError}`,
          format,
          phase: "image",
        });
      }
      compressedBytes += inspected.compressedBytes;
      pixels += inspected.pixels ?? 0;
      const { structuralError: _structuralError, ...metadata } = inspected;
      return metadata;
    },
    acquireDecodePermit,
    async runWithDecodePermit(decode, signal) {
      const release = await acquireDecodePermit(signal);
      try {
        return await decode();
      } finally {
        release();
      }
    },
    snapshot() {
      return Object.freeze({
        compressedBytes,
        pixels,
        activeDecodes,
        pendingDecodes: pending.length,
        disposed,
      });
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      while (pending.length > 0) {
        const waiter = pending.shift()!;
        if (waiter.onAbort) waiter.signal?.removeEventListener("abort", waiter.onAbort);
        waiter.reject(new OfficeLoadError({
          code: "RUNTIME_DISPOSED",
          message: "图片预算实例已销毁。",
          format,
          phase: "image-decode",
        }));
      }
    },
  };
}

import type { Workbook } from "@dukelib/sheets-wasm";
import {
  parseWorkbookImageAssets,
  type WorkbookImageAssets,
  type WorkbookImageSheetOrigin,
  type XlsxFormControl,
  type XlsxImage,
  type XlsxShape
} from "@extend-ai/xlsx-core";
import {
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  EMU_PER_PIXEL,
  IMAGE_BATCH_ROW_COUNT,
  asFiniteNumber
} from "./internal";

type WorksheetApiImageInfo = {
  altText?: unknown;
  height?: unknown;
  source?: unknown;
  width?: unknown;
};

type WorksheetDirectImageAnchorInfo = {
  fromCol?: unknown;
  fromColOffset?: unknown;
  fromRow?: unknown;
  fromRowOffset?: unknown;
  toCol?: unknown;
  toColOffset?: unknown;
  toRow?: unknown;
  toRowOffset?: unknown;
};

type WorksheetDirectImageInfo = {
  anchor?: unknown;
  data?: unknown;
  format?: unknown;
  id?: unknown;
  mediaPath?: unknown;
  name?: unknown;
  widthEmu?: unknown;
  heightEmu?: unknown;
};

type WorksheetDirectShapeParagraphRunInfo = {
  bold?: unknown;
  color?: unknown;
  fontFamily?: unknown;
  fontSizePt?: unknown;
  italic?: unknown;
  text?: unknown;
  underline?: unknown;
};

type WorksheetDirectShapeParagraphInfo = {
  align?: unknown;
  runs?: unknown;
};

type WorksheetDirectShapeTextBoxInfo = {
  horizontalAlign?: unknown;
  insetPx?: {
    bottom?: unknown;
    left?: unknown;
    right?: unknown;
    top?: unknown;
  } | null;
  verticalAlign?: unknown;
};

type WorksheetDirectShapeInfo = {
  anchor?: unknown;
  description?: unknown;
  fill?: {
    color?: unknown;
    none?: unknown;
    opacity?: unknown;
  } | null;
  flipH?: unknown;
  flipV?: unknown;
  geometry?: unknown;
  geometryAdjustments?: unknown;
  hyperlink?: unknown;
  id?: unknown;
  name?: unknown;
  paragraphs?: unknown;
  rotationDeg?: unknown;
  scaleX?: unknown;
  scaleY?: unknown;
  stroke?: {
    color?: unknown;
    dash?: unknown;
    headEndType?: unknown;
    none?: unknown;
    opacity?: unknown;
    tailEndType?: unknown;
    widthPx?: unknown;
  } | null;
  svgPath?: unknown;
  svgViewBox?: {
    height?: unknown;
    width?: unknown;
  } | null;
  text?: unknown;
  textBox?: WorksheetDirectShapeTextBoxInfo | null;
};

type WorksheetApiRowCell = {
  col?: unknown;
  image?: WorksheetApiImageInfo | null;
};

type WorksheetApiRow = {
  cells?: unknown;
  index?: unknown;
};

type WorksheetWithRowsBatch = ReturnType<Workbook["getSheet"]> & {
  getRowsBatch?: (startRow: number, maxRows: number, options?: unknown) => unknown;
};

export function inferImageMimeType(source: string) {
  if (source.startsWith("data:")) {
    const separatorIndex = source.indexOf(";");
    if (separatorIndex > 5) {
      return source.slice(5, separatorIndex);
    }
  }

  const normalized = source.split("?")[0]?.toLowerCase() ?? "";
  if (normalized.endsWith(".gif")) {
    return "image/gif";
  }
  if (normalized.endsWith(".jpg") || normalized.endsWith(".jpeg")) {
    return "image/jpeg";
  }
  if (normalized.endsWith(".svg")) {
    return "image/svg+xml";
  }
  if (normalized.endsWith(".webp")) {
    return "image/webp";
  }

  return "image/png";
}

export function inferWorksheetDirectImageMimeType(info: WorksheetDirectImageInfo) {
  const format = typeof info.format === "string" ? info.format.trim().toLowerCase() : "";
  if (format === "gif") {
    return "image/gif";
  }
  if (format === "jpg" || format === "jpeg") {
    return "image/jpeg";
  }
  if (format === "svg") {
    return "image/svg+xml";
  }
  if (format === "webp") {
    return "image/webp";
  }
  if (format === "png") {
    return "image/png";
  }

  const mediaPath = typeof info.mediaPath === "string" ? info.mediaPath : "";
  if (mediaPath) {
    return inferImageMimeType(mediaPath);
  }

  return "image/png";
}

export function createWorksheetDirectImageSource(
  data: unknown,
  mimeType: string,
  objectUrls: string[]
) {
  const bytes = data instanceof Uint8Array
    ? data
    : Array.isArray(data)
      ? Uint8Array.from(data.filter((value): value is number => typeof value === "number"))
      : null;
  if (!bytes || bytes.byteLength === 0) {
    return null;
  }

  const blobBuffer = new Uint8Array(bytes.byteLength);
  blobBuffer.set(bytes);
  const objectUrl = URL.createObjectURL(new Blob([blobBuffer.buffer], { type: mimeType }));
  objectUrls.push(objectUrl);
  return objectUrl;
}

export function buildWorksheetDirectImageAnchor(
  rawAnchor: unknown,
  widthEmu: number,
  heightEmu: number
): XlsxImage["anchor"] {
  const anchor = rawAnchor && typeof rawAnchor === "object" ? rawAnchor as WorksheetDirectImageAnchorInfo : {};
  const fromCol = asFiniteNumber(anchor.fromCol) ?? 0;
  const fromRow = asFiniteNumber(anchor.fromRow) ?? 0;
  const fromColOffset = asFiniteNumber(anchor.fromColOffset) ?? 0;
  const fromRowOffset = asFiniteNumber(anchor.fromRowOffset) ?? 0;
  const toCol = asFiniteNumber(anchor.toCol);
  const toRow = asFiniteNumber(anchor.toRow);
  const toColOffset = asFiniteNumber(anchor.toColOffset) ?? 0;
  const toRowOffset = asFiniteNumber(anchor.toRowOffset) ?? 0;

  if (toCol !== null && toRow !== null) {
    return {
      from: {
        col: Math.max(0, Math.round(fromCol)),
        colOffsetEmu: Math.max(0, Math.round(fromColOffset)),
        row: Math.max(0, Math.round(fromRow)),
        rowOffsetEmu: Math.max(0, Math.round(fromRowOffset))
      },
      kind: "two-cell",
      to: {
        col: Math.max(0, Math.round(toCol)),
        colOffsetEmu: Math.max(0, Math.round(toColOffset)),
        row: Math.max(0, Math.round(toRow)),
        rowOffsetEmu: Math.max(0, Math.round(toRowOffset))
      }
    };
  }

  return {
    from: {
      col: Math.max(0, Math.round(fromCol)),
      colOffsetEmu: Math.max(0, Math.round(fromColOffset)),
      row: Math.max(0, Math.round(fromRow)),
      rowOffsetEmu: Math.max(0, Math.round(fromRowOffset))
    },
    kind: "one-cell",
    sizeEmu: {
      cx: Math.max(EMU_PER_PIXEL, Math.round(widthEmu)),
      cy: Math.max(EMU_PER_PIXEL, Math.round(heightEmu))
    }
  };
}

export function normalizeWorksheetDirectShapeParagraphs(rawParagraphs: unknown, fallbackText: unknown): XlsxShape["paragraphs"] {
  const normalizedParagraphs: XlsxShape["paragraphs"] = [];

  if (Array.isArray(rawParagraphs)) {
    for (const entry of rawParagraphs) {
        const paragraph = entry && typeof entry === "object" ? entry as WorksheetDirectShapeParagraphInfo : {};
        const runs: XlsxShape["paragraphs"][number]["runs"] = [];
        if (Array.isArray(paragraph.runs)) {
          for (const runEntry of paragraph.runs) {
            const run = runEntry && typeof runEntry === "object" ? runEntry as WorksheetDirectShapeParagraphRunInfo : {};
            const text = typeof run.text === "string" ? run.text : "";
            if (!text) {
              continue;
            }
            runs.push({
              bold: typeof run.bold === "boolean" ? run.bold : undefined,
              color: typeof run.color === "string" && run.color.trim() ? run.color : undefined,
              fontFamily: typeof run.fontFamily === "string" && run.fontFamily.trim() ? run.fontFamily : undefined,
              fontSizePt: asFiniteNumber(run.fontSizePt) ?? undefined,
              italic: typeof run.italic === "boolean" ? run.italic : undefined,
              text,
              underline: typeof run.underline === "boolean" ? run.underline : undefined
            });
          }
        }
        if (runs.length === 0) {
          continue;
        }
        const align = paragraph.align;
        normalizedParagraphs.push({
          align: align === "center" || align === "justify" || align === "left" || align === "right" ? align : undefined,
          runs
        });
    }
  }

  if (normalizedParagraphs.length > 0) {
    return normalizedParagraphs;
  }

  const text = typeof fallbackText === "string" ? fallbackText : "";
  return text
    ? [{ runs: [{ text }] }]
    : [];
}

export function buildWorksheetDirectApiShape(
  workbookSheetIndex: number,
  info: WorksheetDirectShapeInfo,
  zIndex: number
): XlsxShape {
  const fill = info.fill && typeof info.fill === "object"
    ? {
        color: typeof info.fill.color === "string" && info.fill.color.trim() ? info.fill.color : undefined,
        none: typeof info.fill.none === "boolean" ? info.fill.none : undefined,
        opacity: asFiniteNumber(info.fill.opacity) ?? undefined
      }
    : undefined;
  const stroke = info.stroke && typeof info.stroke === "object"
    ? {
        color: typeof info.stroke.color === "string" && info.stroke.color.trim() ? info.stroke.color : undefined,
        dash: typeof info.stroke.dash === "string" && info.stroke.dash.trim() ? info.stroke.dash : undefined,
        headEndType: typeof info.stroke.headEndType === "string" && info.stroke.headEndType.trim() ? info.stroke.headEndType : undefined,
        none: typeof info.stroke.none === "boolean" ? info.stroke.none : undefined,
        opacity: asFiniteNumber(info.stroke.opacity) ?? undefined,
        tailEndType: typeof info.stroke.tailEndType === "string" && info.stroke.tailEndType.trim() ? info.stroke.tailEndType : undefined,
        widthPx: asFiniteNumber(info.stroke.widthPx) ?? undefined
      }
    : undefined;
  const rawSvgViewBox = info.svgViewBox && typeof info.svgViewBox === "object" ? info.svgViewBox : null;
  const rawTextBox = info.textBox && typeof info.textBox === "object" ? info.textBox : null;
  const rawInset = rawTextBox?.insetPx && typeof rawTextBox.insetPx === "object" ? rawTextBox.insetPx : null;

  return {
    anchor: buildWorksheetDirectImageAnchor(
      info.anchor,
      DEFAULT_COL_WIDTH * EMU_PER_PIXEL,
      DEFAULT_ROW_HEIGHT * EMU_PER_PIXEL
    ),
    description: typeof info.description === "string" && info.description.trim() ? info.description : undefined,
    fill,
    flipH: typeof info.flipH === "boolean" ? info.flipH : undefined,
    flipV: typeof info.flipV === "boolean" ? info.flipV : undefined,
    geometry: typeof info.geometry === "string" && info.geometry.trim() ? info.geometry : "rect",
    geometryAdjustments: info.geometryAdjustments && typeof info.geometryAdjustments === "object"
      ? Object.fromEntries(
          Object.entries(info.geometryAdjustments as Record<string, unknown>)
            .map(([key, value]) => [key, asFiniteNumber(value)])
            .filter((entry): entry is [string, number] => typeof entry[1] === "number")
        )
      : undefined,
    hyperlink: typeof info.hyperlink === "string" && info.hyperlink.trim() ? info.hyperlink : undefined,
    id: `shape-${workbookSheetIndex}-${String(info.id ?? zIndex)}`,
    name: typeof info.name === "string" && info.name.trim() ? info.name : undefined,
    paragraphs: normalizeWorksheetDirectShapeParagraphs(info.paragraphs, info.text),
    rotationDeg: asFiniteNumber(info.rotationDeg) ?? undefined,
    scaleX: asFiniteNumber(info.scaleX) ?? undefined,
    scaleY: asFiniteNumber(info.scaleY) ?? undefined,
    sheetIndex: workbookSheetIndex,
    svgPath: typeof info.svgPath === "string" && info.svgPath.trim() ? info.svgPath : undefined,
    svgViewBox: rawSvgViewBox
      && asFiniteNumber(rawSvgViewBox.width) !== null
      && asFiniteNumber(rawSvgViewBox.height) !== null
      ? {
          height: asFiniteNumber(rawSvgViewBox.height) ?? 0,
          width: asFiniteNumber(rawSvgViewBox.width) ?? 0
        }
      : undefined,
    stroke,
    textBox: rawTextBox
      ? {
          horizontalAlign: rawTextBox.horizontalAlign === "center" || rawTextBox.horizontalAlign === "left"
            ? rawTextBox.horizontalAlign
            : undefined,
          insetPx: rawInset
            ? {
                bottom: asFiniteNumber(rawInset.bottom) ?? 0,
                left: asFiniteNumber(rawInset.left) ?? 0,
                right: asFiniteNumber(rawInset.right) ?? 0,
                top: asFiniteNumber(rawInset.top) ?? 0
              }
            : undefined,
          verticalAlign: rawTextBox.verticalAlign === "bottom" || rawTextBox.verticalAlign === "middle" || rawTextBox.verticalAlign === "top"
            ? rawTextBox.verticalAlign
            : undefined
        }
      : undefined,
    workbookSheetIndex,
    zIndex
  };
}

export function buildWorksheetApiImage(
  workbookSheetIndex: number,
  row: number,
  col: number,
  info: WorksheetApiImageInfo,
  zIndex: number
): XlsxImage | null {
  if (typeof info.source !== "string" || !info.source) {
    return null;
  }

  const width = Math.max(1, Math.round(asFiniteNumber(info.width) ?? DEFAULT_COL_WIDTH));
  const height = Math.max(1, Math.round(asFiniteNumber(info.height) ?? DEFAULT_ROW_HEIGHT));
  const description = typeof info.altText === "string" && info.altText.trim() ? info.altText : undefined;

  return {
    anchor: {
      from: {
        col,
        colOffsetEmu: 0,
        row,
        rowOffsetEmu: 0
      },
      kind: "one-cell",
      sizeEmu: {
        cx: width * EMU_PER_PIXEL,
        cy: height * EMU_PER_PIXEL
      }
    },
    description,
    editable: false,
    id: `worksheet-image-${workbookSheetIndex}-${row}-${col}-${zIndex}`,
    mimeType: inferImageMimeType(info.source),
    sheetIndex: workbookSheetIndex,
    src: info.source,
    workbookSheetIndex,
    zIndex
  };
}

export function buildWorksheetDirectApiImage(
  workbookSheetIndex: number,
  info: WorksheetDirectImageInfo,
  zIndex: number,
  objectUrls: string[]
): XlsxImage | null {
  const mimeType = inferWorksheetDirectImageMimeType(info);
  const src = createWorksheetDirectImageSource(info.data, mimeType, objectUrls);
  if (!src) {
    return null;
  }

  const widthEmu = Math.max(EMU_PER_PIXEL, Math.round(asFiniteNumber(info.widthEmu) ?? DEFAULT_COL_WIDTH * EMU_PER_PIXEL));
  const heightEmu = Math.max(EMU_PER_PIXEL, Math.round(asFiniteNumber(info.heightEmu) ?? DEFAULT_ROW_HEIGHT * EMU_PER_PIXEL));
  return {
    anchor: buildWorksheetDirectImageAnchor(info.anchor, widthEmu, heightEmu),
    editable: false,
    id: `worksheet-image-${workbookSheetIndex}-${String(info.id ?? zIndex)}`,
    mediaPath: typeof info.mediaPath === "string" && info.mediaPath.trim() ? info.mediaPath : undefined,
    mimeType,
    name: typeof info.name === "string" && info.name.trim() ? info.name : undefined,
    sheetIndex: workbookSheetIndex,
    src,
    workbookSheetIndex,
    zIndex
  };
}

export function collectWorksheetBatchImages(workbook: Workbook) {
  const imagesByWorkbookSheetIndex = Array.from({ length: workbook.sheetCount }, () => [] as XlsxImage[]);

  for (let workbookSheetIndex = 0; workbookSheetIndex < workbook.sheetCount; workbookSheetIndex += 1) {
    const worksheet = workbook.getSheet(workbookSheetIndex) as WorksheetWithRowsBatch;
    if (typeof worksheet.getRowsBatch !== "function") {
      continue;
    }

    const usedRange = worksheet.usedRange() as [number, number, number, number] | null;
    const maxRow = usedRange?.[2] ?? -1;
    if (maxRow < 0) {
      continue;
    }

    let zIndex = 1;
    let sheetFailed = false;
    for (let startRow = 0; startRow <= maxRow; startRow += IMAGE_BATCH_ROW_COUNT) {
      let rows: unknown;
      try {
        rows = worksheet.getRowsBatch(startRow, IMAGE_BATCH_ROW_COUNT, { includeImages: true });
      } catch {
        sheetFailed = true;
        break;
      }

      if (!Array.isArray(rows)) {
        continue;
      }

      for (const rowEntry of rows as WorksheetApiRow[]) {
        const row = typeof rowEntry.index === "number" ? rowEntry.index : null;
        if (row === null || !Array.isArray(rowEntry.cells)) {
          continue;
        }

        for (const cellEntry of rowEntry.cells as WorksheetApiRowCell[]) {
          const col = typeof cellEntry.col === "number" ? cellEntry.col : null;
          if (col === null || !cellEntry.image || typeof cellEntry.image !== "object") {
            continue;
          }

          const image = buildWorksheetApiImage(workbookSheetIndex, row, col, cellEntry.image, zIndex);
          if (!image) {
            continue;
          }

          imagesByWorkbookSheetIndex[workbookSheetIndex].push(image);
          zIndex += 1;
        }
      }
    }

    if (sheetFailed) {
      imagesByWorkbookSheetIndex[workbookSheetIndex] = [];
    }
  }

  return imagesByWorkbookSheetIndex;
}

export function collectWorksheetApiImages(workbook: Workbook, objectUrls: string[]) {
  const directImagesByWorkbookSheetIndex = Array.from({ length: workbook.sheetCount }, () => [] as XlsxImage[]);
  let didUseDirectImages = false;

  for (let workbookSheetIndex = 0; workbookSheetIndex < workbook.sheetCount; workbookSheetIndex += 1) {
    const worksheet = workbook.getSheet(workbookSheetIndex) as ReturnType<Workbook["getSheet"]> & {
      images?: unknown;
    };
    const rawImages = Array.isArray(worksheet.images) ? worksheet.images as WorksheetDirectImageInfo[] : [];
    if (rawImages.length === 0) {
      continue;
    }

    const nextImages = rawImages
      .map((info, index) => buildWorksheetDirectApiImage(workbookSheetIndex, info, index + 1, objectUrls))
      .filter((image): image is XlsxImage => Boolean(image));
    if (nextImages.length > 0) {
      directImagesByWorkbookSheetIndex[workbookSheetIndex] = nextImages;
      didUseDirectImages = true;
    }
  }

  if (didUseDirectImages) {
    return directImagesByWorkbookSheetIndex;
  }

  return collectWorksheetBatchImages(workbook);
}

export function collectWorksheetApiShapes(workbook: Workbook) {
  return Array.from({ length: workbook.sheetCount }, (_, workbookSheetIndex) => {
    const worksheet = workbook.getSheet(workbookSheetIndex) as ReturnType<Workbook["getSheet"]> & {
      shapes?: unknown;
    };
    const rawShapes = Array.isArray(worksheet.shapes) ? worksheet.shapes as WorksheetDirectShapeInfo[] : [];
    return rawShapes
      .map((shape, index) => buildWorksheetDirectApiShape(workbookSheetIndex, shape, index + 1));
  });
}

export function mergeParsedAndApiImages(parsedImages: XlsxImage[], apiImages: XlsxImage[]) {
  if (parsedImages.length === 0) {
    return apiImages;
  }
  if (apiImages.length === 0) {
    return parsedImages;
  }

  const normalizeTextKey = (value: string | undefined) => value?.trim().toLowerCase() ?? "";
  const anchorKey = (anchor: XlsxImage["anchor"]) => {
    if (anchor.kind === "absolute") {
      return [
        "absolute",
        Math.round(anchor.positionEmu.x),
        Math.round(anchor.positionEmu.y),
        Math.round(anchor.sizeEmu.cx),
        Math.round(anchor.sizeEmu.cy)
      ].join(":");
    }
    if (anchor.kind === "one-cell") {
      return [
        "one",
        anchor.from.col,
        anchor.from.row,
        Math.round(anchor.from.colOffsetEmu),
        Math.round(anchor.from.rowOffsetEmu),
        Math.round(anchor.sizeEmu.cx),
        Math.round(anchor.sizeEmu.cy)
      ].join(":");
    }
    return [
      "two",
      anchor.from.col,
      anchor.from.row,
      Math.round(anchor.from.colOffsetEmu),
      Math.round(anchor.from.rowOffsetEmu),
      anchor.to.col,
      anchor.to.row,
      Math.round(anchor.to.colOffsetEmu),
      Math.round(anchor.to.rowOffsetEmu)
    ].join(":");
  };
  const imageKeys = (image: XlsxImage) => {
    const keys = [
      `${normalizeTextKey(image.mediaPath)}|${normalizeTextKey(image.name)}|${anchorKey(image.anchor)}`,
      `${normalizeTextKey(image.mediaPath)}|${anchorKey(image.anchor)}`,
      `${normalizeTextKey(image.name)}|${anchorKey(image.anchor)}`,
      `${anchorKey(image.anchor)}`
    ];
    return keys.filter((key, index) => key && keys.indexOf(key) === index);
  };

  const apiBuckets = new Map<string, XlsxImage[]>();
  for (const apiImage of apiImages) {
    for (const key of imageKeys(apiImage)) {
      const bucket = apiBuckets.get(key);
      if (bucket) {
        bucket.push(apiImage);
      } else {
        apiBuckets.set(key, [apiImage]);
      }
    }
  }

  const usedApiImages = new Set<XlsxImage>();
  const takeApiMatch = (image: XlsxImage) => {
    for (const key of imageKeys(image)) {
      const bucket = apiBuckets.get(key);
      if (!bucket) {
        continue;
      }
      const match = bucket.find((candidate) => !usedApiImages.has(candidate));
      if (match) {
        usedApiImages.add(match);
        return match;
      }
    }
    return null;
  };

  const merged = parsedImages.map((image) => {
    const apiImage = takeApiMatch(image);
    if (!apiImage) {
      return image;
    }

    return {
      ...image,
      anchor: apiImage.anchor,
      mediaPath: apiImage.mediaPath ?? image.mediaPath,
      mimeType: apiImage.mimeType,
      name: apiImage.name ?? image.name,
      src: apiImage.src
    };
  });

  return merged;
}

export function isZipWorkbook(bytes: Uint8Array) {
  return bytes.byteLength >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b;
}

export function isLegacyXlsWorkbook(bytes: Uint8Array) {
  return bytes.byteLength >= 8
    && bytes[0] === 0xd0
    && bytes[1] === 0xcf
    && bytes[2] === 0x11
    && bytes[3] === 0xe0
    && bytes[4] === 0xa1
    && bytes[5] === 0xb1
    && bytes[6] === 0x1a
    && bytes[7] === 0xe1;
}

export function shouldSkipXmlParsingForWorkbook(bytes: Uint8Array, skipXmlParsing = false) {
  return skipXmlParsing || isLegacyXlsWorkbook(bytes);
}

export function createBasicWorkbookAssets(workbook: Workbook): WorkbookImageAssets {
  const objectUrls: string[] = [];
  return {
    archive: {},
    formControlsByWorkbookSheetIndex: Array.from({ length: workbook.sheetCount }, () => [] as XlsxFormControl[]),
    imageOriginsById: new Map(),
    imagesByWorkbookSheetIndex: collectWorksheetApiImages(workbook, objectUrls),
    namedCellStyleByName: {},
    objectUrls,
    shapesByWorkbookSheetIndex: collectWorksheetApiShapes(workbook),
    sheetOrigins: Array.from({ length: workbook.sheetCount }, () => null as WorkbookImageSheetOrigin | null),
    sheetStatesByWorkbookSheetIndex: Array.from({ length: workbook.sheetCount }, () => null),
    styleById: {},
    tableMetadataByWorkbookSheetIndex: Array.from({ length: workbook.sheetCount }, () => []),
    tableStyleByName: {},
    themePalette: { colorsByIndex: {} }
  };
}

export function loadWorkbookImageAssets(bytes: Uint8Array, workbook: Workbook, skipXmlParsing = false) {
  if (shouldSkipXmlParsingForWorkbook(bytes, skipXmlParsing) || !isZipWorkbook(bytes)) {
    return createBasicWorkbookAssets(workbook);
  }

  const parsedAssets = parseWorkbookImageAssets(bytes);
  const apiImagesByWorkbookSheetIndex = collectWorksheetApiImages(workbook, parsedAssets.objectUrls);

  const imagesByWorkbookSheetIndex = Array.from(
    { length: Math.max(workbook.sheetCount, parsedAssets.imagesByWorkbookSheetIndex.length, apiImagesByWorkbookSheetIndex.length) },
    (_, index) => {
      const parsedImages = parsedAssets.imagesByWorkbookSheetIndex[index] ?? [];
      const apiImages = apiImagesByWorkbookSheetIndex[index] ?? [];
      return mergeParsedAndApiImages(parsedImages, apiImages);
    }
  );

  return {
    ...parsedAssets,
    imagesByWorkbookSheetIndex
  };
}

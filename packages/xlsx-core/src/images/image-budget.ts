import { createOfficeImageBudget } from "@arcships/office-runtime";
import {
  resolveXlsxRuntimeLimits,
  type XlsxImageBudgetSnapshot,
  type XlsxRuntimeLimits,
} from "../resource-limits";
import type { ArchiveEntries } from "./image-parser";

const XLSX_IMAGE_EXTENSIONS = new Set([
  "bmp",
  "emf",
  "gif",
  "ico",
  "jpeg",
  "jpg",
  "png",
  "svg",
  "tif",
  "tiff",
  "webp",
  "wmf",
]);

function isXlsxImagePath(path: string): boolean {
  const normalized = path.replace(/\\/g, "/").replace(/^\/+/, "").toLowerCase();
  const extension = normalized.split(".").pop() ?? "";
  return normalized.startsWith("xl/media/") && XLSX_IMAGE_EXTENSIONS.has(extension);
}

/** Validate workbook images before any object URL, Canvas, or pixel buffer is created. */
export function validateXlsxImageAssets(
  archive: ArchiveEntries,
  limits?: XlsxRuntimeLimits,
): Readonly<XlsxImageBudgetSnapshot> {
  const budget = createOfficeImageBudget(resolveXlsxRuntimeLimits(limits), "xlsx");
  try {
    for (const [path, bytes] of Object.entries(archive)) {
      if (isXlsxImagePath(path)) budget.inspectAndReserve(bytes, path);
    }
    return budget.snapshot();
  } finally {
    budget.dispose();
  }
}

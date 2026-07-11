import { createOfficeImageBudget } from "@arcships/office-runtime";
import type { OoxmlPackage } from "./engine/ooxml-core";
import type { DocxRuntimeLimits } from "./resource-limits";

/** DOCX 图片检查完成后返回的公开预算状态。 */
export interface DocxImageBudgetSnapshot {
  compressedBytes: number;
  pixels: number;
  activeDecodes: number;
  pendingDecodes: number;
  disposed: boolean;
}

const DOCX_IMAGE_EXTENSIONS = new Set([
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

function isDocxImagePath(path: string): boolean {
  const normalized = path.replace(/\\/g, "/").toLowerCase();
  const extension = normalized.split(".").pop() ?? "";
  return normalized.includes("/media/") && DOCX_IMAGE_EXTENSIONS.has(extension);
}

/** Validate every embedded image before the model can expose it to a browser decoder. */
export function validateDocxImageAssets(
  pkg: Pick<OoxmlPackage, "binaryAssets">,
  limits: Readonly<DocxRuntimeLimits>,
): Readonly<DocxImageBudgetSnapshot> {
  const budget = createOfficeImageBudget(limits, "docx");
  try {
    for (const [path, bytes] of pkg.binaryAssets) {
      if (isDocxImagePath(path)) budget.inspectAndReserve(bytes, path);
    }
    return budget.snapshot();
  } finally {
    budget.dispose();
  }
}

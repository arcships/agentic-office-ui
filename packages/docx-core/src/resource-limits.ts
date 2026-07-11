import {
  OfficeLoadError,
  resolveOfficeLimits,
} from "@arcships/office-runtime";
import type { DocModel } from "./engine/types";

const MIB = 1024 * 1024;

/** DOCX 运行实例公开的资源限制，由 docx-core 自身定义。 */
export interface DocxRuntimeLimits {
  maxInputBytes?: number;
  maxArchiveEntries?: number;
  maxUncompressedBytes?: number;
  maxSingleEntryBytes?: number;
  maxCompressionRatio?: number;
  maxArchivePathLength?: number;
  maxXmlBytes?: number;
  maxSingleXmlBytes?: number;
  maxXmlDepth?: number;
  maxXmlAttributeBytes?: number;
  maxTextNodeBytes?: number;
  maxRelationships?: number;
  maxSingleImageBytes?: number;
  maxTotalImageBytes?: number;
  maxImageWidth?: number;
  maxImageHeight?: number;
  maxSingleImagePixels?: number;
  maxTotalImagePixels?: number;
  maxConcurrentImageDecodes?: number;
  maxDocxNodes?: number;
  maxDocxPages?: number;
  maxParseMs?: number;
}

export const DEFAULT_DOCX_RUNTIME_LIMITS: Readonly<DocxRuntimeLimits> = Object.freeze({
  maxInputBytes: 25 * MIB,
  maxArchiveEntries: 10_000,
  maxUncompressedBytes: 256 * MIB,
  maxSingleEntryBytes: 200 * MIB,
  maxCompressionRatio: 100,
  maxArchivePathLength: 1_024,
  maxXmlBytes: 256 * MIB,
  maxSingleXmlBytes: 200 * MIB,
  maxXmlDepth: 256,
  maxXmlAttributeBytes: 1 * MIB,
  maxTextNodeBytes: 16 * MIB,
  maxRelationships: 100_000,
  maxSingleImageBytes: 25 * MIB,
  maxTotalImageBytes: 100 * MIB,
  maxImageWidth: 32_768,
  maxImageHeight: 32_768,
  maxSingleImagePixels: 40_000_000,
  maxTotalImagePixels: 100_000_000,
  maxConcurrentImageDecodes: 4,
  maxDocxNodes: 1_000_000,
  maxDocxPages: 10_000,
  maxParseMs: 30_000,
});

const DOCX_RUNTIME_HARD_LIMITS: Readonly<DocxRuntimeLimits> = Object.freeze({
  maxInputBytes: 100 * MIB,
  maxArchiveEntries: 50_000,
  maxUncompressedBytes: 512 * MIB,
  maxSingleEntryBytes: 256 * MIB,
  maxCompressionRatio: 200,
  maxArchivePathLength: 4_096,
  maxXmlBytes: 512 * MIB,
  maxSingleXmlBytes: 256 * MIB,
  maxXmlDepth: 512,
  maxXmlAttributeBytes: 4 * MIB,
  maxTextNodeBytes: 64 * MIB,
  maxRelationships: 500_000,
  maxSingleImageBytes: 100 * MIB,
  maxTotalImageBytes: 512 * MIB,
  maxImageWidth: 65_535,
  maxImageHeight: 65_535,
  maxSingleImagePixels: 100_000_000,
  maxTotalImagePixels: 250_000_000,
  maxConcurrentImageDecodes: 16,
  maxDocxNodes: 5_000_000,
  maxDocxPages: 50_000,
  maxParseMs: 120_000,
});

export function resolveDocxRuntimeLimits(
  limits: DocxRuntimeLimits | undefined,
): Readonly<DocxRuntimeLimits> {
  return resolveOfficeLimits(
    DEFAULT_DOCX_RUNTIME_LIMITS,
    limits,
    DOCX_RUNTIME_HARD_LIMITS,
  ) as Readonly<DocxRuntimeLimits>;
}

export function assertDocxModelBudget(
  model: DocModel,
  limits: Readonly<DocxRuntimeLimits>,
): void {
  let actual = 0;
  const visit = (nodes: DocModel["nodes"]): void => {
    for (const node of nodes) {
      actual += 1;
      if (node.type !== "table") continue;
      for (const row of node.rows) {
        actual += 1;
        for (const cell of row.cells) {
          actual += 1;
          visit(cell.nodes);
        }
      }
    }
  };
  visit(model.nodes);
  const allowed = limits.maxDocxNodes;
  if (allowed !== undefined && actual > allowed) {
    throw new OfficeLoadError({
      code: "LIMIT_EXCEEDED",
      message: `DOCX 模型节点数 ${actual} 超过允许值 ${allowed}。`,
      format: "docx",
      phase: "model",
      limit: "maxDocxNodes",
      actual,
      allowed,
    });
  }
  const actualPages = model.metadata.documentPageCount;
  const allowedPages = limits.maxDocxPages;
  if (actualPages !== undefined && allowedPages !== undefined && actualPages > allowedPages) {
    throw new OfficeLoadError({
      code: "LIMIT_EXCEEDED",
      message: `DOCX 页数 ${actualPages} 超过允许值 ${allowedPages}。`,
      format: "docx",
      phase: "layout",
      limit: "maxDocxPages",
      actual: actualPages,
      allowed: allowedPages,
    });
  }
}

export function assertDocxParseTime(
  actual: number,
  limits: Readonly<DocxRuntimeLimits>,
): void {
  const allowed = limits.maxParseMs;
  if (allowed === undefined || actual <= allowed) return;
  throw new OfficeLoadError({
    code: "TIMEOUT",
    message: `DOCX 解析耗时超过 ${allowed} 毫秒。`,
    format: "docx",
    phase: "parse",
    limit: "maxParseMs",
    actual,
    allowed,
  });
}

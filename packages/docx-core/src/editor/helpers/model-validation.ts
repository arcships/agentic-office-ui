import type { DocModel } from "../../engine/types";

/** A clear public error for incomplete editor starter models. */
export class DocxModelValidationError extends Error {
  readonly code = "INVALID_DOC_MODEL";
  readonly missingFields: readonly string[];

  constructor(missingFields: readonly string[]) {
    super(`DOCX editor model is missing required fields: ${missingFields.join(", ")}`);
    this.name = "DocxModelValidationError";
    this.missingFields = missingFields;
  }
}

/**
 * Reject incomplete JavaScript input before cloneDocModel reaches a low-level
 * spread or map operation. TypeScript callers are still checked statically.
 */
export function assertValidDocxModel(model: unknown): asserts model is DocModel {
  const candidate = model as {
    nodes?: unknown;
    metadata?: {
      sourceParts?: unknown;
      warnings?: unknown;
      headerSections?: unknown;
      footerSections?: unknown;
      paragraphStyles?: unknown;
    };
  } | null;
  const metadata = candidate?.metadata;
  const missingFields: string[] = [];

  if (!Array.isArray(candidate?.nodes)) missingFields.push("nodes");
  if (!metadata || typeof metadata !== "object") {
    missingFields.push("metadata");
  } else {
    if (!Number.isFinite(metadata.sourceParts)) missingFields.push("metadata.sourceParts");
    if (!Array.isArray(metadata.warnings)) missingFields.push("metadata.warnings");
    if (!Array.isArray(metadata.headerSections)) missingFields.push("metadata.headerSections");
    if (!Array.isArray(metadata.footerSections)) missingFields.push("metadata.footerSections");
    if (!Array.isArray(metadata.paragraphStyles)) missingFields.push("metadata.paragraphStyles");
  }

  if (missingFields.length > 0) {
    throw new DocxModelValidationError(missingFields);
  }
}

// Stage 2–4 object index for PDF.
//
// Native links, form widgets and annotations are always available; visual
// objects appear only when a host registers a recognizer, and stage 4
// behaviors are references without geometry.

import type {
  OfficeObjectDescriptor,
  OfficeObjectKind,
  OfficeObjectProvider,
  OfficeObjectRegistry,
} from "@arcships/office-interaction";
import { createOfficeObjectRegistry } from "@arcships/office-interaction";
import type { PdfRenderAnnotation } from "./pdf-render-runtime";
import type {
  PdfBehaviorRecord,
  PdfVisualObjectDetection,
} from "./pdf-native-objects";
import { PDF_BEHAVIOR_REFERENCE_KIND } from "./pdf-native-objects";
import type { PdfReferenceContext } from "./pdf-reference-adapter";
import {
  describePdfBehavior,
  describePdfNativeObject,
  describePdfVisualObject,
} from "./pdf-reference-adapter";

export const PDF_STAGE2_KINDS: readonly OfficeObjectKind[] = Object.freeze([
  "link",
  "form-control",
  "annotation",
]);

export const PDF_VISUAL_KINDS: readonly OfficeObjectKind[] = Object.freeze([
  "text-block",
  "image",
  "table",
  "table-row",
  "table-cell",
]);

export const PDF_STAGE4_KINDS: readonly OfficeObjectKind[] = Object.freeze([
  "action",
  "form-rule",
  "annotation",
  "attachment",
  "layer",
  "signature",
  "permission",
]);

export type PdfVisualObjectRecord = PdfVisualObjectDetection & {
  pageIndex: number;
  /** Recognizer that produced the detection; namespaces the visual locator. */
  providerId?: string;
};

export const PDF_NATIVE_OBJECT_PROVIDER: OfficeObjectProvider<PdfReferenceContext> = {
  id: "pdf.native-objects",
  format: "pdf",
  kinds: PDF_STAGE2_KINDS,
  list(context) {
    return (context.annotations ?? []).flatMap((annotation: PdfRenderAnnotation) => {
      try {
        return [describePdfNativeObject(context, annotation)];
      } catch {
        return [];
      }
    });
  },
};

export const PDF_VISUAL_OBJECT_PROVIDER: OfficeObjectProvider<PdfReferenceContext> = {
  id: "pdf.visual-objects",
  format: "pdf",
  kinds: PDF_VISUAL_KINDS,
  list(context) {
    return (context.visualObjects ?? []).flatMap((record: PdfVisualObjectRecord) => {
      try {
        return [describePdfVisualObject(context, record.providerId ?? "visual", record.pageIndex, record)];
      } catch {
        return [];
      }
    });
  },
};

export const PDF_BEHAVIOR_OBJECT_PROVIDER: OfficeObjectProvider<PdfReferenceContext> = {
  id: "pdf.behaviors",
  format: "pdf",
  kinds: PDF_STAGE4_KINDS,
  list(context) {
    return (context.behaviors ?? []).flatMap((record: PdfBehaviorRecord) => {
      try {
        return [describePdfBehavior(context, record)];
      } catch {
        return [];
      }
    });
  },
};

/** Registry with the built-in PDF object families registered. */
export function createPdfObjectRegistry(): OfficeObjectRegistry<PdfReferenceContext> {
  return createOfficeObjectRegistry<PdfReferenceContext>([
    PDF_NATIVE_OBJECT_PROVIDER,
    PDF_VISUAL_OBJECT_PROVIDER,
    PDF_BEHAVIOR_OBJECT_PROVIDER,
  ]);
}

/**
 * Build a visual-object descriptor for a detection that carries no provider id.
 * The returned descriptor is what a host shows in an object list or overlay.
 */
export function pdfVisualObjectDescriptor(
  context: PdfReferenceContext,
  providerId: string,
  record: PdfVisualObjectRecord,
): OfficeObjectDescriptor {
  return describePdfVisualObject(context, providerId, record.pageIndex, record);
}

/** Stage 4 behavior reference kind for a record, exposed for host panels. */
export function pdfBehaviorReferenceKind(record: PdfBehaviorRecord): OfficeObjectKind {
  return PDF_BEHAVIOR_REFERENCE_KIND[record.behavior];
}

/** Convenience descriptor for a native PDF object. */
export function pdfNativeObjectDescriptor(
  context: PdfReferenceContext,
  annotation: PdfRenderAnnotation,
): OfficeObjectDescriptor {
  return describePdfNativeObject(context, annotation);
}

/** Wrap a detection with its page so the index can address it. */
export function withPdfVisualPage(
  pageIndex: number,
  detection: PdfVisualObjectDetection,
): PdfVisualObjectRecord {
  return { ...detection, pageIndex };
}

/** A stage 4 behavior record with no geometry, for host timeline/panel lists. */
export function pdfBehaviorObjectDescriptor(
  context: PdfReferenceContext,
  record: PdfBehaviorRecord,
): OfficeObjectDescriptor {
  return describePdfBehavior(context, record);
}

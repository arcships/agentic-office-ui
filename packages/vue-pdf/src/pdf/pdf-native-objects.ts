// PDF native objects, pluggable visual recognizers and stage 4 behaviors.
//
// Native links, form widgets and annotations come from the PDF runtime. Visual
// objects (text blocks, images, tables) come from an optional provider that the
// host plugs in; the library never ships OCR. Behaviors are references with a
// native id and a page scope, never a canvas outline.

import type {
  JsonValue,
  NormalizedRect,
  OfficeDocumentRevision,
  OfficeObjectKind,
  OfficeObjectReliability,
  OfficeMessages,
  PdfObjectLocator,
} from "@arcships/office-interaction"
import { formatOfficeMessage, officeMessagesEn } from "@arcships/office-interaction"
import type { PdfRenderAnnotation } from "./pdf-render-runtime"

/** Provider result for one visual object on one page. */
export interface PdfVisualObjectDetection {
  objectId: string
  kind: "text-block" | "image" | "table" | "table-row" | "table-cell"
  region: NormalizedRect
  label?: string
  parentObjectId?: string
  content?: { text?: string }
  reliability: OfficeObjectReliability
}

/**
 * Optional visual recognizer. The Surface never bundles one; a host passes an
 * implementation and the PDF adapter turns its detections into visual locators.
 */
export interface PdfVisualObjectProvider {
  readonly id: string
  recognize(input: {
    revision: OfficeDocumentRevision
    pageIndex: number
    image: ImageBitmap
    signal: AbortSignal
  }): Promise<readonly PdfVisualObjectDetection[]>
}

export type PdfBehaviorKind =
  | "link-action"
  | "form-rule"
  | "annotation-state"
  | "attachment"
  | "layer"
  | "signature"
  | "permission"

export const PDF_BEHAVIOR_REFERENCE_KIND: Record<PdfBehaviorKind, OfficeObjectKind> = {
  "link-action": "action",
  "form-rule": "form-rule",
  "annotation-state": "annotation",
  attachment: "attachment",
  layer: "layer",
  signature: "signature",
  permission: "permission",
}

/** Localized labels for stage 4 PDF behavior records. */
export interface PdfBehaviorLabels {
  document: string
  page: (pageIndex: number) => string
  action: string
  /** Template `{target}` already applied by the builder. */
  linkAction: (target: string) => string
  formField: (name: string) => string
  annotation: (text: string) => string
  annotationById: (id: string) => string
  attachment: (name: string) => string
  layer: (name: string) => string
  signature: (name: string) => string
  permission: string
}

/** Build behavior labels from the shared catalog (English by default). */
export function createPdfBehaviorLabels(
  common: OfficeMessages["common"] = officeMessagesEn.common,
): PdfBehaviorLabels {
  const descriptors = common.descriptors
  return {
    document: descriptors.pdf,
    page: (pageIndex) => formatOfficeMessage(descriptors.page, { number: pageIndex + 1 }),
    action: descriptors.action,
    linkAction: (target) => formatOfficeMessage(descriptors.actionTarget, { target }),
    formField: (name) => formatOfficeMessage(descriptors.formFieldName, { name }),
    annotation: (text) => formatOfficeMessage(descriptors.annotationText, { text }),
    annotationById: (id) => formatOfficeMessage(descriptors.annotationById, { id }),
    attachment: (name) => formatOfficeMessage(descriptors.attachment, { name }),
    layer: (name) => formatOfficeMessage(descriptors.layer, { name }),
    signature: (name) => formatOfficeMessage(descriptors.signature, { name }),
    permission: descriptors.permission,
  }
}

export interface PdfBehaviorRecord {
  behavior: PdfBehaviorKind
  /** Deterministic id inside the page (or document) scope. */
  instanceId: string
  pageIndex?: number
  label: string
  path: readonly { kind: OfficeObjectKind; label: string }[]
  content?: { text?: string }
  facets: Readonly<Record<string, JsonValue>>
}

export interface PdfAttachmentSource {
  id: string
  name: string
  description?: string
}

export interface PdfLayerSource {
  id: string
  name: string
  visible?: boolean
}

export interface PdfSignatureSource {
  id: string
  name?: string
  signer?: string
  valid?: boolean
}

export interface PdfPermissionsSource {
  printing?: boolean
  copying?: boolean
  modifying?: boolean
  annotating?: boolean
  formFilling?: boolean
  documentAssembly?: boolean
}

export interface PdfBehaviorSource {
  annotations?: readonly PdfRenderAnnotation[]
  attachments?: readonly PdfAttachmentSource[]
  layers?: readonly PdfLayerSource[]
  signatures?: readonly PdfSignatureSource[]
  permissions?: PdfPermissionsSource | null
}

export function pdfNativeObjectType(
  annotation: Pick<PdfRenderAnnotation, "kind">,
): "link" | "form" | "annotation" {
  if (annotation.kind === "link") return "link"
  if (annotation.kind === "widget") return "form"
  return "annotation"
}

function actionLabel(annotation: PdfRenderAnnotation, labels: PdfBehaviorLabels): string {
  if (annotation.url) return labels.linkAction(annotation.url)
  if (annotation.targetPageIndex !== undefined) {
    return labels.linkAction(formatOfficeMessage(officeMessagesEn.common.descriptors.page, {
      number: annotation.targetPageIndex + 1,
    }))
  }
  return labels.action
}

function fieldLabel(annotation: PdfRenderAnnotation, labels: PdfBehaviorLabels): string {
  const name = annotation.fieldName?.trim() || annotation.id
  return labels.formField(name)
}

function annotationLabel(annotation: PdfRenderAnnotation, labels: PdfBehaviorLabels): string {
  const text = annotation.contents?.replace(/\s+/gu, " ").trim()
  return text ? labels.annotation(text.slice(0, 80)) : labels.annotationById(annotation.id)
}

/**
 * Enumerate stage 4 behaviors from the PDF runtime data plus optional
 * document-level sources (attachments, layers, signatures, permissions).
 */
export function createPdfBehaviorRecords(
  source: PdfBehaviorSource,
  labels: PdfBehaviorLabels = createPdfBehaviorLabels(),
): readonly PdfBehaviorRecord[] {
  const records: PdfBehaviorRecord[] = []
  for (const annotation of source.annotations ?? []) {
    if (annotation.kind === "link") {
      records.push({
        behavior: "link-action",
        instanceId: `pdf.link-action:${annotation.pageIndex}:${annotation.id}`,
        pageIndex: annotation.pageIndex,
        label: actionLabel(annotation, labels),
        path: [
          { kind: "document", label: labels.document },
          { kind: "page", label: labels.page(annotation.pageIndex) },
          { kind: "action", label: actionLabel(annotation, labels) },
        ],
        facets: {
          annotationId: annotation.id,
          ...(annotation.url ? { url: annotation.url } : {}),
          ...(annotation.targetPageIndex === undefined ? {} : { targetPageIndex: annotation.targetPageIndex }),
          ...(annotation.blocked === undefined ? {} : { blocked: annotation.blocked }),
        },
      })
      continue
    }
    if (annotation.kind === "widget") {
      records.push({
        behavior: "form-rule",
        instanceId: `pdf.form-rule:${annotation.pageIndex}:${annotation.id}`,
        pageIndex: annotation.pageIndex,
        label: fieldLabel(annotation, labels),
        path: [
          { kind: "document", label: labels.document },
          { kind: "page", label: labels.page(annotation.pageIndex) },
          { kind: "form-rule", label: fieldLabel(annotation, labels) },
        ],
        facets: {
          annotationId: annotation.id,
          ...(annotation.fieldName ? { fieldName: annotation.fieldName } : {}),
          ...(annotation.fieldType ? { fieldType: annotation.fieldType } : {}),
          readOnly: annotation.readOnly === true,
          ...(annotation.options
            ? { options: annotation.options.map((option) => option.label) }
            : {}),
        },
      })
      continue
    }
    records.push({
      behavior: "annotation-state",
      instanceId: `pdf.annotation-state:${annotation.pageIndex}:${annotation.id}`,
      pageIndex: annotation.pageIndex,
      label: annotationLabel(annotation, labels),
      path: [
        { kind: "document", label: labels.document },
        { kind: "page", label: labels.page(annotation.pageIndex) },
        { kind: "annotation", label: annotationLabel(annotation, labels) },
      ],
      ...(annotation.contents ? { content: { text: annotation.contents } } : {}),
      facets: {
        annotationId: annotation.id,
        ...(annotation.flags ? { flags: [...annotation.flags] } : {}),
        ...(annotation.color ? { color: annotation.color } : {}),
        ...(annotation.opacity === undefined ? {} : { opacity: annotation.opacity }),
      },
    })
  }
  for (const attachment of source.attachments ?? []) {
    records.push({
      behavior: "attachment",
      instanceId: `pdf.attachment:${attachment.id}`,
      label: labels.attachment(attachment.name),
      path: [
        { kind: "document", label: labels.document },
        { kind: "attachment", label: labels.attachment(attachment.name) },
      ],
      ...(attachment.description ? { content: { text: attachment.description } } : {}),
      facets: { attachmentId: attachment.id, name: attachment.name },
    })
  }
  for (const layer of source.layers ?? []) {
    records.push({
      behavior: "layer",
      instanceId: `pdf.layer:${layer.id}`,
      label: labels.layer(layer.name),
      path: [
        { kind: "document", label: labels.document },
        { kind: "layer", label: labels.layer(layer.name) },
      ],
      facets: {
        layerId: layer.id,
        name: layer.name,
        ...(layer.visible === undefined ? {} : { visible: layer.visible }),
      },
    })
  }
  for (const signature of source.signatures ?? []) {
    const name = signature.name?.trim() || signature.id
    records.push({
      behavior: "signature",
      instanceId: `pdf.signature:${signature.id}`,
      label: labels.signature(name),
      path: [
        { kind: "document", label: labels.document },
        { kind: "signature", label: labels.signature(name) },
      ],
      facets: {
        signatureId: signature.id,
        ...(signature.signer ? { signer: signature.signer } : {}),
        ...(signature.valid === undefined ? {} : { valid: signature.valid }),
      },
    })
  }
  if (source.permissions) {
    records.push({
      behavior: "permission",
      instanceId: "pdf.permission:document",
      label: labels.permission,
      path: [
        { kind: "document", label: labels.document },
        { kind: "permission", label: labels.permission },
      ],
      facets: { ...source.permissions },
    })
  }
  return Object.freeze(records)
}

/** Locator value for a stage 4 PDF behavior record. */
export function pdfBehaviorLocatorValue(record: PdfBehaviorRecord): PdfObjectLocator {
  return {
    kind: "behavior",
    ...(record.pageIndex === undefined ? {} : { pageIndex: record.pageIndex }),
    behavior: record.behavior,
    nativeId: record.instanceId,
  }
}

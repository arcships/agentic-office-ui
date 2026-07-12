import type { PptxDocumentSession, PptxPreviewSession } from "@arcships/pptx-core/browser"
import type { ShallowRef } from "vue"
import type { UsePptxDocumentReturn } from "../headless-types"

export interface PptxDocumentBinding {
  session: ShallowRef<PptxPreviewSession | null>
  setActiveIndex(index: number): void
}

const bindings = new WeakMap<UsePptxDocumentReturn, PptxDocumentBinding>()

export function registerDocumentBinding(
  document: UsePptxDocumentReturn,
  binding: PptxDocumentBinding,
): void {
  bindings.set(document, binding)
}

export function getDocumentBinding(document: UsePptxDocumentReturn): PptxDocumentBinding | undefined {
  return bindings.get(document)
}

export function isDocumentSession(
  session: PptxPreviewSession | null,
): session is PptxDocumentSession {
  return Boolean(session && "createPlaybackController" in session)
}

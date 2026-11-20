import type { PdfSource, PdfUrlPolicy } from "./pdf-url-policy"
import type { PdfPasswordHandler, PdfRenderRuntime } from "./pdf/pdf-render-runtime"
import type { PdfPageLayout } from "./pdf/pdf-page-layout"
import type { OfficeLocale, OfficeMessagesOverride } from "@arcships/office-interaction"

export interface PdfViewerProps {
  src?: string
  /** Preferred source input. When both source and src are set, source wins. */
  source?: PdfSource
  /** Explicit URL policy required for URL sources. Blob/File sources do not fetch. */
  urlPolicy?: PdfUrlPolicy
  /** Optional caller-owned runtime. When omitted, this viewer owns a dedicated PDF Worker. */
  runtime?: PdfRenderRuntime
  /** Overrides the bundled PDFium WASM URL for the viewer-owned runtime. */
  pdfiumWasmUrl?: string
  /** Maximum accepted PDF source size in bytes. Defaults to 50 MiB. */
  maxFileSize?: number
  /** Password for encrypted PDF documents. */
  password?: string
  /** Invoked when an encrypted PDF needs a password. */
  onPassword?: PdfPasswordHandler
  fileName?: string
  defaultZoom?: number
  /**
   * Page layout: `continuous` scrolling, `single` page at a time, or two-page
   * `spread`s. Defaults to `single`. Switching modes keeps the active page.
   */
  pageLayout?: PdfPageLayout | null
  /** Page layout used while `pageLayout` is uncontrolled. Defaults to single. */
  defaultPageLayout?: PdfPageLayout
  className?: string
  showToolbar?: boolean
  showDownload?: boolean
  showRotateControls?: boolean
  /** Render the print control in the toolbar. Defaults to true. */
  showPrint?: boolean
  /**
   * Render a visible `OfficeStatusBar` footer with the page, file and state.
   * Defaults to false; the visually hidden live region stays mounted either way.
   */
  showStatusBar?: boolean
  /** Optional minimum height; the viewer itself does not impose one. */
  minHeight?: string | number
  /** Dark viewer chrome and page surround. Defaults to false. */
  isDark?: boolean
  locale?: OfficeLocale
  messages?: OfficeMessagesOverride
}

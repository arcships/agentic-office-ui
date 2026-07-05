import type { UseXlsxViewerControllerOptions, XlsxViewerController } from "./composables"

export interface XlsxViewerProps {
  file?: ArrayBuffer
  src?: string
  controller?: XlsxViewerController
  isDark?: boolean
  className?: string
  readOnly?: boolean
  showToolbar?: boolean
  showSheetTabs?: boolean
}

export interface XlsxViewerProviderProps extends UseXlsxViewerControllerOptions {
  controller?: XlsxViewerController
  isDark?: boolean
}

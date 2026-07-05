export interface PdfViewerProps {
  src?: string
  fileName?: string
  defaultZoom?: number
  className?: string
  showToolbar?: boolean
  showDownload?: boolean
  showRotateControls?: boolean
}

export interface SignaturePadProps {
  width?: number
  height?: number
  penColor?: string
  backgroundColor?: string
  className?: string
}

export interface FileUploadProps {
  accept?: string
  multiple?: boolean
  maxSize?: number
  disabled?: boolean
  className?: string
}

export interface FileThumbnailFile {
  name: string
  type: string
  url?: string
}

export interface FileThumbnailProps {
  file: FileThumbnailFile
  size?: "sm" | "md" | "lg"
  className?: string
}

export interface BoundingBoxField {
  id: string
  label: string
  page: number
  rect: [number, number, number, number]
  value?: string
  confidence?: number
}

export interface BoundingBoxCitationsProps {
  file: string
  fields: BoundingBoxField[]
  className?: string
}

export interface LayoutBlock {
  id: string
  bbox: [number, number, number, number]
  kind: string
  text?: string
  confidence?: number
  parentId?: string
}

export interface ParsedOcrOutput {
  width: number
  height: number
  blocks: LayoutBlock[]
}

export interface LayoutBlocksProps {
  file: string
  output: ParsedOcrOutput
  className?: string
}

export interface SpinnerProps {
  size?: "sm" | "md" | "lg"
  className?: string
}

export interface TooltipProps {
  content?: string
  side?: "top" | "bottom" | "left" | "right"
  align?: "start" | "center" | "end"
  delayMs?: number
  className?: string
}

/**
 * @extend-ai/vue-extend
 *
 * Vue 3 UI components for the Extend platform.
 * Uses Tailwind CSS with design-token classes.
 * Does not depend on shadcn-vue.
 */

// PDF viewer
export { default as PdfViewer } from "./components/PdfViewer.vue"

// Signature pad
export { default as SignaturePad } from "./components/SignaturePad.vue"

// File upload
export { default as FileUpload } from "./components/FileUpload.vue"

// File thumbnail
export { default as FileThumbnail } from "./components/FileThumbnail.vue"

// Bounding box citations
export { default as BoundingBoxCitations } from "./components/BoundingBoxCitations.vue"

// Layout blocks
export { default as LayoutBlocks } from "./components/LayoutBlocks.vue"

// Base UI
export { default as Spinner } from "./components/Spinner.vue"
export { default as Tooltip } from "./components/Tooltip.vue"

export type * from "./types"

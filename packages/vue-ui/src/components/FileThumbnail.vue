// Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn.
<script setup lang="ts">
import {
  computed,
  onBeforeUnmount,
  onMounted,
  ref,
  useSlots,
  watch,
  type Component,
} from "vue"
import {
  File,
  FileArchive,
  FileAudio,
  FileImage,
  FileSpreadsheet,
  FileText,
  FileVideo,
} from "@lucide/vue"
import type { FileThumbnailProps } from "../types"
import { useOfficeUiMessages } from "../i18n"

const props = withDefaults(defineProps<FileThumbnailProps>(), {
  size: "md",
  className: "",
  previewImageUrl: undefined,
  previewAspectRatio: undefined,
  previewClassName: "",
  isLoading: false,
  hasError: false,
  showTypeGlyph: false,
  showMeta: false,
  fill: false,
  previewFill: false,
  alt: undefined,
})

const slots = useSlots()
const m = useOfficeUiMessages()

// Preview URLs that already completed their reveal. View switches remount
// thumbnails; only the first load should animate the blur-in.
const revealedPreviewImageUrls = new Set<string>()

const namedSizes = { sm: "32px", md: "48px", lg: "64px" } as const

const imageRef = ref<HTMLImageElement | null>(null)
const loadedPreviewImageUrl = ref<string | null>(null)
const failedPreviewImageUrl = ref<string | null>(null)
let revealFrame = 0
let revealFrame2 = 0

const sizeValue = computed(() => {
  if (typeof props.size === "number") return `${props.size}px`
  if (props.size in namedSizes) return namedSizes[props.size as keyof typeof namedSizes]
  return props.size
})

const rootClasses = computed(() => [
  "file-thumbnail",
  "office-ui",
  props.fill ? "file-thumbnail--fill" : "",
  props.className,
])

const rootStyle = computed(() => (props.fill ? undefined : { width: sizeValue.value }))

const previewStyle = computed(() =>
  props.previewAspectRatio ? { aspectRatio: String(props.previewAspectRatio) } : undefined,
)

const previewImageUrl = computed(() => {
  if (props.previewImageUrl !== undefined) return props.previewImageUrl
  return props.file.type.startsWith("image/") ? props.file.url ?? null : null
})

const failedImage = computed(
  () => Boolean(previewImageUrl.value) && failedPreviewImageUrl.value === previewImageUrl.value,
)

const imageLoading = computed(() => {
  const url = previewImageUrl.value
  if (!url || failedImage.value) return false
  return loadedPreviewImageUrl.value !== url && !revealedPreviewImageUrls.has(url)
})

const showLoading = computed(() => props.isLoading || imageLoading.value)

const hasPreviewContent = computed(() => Boolean(slots.preview))

const showFallback = computed(
  () =>
    !showLoading.value &&
    (props.hasError ||
      failedImage.value ||
      (!previewImageUrl.value && !hasPreviewContent.value)),
)

const extension = computed(() => {
  const parts = props.file.name.split(".")
  return parts.length > 1 ? parts[parts.length - 1]!.toLowerCase() : ""
})

const typeGlyph = computed<Component>(() => {
  const type = props.file.type.toLowerCase()
  if (type.startsWith("image/")) return FileImage
  if (type === "application/pdf") return FileText
  if (type.includes("sheet") || type.includes("excel") || type.includes("csv")) {
    return FileSpreadsheet
  }
  if (type.includes("word") || type.includes("document")) return FileText
  if (type.startsWith("video/")) return FileVideo
  if (type.startsWith("audio/")) return FileAudio
  if (type.includes("zip") || type.includes("rar") || type.includes("archive")) {
    return FileArchive
  }

  const byExtension: Record<string, Component> = {
    pdf: FileText,
    doc: FileText,
    docx: FileText,
    txt: FileText,
    rtf: FileText,
    xls: FileSpreadsheet,
    xlsx: FileSpreadsheet,
    csv: FileSpreadsheet,
    png: FileImage,
    jpg: FileImage,
    jpeg: FileImage,
    gif: FileImage,
    webp: FileImage,
    svg: FileImage,
    zip: FileArchive,
    rar: FileArchive,
    "7z": FileArchive,
    mp4: FileVideo,
    mov: FileVideo,
    webm: FileVideo,
    mp3: FileAudio,
    wav: FileAudio,
    m4a: FileAudio,
  }
  return byExtension[extension.value] ?? File
})

const imageAlt = computed(() => props.alt ?? props.file.name)

function requestFrame(callback: () => void): number {
  if (typeof requestAnimationFrame === "function") return requestAnimationFrame(callback)
  if (typeof window !== "undefined") return window.setTimeout(callback, 0)
  return setTimeout(callback, 0) as unknown as number
}

function cancelFrame(handle: number): void {
  if (!handle) return
  if (typeof cancelAnimationFrame === "function") cancelAnimationFrame(handle)
  else if (typeof window !== "undefined") window.clearTimeout(handle)
  else clearTimeout(handle)
}

function cancelReveal(): void {
  cancelFrame(revealFrame)
  cancelFrame(revealFrame2)
  revealFrame = 0
  revealFrame2 = 0
}

function markImageLoaded(image: HTMLImageElement | null, url: string | null): void {
  if (!url) return
  const didLoad = Boolean(image && image.naturalWidth > 0 && image.naturalHeight > 0)
  if (!didLoad) {
    failedPreviewImageUrl.value = url
    return
  }

  failedPreviewImageUrl.value = null
  revealedPreviewImageUrls.add(url)
  cancelReveal()
  // Two frames mirror the reference: the browser paints the decoded image
  // before the overlay is removed, so the reveal never flashes an empty tile.
  revealFrame = requestFrame(() => {
    revealFrame2 = requestFrame(() => {
      loadedPreviewImageUrl.value = url
      revealFrame = 0
      revealFrame2 = 0
    })
  })
}

function onImageLoad(event: Event): void {
  markImageLoaded(event.currentTarget as HTMLImageElement, previewImageUrl.value)
}

function onImageError(): void {
  const url = previewImageUrl.value
  if (!url) return
  revealedPreviewImageUrls.delete(url)
  cancelReveal()
  failedPreviewImageUrl.value = url
  if (loadedPreviewImageUrl.value === url) loadedPreviewImageUrl.value = null
}

watch(
  previewImageUrl,
  () => {
    cancelReveal()
    loadedPreviewImageUrl.value = null
    if (failedPreviewImageUrl.value !== previewImageUrl.value) {
      failedPreviewImageUrl.value = null
    }
  },
  { flush: "post" },
)

onMounted(() => {
  const image = imageRef.value
  if (image && previewImageUrl.value && image.complete) {
    markImageLoaded(image, previewImageUrl.value)
  }
})

onBeforeUnmount(cancelReveal)
</script>

<template>
  <div :class="rootClasses" :style="rootStyle" data-slot="file-thumbnail">
    <div
      :class="[
        'file-thumbnail__preview',
        props.previewFill ? 'file-thumbnail__preview--fill' : '',
        props.previewClassName,
      ]"
      :style="previewStyle"
      data-slot="file-thumbnail-preview"
    >
      <img
        v-if="previewImageUrl && !failedImage"
        ref="imageRef"
        :src="previewImageUrl"
        :alt="imageAlt"
        draggable="false"
        loading="lazy"
        decoding="async"
        :class="['file-thumbnail__image', { 'is-loading': showLoading }]"
        @load="onImageLoad"
        @error="onImageError"
      />
      <div
        v-if="hasPreviewContent"
        :class="['file-thumbnail__content', { 'is-loading': showLoading }]"
        data-slot="file-thumbnail-content"
      >
        <slot name="preview" />
      </div>
      <div v-if="showLoading" class="file-thumbnail__loading" aria-hidden="true">
        <div class="file-thumbnail__loading-base" />
        <div class="file-thumbnail__loading-pulse" />
      </div>
      <div
        v-if="showFallback"
        class="file-thumbnail__fallback"
        data-slot="file-thumbnail-fallback"
        aria-hidden="true"
      >
        <component
          :is="typeGlyph"
          v-if="props.showTypeGlyph"
          class="file-thumbnail__glyph"
          :size="24"
        />
      </div>
    </div>
    <div v-if="props.showMeta" class="file-thumbnail__meta" data-slot="file-thumbnail-meta">
      <p class="file-thumbnail__name">{{ props.file.name }}</p>
      <p class="file-thumbnail__type">
        {{ props.file.type || m.vueUi.fileThumbnail.unknownType }}
      </p>
    </div>
  </div>
</template>

<style scoped>
.file-thumbnail {
  flex: 0 0 auto;
  box-sizing: border-box;
  width: 3rem;
  overflow: hidden;
  border: 1px solid var(--office-border);
  border-radius: var(--office-radius-lg);
  background: var(--office-background);
  color: var(--office-foreground);
}
.file-thumbnail,
.file-thumbnail * {
  box-sizing: border-box;
}
.file-thumbnail__preview {
  position: relative;
  width: 100%;
  aspect-ratio: 1 / 1;
  overflow: hidden;
  background: var(--office-muted);
  contain: layout paint;
}
.file-thumbnail--fill {
  width: 100%;
  height: 100%;
}
.file-thumbnail__preview--fill {
  width: auto;
  height: 100%;
}
.file-thumbnail__image {
  position: absolute;
  inset: 0;
  display: block;
  width: 100%;
  height: 100%;
  max-width: 100%;
  object-fit: cover;
  opacity: 1;
  filter: none;
  transition:
    opacity 160ms cubic-bezier(0.22, 1, 0.36, 1),
    filter 160ms cubic-bezier(0.22, 1, 0.36, 1);
}
.file-thumbnail__image.is-loading {
  opacity: 0;
  filter: blur(4px);
}
.file-thumbnail__content {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  opacity: 1;
  filter: none;
  transition:
    opacity 160ms cubic-bezier(0.22, 1, 0.36, 1),
    filter 160ms cubic-bezier(0.22, 1, 0.36, 1);
}
.file-thumbnail__content.is-loading {
  opacity: 0;
  filter: blur(4px);
}
.file-thumbnail__loading {
  position: absolute;
  inset: 0;
  z-index: 10;
  overflow: hidden;
  background: var(--office-muted);
}
.file-thumbnail__loading-base {
  position: absolute;
  inset: 0;
  background: var(--office-muted);
}
.file-thumbnail__loading-pulse {
  position: absolute;
  inset: 0;
  background: color-mix(in oklab, var(--office-background) 55%, transparent);
  animation: office-file-thumbnail-pulse 2s cubic-bezier(0.4, 0, 0.6, 1) infinite;
}
.file-thumbnail__fallback {
  position: absolute;
  inset: 0;
  display: grid;
  place-items: center;
  background: var(--office-muted);
}
.file-thumbnail__glyph {
  width: 40%;
  height: 40%;
  color: var(--office-muted-foreground);
}
.file-thumbnail__meta {
  min-width: 0;
  padding: var(--office-ui-space-2);
  border-top: 1px solid var(--office-border);
}
.file-thumbnail__name,
.file-thumbnail__type {
  margin: 0;
  overflow: hidden;
  white-space: nowrap;
  text-overflow: ellipsis;
}
.file-thumbnail__name {
  font-size: var(--office-text-sm);
  line-height: var(--office-leading-sm);
  font-weight: 500;
}
.file-thumbnail__type {
  margin-top: var(--office-ui-space-0-5);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  line-height: var(--office-leading-xs);
}
@keyframes office-file-thumbnail-pulse {
  50% {
    opacity: 0.5;
  }
}
@media (prefers-reduced-motion: reduce) {
  .file-thumbnail__image,
  .file-thumbnail__content {
    transition: none;
  }
  .file-thumbnail__loading-pulse {
    animation: none;
  }
}
</style>

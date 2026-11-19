<template>
  <Teleport v-for="box of boxes" :key="box.key" :to="box.surface">
    <div
      :data-docx-block-highlight="box.blockId"
      :data-docx-word-highlight="box.wordId"
      :data-docx-block-type="box.type"
      :data-docx-highlight-active="box.active || undefined"
      aria-hidden="true"
      :style="box.style"
    />
  </Teleport>
</template>

<script setup lang="ts">
import { computed } from "vue"
import type { DocumentBlock } from "@arcships/docx-core"
import type { DocxBlockHighlightsProps } from "./block-highlight-types"
import { useDocxBlockGeometry } from "../composables/useDocxBlockGeometry"

/**
 * Absolute highlight boxes for document blocks (or words) on mounted pages.
 *
 * Geometry is measured in page CSS pixels and the box is teleported into the
 * page surface, so the surface's zoom transform applies to the overlay without
 * any extra math. `pageIndex` limits the overlay to one page and `blockIds`
 * selects blocks from `blocks` (or treats the ids as source ids when no block
 * list is supplied).
 */
const props = withDefaults(
  defineProps<DocxBlockHighlightsProps>(),
  {
    root: null,
    block: undefined,
    blocks: undefined,
    blockIds: undefined,
    pageIndex: undefined,
    color: "#2563eb",
    colorForBlock: undefined,
    onGeometry: undefined,
  },
)

function syntheticBlock(id: string): DocumentBlock {
  return {
    object: "block",
    id,
    type: "text",
    content: "",
    details: { kind: "node" },
    sourceRefs: [
      {
        documentId: "",
        nodeId: id,
        partName: "",
        story: "body",
        offsetEncoding: "utf-16",
      },
    ],
    children: [],
  }
}

const targets = computed<readonly DocumentBlock[]>(() => {
  const unique = new Map<string, DocumentBlock>()
  const ids = props.blockIds
  if (ids && ids.length > 0) {
    if (props.blocks) {
      const byId = new Map(props.blocks.map((item) => [item.id, item]))
      for (const id of ids) {
        const block = byId.get(id)
        if (block) unique.set(id, block)
        else unique.set(id, syntheticBlock(id))
      }
    } else {
      for (const id of ids) unique.set(id, syntheticBlock(id))
    }
  } else {
    for (const block of props.blocks ?? []) unique.set(block.id, block)
  }
  if (props.block) unique.set(props.block.id, props.block)
  return [...unique.values()]
})

const { geometries, surfaces, refresh } = useDocxBlockGeometry({
  root: () => props.root,
  blocks: () => targets.value,
  onGeometry: (value) => props.onGeometry?.(value),
})

defineExpose({ refresh, geometries })

interface HighlightBox {
  key: string
  surface: HTMLElement
  blockId: string
  wordId?: string
  type: string
  active: boolean
  style: Record<string, string | number>
}

const boxes = computed<HighlightBox[]>(() => {
  const result: HighlightBox[] = []
  for (const block of targets.value) {
    const geometry = geometries.value.get(block.id)
    if (!geometry) continue
    const tint = props.colorForBlock?.(block) ?? props.color
    const active = block.id === props.block?.id
    const word = block.details.kind === "word"
    geometry.fragments.forEach((fragment, index) => {
      if (props.pageIndex !== undefined && fragment.pageIndex !== props.pageIndex) return
      const surface = surfaces.value.get(fragment.pageIndex)
      if (!surface) return
      const box = fragment.boundingBox
      result.push({
        key: `${block.id}-${fragment.pageIndex}-${index}`,
        surface,
        blockId: block.id,
        wordId: word ? block.id : undefined,
        type: block.type,
        active,
        style: {
          position: "absolute",
          pointerEvents: "none",
          zIndex: active ? 101 : 100,
          left: `${box.left}px`,
          top: `${box.top}px`,
          width: `${box.right - box.left}px`,
          height: `${box.bottom - box.top}px`,
          boxSizing: "border-box",
          border: `${active && props.blocks ? 2 : 1}px solid ${tint}`,
          backgroundColor: `${tint}${active || !props.blocks ? "20" : "08"}`,
        },
      })
    })
  }
  return result
})

</script>

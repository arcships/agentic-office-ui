<!-- Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. -->
<script setup lang="ts">
import { computed } from "vue"
import {
  ScrollAreaCorner,
  ScrollAreaRoot,
  ScrollAreaScrollbar,
  ScrollAreaThumb,
  ScrollAreaViewport,
} from "reka-ui"
import { officeClass } from "./types"

export interface OfficeScrollAreaProps {
  /** Which scrollbars to render. */
  orientation?: "vertical" | "horizontal" | "both"
  /** Reka visibility strategy; matches the reference `auto` behaviour. */
  type?: "auto" | "always" | "scroll" | "hover" | "glimpse"
  scrollHideDelay?: number
  /** Extra classes merged on the root. */
  class?: string
  /** Extra classes merged on the viewport. */
  viewportClass?: string
}

const props = withDefaults(defineProps<OfficeScrollAreaProps>(), {
  orientation: "vertical",
  type: "auto",
  scrollHideDelay: 600,
  class: "",
  viewportClass: "",
})

const classes = computed(() =>
  officeClass("office-ui", "office-scroll-area", props.class),
)
</script>

<template>
  <ScrollAreaRoot
    :type="props.type"
    :scroll-hide-delay="props.scrollHideDelay"
    :class="classes"
    data-slot="scroll-area"
  >
    <ScrollAreaViewport
      :class="officeClass('office-scroll-area__viewport', props.viewportClass)"
      data-slot="scroll-area-viewport"
    >
      <slot />
    </ScrollAreaViewport>
    <ScrollAreaScrollbar
      v-if="props.orientation !== 'horizontal'"
      orientation="vertical"
      class="office-scroll-area__scrollbar"
      data-slot="scroll-area-scrollbar"
    >
      <ScrollAreaThumb class="office-scroll-area__thumb" data-slot="scroll-area-thumb" />
    </ScrollAreaScrollbar>
    <ScrollAreaScrollbar
      v-if="props.orientation !== 'vertical'"
      orientation="horizontal"
      class="office-scroll-area__scrollbar"
      data-slot="scroll-area-scrollbar"
    >
      <ScrollAreaThumb class="office-scroll-area__thumb" data-slot="scroll-area-thumb" />
    </ScrollAreaScrollbar>
    <ScrollAreaCorner
      v-if="props.orientation === 'both'"
      class="office-scroll-area__corner"
      data-slot="scroll-area-corner"
    />
  </ScrollAreaRoot>
</template>

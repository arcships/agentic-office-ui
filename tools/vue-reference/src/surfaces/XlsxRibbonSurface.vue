<script setup lang="ts">
/**
 * Issue #170 — XLSX ribbon format-row parity subject.
 *
 * Layout source: `tools/extend-reference/src/vendor/extend/xlsx-ribbon.tsx`
 * (Extend UI `EditorToolbar` format row, pinned 70fe70c6). This surface renders
 * the same resting format row — undo/redo, font family + size, B/I/U/S, text /
 * fill colour, horizontal + vertical alignment, wrap text, number format, cell
 * borders and merge — with the production primitives (`OfficeToolbar`,
 * `OfficeToolbarGroup`, `OfficeToolbarSeparator`, `OfficeButton`,
 * `OfficeIconButton`, `OfficeToggle`, `OfficeIcon`, `HugeIcon`) and the
 * `--office-*` tokens.
 *
 * The triggers are static: upstream renders the closed `SelectTrigger` and the
 * parity matrix never opens a select, so this surface paints the closed trigger
 * and keeps the keyboard/native listbox behaviour out of the harness (the
 * production ribbon's own native selects are covered by its component tests).
 *
 * Not covered here because the upstream format row has no such command:
 * clipboard (cut/copy/paste), structure (insert/delete rows and columns), the
 * authoring cluster (chart, comment, hyperlink, table, conditional formatting,
 * validation, freeze panes, name manager) and the read-only switch. Those live
 * in `packages/vue-xlsx/src/components/XlsxRibbon.vue` and are covered by
 * `tests/component/xlsx-*.test.mjs` and the `xlsx-chrome` capture.
 */
import {
  AlignBoxBottomCenterIcon,
  AlignBoxMiddleCenterIcon,
  AlignBoxTopCenterIcon,
  BorderAll01Icon,
  ChevronDownIcon,
  PaintBucketIcon,
  PathfinderMergeIcon,
  Redo02Icon,
  TextAlignCenterIcon,
  TextAlignLeft01Icon,
  TextAlignRight01Icon,
  TextBoldIcon,
  TextColorIcon,
  TextItalicIcon,
  TextNumberSignIcon,
  TextStrikethroughIcon,
  TextUnderlineIcon,
  TextWrapIcon,
  Undo02Icon,
} from "@hugeicons/core-free-icons";
import {
  HugeIcon,
  OfficeButton,
  OfficeIcon,
  OfficeIconButton,
  OfficeToggle,
  OfficeToolbar,
  OfficeToolbarGroup,
  OfficeToolbarSeparator,
} from "@arcships/vue-ui";
import { fixtures } from "../fixtures";
import ParityCell from "../parity/ParityCell.vue";

defineProps<{ cellWidth: number }>();

const ribbon = fixtures.xlsxRibbon;

const states = [
  { state: "default", matrixState: "default", disabled: false },
  { state: "hover", matrixState: "hover", disabled: false },
  { state: "focus", matrixState: "focus", disabled: false },
  { state: "disabled", matrixState: "empty", disabled: true },
];

const numberFormatLabels: Record<string, string> = {
  currency: "Currency",
  date: "Date",
  general: "General",
  number: "Number",
  percent: "Percent",
};
</script>

<template>
  <ParityCell
    v-for="entry in states"
    :key="entry.state"
    :state="entry.state"
    :matrix-state="entry.matrixState"
    :width="cellWidth"
  >
    <OfficeToolbar class="xlsx-ribbon-parity" data-parity-xlsx-ribbon>
      <OfficeToolbarGroup class="xlsx-ribbon-parity__row" gap="md">
        <OfficeToolbarGroup class="xlsx-ribbon-parity__cluster" gap="sm" :wrap="false">
          <OfficeIconButton
            label="Undo"
            size="icon-sm"
            :disabled="entry.disabled || !ribbon.canUndo"
          >
            <HugeIcon :icon="Undo02Icon" class="xlsx-ribbon-parity__icon" />
          </OfficeIconButton>
          <OfficeIconButton
            label="Redo"
            size="icon-sm"
            :disabled="entry.disabled || !ribbon.canRedo"
          >
            <HugeIcon :icon="Redo02Icon" class="xlsx-ribbon-parity__icon" />
          </OfficeIconButton>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator class="xlsx-ribbon-parity__separator" />

        <OfficeToolbarGroup class="xlsx-ribbon-parity__cluster" gap="sm" :wrap="false">
          <button
            type="button"
            class="xlsx-ribbon-parity__trigger xlsx-ribbon-parity__trigger--font"
            aria-label="Font family"
            :disabled="entry.disabled"
          >
            <span class="xlsx-ribbon-parity__value">{{ ribbon.fontFamily }}</span>
            <OfficeIcon name="chevron-down" class="xlsx-ribbon-parity__chevron" />
          </button>
          <button
            type="button"
            class="xlsx-ribbon-parity__trigger xlsx-ribbon-parity__trigger--size"
            aria-label="Font size"
            :disabled="entry.disabled"
          >
            <span class="xlsx-ribbon-parity__value">{{ ribbon.fontSize }} pt</span>
            <OfficeIcon name="chevron-down" class="xlsx-ribbon-parity__chevron" />
          </button>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator class="xlsx-ribbon-parity__separator" />

        <OfficeToolbarGroup class="xlsx-ribbon-parity__cluster" gap="sm" :wrap="false">
          <OfficeToggle
            class="xlsx-ribbon-parity__toggle"
            label="Bold"
            size="sm"
            :disabled="entry.disabled"
            :model-value="false"
          >
            <HugeIcon :icon="TextBoldIcon" class="xlsx-ribbon-parity__toggle-icon" />
          </OfficeToggle>
          <OfficeToggle
            class="xlsx-ribbon-parity__toggle"
            label="Italic"
            size="sm"
            :disabled="entry.disabled"
            :model-value="false"
          >
            <HugeIcon :icon="TextItalicIcon" class="xlsx-ribbon-parity__toggle-icon" />
          </OfficeToggle>
          <OfficeToggle
            class="xlsx-ribbon-parity__toggle"
            label="Underline"
            size="sm"
            :disabled="entry.disabled"
            :model-value="false"
          >
            <HugeIcon :icon="TextUnderlineIcon" class="xlsx-ribbon-parity__toggle-icon" />
          </OfficeToggle>
          <OfficeToggle
            class="xlsx-ribbon-parity__toggle"
            label="Strikethrough"
            size="sm"
            :disabled="entry.disabled"
            :model-value="false"
          >
            <HugeIcon :icon="TextStrikethroughIcon" class="xlsx-ribbon-parity__toggle-icon" />
          </OfficeToggle>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator class="xlsx-ribbon-parity__separator" />

        <OfficeToolbarGroup class="xlsx-ribbon-parity__cluster" gap="sm" :wrap="false">
          <OfficeIconButton
            label="Text color"
            size="icon-sm"
            :disabled="entry.disabled"
          >
            <HugeIcon :icon="TextColorIcon" class="xlsx-ribbon-parity__icon" />
            <span
              aria-hidden="true"
              class="xlsx-ribbon-parity__swatch"
              :style="{ backgroundColor: ribbon.textColor }"
            />
          </OfficeIconButton>
          <OfficeIconButton
            label="Fill color"
            size="icon-sm"
            :disabled="entry.disabled"
          >
            <HugeIcon :icon="PaintBucketIcon" class="xlsx-ribbon-parity__icon" />
            <span
              aria-hidden="true"
              class="xlsx-ribbon-parity__swatch"
              :style="{ backgroundColor: ribbon.fillColor }"
            />
          </OfficeIconButton>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator class="xlsx-ribbon-parity__separator" />

        <OfficeToolbarGroup class="xlsx-ribbon-parity__cluster" gap="sm" :wrap="false">
          <OfficeToolbarGroup class="xlsx-ribbon-parity__cluster" gap="sm" :wrap="false">
            <OfficeToggle
              class="xlsx-ribbon-parity__toggle"
              label="Align left"
              size="sm"
              :disabled="entry.disabled"
              :model-value="ribbon.horizontalAlignment === 'left'"
            >
              <HugeIcon :icon="TextAlignLeft01Icon" class="xlsx-ribbon-parity__toggle-icon" />
            </OfficeToggle>
            <OfficeToggle
              class="xlsx-ribbon-parity__toggle"
              label="Align center"
              size="sm"
              :disabled="entry.disabled"
              :model-value="ribbon.horizontalAlignment === 'center'"
            >
              <HugeIcon :icon="TextAlignCenterIcon" class="xlsx-ribbon-parity__toggle-icon" />
            </OfficeToggle>
            <OfficeToggle
              class="xlsx-ribbon-parity__toggle"
              label="Align right"
              size="sm"
              :disabled="entry.disabled"
              :model-value="ribbon.horizontalAlignment === 'right'"
            >
              <HugeIcon :icon="TextAlignRight01Icon" class="xlsx-ribbon-parity__toggle-icon" />
            </OfficeToggle>
          </OfficeToolbarGroup>
          <OfficeToolbarGroup class="xlsx-ribbon-parity__cluster" gap="sm" :wrap="false">
            <OfficeToggle
              class="xlsx-ribbon-parity__toggle"
              label="Align top"
              size="sm"
              :disabled="entry.disabled"
              :model-value="ribbon.verticalAlignment === 'top'"
            >
              <HugeIcon :icon="AlignBoxTopCenterIcon" class="xlsx-ribbon-parity__toggle-icon" />
            </OfficeToggle>
            <OfficeToggle
              class="xlsx-ribbon-parity__toggle"
              label="Align middle"
              size="sm"
              :disabled="entry.disabled"
              :model-value="ribbon.verticalAlignment === 'center'"
            >
              <HugeIcon :icon="AlignBoxMiddleCenterIcon" class="xlsx-ribbon-parity__toggle-icon" />
            </OfficeToggle>
            <OfficeToggle
              class="xlsx-ribbon-parity__toggle"
              label="Align bottom"
              size="sm"
              :disabled="entry.disabled"
              :model-value="ribbon.verticalAlignment === 'bottom'"
            >
              <HugeIcon :icon="AlignBoxBottomCenterIcon" class="xlsx-ribbon-parity__toggle-icon" />
            </OfficeToggle>
          </OfficeToolbarGroup>
          <OfficeIconButton
            label="Wrap text"
            size="icon-sm"
            :disabled="entry.disabled"
          >
            <HugeIcon :icon="TextWrapIcon" class="xlsx-ribbon-parity__icon" />
          </OfficeIconButton>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator class="xlsx-ribbon-parity__separator" />

        <OfficeToolbarGroup class="xlsx-ribbon-parity__cluster" gap="sm" :wrap="false">
          <button
            type="button"
            class="xlsx-ribbon-parity__trigger xlsx-ribbon-parity__trigger--number"
            aria-label="Number format"
            :disabled="entry.disabled"
          >
            <HugeIcon
              :icon="TextNumberSignIcon"
              class="xlsx-ribbon-parity__format-icon"
            />
            <span class="xlsx-ribbon-parity__value">{{ numberFormatLabels[ribbon.numberFormat] }}</span>
            <OfficeIcon name="chevron-down" class="xlsx-ribbon-parity__chevron" />
          </button>
          <OfficeIconButton
            label="Cell borders"
            size="icon-sm"
            :disabled="entry.disabled"
          >
            <HugeIcon :icon="BorderAll01Icon" class="xlsx-ribbon-parity__icon" />
            <span
              aria-hidden="true"
              class="xlsx-ribbon-parity__swatch"
              :style="{ backgroundColor: ribbon.borderColor }"
            />
          </OfficeIconButton>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator class="xlsx-ribbon-parity__separator" />

        <div class="xlsx-ribbon-parity__merge">
          <OfficeButton
            class="xlsx-ribbon-parity__merge-main"
            size="sm"
            :variant="ribbon.isMerged ? 'secondary' : 'ghost'"
            :aria-label="'Merge & Center'"
            :disabled="entry.disabled || !ribbon.canMerge"
          >
            <HugeIcon :icon="PathfinderMergeIcon" class="xlsx-ribbon-parity__icon" />
            <span class="xlsx-ribbon-parity__merge-label">Merge &amp; Center</span>
          </OfficeButton>
          <OfficeButton
            class="xlsx-ribbon-parity__merge-toggle"
            size="icon-sm"
            :variant="ribbon.isMerged ? 'secondary' : 'ghost'"
            aria-label="Merge cells options"
            :disabled="entry.disabled"
          >
            <HugeIcon :icon="ChevronDownIcon" class="xlsx-ribbon-parity__merge-chevron" />
          </OfficeButton>
        </div>
      </OfficeToolbarGroup>
    </OfficeToolbar>
  </ParityCell>
</template>

<style scoped>
/*
 * Upstream's row is the page type scale (16px / 24px) while `OfficeToolbar`
 * sets `text-sm`; keep the row metrics at the reference values.
 */
.xlsx-ribbon-parity.office-toolbar {
  font-size: var(--office-text-base);
  line-height: var(--office-leading-base);
  flex: 0 0 auto;
}

/* `flex gap-2` on the reference row; each cluster is `shrink-0`. */
.xlsx-ribbon-parity__row {
  flex: 0 1 auto;
  gap: var(--office-ui-space-2);
}

.xlsx-ribbon-parity__cluster {
  flex: 0 0 auto;
  flex-wrap: nowrap;
}

/* Upstream `ToolbarSeparator` is `mx-1 h-4 self-center`. */
.xlsx-ribbon-parity__separator.office-separator {
  block-size: var(--office-ui-space-4);
  margin-inline: var(--office-ui-space-1);
}

/*
 * Upstream icon buttons carry an explicit `size-4`: 16px at every breakpoint.
 * `.office-button svg` is 18px below 640px, so pin the reference size here.
 */
.xlsx-ribbon-parity .xlsx-ribbon-parity__icon {
  inline-size: 1rem;
  block-size: 1rem;
}

/* Upstream toggle items are `size-8` / `sm:size-7` squares with an 80% glyph.
 * The Reka `Toggle` root does not receive the parent's scope attribute, so the
 * square override is anchored on the cluster (which does). */
.xlsx-ribbon-parity__cluster :deep(.xlsx-ribbon-parity__toggle.office-toggle) {
  inline-size: var(--office-ui-height-sm);
  block-size: var(--office-ui-height-sm);
  padding-inline: 0;
}

.xlsx-ribbon-parity__toggle .xlsx-ribbon-parity__toggle-icon {
  inline-size: 1rem;
  block-size: 1rem;
  opacity: 0.8;
}

/* ColorPicker swatch trigger: `absolute right-1 bottom-1 h-1 w-4 rounded-full`. */
.xlsx-ribbon-parity__swatch {
  position: absolute;
  inset-inline-end: var(--office-ui-space-1);
  inset-block-end: var(--office-ui-space-1);
  inline-size: var(--office-ui-space-4);
  block-size: var(--office-ui-space-1);
  border: var(--office-border-width) solid var(--office-background);
  border-radius: var(--office-radius-full);
}

/*
 * Closed `SelectTrigger size="sm"`: 32px / 16px type below 640px, 28px / 14px
 * above, 9px inline padding, 6px gap, `rounded-lg`, `border-input`, the
 * `shadow-xs` surface and the light/dark 1px inset edge.
 */
.xlsx-ribbon-parity__trigger {
  position: relative;
  display: flex;
  flex: none;
  align-items: center;
  justify-content: space-between;
  gap: 0.375rem;
  block-size: 2rem;
  margin: 0;
  padding: 0 0.5625rem;
  border: var(--office-border-width) solid var(--office-input);
  border-radius: var(--office-radius-lg);
  background: light-dark(
    var(--office-background),
    color-mix(in oklab, var(--office-input) 32%, transparent)
  );
  /* Upstream is `not-dark:bg-clip-padding`, matching the zoom trigger. */
  background-clip: padding-box;
  color: var(--office-foreground);
  font: inherit;
  font-size: var(--office-text-base);
  line-height: var(--office-leading-base);
  text-align: start;
  /* `XLSX_EDITOR_SELECT_CHROME_CLASS` sets `shadow-none before:shadow-none`. */
  box-shadow: none;
  /* Upstream trigger carries `outline-none`; the focus ring is the box shadow. */
  outline: none;
  cursor: pointer;
}

[data-office-theme="dark"] .xlsx-ribbon-parity__trigger {
  background-clip: border-box;
}

.xlsx-ribbon-parity__trigger:focus-visible {
  border-color: var(--office-ring);
  box-shadow: var(--office-ui-ring-shadow);
}

.xlsx-ribbon-parity__trigger:disabled {
  cursor: not-allowed;
  box-shadow: var(--office-ui-empty-shadow);
  opacity: 0.64;
}

.xlsx-ribbon-parity__trigger--font {
  inline-size: 132px;
  min-inline-size: 132px;
}

.xlsx-ribbon-parity__trigger--size {
  inline-size: 78px;
  min-inline-size: 78px;
}

.xlsx-ribbon-parity__trigger--number {
  inline-size: 132px;
  min-inline-size: 132px;
}

.xlsx-ribbon-parity__value {
  flex: 1 1 0%;
  min-inline-size: 0;
  overflow: hidden;
  font-size: inherit;
  line-height: inherit;
  text-overflow: ellipsis;
  white-space: nowrap;
}

/* Upstream `selectTriggerIconClassName`: `-me-1 size-4.5 sm:size-4 opacity-80`. */
.xlsx-ribbon-parity .xlsx-ribbon-parity__chevron {
  flex: none;
  inline-size: 1.125rem;
  block-size: 1.125rem;
  margin-inline-end: calc(var(--office-ui-space-1) * -1);
  opacity: 0.8;
}

.xlsx-ribbon-parity .xlsx-ribbon-parity__format-icon {
  flex: none;
  inline-size: 1rem;
  block-size: 1rem;
  opacity: 0.8;
}

@media (min-width: 640px) {
  .xlsx-ribbon-parity__trigger {
    block-size: var(--office-ui-height-sm);
    font-size: var(--office-text-sm);
    line-height: var(--office-leading-sm);
  }

  .xlsx-ribbon-parity .xlsx-ribbon-parity__chevron {
    inline-size: 1rem;
    block-size: 1rem;
  }
}

/*
 * Merge split button: upstream keeps a 28px row (`h-7`) at every breakpoint,
 * drops the 1px border on both halves and only shows the label at `lg`.
 */
.xlsx-ribbon-parity__merge {
  display: inline-flex;
  flex: 0 0 auto;
  overflow: hidden;
  border: var(--office-border-width) solid transparent;
  border-radius: var(--office-radius-lg);
}

.xlsx-ribbon-parity__merge-main.office-button {
  block-size: 1.75rem;
  padding-inline: 0.625rem;
  border-width: 0;
  border-start-end-radius: 0;
  border-end-end-radius: 0;
}

.xlsx-ribbon-parity__merge-toggle.office-button {
  block-size: 1.75rem;
  inline-size: 1.5rem;
  padding-inline: 0;
  border-width: 0;
  border-start-start-radius: 0;
  border-end-start-radius: 0;
}

.xlsx-ribbon-parity .xlsx-ribbon-parity__merge-label {
  display: none;
  font-size: var(--office-text-xs);
  font-weight: 500;
  line-height: 1;
}

@media (min-width: 640px) {
  /* `w-6` wins below 640px, `sm:size-7` (28px) above it. */
  .xlsx-ribbon-parity__merge-toggle.office-button {
    inline-size: 1.75rem;
  }
}

@media (min-width: 1024px) {
  .xlsx-ribbon-parity .xlsx-ribbon-parity__merge-label {
    display: inline;
  }
}

.xlsx-ribbon-parity .xlsx-ribbon-parity__merge-chevron {
  inline-size: 0.875rem;
  block-size: 0.875rem;
  opacity: 1;
}
</style>

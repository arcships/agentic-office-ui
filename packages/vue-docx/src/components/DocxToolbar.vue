<template>
  <div class="docx-toolbar-shell" :class="{ 'docx-toolbar-shell--dark': toolbarTheme === 'dark' }" data-testid="docx-editor-toolbar">
    <div class="docx-toolbar-topbar">
      <div class="docx-toolbar-topbar__leading">
        <button
          type="button"
          class="docx-toolbar-btn"
          :class="{ active: showThumbnails }"
          :aria-label="m.docx.toolbar.sidebarToggle"
          :aria-pressed="showThumbnails"
          :title="m.docx.toolbar.pageThumbnails"
          @click="emit('toggleThumbnails')"
        >
          <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true">
            <rect x="3" y="4" width="18" height="16" rx="2" />
            <path d="M8 4v16M5.5 8h.01M5.5 12h.01M5.5 16h.01" />
          </svg>
        </button>
        <span class="docx-toolbar-filename" :title="controller.fileName || 'document.docx'">
          {{ controller.fileName || "document.docx" }}
        </span>
        <div class="docx-toolbar-page-control">
          <span>{{ m.common.page.label }}</span>
          <input
            :value="currentPage"
            :aria-label="m.common.page.current"
            inputmode="numeric"
            :disabled="(totalPages ?? 0) < 1"
            @change="onPageChange"
          />
          <span>{{ m.common.page.of }} {{ totalPages || "–" }}</span>
        </div>
      </div>

      <div class="docx-toolbar-topbar__actions">
        <label class="docx-toolbar-toggle" :title="m.docx.toolbar.trackedChanges">
          <input type="checkbox" :checked="controller.showTrackedChanges" @change="controller.toggleShowTrackedChanges()" />
          <span>{{ m.docx.editor.edits }}</span>
        </label>
        <label class="docx-toolbar-toggle" :title="m.docx.toolbar.comments">
          <input type="checkbox" :checked="controller.showComments" @change="controller.toggleShowComments()" />
          <span>{{ m.docx.toolbar.comments }}</span>
        </label>
        <button
          type="button"
          class="docx-toolbar-action"
          data-testid="editor-accept-all-changes"
          :title="m.docx.editor.acceptAllTitle"
          :disabled="readOnly || !hasTrackedChanges"
          @click="controller.acceptAllTrackedChanges()"
        >
          <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="m5 13 4 4L19 7" /></svg>
          <span>{{ m.docx.editor.acceptAll }}</span>
        </button>
        <button
          type="button"
          class="docx-toolbar-action"
          data-testid="editor-reject-all-changes"
          :title="m.docx.editor.rejectAllTitle"
          :disabled="readOnly || !hasTrackedChanges"
          @click="controller.rejectAllTrackedChanges()"
        >
          <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 6l12 12M18 6 6 18" /></svg>
          <span>{{ m.docx.editor.rejectAll }}</span>
        </button>
        <span v-if="readOnly" class="docx-toolbar-readonly">{{ m.common.readOnly }}</span>
        <button type="button" class="docx-toolbar-btn" :title="m.docx.editor.toggleTheme" @click="controller.setDocumentTheme(controller.documentTheme === 'light' ? 'dark' : 'light')">
          <svg v-if="controller.documentTheme === 'light'" class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true">
            <path d="M20.5 14.2A8.5 8.5 0 0 1 9.8 3.5 8.5 8.5 0 1 0 20.5 14.2Z" />
          </svg>
          <svg v-else class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true">
            <circle cx="12" cy="12" r="4" /><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2" />
          </svg>
        </button>
        <span class="docx-toolbar-separator" />
        <label class="docx-toolbar-zoom" :title="m.docx.editor.zoom">
          <select
            data-testid="editor-zoom"
            class="docx-toolbar-select docx-toolbar-select-sm"
            :value="zoomPercent"
            @change="onZoomChange"
            :title="m.docx.editor.zoom"
            :aria-label="m.docx.editor.zoom"
          >
            <option v-for="z of zoomSelectOptions" :key="z" :value="z">{{ z }}%</option>
          </select>
        </label>
        <button data-testid="editor-new" type="button" class="docx-toolbar-action" :title="m.docx.editor.newTitle" :disabled="readOnly" @click="controller.newDocument()">
          <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 3h9l4 4v14H6zM15 3v4h4M9 12h7M12.5 8.5v7" /></svg>
          <span>{{ m.docx.editor.new }}</span>
        </button>
        <button data-testid="editor-import" type="button" class="docx-toolbar-action" :title="m.docx.editor.importTitle" :disabled="readOnly" @click="onImport">
          <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 15V3M7.5 7.5 12 3l4.5 4.5M5 13v6h14v-6" /></svg>
          <span>{{ m.docx.editor.import }}</span>
        </button>
        <button data-testid="editor-export" type="button" class="docx-toolbar-action docx-toolbar-action--primary" :title="m.docx.editor.exportTitle" @click="controller.exportDocx()">
          <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4v12M7.5 11.5 12 16l4.5-4.5M5 19h14" /></svg>
          <span>{{ m.docx.editor.export }}</span>
        </button>
      </div>
    </div>

    <fieldset class="docx-toolbar" :disabled="readOnly">
    <!-- Undo/Redo -->
    <div class="docx-toolbar-group">
      <button data-testid="editor-undo" class="docx-toolbar-btn" :disabled="!controller.canUndo" :title="m.docx.editor.undo" @click="controller.undo()">
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M9 7 4 12l5 5M5 12h8a6 6 0 0 1 6 6" /></svg>
      </button>
      <button data-testid="editor-redo" class="docx-toolbar-btn" :disabled="!controller.canRedo" :title="m.docx.editor.redo" @click="controller.redo()">
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="m15 7 5 5-5 5M19 12h-8a6 6 0 0 0-6 6" /></svg>
      </button>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Font family -->
    <div class="docx-toolbar-group">
      <select
        class="docx-toolbar-select"
        :value="selectedRunStyle?.fontFamily ?? ''"
        @change="onFontFamilyChange"
        :title="m.docx.editor.fontFamily"
      >
        <option value="">{{ m.docx.editor.font }}</option>
        <option value="Calibri">Calibri</option>
        <option value="Arial">Arial</option>
        <option value="Times New Roman">Times New Roman</option>
        <option value="Georgia">Georgia</option>
        <option value="Helvetica">Helvetica</option>
        <option value="Courier New">Courier New</option>
      </select>
    </div>

    <!-- Font size -->
    <div class="docx-toolbar-group">
      <select
        class="docx-toolbar-select docx-toolbar-select-sm"
        :value="selectedRunStyle?.fontSizePt ?? ''"
        @change="onFontSizeChange"
        :title="m.docx.editor.fontSize"
      >
        <option value="">{{ m.docx.editor.size }}</option>
        <option v-for="size of fontSizes" :key="size" :value="size">{{ size }}</option>
      </select>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Text style toggles -->
    <div class="docx-toolbar-group">
      <button
        data-testid="editor-bold"
        type="button"
        class="docx-toolbar-btn"
        :class="{ active: selectedRunStyle?.bold }"
        :title="m.docx.editor.bold"
        :aria-label="m.docx.editor.bold"
        :aria-pressed="Boolean(selectedRunStyle?.bold)"
        @click="controller.toggleBold()"
      >
        <strong>B</strong>
      </button>
      <button
        data-testid="editor-italic"
        type="button"
        class="docx-toolbar-btn"
        :class="{ active: selectedRunStyle?.italic }"
        :title="m.docx.editor.italic"
        :aria-label="m.docx.editor.italic"
        :aria-pressed="Boolean(selectedRunStyle?.italic)"
        @click="controller.toggleItalic()"
      >
        <em>I</em>
      </button>
      <button
        data-testid="editor-underline"
        type="button"
        class="docx-toolbar-btn"
        :class="{ active: selectedRunStyle?.underline }"
        :title="m.docx.editor.underline"
        :aria-label="m.docx.editor.underline"
        :aria-pressed="Boolean(selectedRunStyle?.underline)"
        @click="controller.toggleUnderline()"
      >
        <u>U</u>
      </button>
      <button
        data-testid="editor-strike"
        type="button"
        class="docx-toolbar-btn"
        :class="{ active: selectedRunStyle?.strike }"
        :title="m.docx.editor.strikethrough"
        :aria-label="m.docx.editor.strikethrough"
        :aria-pressed="Boolean(selectedRunStyle?.strike)"
        @click="controller.toggleStrike()"
      >
        <s>S</s>
      </button>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Superscript / Subscript -->
    <div class="docx-toolbar-group">
      <button
        type="button"
        class="docx-toolbar-btn"
        :class="{ active: selectedRunStyle?.verticalAlign === 'superscript' }"
        :title="m.docx.editor.superscript"
        :aria-label="m.docx.editor.superscript"
        :aria-pressed="selectedRunStyle?.verticalAlign === 'superscript'"
        @click="controller.toggleSuperscript()"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="m4 7 7 10M11 7 4 17M15 8c.4-1.3 1.4-2 2.8-2 1.3 0 2.2.7 2.2 1.8 0 1.6-1.7 2.3-4.8 4.2H20" /></svg>
      </button>
      <button
        type="button"
        class="docx-toolbar-btn"
        :class="{ active: selectedRunStyle?.verticalAlign === 'subscript' }"
        :title="m.docx.editor.subscript"
        :aria-label="m.docx.editor.subscript"
        :aria-pressed="selectedRunStyle?.verticalAlign === 'subscript'"
        @click="controller.toggleSubscript()"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="m4 6 7 10M11 6 4 16M15 15c.4-1.3 1.4-2 2.8-2 1.3 0 2.2.7 2.2 1.8 0 1.6-1.7 2.3-4.8 4.2H20" /></svg>
      </button>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Text color + Highlight color -->
    <div class="docx-toolbar-group">
      <div class="docx-toolbar-color-picker" :title="m.docx.editor.textColor">
        <input
          type="color"
          class="docx-toolbar-color-input"
          :value="textColorValue"
          :aria-label="m.docx.editor.textColor"
          @input="onTextColorChange"
        />
        <span class="docx-toolbar-color-label">A</span>
      </div>
      <div class="docx-toolbar-color-picker" :title="m.docx.editor.highlightColor">
        <input
          type="color"
          class="docx-toolbar-color-input"
          :value="highlightColorValue"
          :aria-label="m.docx.editor.highlightColor"
          @input="onHighlightColorChange"
        />
        <span class="docx-toolbar-color-label docx-toolbar-highlight-label">H</span>
      </div>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Hyperlink -->
    <div class="docx-toolbar-group">
      <button
        data-testid="editor-link"
        type="button"
        class="docx-toolbar-btn"
        :class="{ active: !!controller.selectedLink }"
        :title="m.docx.editor.link"
        :aria-label="m.docx.editor.link"
        :aria-pressed="Boolean(controller.selectedLink)"
        @click="onLinkClick"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M9.5 14.5 14.5 9.5M7.3 16.7l-1 1a3.5 3.5 0 0 1-5-5l4-4a3.5 3.5 0 0 1 5 0M16.7 7.3l1-1a3.5 3.5 0 0 1 5 5l-4 4a3.5 3.5 0 0 1-5 0" /></svg>
      </button>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Alignment -->
    <div class="docx-toolbar-group">
      <button
        data-testid="editor-align-left"
        type="button"
        class="docx-toolbar-btn"
        :class="{ active: paragraphAlignment === 'left' }"
        :title="m.docx.editor.alignLeft"
        :aria-label="m.docx.editor.alignLeft"
        :aria-pressed="paragraphAlignment === 'left'"
        @click="controller.setAlignment('left')"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M4 10h11M4 14h16M4 18h9" /></svg>
      </button>
      <button
        data-testid="editor-align-center"
        type="button"
        class="docx-toolbar-btn"
        :class="{ active: paragraphAlignment === 'center' }"
        :title="m.docx.editor.alignCenter"
        :aria-label="m.docx.editor.alignCenter"
        :aria-pressed="paragraphAlignment === 'center'"
        @click="controller.setAlignment('center')"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M7 10h10M4 14h16M8 18h8" /></svg>
      </button>
      <button
        data-testid="editor-align-right"
        type="button"
        class="docx-toolbar-btn"
        :class="{ active: paragraphAlignment === 'right' }"
        :title="m.docx.editor.alignRight"
        :aria-label="m.docx.editor.alignRight"
        :aria-pressed="paragraphAlignment === 'right'"
        @click="controller.setAlignment('right')"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M9 10h11M4 14h16M11 18h9" /></svg>
      </button>
      <button
        data-testid="editor-align-justify"
        type="button"
        class="docx-toolbar-btn"
        :class="{ active: paragraphAlignment === 'justify' }"
        :title="m.docx.editor.justify"
        :aria-label="m.docx.editor.justify"
        :aria-pressed="paragraphAlignment === 'justify'"
        @click="controller.setAlignment('justify')"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M4 10h16M4 14h16M4 18h16" /></svg>
      </button>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Lists -->
    <div class="docx-toolbar-group">
      <button
        type="button"
        class="docx-toolbar-btn"
        :class="{ active: controller.hasUnorderedList }"
        :title="m.docx.editor.bulletList"
        :aria-label="m.docx.editor.bulletList"
        :aria-pressed="controller.hasUnorderedList"
        @click="controller.toggleList('unordered')"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><circle cx="5" cy="7" r="1" /><circle cx="5" cy="12" r="1" /><circle cx="5" cy="17" r="1" /><path d="M9 7h11M9 12h11M9 17h11" /></svg>
      </button>
      <button
        type="button"
        class="docx-toolbar-btn"
        :class="{ active: controller.hasOrderedList }"
        :title="m.docx.editor.numberedList"
        :aria-label="m.docx.editor.numberedList"
        :aria-pressed="controller.hasOrderedList"
        @click="controller.toggleList('ordered')"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 5h1v4M3.5 9H6M3.5 13h2L3.5 17H6M10 7h10M10 15h10" /></svg>
      </button>
      <button
        data-testid="editor-indent"
        class="docx-toolbar-btn"
        :title="m.docx.editor.increaseIndentTitle"
        :aria-label="m.docx.editor.increaseIndent"
        :disabled="!canIncreaseListDepth"
        @click="controller.adjustSelectedListDepth(1)"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M4 12h16M4 18h16M9 6l-4 3 4 3" /></svg>
      </button>
      <button
        data-testid="editor-outdent"
        class="docx-toolbar-btn"
        :title="m.docx.editor.decreaseIndentTitle"
        :aria-label="m.docx.editor.decreaseIndent"
        :disabled="!canDecreaseListDepth"
        @click="controller.adjustSelectedListDepth(-1)"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M4 12h16M4 18h16M5 6l4 3-4 3" /></svg>
      </button>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Line spacing -->
    <div class="docx-toolbar-group">
      <select
        class="docx-toolbar-select docx-toolbar-select-sm"
        :value="lineSpacingValue"
        @change="onLineSpacingChange"
        :title="m.docx.editor.lineSpacing"
      >
        <option value="1">1.00</option>
        <option value="1.15">1.15</option>
        <option value="1.2">1.20</option>
        <option value="1.5">1.50</option>
        <option value="2">2.00</option>
        <option value="2.5">2.50</option>
        <option value="3">3.00</option>
      </select>
    </div>

    <!-- Heading styles -->
    <div class="docx-toolbar-group">
      <select
        class="docx-toolbar-select"
        :value="controller.selectedParagraphStyleId ?? ''"
        @change="onStyleChange"
        :title="m.docx.editor.paragraphStyle"
      >
        <option value="">{{ m.docx.editor.normal }}</option>
        <option
          v-for="style of controller.availableParagraphStyles"
          :key="style.id"
          :value="style.id"
        >
          {{ style.name }}
        </option>
      </select>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Borders -->
    <div class="docx-toolbar-group">
      <select
        class="docx-toolbar-select"
        :value="activeBorderPresetValue"
        @change="onBorderPresetChange"
        :title="m.docx.editor.borders"
      >
        <option value="">{{ m.docx.editor.borders }}</option>
        <option v-for="preset of borderPresets" :key="preset.value" :value="preset.value">
          {{ preset.label }}
        </option>
      </select>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Insert image + Insert table + New comment -->
    <div class="docx-toolbar-group">
      <button class="docx-toolbar-btn" :title="m.docx.editor.insertImage" @click="onInsertImage">
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="16" rx="2" /><circle cx="9" cy="10" r="2" /><path d="m5 18 5-5 3 3 2-2 4 4" /></svg>
      </button>
      <button class="docx-toolbar-btn" :title="m.docx.editor.insertTable" @click="controller.insertTable()">
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="16" rx="1" /><path d="M3 10h18M3 15h18M9 4v16M15 4v16" /></svg>
      </button>
      <button
        data-testid="editor-add-comment"
        type="button"
        class="docx-toolbar-btn"
        :class="{ active: Boolean(controller.commentComposer) }"
        :disabled="readOnly || !canAddComment"
        :title="m.docx.editor.addCommentTitle"
        :aria-label="m.docx.editor.addComment"
        @click="onAddComment"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M21 12a8 8 0 0 1-8 8H7l-4 3V12a8 8 0 0 1 8-8h2a8 8 0 0 1 8 8Z" /><path d="M8 11h8M8 15h5" /></svg>
      </button>
    </div>

    <div class="docx-toolbar-separator" />

    <!-- Page setup (margins / paper size / orientation) -->
    <div class="docx-toolbar-group">
      <button
        type="button"
        class="docx-toolbar-btn"
        :class="{ active: pageSetupOpen }"
        data-testid="editor-page-setup"
        :title="m.docx.pageSetup.open"
        :aria-label="m.docx.pageSetup.open"
        :aria-expanded="pageSetupOpen"
        @click="pageSetupOpen = true"
      >
        <svg class="docx-toolbar-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M6 3h12v18H6zM3 7h3M18 7h3M3 12h3M18 12h3M3 17h3M18 17h3" /></svg>
      </button>
    </div>

    <DocxPageSetupDialog
      v-if="pageSetupOpen"
      :section-index="pageSetup.activeSectionIndex.value"
      :sections="pageSetupSectionOptions"
      :page-size="pageSetup.pageSize.value"
      :margins="pageSetup.margins.value"
      :read-only="readOnly"
      :is-dark="toolbarTheme === 'dark'"
      @apply="onPageSetupApply"
      @update:section-index="pageSetup.selectSection($event)"
      @close="pageSetupOpen = false"
    />

    <!-- Table structure commands for the selected cell/range -->
    <div
      v-if="selectedTableCell"
      class="docx-toolbar-group docx-toolbar-group--table"
      data-testid="editor-table-actions"
    >
      <button data-testid="editor-table-row-above" class="docx-toolbar-btn" :title="m.docx.contextMenu.insertRowAbove" @click="insertSelectedTableRow('above')">R↑</button>
      <button data-testid="editor-table-row-below" class="docx-toolbar-btn" :title="m.docx.contextMenu.insertRowBelow" @click="insertSelectedTableRow('below')">R↓</button>
      <button data-testid="editor-table-column-left" class="docx-toolbar-btn" :title="m.docx.contextMenu.insertColumnLeft" @click="insertSelectedTableColumn('left')">C←</button>
      <button data-testid="editor-table-column-right" class="docx-toolbar-btn" :title="m.docx.contextMenu.insertColumnRight" @click="insertSelectedTableColumn('right')">C→</button>
      <button data-testid="editor-table-delete-row" class="docx-toolbar-btn" :title="m.docx.contextMenu.deleteRow" @click="deleteSelectedTableRow()">R×</button>
      <button data-testid="editor-table-delete-column" class="docx-toolbar-btn" :title="m.docx.contextMenu.deleteColumn" @click="deleteSelectedTableColumn()">C×</button>
      <button
        data-testid="editor-table-merge"
        class="docx-toolbar-btn"
        :title="m.docx.contextMenu.mergeCells"
        :disabled="!canMergeSelectedTableCells"
        @click="mergeSelectedTableCells()"
      >⧉</button>
      <button
        data-testid="editor-table-split"
        class="docx-toolbar-btn"
        :title="m.docx.contextMenu.splitCell"
        :disabled="!canSplitSelectedTableCell"
        @click="splitSelectedTableCell()"
      >⧅</button>
      <button data-testid="editor-table-delete" class="docx-toolbar-btn" :title="m.docx.contextMenu.deleteTable" @click="deleteSelectedTable()">T×</button>
    </div>


    </fieldset>
  </div>
</template>

<script setup lang="ts">
import { computed, ref, watch } from "vue"
import { sanitizeOfficeLink } from "@arcships/office-runtime/link-policy"
import type { DocxEditorController, DocxBorderPreset, HeadingLevel } from "@arcships/docx-core"
import { DOCX_FILE_ACCEPT } from "../composables/docx-file-types"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { useDocxMessages } from "../i18n"
import DocxPageSetupDialog from "./DocxPageSetupDialog.vue"
import { useDocxPageSetup } from "../composables/useDocxPageSetup"
import type { DocxPageSetupApplyPatch } from "./page-setup-dialog-types"

const props = defineProps<{
  controller: DocxEditorController
  zoom?: number
  readOnly?: boolean
  showThumbnails?: boolean
  currentPage?: number
  totalPages?: number
  /** Toolbar appearance. Defaults to the controller document theme. */
  theme?: "light" | "dark"
}>()

const emit = defineEmits<{
  (e: "update:zoom", value: number): void
  (e: "toggleThumbnails"): void
  (e: "selectPage", page: number): void
}>()

const toolbarTheme = computed(() => props.theme ?? props.controller.documentTheme)

// ── Comment authoring ───────────────────────────────────────────────
// A new comment needs a selection (or a caret) in the document; the composer
// itself is rendered in the annotation pane beside the anchored page.
const canAddComment = computed(() => Boolean(props.controller.activeTextRange))

function onAddComment(): void {
  props.controller.beginComment(props.controller.activeTextRange)
}

// ── Font sizes ──────────────────────────────────────────────────────
const fontSizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48]
const m = useDocxMessages()

// ── Page setup ─────────────────────────────────────────────────────
// Margins/paper size/orientation live on the section that holds the caret. The
// dialog edits that section and offers a picker when the document has more
// than one.
const pageSetupOpen = ref(false)
const pageSetup = useDocxPageSetup(props.controller)
const pageSetupSectionOptions = computed(() =>
  pageSetup.sections.value.map((section) => ({
    index: section.index,
    sectionNumber: section.sectionNumber,
  }))
)

function onPageSetupApply(patch: DocxPageSetupApplyPatch): void {
  pageSetup.applyPageSetup(patch)
  pageSetupOpen.value = false
}

// ── Zoom ────────────────────────────────────────────────────────────
const zoomLevels = [50, 60, 70, 80, 90, 100, 110, 120, 130, 150, 175, 200]
const zoomPercent = ref(props.zoom ?? 100)
const zoomSelectOptions = computed(() => {
  const current = zoomPercent.value
  if (!Number.isFinite(current) || zoomLevels.includes(current)) return zoomLevels
  return [...zoomLevels, current].sort((a, b) => a - b)
})

watch(() => props.zoom, (v) => { zoomPercent.value = v ?? 100 })

function onZoomChange(event: Event): void {
  const target = event.target as HTMLSelectElement
  const value = Number(target.value)
  zoomPercent.value = value
  emit("update:zoom", value)
}

function onPageChange(event: Event): void {
  const value = Number((event.target as HTMLInputElement).value)
  const total = Math.max(1, props.totalPages ?? 1)
  if (!Number.isInteger(value)) return
  emit("selectPage", Math.min(Math.max(value, 1), total))
}

// ── Derived values ──────────────────────────────────────────────────
const selectedRunStyle = computed(() => props.controller.selectedRunStyle)
const paragraphAlignment = computed(() => props.controller.selectedParagraph?.style?.align ?? "left")

// Revisions are only counted while they are visible, so `null` means "unknown"
// rather than "none": the buttons stay enabled so a document whose edits are
// currently hidden can still be cleaned up.
const hasTrackedChanges = computed(() => props.controller.trackedChangeCount !== 0)

const textColorValue = computed(() => {
  const color = selectedRunStyle.value?.color
  if (color && /^#[0-9a-fA-F]{6}$/.test(color)) return color
  return "#000000"
})

const highlightColorValue = computed(() => {
  const highlight = selectedRunStyle.value?.highlight
  if (highlight && /^#[0-9a-fA-F]{6}$/.test(highlight)) return highlight
  return "#ffff00"
})

const lineSpacingValue = computed(() => {
  const ls = props.controller.selectedLineSpacing
  const multiple = ls.multiple
  const presets = [1, 1.15, 1.2, 1.5, 2, 2.5, 3]
  const closest = presets.reduce((prev, curr) =>
    Math.abs(curr - multiple) < Math.abs(prev - multiple) ? curr : prev
  )
  return String(closest)
})

const selectedListLevel = computed<number | undefined>(() => {
  const numbering = props.controller.selectedParagraph?.style?.numbering
  if (!numbering || numbering.numId <= 0) return undefined
  return Math.max(0, numbering.ilvl ?? 0)
})
const canIncreaseListDepth = computed(
  () => selectedListLevel.value !== undefined && selectedListLevel.value < 8
)
const canDecreaseListDepth = computed(() => (selectedListLevel.value ?? 0) > 0)

// ── Table commands ──────────────────────────────────────────────────
const selectedTableCell = computed(() => {
  const selection = props.controller.selection
  if (selection.kind !== "table-cell" || selection.tableRuntimeKey) {
    return undefined
  }
  return selection
})

const selectedTableRange = computed(() => selectedTableCell.value?.cellRange)

const canMergeSelectedTableCells = computed(() => {
  const range = selectedTableRange.value
  return Boolean(
    range &&
      (range.anchorRowIndex !== range.focusRowIndex ||
        range.anchorCellIndex !== range.focusCellIndex)
  )
})

const canSplitSelectedTableCell = computed(() => {
  const selection = selectedTableCell.value
  if (!selection) return false
  const table = props.controller.model.nodes[selection.tableIndex]
  if (!table || table.type !== "table") return false
  const cell = table.rows[selection.rowIndex]?.cells[selection.cellIndex]
  return Boolean(
    cell &&
      ((cell.style?.gridSpan ?? 1) > 1 ||
        (cell.style?.rowSpan ?? 1) > 1 ||
        cell.style?.vMergeContinuation)
  )
})

function insertSelectedTableRow(direction: "above" | "below"): void {
  const selection = selectedTableCell.value
  if (!selection) return
  props.controller.insertTableRow(selection.tableIndex, selection.rowIndex, direction)
}

function insertSelectedTableColumn(direction: "left" | "right"): void {
  const selection = selectedTableCell.value
  if (!selection) return
  props.controller.insertTableColumn(
    selection.tableIndex,
    selection.cellIndex,
    direction,
    selection.rowIndex
  )
}

function deleteSelectedTableRow(): void {
  const selection = selectedTableCell.value
  if (!selection) return
  props.controller.deleteTableRow(selection.tableIndex, selection.rowIndex)
}

function deleteSelectedTableColumn(): void {
  const selection = selectedTableCell.value
  if (!selection) return
  props.controller.deleteTableColumn(
    selection.tableIndex,
    selection.cellIndex,
    selection.rowIndex
  )
}

function mergeSelectedTableCells(): void {
  const selection = selectedTableCell.value
  if (!selection) return
  props.controller.mergeTableCells(
    selection.tableIndex,
    selectedTableRange.value ?? {
      tableIndex: selection.tableIndex,
      anchorRowIndex: selection.rowIndex,
      anchorCellIndex: selection.cellIndex,
      focusRowIndex: selection.rowIndex,
      focusCellIndex: selection.cellIndex,
    }
  )
}

function splitSelectedTableCell(): void {
  const selection = selectedTableCell.value
  if (!selection) return
  props.controller.splitTableCell(
    selection.tableIndex,
    selection.rowIndex,
    selection.cellIndex
  )
}

function deleteSelectedTable(): void {
  const selection = selectedTableCell.value
  if (!selection) return
  props.controller.deleteTable(selection.tableIndex)
}

// ── Border presets ──────────────────────────────────────────────────
const borderPresets = computed<{ value: DocxBorderPreset; label: string }[]>(() => [
  { value: "bottom", label: m.value.docx.editor.borderPresets.bottom },
  { value: "top", label: m.value.docx.editor.borderPresets.top },
  { value: "left", label: m.value.docx.editor.borderPresets.left },
  { value: "right", label: m.value.docx.editor.borderPresets.right },
  { value: "none", label: m.value.docx.editor.borderPresets.none },
  { value: "all", label: m.value.docx.editor.borderPresets.all },
  { value: "outside", label: m.value.docx.editor.borderPresets.outside },
  { value: "inside", label: m.value.docx.editor.borderPresets.inside },
  { value: "inside-horizontal", label: m.value.docx.editor.borderPresets.insideHorizontal },
  { value: "inside-vertical", label: m.value.docx.editor.borderPresets.insideVertical },
  { value: "diagonal-down", label: m.value.docx.editor.borderPresets.diagonalDown },
  { value: "diagonal-up", label: m.value.docx.editor.borderPresets.diagonalUp },
  { value: "horizontal-line", label: m.value.docx.editor.borderPresets.horizontalLine },
])

const activeBorderPresetValue = computed(() => {
  const presets = props.controller.activeBorderPresets
  for (const p of borderPresets.value) {
    if (presets[p.value]) return p.value
  }
  return ""
})

// ── Event handlers ──────────────────────────────────────────────────
function onFontFamilyChange(event: Event): void {
  const target = event.target as HTMLSelectElement
  if (target.value) props.controller.setFontFamily(target.value)
}

function onFontSizeChange(event: Event): void {
  const target = event.target as HTMLSelectElement
  const value = Number(target.value)
  if (value) props.controller.setFontSize(value)
}

function onTextColorChange(event: Event): void {
  const target = event.target as HTMLInputElement
  props.controller.setTextColor(target.value || undefined)
}

function onHighlightColorChange(event: Event): void {
  const target = event.target as HTMLInputElement
  props.controller.setHighlight(target.value || undefined)
}

function onLinkClick(): void {
  const existing = props.controller.selectedLink
  if (existing) {
    if (confirm(formatOfficeMessage(m.value.docx.editor.removeLink, { url: existing }))) {
      props.controller.setLink(undefined)
    }
    return
  }
  const url = window.prompt(m.value.docx.editor.enterUrl, "https://")
  if (!url) return
  const safeUrl = sanitizeOfficeLink(url)
  if (!safeUrl) {
    window.alert(m.value.docx.editor.unsupportedLink)
    return
  }
  props.controller.setLink(safeUrl)
}

function onLineSpacingChange(event: Event): void {
  const target = event.target as HTMLSelectElement
  props.controller.setLineSpacing(Number(target.value))
}

function onBorderPresetChange(event: Event): void {
  const target = event.target as HTMLSelectElement
  if (target.value) props.controller.applyBorderPreset(target.value as DocxBorderPreset)
}

function onStyleChange(event: Event): void {
  const target = event.target as HTMLSelectElement
  const value = target.value || undefined
  if (value && value.startsWith("Heading")) {
    const level = parseInt(value.replace("Heading ", ""))
    if (level >= 1 && level <= 6) {
      props.controller.setHeading(level as HeadingLevel)
    } else {
      props.controller.setParagraphStyle(value)
    }
  } else {
    props.controller.setParagraphStyle(value)
  }
}

function onInsertImage(): void {
  const input = document.createElement("input")
  input.type = "file"
  input.accept = "image/*"
  input.onchange = async () => {
    const file = input.files?.[0]
    if (file) {
      await props.controller.insertImageFile(file)
    }
  }
  input.click()
}

function onImport(): void {
  const input = document.createElement("input")
  input.type = "file"
  input.accept = DOCX_FILE_ACCEPT
  input.onchange = async () => {
    const file = input.files?.[0]
    if (file) {
      await props.controller.importDocxFile(file)
    }
  }
  input.click()
}
</script>

<style scoped>
.docx-toolbar-shell {
  background: #fff;
  border-bottom: 1px solid #e4e4e7;
  color: #18181b;
  flex: 0 0 auto;
  min-width: 0;
}
.docx-toolbar-topbar {
  align-items: center;
  border-bottom: 1px solid #e4e4e7;
  display: flex;
  gap: 12px;
  justify-content: space-between;
  min-height: 44px;
  padding: 6px 12px;
}
.docx-toolbar-topbar__leading,
.docx-toolbar-topbar__actions,
.docx-toolbar-page-control {
  align-items: center;
  display: flex;
}
.docx-toolbar-topbar__leading { gap: 9px; min-width: 0; }
.docx-toolbar-topbar__actions { gap: 4px; }
.docx-toolbar-filename {
  font-size: 13px;
  font-weight: 600;
  max-width: 220px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.docx-toolbar-page-control {
  color: #52525b;
  font-size: 12px;
  gap: 5px;
  white-space: nowrap;
}
.docx-toolbar-page-control input {
  background: #fff;
  border: 1px solid #d4d4d8;
  border-radius: 6px;
  color: #18181b;
  font: inherit;
  height: 28px;
  padding: 0 4px;
  text-align: center;
  width: 38px;
}
.docx-toolbar-action {
  align-items: center;
  background: #fff;
  border: 1px solid #d4d4d8;
  border-radius: 7px;
  color: #27272a;
  cursor: pointer;
  display: inline-flex;
  font: inherit;
  font-size: 12px;
  font-weight: 500;
  gap: 6px;
  height: 32px;
  padding: 0 10px;
}
.docx-toolbar-action:hover { background: #f4f4f5; }
.docx-toolbar-action--primary { background: #18181b; border-color: #18181b; color: #fff; }
.docx-toolbar-action--primary:hover { background: #27272a; }
.docx-toolbar-icon {
  fill: none;
  height: 16px;
  stroke: currentColor;
  stroke-linecap: round;
  stroke-linejoin: round;
  stroke-width: 1.7;
  width: 16px;
}
.docx-toolbar-readonly {
  background: #f4f4f5;
  border-radius: 999px;
  color: #71717a;
  font-size: 11px;
  font-weight: 600;
  padding: 4px 8px;
}
.docx-toolbar {
  display: flex;
  align-items: center;
  gap: 2px;
  padding: 6px 12px;
  background: #fff;
  border: 0;
  border-bottom: 0;
  overflow-x: visible;
  flex-shrink: 0;
  flex-wrap: wrap;
  row-gap: 4px;
  margin: 0;
  min-inline-size: 0;
}
.docx-toolbar-group {
  display: flex;
  align-items: center;
  gap: 1px;
  flex-shrink: 0;
}
.docx-toolbar-btn {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 32px;
  height: 32px;
  border: none;
  background: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 13px;
  color: #374151;
}
.docx-toolbar-btn:hover {
  background: #f3f4f6;
}
.docx-toolbar-btn.active {
  background: #e5e7eb;
  color: #2563eb;
}
.docx-toolbar-btn:disabled {
  opacity: 0.3;
  cursor: default;
}
.docx-toolbar-btn:focus-visible,
.docx-toolbar-action:focus-visible,
.docx-toolbar-select:focus-visible,
.docx-toolbar-page-control input:focus-visible {
  outline: 2px solid #2563eb;
  outline-offset: 2px;
}
.docx-toolbar-separator {
  width: 1px;
  height: 24px;
  background: #e5e7eb;
  margin: 0 4px;
  flex-shrink: 0;
}
.docx-toolbar-select {
  height: 32px;
  padding: 0 8px;
  border: 1px solid #e5e7eb;
  border-radius: 4px;
  font-size: 13px;
  background: #fff;
  cursor: pointer;
}
.docx-toolbar-select-sm {
  width: 60px;
  padding: 0 4px;
}
.docx-toolbar-spacer {
  flex: 1;
}
.docx-toolbar-zoom {
  align-items: center;
  display: flex;
}
.docx-toolbar-color-picker {
  position: relative;
  width: 32px;
  height: 32px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 4px;
  cursor: pointer;
}
.docx-toolbar-color-picker:hover {
  background: #f3f4f6;
}
.docx-toolbar-color-input {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  opacity: 0;
  cursor: pointer;
}
.docx-toolbar-color-label {
  font-size: 14px;
  font-weight: 700;
  color: #374151;
  pointer-events: none;
}
.docx-toolbar-highlight-label {
  font-size: 12px;
}
.docx-toolbar-toggle {
  display: flex;
  align-items: center;
  gap: 4px;
  font-size: 12px;
  cursor: pointer;
  user-select: none;
  padding: 2px 6px;
  border-radius: 4px;
}
.docx-toolbar-toggle input {
  appearance: none;
  background: #d4d4d8;
  border-radius: 999px;
  cursor: pointer;
  height: 16px;
  margin: 0;
  position: relative;
  transition: background .15s ease;
  width: 28px;
}
.docx-toolbar-toggle input::after {
  background: #fff;
  border-radius: 50%;
  box-shadow: 0 1px 2px rgba(0, 0, 0, .2);
  content: "";
  height: 12px;
  left: 2px;
  position: absolute;
  top: 2px;
  transition: transform .15s ease;
  width: 12px;
}
.docx-toolbar-toggle input:checked { background: #2563eb; }
.docx-toolbar-toggle input:checked::after { transform: translateX(12px); }
.docx-toolbar-toggle:hover {
  background: #f3f4f6;
}
.docx-toolbar-toggle-label {
  color: #6b7280;
  font-weight: 500;
}

.docx-toolbar-shell--dark {
  background: #18181b;
  border-bottom-color: #3f3f46;
  color: #f4f4f5;
}
.docx-toolbar-shell--dark .docx-toolbar-topbar { border-bottom-color: #3f3f46; }
.docx-toolbar-shell--dark .docx-toolbar { background: #18181b; }
.docx-toolbar-shell--dark .docx-toolbar-page-control { color: #a1a1aa; }
.docx-toolbar-shell--dark .docx-toolbar-page-control input,
.docx-toolbar-shell--dark .docx-toolbar-action,
.docx-toolbar-shell--dark .docx-toolbar-select {
  background: #27272a;
  border-color: #3f3f46;
  color: #f4f4f5;
}
.docx-toolbar-shell--dark .docx-toolbar-action:hover,
.docx-toolbar-shell--dark .docx-toolbar-btn:hover,
.docx-toolbar-shell--dark .docx-toolbar-color-picker:hover,
.docx-toolbar-shell--dark .docx-toolbar-toggle:hover { background: #27272a; }
.docx-toolbar-shell--dark .docx-toolbar-action--primary { background: #f4f4f5; border-color: #f4f4f5; color: #18181b; }
.docx-toolbar-shell--dark .docx-toolbar-action--primary:hover { background: #d4d4d8; }
.docx-toolbar-shell--dark .docx-toolbar-btn { color: #d4d4d8; }
.docx-toolbar-shell--dark .docx-toolbar-btn.active { background: #3f3f46; color: #93c5fd; }
.docx-toolbar-shell--dark .docx-toolbar-separator { background: #3f3f46; }
.docx-toolbar-shell--dark .docx-toolbar-readonly { background: #27272a; color: #a1a1aa; }
.docx-toolbar-shell--dark .docx-toolbar-color-label { color: #d4d4d8; }
.docx-toolbar-shell--dark .docx-toolbar-toggle-label { color: #a1a1aa; }

@container docx-shell (max-width: 900px) {
  .docx-toolbar-topbar { align-items: stretch; flex-direction: column; gap: 5px; }
  .docx-toolbar-topbar__actions { justify-content: flex-end; overflow-x: auto; }
  .docx-toolbar-filename { flex: 1; max-width: none; }
}

@container docx-shell (max-width: 560px) {
  .docx-toolbar-topbar__actions .docx-toolbar-toggle { display: none; }
  .docx-toolbar-action span { display: none; }
  .docx-toolbar-action { padding: 0; width: 32px; }
}

@media (pointer: coarse) {
  .docx-toolbar-action,
  .docx-toolbar-select,
  .docx-toolbar-page-control input {
    min-height: 44px;
    min-width: 44px;
  }
  .docx-toolbar-toggle { min-height: 44px; }
}
</style>

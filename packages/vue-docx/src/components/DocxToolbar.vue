<template>
  <div class="docx-toolbar-shell" :class="{ 'docx-toolbar-shell--dark': toolbarTheme === 'dark' }" data-testid="docx-editor-toolbar">
    <OfficeThemeProvider :theme="toolbarTheme">
    <OfficeToolbar class="docx-toolbar-topbar">
      <template #leading>
        <OfficeToolbarGroup gap="sm" class="docx-toolbar-topbar__leading">
          <OfficeButton
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :class="showThumbnails ? 'active' : undefined"
            data-testid="editor-toggle-thumbnails"
            :aria-label="m.docx.toolbar.sidebarToggle"
            :aria-pressed="showThumbnails"
            :title="m.docx.toolbar.pageThumbnails"
            @click="emit('toggleThumbnails')"
          >
            <OfficeIcon name="panel-left" />
          </OfficeButton>
          <span class="docx-toolbar-filename" :title="controller.fileName || 'document.docx'">
            {{ controller.fileName || "document.docx" }}
          </span>
          <OfficePageControl
            class="docx-toolbar-page-control"
            :model-value="currentPage ?? 1"
            :total-pages="totalPages ?? 0"
            :native-input="true"
            input-test-id="docx-editor-page-current"
            @update:model-value="emit('selectPage', $event)"
          />
        </OfficeToolbarGroup>
      </template>

      <template #trailing>
        <OfficeToolbarGroup gap="sm" class="docx-toolbar-topbar__actions">
          <slot name="find" />
          <OfficeButton
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            data-testid="editor-find-toggle"
            :class="findOpen ? 'active' : undefined"
            :aria-pressed="findOpen"
            :aria-label="m.docx.editor.find"
            :title="m.docx.editor.find"
            @click="emit('toggleFind')"
          >
            <OfficeIcon name="search" />
          </OfficeButton>
          <OfficeButton
            variant="outline"
            size="sm"
            class="docx-toolbar-toggle"
            :aria-pressed="controller.showTrackedChanges"
            :label="m.docx.editor.edits"
            :title="m.docx.toolbar.trackedChanges"
            @click="controller.toggleShowTrackedChanges()"
          >{{ m.docx.editor.edits }}</OfficeButton>
          <label class="docx-toolbar-revision" :title="m.docx.toolbar.revisionDisplay">
            <select
              data-testid="editor-revision-display"
              class="docx-toolbar-select docx-toolbar-select-sm"
              :value="controller.revisionDisplay"
              :aria-label="m.docx.toolbar.revisionDisplay"
              @change="onRevisionDisplayChange"
            >
              <option value="inline">{{ m.docx.toolbar.revisionDisplayInline }}</option>
              <option value="balloons">{{ m.docx.toolbar.revisionDisplayBalloons }}</option>
              <option value="none">{{ m.docx.toolbar.revisionDisplayNone }}</option>
            </select>
          </label>
          <OfficeButton
            variant="outline"
            size="sm"
            class="docx-toolbar-toggle"
            :aria-pressed="controller.showComments"
            :label="m.docx.toolbar.comments"
            :title="m.docx.toolbar.comments"
            @click="controller.toggleShowComments()"
          >{{ m.docx.toolbar.comments }}</OfficeButton>
          <OfficeButton
            variant="outline"
            size="xs"
            class="docx-toolbar-action"
            data-testid="editor-accept-all-changes"
            :title="m.docx.editor.acceptAllTitle"
            :disabled="readOnly || !hasTrackedChanges"
            @click="controller.acceptAllTrackedChanges()"
          >
            <OfficeIcon name="check" />
            <span>{{ m.docx.editor.acceptAll }}</span>
          </OfficeButton>
          <OfficeButton
            variant="outline"
            size="xs"
            class="docx-toolbar-action"
            data-testid="editor-reject-all-changes"
            :title="m.docx.editor.rejectAllTitle"
            :disabled="readOnly || !hasTrackedChanges"
            @click="controller.rejectAllTrackedChanges()"
          >
            <OfficeIcon name="x" />
            <span>{{ m.docx.editor.rejectAll }}</span>
          </OfficeButton>
          <OfficeBadge v-if="readOnly" variant="secondary" size="sm" class="docx-toolbar-readonly">
            {{ m.common.readOnly }}
          </OfficeBadge>
          <OfficeIconButton
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :label="m.docx.editor.toggleTheme"
            :title="m.docx.editor.toggleTheme"
            :icon="controller.documentTheme === 'light' ? 'moon' : 'sun'"
            @click="controller.setDocumentTheme(controller.documentTheme === 'light' ? 'dark' : 'light')"
          />
          <OfficeToolbarSeparator />
          <OfficeZoomControl
            :model-value="zoomPercent"
            :options="zoomSelectOptions"
            :show-fit="false"
            :native-select="true"
            select-test-id="editor-zoom"
            :select-title="m.docx.editor.zoom"
            :label="m.docx.editor.zoom"
            @update:model-value="onZoomValue"
          />
          <OfficeButton
            data-testid="editor-new"
            type="button"
            variant="outline"
            size="sm"
            class="docx-toolbar-action"
            :title="m.docx.editor.newTitle"
            :disabled="readOnly && !controller.viewOnlyReason"
            @click="controller.newDocument()"
          >
            <OfficeIcon name="file-plus" />
            <span>{{ m.docx.editor.new }}</span>
          </OfficeButton>
          <OfficeButton
            data-testid="editor-import"
            type="button"
            variant="outline"
            size="sm"
            class="docx-toolbar-action"
            :title="m.docx.editor.importTitle"
            :disabled="readOnly && !controller.viewOnlyReason"
            @click="onImport"
          >
            <OfficeIcon name="upload" />
            <span>{{ m.docx.editor.import }}</span>
          </OfficeButton>
          <!-- Export stays a native action: it never gets a disabled attribute,
               even while the document is read-only. -->
          <button
            data-testid="editor-export"
            type="button"
            class="office-ui office-button office-button--default office-button--sm docx-toolbar-action docx-toolbar-action--primary"
            :title="m.docx.editor.exportTitle"
            @click="controller.exportDocx()"
          >
            <span class="office-button__content">
              <OfficeIcon name="download" />
              <span>{{ m.docx.editor.export }}</span>
            </span>
          </button>
        </OfficeToolbarGroup>
      </template>
    </OfficeToolbar>

    <fieldset class="docx-toolbar-shell__formatting" :disabled="readOnly">
      <OfficeToolbar class="docx-toolbar">
        <!-- Undo/Redo -->
        <OfficeToolbarGroup gap="sm">
          <OfficeButton
            data-testid="editor-undo"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :disabled="!controller.canUndo"
            :title="m.docx.editor.undo"
            @click="controller.undo()"
          >
            <OfficeIcon name="undo-2" />
          </OfficeButton>
          <OfficeButton
            data-testid="editor-redo"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :disabled="!controller.canRedo"
            :title="m.docx.editor.redo"
            @click="controller.redo()"
          >
            <OfficeIcon name="redo-2" />
          </OfficeButton>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator />

        <!-- Font family -->
        <OfficeToolbarGroup gap="sm">
          <select
            class="docx-toolbar-select"
            :value="selectedRunStyle?.fontFamily ?? ''"
            @change="onFontFamilyChange"
            :title="m.docx.editor.fontFamily"
            :aria-label="m.docx.editor.fontFamily"
          >
            <option value="">{{ m.docx.editor.font }}</option>
            <option value="Calibri">Calibri</option>
            <option value="Arial">Arial</option>
            <option value="Times New Roman">Times New Roman</option>
            <option value="Georgia">Georgia</option>
            <option value="Helvetica">Helvetica</option>
            <option value="Courier New">Courier New</option>
          </select>
          <select
            class="docx-toolbar-select docx-toolbar-select-sm"
            :value="selectedRunStyle?.fontSizePt ?? ''"
            @change="onFontSizeChange"
            :title="m.docx.editor.fontSize"
            :aria-label="m.docx.editor.fontSize"
          >
            <option value="">{{ m.docx.editor.size }}</option>
            <option v-for="size of fontSizes" :key="size" :value="size">{{ size }}</option>
          </select>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator />

        <!-- Text style toggles -->
        <OfficeToolbarGroup gap="sm">
          <OfficeButton
            data-testid="editor-bold"
            type="button"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn docx-toolbar-btn--text"
            :class="selectedRunStyle?.bold ? 'active' : undefined"
            :title="m.docx.editor.bold"
            :aria-label="m.docx.editor.bold"
            :aria-pressed="Boolean(selectedRunStyle?.bold)"
            @click="controller.toggleBold()"
          >
            <strong>B</strong>
          </OfficeButton>
          <OfficeButton
            data-testid="editor-italic"
            type="button"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn docx-toolbar-btn--text"
            :class="selectedRunStyle?.italic ? 'active' : undefined"
            :title="m.docx.editor.italic"
            :aria-label="m.docx.editor.italic"
            :aria-pressed="Boolean(selectedRunStyle?.italic)"
            @click="controller.toggleItalic()"
          >
            <em>I</em>
          </OfficeButton>
          <OfficeButton
            data-testid="editor-underline"
            type="button"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn docx-toolbar-btn--text"
            :class="selectedRunStyle?.underline ? 'active' : undefined"
            :title="m.docx.editor.underline"
            :aria-label="m.docx.editor.underline"
            :aria-pressed="Boolean(selectedRunStyle?.underline)"
            @click="controller.toggleUnderline()"
          >
            <u>U</u>
          </OfficeButton>
          <OfficeButton
            data-testid="editor-strike"
            type="button"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn docx-toolbar-btn--text"
            :class="selectedRunStyle?.strike ? 'active' : undefined"
            :title="m.docx.editor.strikethrough"
            :aria-label="m.docx.editor.strikethrough"
            :aria-pressed="Boolean(selectedRunStyle?.strike)"
            @click="controller.toggleStrike()"
          >
            <s>S</s>
          </OfficeButton>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator />

        <!-- Superscript / Subscript -->
        <OfficeToolbarGroup gap="sm">
          <OfficeButton
            type="button"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :class="selectedRunStyle?.verticalAlign === 'superscript' ? 'active' : undefined"
            :title="m.docx.editor.superscript"
            :aria-label="m.docx.editor.superscript"
            :aria-pressed="selectedRunStyle?.verticalAlign === 'superscript'"
            @click="controller.toggleSuperscript()"
          >
            <OfficeIcon name="superscript" />
          </OfficeButton>
          <OfficeButton
            type="button"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :class="selectedRunStyle?.verticalAlign === 'subscript' ? 'active' : undefined"
            :title="m.docx.editor.subscript"
            :aria-label="m.docx.editor.subscript"
            :aria-pressed="selectedRunStyle?.verticalAlign === 'subscript'"
            @click="controller.toggleSubscript()"
          >
            <OfficeIcon name="subscript" />
          </OfficeButton>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator />

        <!-- Text color + Highlight color -->
        <OfficeToolbarGroup gap="sm">
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
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator />

        <!-- Hyperlink -->
        <OfficeToolbarGroup gap="sm">
          <OfficeButton
            data-testid="editor-link"
            type="button"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :class="controller.selectedLink ? 'active' : undefined"
            :title="m.docx.editor.link"
            :aria-label="m.docx.editor.link"
            :aria-pressed="Boolean(controller.selectedLink)"
            @click="onLinkClick"
          >
            <OfficeIcon name="link" />
          </OfficeButton>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator />

        <!-- Alignment -->
        <OfficeToolbarGroup gap="sm">
          <OfficeButton
            data-testid="editor-align-left"
            type="button"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :class="paragraphAlignment === 'left' ? 'active' : undefined"
            :title="m.docx.editor.alignLeft"
            :aria-label="m.docx.editor.alignLeft"
            :aria-pressed="paragraphAlignment === 'left'"
            @click="controller.setAlignment('left')"
          >
            <OfficeIcon name="align-left" />
          </OfficeButton>
          <OfficeButton
            data-testid="editor-align-center"
            type="button"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :class="paragraphAlignment === 'center' ? 'active' : undefined"
            :title="m.docx.editor.alignCenter"
            :aria-label="m.docx.editor.alignCenter"
            :aria-pressed="paragraphAlignment === 'center'"
            @click="controller.setAlignment('center')"
          >
            <OfficeIcon name="align-center" />
          </OfficeButton>
          <OfficeButton
            data-testid="editor-align-right"
            type="button"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :class="paragraphAlignment === 'right' ? 'active' : undefined"
            :title="m.docx.editor.alignRight"
            :aria-label="m.docx.editor.alignRight"
            :aria-pressed="paragraphAlignment === 'right'"
            @click="controller.setAlignment('right')"
          >
            <OfficeIcon name="align-right" />
          </OfficeButton>
          <OfficeButton
            data-testid="editor-align-justify"
            type="button"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :class="paragraphAlignment === 'justify' ? 'active' : undefined"
            :title="m.docx.editor.justify"
            :aria-label="m.docx.editor.justify"
            :aria-pressed="paragraphAlignment === 'justify'"
            @click="controller.setAlignment('justify')"
          >
            <OfficeIcon name="align-justify" />
          </OfficeButton>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator />

        <!-- Lists -->
        <OfficeToolbarGroup gap="sm">
          <OfficeButton
            type="button"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :class="controller.hasUnorderedList ? 'active' : undefined"
            :title="m.docx.editor.bulletList"
            :aria-label="m.docx.editor.bulletList"
            :aria-pressed="controller.hasUnorderedList"
            @click="controller.toggleList('unordered')"
          >
            <OfficeIcon name="list" />
          </OfficeButton>
          <OfficeButton
            type="button"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :class="controller.hasOrderedList ? 'active' : undefined"
            :title="m.docx.editor.numberedList"
            :aria-label="m.docx.editor.numberedList"
            :aria-pressed="controller.hasOrderedList"
            @click="controller.toggleList('ordered')"
          >
            <OfficeIcon name="list-ordered" />
          </OfficeButton>
          <OfficeButton
            data-testid="editor-indent"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :title="m.docx.editor.increaseIndentTitle"
            :aria-label="m.docx.editor.increaseIndent"
            :disabled="!canIncreaseListDepth"
            @click="controller.adjustSelectedListDepth(1)"
          >
            <OfficeIcon name="indent-increase" />
          </OfficeButton>
          <OfficeButton
            data-testid="editor-outdent"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :title="m.docx.editor.decreaseIndentTitle"
            :aria-label="m.docx.editor.decreaseIndent"
            :disabled="!canDecreaseListDepth"
            @click="controller.adjustSelectedListDepth(-1)"
          >
            <OfficeIcon name="indent-decrease" />
          </OfficeButton>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator />

        <!-- Body columns of the caret's section -->
        <OfficeToolbarGroup gap="sm">
          <select
            data-testid="editor-columns"
            class="docx-toolbar-select docx-toolbar-select-sm"
            :value="pageSetup.columns.value.count"
            :title="m.docx.editor.columns"
            :aria-label="m.docx.editor.columns"
            @change="onColumnsChange"
          >
            <option
              v-for="option of columnOptions"
              :key="option.count"
              :value="option.count"
            >
              {{ option.label }}
            </option>
          </select>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator />

        <!-- Line spacing -->
        <OfficeToolbarGroup gap="sm">
          <select
            class="docx-toolbar-select docx-toolbar-select-sm"
            :value="lineSpacingValue"
            @change="onLineSpacingChange"
            :title="m.docx.editor.lineSpacing"
            :aria-label="m.docx.editor.lineSpacing"
          >
            <option value="1">1.00</option>
            <option value="1.15">1.15</option>
            <option value="1.2">1.20</option>
            <option value="1.5">1.50</option>
            <option value="2">2.00</option>
            <option value="2.5">2.50</option>
            <option value="3">3.00</option>
          </select>
        </OfficeToolbarGroup>

        <!-- Heading styles -->
        <OfficeToolbarGroup
          ref="stylePickerEl"
          gap="sm"
          class="docx-toolbar-style-picker"
          data-testid="editor-paragraph-style-picker"
          @pointerenter="onStylePickerEnter"
          @pointerleave="onStylePickerLeave"
        >
          <select
            class="docx-toolbar-select"
            data-testid="editor-paragraph-style"
            :value="controller.selectedParagraphStyleId ?? ''"
            @change="onStyleChange"
            @mousedown="closeStylePreview"
            @focus="closeStylePreview"
            :title="m.docx.editor.paragraphStyle"
            :aria-label="m.docx.editor.paragraphStyle"
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
          <DocxStylePreviewPopover
            v-if="showStylePreview !== false && stylePreviewOpen"
            :open="stylePreviewOpen"
            :styles="controller.availableParagraphStyles"
            :preview-style-id="stylePreviewStyleId"
            :document-theme="toolbarTheme"
            :x="stylePreviewPosition.x"
            :y="stylePreviewPosition.y"
            @hover="stylePreviewStyleId = $event"
            @select="onStylePreviewSelect"
            @close="closeStylePreview"
            @pointerenter="cancelStylePreviewClose"
            @pointerleave="scheduleStylePreviewClose"
          />
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator />

        <!-- Borders -->
        <OfficeToolbarGroup gap="sm">
          <select
            class="docx-toolbar-select"
            :value="activeBorderPresetValue"
            @change="onBorderPresetChange"
            :title="m.docx.editor.borders"
            :aria-label="m.docx.editor.borders"
          >
            <option value="">{{ m.docx.editor.borders }}</option>
            <option v-for="preset of borderPresets" :key="preset.value" :value="preset.value">
              {{ preset.label }}
            </option>
          </select>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator />

        <!-- Insert image + Insert table + New comment -->
        <OfficeToolbarGroup gap="sm">
          <OfficeButton
            data-testid="editor-insert-image"
            type="button"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :title="m.docx.editor.insertImage"
            :aria-label="m.docx.editor.insertImage"
            @click="onInsertImage"
          >
            <OfficeIcon name="image" />
          </OfficeButton>
          <OfficeButton
            data-testid="editor-insert-table"
            type="button"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :title="m.docx.editor.insertTable"
            :aria-label="m.docx.editor.insertTable"
            @click="controller.insertTable()"
          >
            <OfficeIcon name="table" />
          </OfficeButton>
          <OfficeButton
            data-testid="editor-add-comment"
            type="button"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :class="controller.commentComposer ? 'active' : undefined"
            :disabled="readOnly || !canAddComment"
            :title="m.docx.editor.addCommentTitle"
            :aria-label="m.docx.editor.addComment"
            @click="onAddComment"
          >
            <OfficeIcon name="message-square-plus" />
          </OfficeButton>
        </OfficeToolbarGroup>

        <OfficeToolbarSeparator />

        <!-- Page setup (margins / paper size / orientation) -->
        <OfficeToolbarGroup gap="sm">
          <OfficeButton
            type="button"
            variant="ghost"
            size="icon-sm"
            class="docx-toolbar-btn"
            :class="pageSetupOpen ? 'active' : undefined"
            data-testid="editor-page-setup"
            :title="m.docx.pageSetup.open"
            :aria-label="m.docx.pageSetup.open"
            :aria-expanded="pageSetupOpen"
            @click="pageSetupOpen = true"
          >
            <OfficeIcon name="settings" />
          </OfficeButton>
        </OfficeToolbarGroup>

        <!-- Table structure commands for the selected cell/range -->
        <OfficeToolbarGroup
          v-if="selectedTableCell"
          gap="sm"
          class="docx-toolbar-group--table"
          data-testid="editor-table-actions"
        >
          <OfficeButton data-testid="editor-table-row-above" variant="ghost" size="xs" class="docx-toolbar-btn" :title="m.docx.contextMenu.insertRowAbove" @click="insertSelectedTableRow('above')">R↑</OfficeButton>
          <OfficeButton data-testid="editor-table-row-below" variant="ghost" size="xs" class="docx-toolbar-btn" :title="m.docx.contextMenu.insertRowBelow" @click="insertSelectedTableRow('below')">R↓</OfficeButton>
          <OfficeButton data-testid="editor-table-column-left" variant="ghost" size="xs" class="docx-toolbar-btn" :title="m.docx.contextMenu.insertColumnLeft" @click="insertSelectedTableColumn('left')">C←</OfficeButton>
          <OfficeButton data-testid="editor-table-column-right" variant="ghost" size="xs" class="docx-toolbar-btn" :title="m.docx.contextMenu.insertColumnRight" @click="insertSelectedTableColumn('right')">C→</OfficeButton>
          <OfficeButton data-testid="editor-table-delete-row" variant="ghost" size="xs" class="docx-toolbar-btn" :title="m.docx.contextMenu.deleteRow" @click="deleteSelectedTableRow()">R×</OfficeButton>
          <OfficeButton data-testid="editor-table-delete-column" variant="ghost" size="xs" class="docx-toolbar-btn" :title="m.docx.contextMenu.deleteColumn" @click="deleteSelectedTableColumn()">C×</OfficeButton>
          <OfficeButton
            data-testid="editor-table-merge"
            variant="ghost"
            size="xs"
            class="docx-toolbar-btn"
            :title="m.docx.contextMenu.mergeCells"
            :disabled="!canMergeSelectedTableCells"
            @click="mergeSelectedTableCells()"
          >⧉</OfficeButton>
          <OfficeButton
            data-testid="editor-table-split"
            variant="ghost"
            size="xs"
            class="docx-toolbar-btn"
            :title="m.docx.contextMenu.splitCell"
            :disabled="!canSplitSelectedTableCell"
            @click="splitSelectedTableCell()"
          >⧅</OfficeButton>
          <OfficeButton
            data-testid="editor-table-select-all"
            variant="ghost"
            size="xs"
            class="docx-toolbar-btn"
            :title="m.docx.contextMenu.selectTable"
            @click="selectWholeTable()"
          >▦</OfficeButton>
          <OfficeButton
            data-testid="editor-table-clear"
            variant="ghost"
            size="xs"
            class="docx-toolbar-btn"
            :title="m.docx.contextMenu.clearContents"
            :disabled="!canClearSelectedTableCells"
            @click="clearSelectedTableCells()"
          >⌫</OfficeButton>
          <OfficeButton data-testid="editor-table-delete" variant="ghost" size="xs" class="docx-toolbar-btn" :title="m.docx.contextMenu.deleteTable" @click="deleteSelectedTable()">T×</OfficeButton>
        </OfficeToolbarGroup>
      </OfficeToolbar>
    </fieldset>

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
    </OfficeThemeProvider>
  </div>
</template>

<script setup lang="ts">
import { computed, onUnmounted, ref, watch } from "vue"
import { sanitizeOfficeLink } from "@arcships/office-runtime/link-policy"
import type { DocxEditorController, DocxBorderPreset, HeadingLevel } from "@arcships/docx-core"
import {
  OfficeBadge,
  OfficeButton,
  OfficeIcon,
  OfficeIconButton,
  OfficePageControl,
  OfficeToolbar,
  OfficeToolbarGroup,
  OfficeToolbarSeparator,
  OfficeThemeProvider,
  OfficeZoomControl,
} from "@arcships/vue-ui"
import { DOCX_FILE_ACCEPT } from "../composables/docx-file-types"
import { formatOfficeMessage } from "@arcships/office-interaction"
import { useDocxMessages } from "../i18n"
import DocxPageSetupDialog from "./DocxPageSetupDialog.vue"
import { useDocxPageSetup } from "../composables/useDocxPageSetup"
import type { DocxPageSetupApplyPatch } from "./page-setup-dialog-types"
import DocxStylePreviewPopover from "./DocxStylePreviewPopover.vue"

const props = withDefaults(
  defineProps<{
    controller: DocxEditorController
    zoom?: number
    readOnly?: boolean
    showThumbnails?: boolean
    currentPage?: number
    totalPages?: number
    /** Toolbar appearance. Defaults to the controller document theme. */
    theme?: "light" | "dark"
    /** Hovering the paragraph-style picker previews each style. */
    showStylePreview?: boolean
    /** True while the host renders the find & replace bar in the `find` slot. */
    findOpen?: boolean
  }>(),
  { showStylePreview: true }
)

const emit = defineEmits<{
  (e: "update:zoom", value: number): void
  (e: "toggleThumbnails"): void
  (e: "selectPage", page: number): void
  (e: "toggleFind"): void
}>()

const toolbarTheme = computed(() => props.theme ?? props.controller.documentTheme)

// ── Paragraph style preview ─────────────────────────────────────────
// A native <select> cannot show what a style looks like: hovering the picker
// opens a preview list whose sample text follows the hovered style. Clicking the
// select (or an item) applies it; the select stays the keyboard/simple path.
const stylePickerEl = ref<InstanceType<typeof OfficeToolbarGroup> | undefined>()
const stylePreviewOpen = ref(false)
const stylePreviewStyleId = ref<string | undefined>()
const stylePreviewPosition = ref({ x: 0, y: 0 })
let stylePreviewHideTimer: ReturnType<typeof setTimeout> | undefined

function stylePickerElement(): HTMLElement | undefined {
  const instance = stylePickerEl.value as unknown
  if (!instance) return undefined
  const root = (instance as { $el?: unknown }).$el ?? instance
  return typeof (root as HTMLElement).getBoundingClientRect === "function"
    ? (root as HTMLElement)
    : undefined
}

function cancelStylePreviewClose(): void {
  if (stylePreviewHideTimer !== undefined) clearTimeout(stylePreviewHideTimer)
  stylePreviewHideTimer = undefined
}

function scheduleStylePreviewClose(): void {
  cancelStylePreviewClose()
  stylePreviewHideTimer = setTimeout(() => {
    stylePreviewOpen.value = false
    stylePreviewHideTimer = undefined
  }, 160)
}

function openStylePreview(): void {
  cancelStylePreviewClose()
  if (stylePreviewOpen.value) return
  stylePreviewStyleId.value = props.controller.selectedParagraphStyleId
  const rect = stylePickerElement()?.getBoundingClientRect?.()
  if (rect) {
    stylePreviewPosition.value = {
      x: Math.round(rect.left),
      y: Math.round(rect.bottom + 6),
    }
  } else {
    stylePreviewPosition.value = { x: 0, y: 0 }
  }
  stylePreviewOpen.value = true
}

function closeStylePreview(): void {
  cancelStylePreviewClose()
  stylePreviewOpen.value = false
  stylePreviewStyleId.value = undefined
}

function onStylePickerEnter(): void {
  if (props.showStylePreview === false) return
  openStylePreview()
}

function onStylePickerLeave(): void {
  scheduleStylePreviewClose()
}

function onStylePreviewSelect(styleId: string | undefined): void {
  applyParagraphStyleId(styleId)
  closeStylePreview()
}

onUnmounted(cancelStylePreviewClose)

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

// ── Section columns ─────────────────────────────────────────────────
// The toolbar offers Word's common 1/2/3-column choices and keeps a larger
// count visible when the imported section already uses one. Changing the
// value applies to the section that holds the caret.
function columnsLabel(count: number): string {
  if (count === 1) return m.value.docx.editor.columnsOne
  if (count === 2) return m.value.docx.editor.columnsTwo
  if (count === 3) return m.value.docx.editor.columnsThree
  return formatOfficeMessage(m.value.docx.editor.columnsMore, { count })
}

const columnOptions = computed(() => {
  const current = Math.max(1, Math.round(pageSetup.columns.value.count))
  const counts = [1, 2, 3]
  if (!counts.includes(current)) counts.push(current)
  return counts
    .sort((left, right) => left - right)
    .map((count) => ({ count, label: columnsLabel(count) }))
})

function onColumnsChange(event: Event): void {
  const value = Number((event.target as HTMLSelectElement | null)?.value)
  if (!Number.isFinite(value)) return
  pageSetup.applyColumns(value)
}

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

function onZoomValue(value: number): void {
  if (!Number.isFinite(value)) return
  zoomPercent.value = value
  emit("update:zoom", value)
}

function onRevisionDisplayChange(event: Event): void {
  const value = (event.target as HTMLSelectElement).value
  if (value === "inline" || value === "balloons" || value === "none") {
    props.controller.setRevisionDisplay(value)
  }
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

const canClearSelectedTableCells = computed(() => {
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

function selectWholeTable(): void {
  const selection = selectedTableCell.value
  if (!selection) return
  props.controller.selectWholeTable(selection.tableIndex)
}

function clearSelectedTableCells(): void {
  props.controller.clearTableSelection()
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

function applyParagraphStyleId(value: string | undefined): void {
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

function onStyleChange(event: Event): void {
  const target = event.target as HTMLSelectElement
  applyParagraphStyleId(target.value || undefined)
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
  background: var(--office-background, #fff);
  border-bottom: 1px solid var(--office-border, #e4e4e7);
  color: var(--office-foreground, #18181b);
  flex: 0 0 auto;
  min-width: 0;
}

/* Host pages may impose a global min-height on buttons/inputs. The editor
 * toolbar owns the 28/32px control contract and opts into 44px only for
 * coarse pointers. */
.docx-toolbar-shell :deep(button),
.docx-toolbar-shell :deep(input),
.docx-toolbar-shell :deep(select) {
  min-block-size: 0;
}
.docx-toolbar-shell--dark {
  background: var(--office-background, #18181b);
  border-bottom-color: var(--office-border, #3f3f46);
  color: var(--office-foreground, #f4f4f5);
}
.docx-toolbar-topbar {
  --office-ui-toolbar-min-height: 2.75rem;
  --office-ui-toolbar-padding-block: var(--office-ui-space-1-5);
  --office-ui-toolbar-padding-inline: var(--office-ui-space-3);
  border-bottom: 1px solid var(--office-border, #e4e4e7);
}
.docx-toolbar-topbar :deep(.office-toolbar__leading),
.docx-toolbar-topbar :deep(.office-toolbar__trailing) {
  min-inline-size: 0;
}
.docx-toolbar-topbar :deep(.office-toolbar__trailing) {
  flex: 0 1 auto;
}
.docx-toolbar-topbar__leading {
  flex: 1 1 auto;
  min-inline-size: 0;
  gap: var(--office-ui-space-2);
}
.docx-toolbar-topbar__actions {
  flex: 0 1 auto;
  justify-content: flex-end;
  max-inline-size: 100%;
  row-gap: var(--office-ui-space-1);
}
.docx-toolbar-filename {
  font-size: 13px;
  font-weight: 600;
  max-width: 220px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.docx-toolbar-page-control {
  flex: 0 0 auto;
}
.docx-toolbar-shell__formatting {
  border: 0;
  margin: 0;
  min-inline-size: 0;
  padding: 0;
}
.docx-toolbar {
  --office-ui-toolbar-min-height: 0;
  --office-ui-toolbar-padding-block: var(--office-ui-space-1-5);
  --office-ui-toolbar-padding-inline: var(--office-ui-space-3);
  --office-ui-toolbar-gap: var(--office-ui-space-0-5);
  border-block-end: 0;
  background: var(--office-background, #fff);
  row-gap: var(--office-ui-space-1);
}
.docx-toolbar-btn {
  color: var(--office-muted-foreground, #374151);
}
.docx-toolbar-btn:hover:not(:disabled) {
  color: var(--office-foreground, #18181b);
}
.docx-toolbar-btn.active,
.docx-toolbar-btn[aria-pressed="true"] {
  background: var(--office-accent, #e5e7eb);
  color: var(--office-foreground, #2563eb);
}
.docx-toolbar-btn--text {
  font-size: 13px;
}
.docx-toolbar-action {
  gap: var(--office-ui-space-1);
  white-space: nowrap;
}
.docx-toolbar-action--primary {
  background: var(--office-primary, #18181b);
  border-color: var(--office-primary, #18181b);
  color: var(--office-primary-foreground, #fff);
}
.docx-toolbar-action--primary:hover {
  background: color-mix(in oklab, var(--office-primary, #18181b) 90%, transparent);
}
.docx-toolbar-readonly {
  flex: 0 0 auto;
}
.docx-toolbar-select {
  height: var(--office-ui-height-sm, 2rem);
  padding: 0 var(--office-ui-space-2, 0.5rem);
  border: 1px solid var(--office-input, #e5e7eb);
  border-radius: var(--office-radius-md, 4px);
  background: var(--office-background, #fff);
  color: var(--office-foreground, #18181b);
  font: inherit;
  font-size: var(--office-text-sm, 13px);
  cursor: pointer;
}
.docx-toolbar-select-sm {
  width: 60px;
  padding: 0 var(--office-ui-space-1, 0.25rem);
}
.docx-toolbar-select:focus-visible {
  outline: var(--office-ring-width, 2px) solid var(--office-ring, #2563eb);
  outline-offset: var(--office-ring-offset-width, 2px);
}
.docx-toolbar-revision {
  align-items: center;
  display: inline-flex;
  min-inline-size: 0;
}
.docx-toolbar-color-picker {
  position: relative;
  width: var(--office-ui-height-sm, 2rem);
  height: var(--office-ui-height-sm, 2rem);
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: var(--office-radius-md, 4px);
  cursor: pointer;
}
.docx-toolbar-color-picker:hover {
  background: var(--office-accent, #f3f4f6);
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
  color: var(--office-foreground, #374151);
  pointer-events: none;
}
.docx-toolbar-highlight-label {
  font-size: 12px;
}
.docx-toolbar-style-picker {
  position: relative;
}

@container docx-shell (max-width: 900px) {
  .docx-toolbar-topbar__leading,
  .docx-toolbar-topbar__actions {
    flex: 1 1 100%;
  }
  .docx-toolbar-topbar :deep(.office-toolbar__trailing) {
    flex: 1 1 100%;
  }
  .docx-toolbar-topbar__actions {
    justify-content: flex-end;
    overflow-x: auto;
  }
  .docx-toolbar-filename { flex: 1; max-width: none; }
}

@container docx-shell (max-width: 560px) {
  .docx-toolbar-topbar__actions .office-toggle { display: none; }
  .docx-toolbar-action span { display: none; }
  .docx-toolbar-action { padding: 0; width: var(--office-ui-height-sm, 2rem); }
}

@media (pointer: coarse) {
  .docx-toolbar-shell :deep(button),
  .docx-toolbar-shell :deep(input),
  .docx-toolbar-shell :deep(select) {
    min-block-size: 44px;
  }
  .docx-toolbar-action,
  .docx-toolbar-select,
  .docx-toolbar-page-control {
    min-height: 44px;
    min-width: 44px;
  }
}
</style>

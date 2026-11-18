<template>
  <div ref="stripRef" class="xlsx-sheettabs" :style="tabsStyle">
    <button
      type="button"
      class="xlsx-sheettabs__nav-btn"
      :disabled="activeTabIndex <= 0"
      :title="m.xlsx.sheetTabs.previous"
      :aria-label="m.xlsx.sheetTabs.previous"
      @click="prevTab"
    >
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="m14.5 6-6 6 6 6" /></svg>
    </button>
    <div
      class="xlsx-sheettabs__list"
      role="tablist"
      aria-orientation="horizontal"
      :aria-label="m.xlsx.sheetTabs.tabListLabel"
    >
      <div
        v-for="(tab, index) in controller.tabs"
        :key="tab.id ?? index"
        class="xlsx-sheettabs__tab-wrapper"
        :class="{
          'xlsx-sheettabs__tab-wrapper--draggable': canMoveSheets,
          'xlsx-sheettabs__tab-wrapper--dragging': draggingTabIndex === index,
        }"
        :data-dragging="draggingTabIndex === index ? 'true' : undefined"
        :draggable="canMoveSheets && renamingTabIndex !== index"
        :ref="(el) => setTabWrapperRef(index, el)"
        @dragstart="onTabDragStart(index, $event)"
        @dragover="onTabDragOver(index, $event)"
        @drop="onTabDrop(index, $event)"
        @dragend="cancelDrag"
        @focusin="onTabFocus(index)"
        @focusout="onTabFocusOut(index)"
        @keydown="onTabDragKeydown"
        @mouseenter="onTabHover(index, $event)"
        @mouseleave="onTabLeave"
        @contextmenu="onTabContextMenu(index, $event)"
      >
        <div
          v-if="dropSlot === index || (dropSlot === controller.tabs.length && index === controller.tabs.length - 1)"
          class="xlsx-sheettabs__drop-indicator"
          :class="{ 'xlsx-sheettabs__drop-indicator--after': dropSlot === controller.tabs.length }"
          data-testid="xlsx-sheet-drop-indicator"
          :data-drop-slot="dropSlot"
          aria-hidden="true"
        />
        <input
          v-if="renamingTabIndex === index"
          ref="renameInputRef"
          class="xlsx-sheettabs__rename-input"
          data-testid="xlsx-sheet-rename-input"
          type="text"
          :aria-invalid="renameError ? 'true' : 'false'"
          :aria-label="m.xlsx.sheetTabs.renameLabel"
          :maxlength="maxSheetNameLength"
          :style="renameInputStyle(index)"
          :value="renameDraft"
          @blur="onRenameBlur"
          @click.stop
          @dblclick.stop
          @input="onRenameInput"
          @keydown.enter.prevent.stop="commitRename"
          @keydown.esc.prevent.stop="cancelRename"
        />
        <button
          v-else
          :ref="(el) => setTabButtonRef(index, el)"
          :id="tabDomId(index)"
          class="xlsx-sheettabs__tab"
          data-testid="xlsx-sheet-tab"
          :data-sheet-name="tab.name"
          :data-visibility="tab.kind === 'sheet' ? tab.visibility ?? 'visible' : undefined"
          role="tab"
          :aria-selected="index === activeTabIndex"
          :aria-controls="panelId"
          :tabindex="tabStopIndex(index)"
          :class="{
            'xlsx-sheettabs__tab--active': index === activeTabIndex,
            'xlsx-sheettabs__tab--hidden': tab.kind === 'sheet' && tab.visibility === 'hidden',
          }"
          :style="tabStyle(index)"
          :title="tabTitle(tab)"
          @click="onTabClick(index)"
          @dblclick="onTabDblClick(tab, index)"
          @keydown="onTabKeydown(index, tab, $event)"
        >
          <span class="xlsx-sheettabs__tab-name">{{ tab.name }}</span>
        </button>
        <button
          v-if="canManageSheets && tab.kind === 'sheet' && tab.visibility === 'hidden'"
          type="button"
          class="xlsx-sheettabs__unhide-btn"
          data-testid="xlsx-unhide-sheet-tab"
          :data-sheet-name="tab.name"
          :title="m.xlsx.sheetTabs.unhide"
          :aria-label="`${m.xlsx.sheetTabs.unhide} ${tab.name}`"
          @click.stop="unhideTabFromStrip(tab)"
        >
          <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M2 12s3.5-6 10-6 10 6 10 6-3.5 6-10 6S2 12 2 12Z" /><circle cx="12" cy="12" r="2.5" /></svg>
        </button>
        <div
          v-if="hoveredTabIndex === index && renamingTabIndex !== index && thumbnailComputed[index]"
          class="xlsx-sheettabs__thumbnail"
          :style="thumbnailPopupStyle"
        >
          <canvas
            :ref="(el) => setThumbnailCanvas(index, el)"
            :height="thumbnailComputed[index]?.height"
            :width="thumbnailComputed[index]?.width"
          />
        </div>
      </div>
      <button
        v-if="!isReadOnly"
        type="button"
        class="xlsx-sheettabs__add-btn"
        data-testid="xlsx-add-sheet"
        :title="m.xlsx.sheetTabs.add"
        :aria-label="m.xlsx.sheetTabs.add"
        @click="controller.addSheet()"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 5v14M5 12h14" /></svg>
      </button>
      <button
        v-if="!isReadOnly"
        type="button"
        class="xlsx-sheettabs__remove-btn"
        data-testid="xlsx-remove-sheet"
        :disabled="!canRemoveActiveSheet"
        :title="m.xlsx.sheetTabs.remove"
        :aria-label="m.xlsx.sheetTabs.remove"
        @click="removeActiveSheet"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 7h14M9 7V4h6v3M8 10v8M12 10v8M16 10v8M7 7l1 14h8l1-14" /></svg>
      </button>
    </div>
    <button
      type="button"
      class="xlsx-sheettabs__nav-btn"
      :disabled="activeTabIndex >= controller.tabs.length - 1"
      :title="m.xlsx.sheetTabs.next"
      :aria-label="m.xlsx.sheetTabs.next"
      @click="nextTab"
    >
      <svg viewBox="0 0 24 24" aria-hidden="true"><path d="m9.5 6 6 6-6 6" /></svg>
    </button>
    <div
      v-if="renameError"
      class="xlsx-sheettabs__rename-error"
      data-testid="xlsx-sheet-rename-error"
      role="alert"
      :style="renameErrorStyle"
    >
      {{ renameError }}
    </div>
    <div
      v-if="moveError"
      class="xlsx-sheettabs__rename-error"
      data-testid="xlsx-sheet-move-error"
      role="alert"
      :style="moveErrorStyle"
    >
      {{ moveError }}
    </div>
    <div
      v-if="visibilityError"
      class="xlsx-sheettabs__rename-error"
      data-testid="xlsx-sheet-visibility-error"
      role="alert"
      :style="visibilityErrorStyle"
    >
      {{ visibilityError }}
    </div>
    <div
      v-if="tabMenu"
      class="xlsx-sheettabs__menu-backdrop"
      data-testid="xlsx-sheet-tab-menu-backdrop"
      @contextmenu.prevent="closeTabMenu"
      @pointerdown="closeTabMenu"
    />
    <div
      v-if="tabMenu"
      class="xlsx-sheettabs__menu"
      data-testid="xlsx-sheet-tab-menu"
      role="menu"
      :aria-label="m.xlsx.sheetTabs.menuLabel"
      :style="tabMenuStyle"
      @keydown.esc="closeTabMenu"
    >
      <button
        type="button"
        class="xlsx-sheettabs__menu-item"
        data-testid="xlsx-rename-sheet"
        role="menuitem"
        @click="renameFromMenu"
      >
        {{ m.xlsx.sheetTabs.rename }}
      </button>
      <button
        type="button"
        class="xlsx-sheettabs__menu-item"
        data-testid="xlsx-hide-sheet"
        role="menuitem"
        :disabled="!canHideMenuTab"
        @click="hideFromMenu"
      >
        {{ m.xlsx.sheetTabs.hide }}
      </button>
      <button
        type="button"
        class="xlsx-sheettabs__menu-item"
        data-testid="xlsx-unhide-menu"
        role="menuitem"
        aria-haspopup="menu"
        :aria-expanded="unhideMenuOpen ? 'true' : 'false'"
        :disabled="controller.hiddenSheets.length === 0"
        @click="toggleUnhideMenu"
      >
        {{ m.xlsx.sheetTabs.unhide }}
      </button>
      <div
        v-if="unhideMenuOpen"
        class="xlsx-sheettabs__submenu"
        data-testid="xlsx-unhide-sheet-menu"
        role="menu"
        :aria-label="m.xlsx.sheetTabs.unhide"
      >
        <button
          v-for="sheet in controller.hiddenSheets"
          :key="sheet.workbookSheetIndex"
          type="button"
          class="xlsx-sheettabs__menu-item xlsx-sheettabs__menu-item--submenu"
          data-testid="xlsx-unhide-sheet"
          :data-sheet-name="sheet.name"
          role="menuitem"
          @click="unhideFromMenu(sheet)"
        >
          {{ sheet.name }}
        </button>
        <span
          v-if="controller.hiddenSheets.length === 0"
          class="xlsx-sheettabs__menu-empty"
        >
          {{ m.xlsx.sheetTabs.noHiddenSheets }}
        </span>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, ref, watch, nextTick, onUnmounted, type CSSProperties } from "vue";
import type {
  XlsxHiddenSheet,
  XlsxSheetNameIssue,
  XlsxViewerController,
  XlsxWorkbookTab,
  XlsxSheetThumbnail,
  UseXlsxViewerThumbnailsOptions,
} from "@arcships/xlsx-core";
import { XLSX_MAX_SHEET_NAME_LENGTH } from "@arcships/xlsx-core";
import { useXlsxViewerThumbnails } from "../composables/useXlsxViewerThumbnails";
import { useXlsxMessages } from "../i18n";
import { nextXlsxDomId } from "./dom-ids";

const props = defineProps<{
  controller: XlsxViewerController;
  isDark?: boolean;
  readOnly?: boolean;
  /** DOM id of the worksheet panel this tab strip controls, used by `aria-controls`. */
  panelId?: string;
  thumbnailOptions?: UseXlsxViewerThumbnailsOptions;
}>();

const activeTabIndex = computed(() => props.controller.activeTabIndex);
const isReadOnly = computed(() => props.controller.isEditingLocked === true || props.controller.readOnly || props.readOnly === true);
// Renaming a protected sheet is allowed, like Excel; only read-only mode blocks it.
const canRenameSheets = computed(() => props.controller.readOnly !== true && props.readOnly !== true);
// Hiding, unhiding and renaming are worksheet management actions; sheet
// protection does not block them, matching Excel.
const canManageSheets = canRenameSheets;
const canMoveSheets = computed(() => canRenameSheets.value && props.controller.tabs.length > 1);
const m = useXlsxMessages();
const canRemoveActiveSheet = computed(() => Boolean(props.controller.activeSheet) && props.controller.sheets.length > 1);
const maxSheetNameLength = XLSX_MAX_SHEET_NAME_LENGTH;

// Accessibility state. `focusedTabIndex` is the roving tab stop; `null` means
// the strip does not own focus, so the active tab is the tab stop instead.
const stripRef = ref<HTMLElement | null>(null);
const focusedTabIndex = ref<number | null>(null);
const tabButtonRefs = ref<Map<number, HTMLElement>>(new Map());
const tabIdPrefix = nextXlsxDomId("xlsx-sheet-tab");

const { prefetchThumbnails, thumbnails: thumbnailComputedRaw } = useXlsxViewerThumbnails(
  () => props.controller,
  props.thumbnailOptions ?? { includeHeaders: true, resolution: { maxHeight: 132, maxWidth: 200 } },
);

const thumbnailComputed = computed(() => (thumbnailComputedRaw as unknown as XlsxSheetThumbnail[]));

const hoveredTabIndex = ref<number | null>(null);
const thumbnailCanvasRefs = ref<Map<number, HTMLCanvasElement | null>>(new Map());
const popupX = ref(0);
const popupY = ref(0);
let popupTimer: ReturnType<typeof setTimeout> | null = null;

// Inline rename state. The tab id, not the index, identifies the editor so a
// sheet added or removed underneath the input cannot retarget the commit.
const renamingTabId = ref<string | null>(null);
const renamingTabIndex = computed(() => {
  if (renamingTabId.value === null) return -1;
  return props.controller.tabs.findIndex((tab) => tab.id === renamingTabId.value);
});
const renameDraft = ref("");
const renameError = ref<string | null>(null);
// The message is positioned against the editor and fixed to the viewport, so
// the scrollable tab list cannot clip it.
const renameAnchor = ref<{ left: number; top: number } | null>(null);
// Vue collects element refs inside `v-for` into an array, so the single
// inline editor is read back through this helper.
const renameInputRef = ref<HTMLInputElement | HTMLInputElement[] | null>(null);
const tabMenu = ref<{ tabId: string; x: number; y: number } | null>(null);
const unhideMenuOpen = ref(false);
const visibilityError = ref<string | null>(null);
const visibilityErrorAnchor = ref<{ left: number; top: number } | null>(null);
const menuTab = computed<XlsxWorkbookTab | null>(() => {
  const menu = tabMenu.value;
  if (!menu) return null;
  return props.controller.tabs.find((tab) => tab.id === menu.tabId) ?? null;
});
const visibleSheetCount = computed(() => (
  props.controller.sheets.filter((sheet) => sheet.visibility === "visible").length
));
const canHideMenuTab = computed(() => {
  const tab = menuTab.value;
  if (!tab) return false;
  if (tab.visibility && tab.visibility !== "visible") return false;
  return resolveTabWorkbookSheetIndex(tab) !== null && visibleSheetCount.value > 1;
});

// Drag-to-reorder state. `dropSlot` is an insertion slot in `tabs` (0 through
// `tabs.length`), not a tab index, so dropping on the right half of a tab
// targets the gap after it.
const draggingTabIndex = ref<number | null>(null);
const dropSlot = ref<number | null>(null);
const moveError = ref<string | null>(null);
const moveErrorAnchor = ref<{ left: number; top: number } | null>(null);
const tabWrapperRefs = ref<Map<number, HTMLElement>>(new Map());
let dragKeydownTarget: Document | null = null;

// Prefetch tuning: one pending idle warm-up at most, for the hovered tab and
// its immediate neighbours. Sweeping the pointer across the strip cancels the
// previous task instead of queueing every tab, and no workspace is read until
// something actually points at it.
const THUMBNAIL_PREFETCH_RADIUS = 1;
type IdleHandle = number | ReturnType<typeof setTimeout>;
const idleHost = globalThis as {
  cancelIdleCallback?: (handle: number) => void;
  requestIdleCallback?: (callback: () => void, options?: { timeout: number }) => number;
};
let prefetchHandle: IdleHandle | null = null;

function cancelScheduledPrefetch(): void {
  if (prefetchHandle === null) return;
  if (typeof prefetchHandle === "number" && idleHost.cancelIdleCallback) {
    idleHost.cancelIdleCallback(prefetchHandle);
  } else {
    clearTimeout(prefetchHandle as ReturnType<typeof setTimeout>);
  }
  prefetchHandle = null;
}

function scheduleHoverPrefetch(index: number): void {
  cancelScheduledPrefetch();
  const tabCount = props.controller.tabs.length;
  const sheetIndices: number[] = [];
  for (let offset = -THUMBNAIL_PREFETCH_RADIUS; offset <= THUMBNAIL_PREFETCH_RADIUS; offset += 1) {
    const candidate = index + offset;
    if (candidate >= 0 && candidate < tabCount) sheetIndices.push(candidate);
  }
  const run = () => {
    prefetchHandle = null;
    prefetchThumbnails(sheetIndices);
  };
  prefetchHandle = typeof idleHost.requestIdleCallback === "function"
    ? idleHost.requestIdleCallback(run, { timeout: 200 })
    : setTimeout(run, 0);
}

const thumbnailPopupStyle = computed<CSSProperties>(() => ({
  left: `${popupX.value}px`,
  position: "fixed",
  top: `${popupY.value}px`,
  zIndex: 9999,
  transform: "translate(-50%, calc(-100% - 8px))",
  backgroundColor: props.isDark ? "#27272a" : "#fff",
  border: `1px solid ${props.isDark ? "#3f3f46" : "#e4e4e7"}`,
  borderRadius: "8px",
  boxShadow: "0 4px 16px rgba(0,0,0,0.12)",
  padding: "8px",
  pointerEvents: "none",
}));

function setThumbnailCanvas(index: number, el: unknown) {
  const canvas = el as HTMLCanvasElement | null;
  thumbnailCanvasRefs.value.set(index, canvas);
}

function onTabHover(index: number, event: MouseEvent) {
  if (popupTimer) clearTimeout(popupTimer);
  scheduleHoverPrefetch(index);
  const rect = (event.currentTarget as HTMLElement).getBoundingClientRect();
  popupX.value = rect.left + rect.width / 2;
  popupY.value = rect.top;
  hoveredTabIndex.value = index;

  nextTick(() => {
    const canvas = thumbnailCanvasRefs.value.get(index);
    const thumb = thumbnailComputed.value[index];
    if (canvas && thumb) {
      thumb.paint(canvas as HTMLCanvasElement);
    }
  });
}

function setTabButtonRef(index: number, el: unknown) {
  const element = el as HTMLElement | null;
  if (element) {
    tabButtonRefs.value.set(index, element);
  } else {
    tabButtonRefs.value.delete(index);
  }
}

function tabDomId(index: number): string {
  return `${tabIdPrefix}-${index}`;
}

/** Roving tabindex: one tab is in the page tab order, the one holding focus. */
function tabStopIndex(index: number): number {
  const tabCount = props.controller.tabs.length;
  const stop = focusedTabIndex.value ?? activeTabIndex.value;
  const normalized = stop >= 0 && stop < tabCount ? stop : 0;
  return index === normalized ? 0 : -1;
}

function focusTab(index: number) {
  tabButtonRefs.value.get(index)?.focus?.();
}

function activeElementIsInStrip(): boolean {
  // Only a tab that received focus arms the follow-focus rule; the prev/next
  // arrow buttons live in the same strip but must keep their own focus.
  if (focusedTabIndex.value === null) return false;
  const root = stripRef.value;
  if (!root || typeof root.contains !== "function") return true;
  const active = typeof document !== "undefined" ? document.activeElement : null;
  return !active || root.contains(active);
}

/**
 * After a switch, the roving tab stop follows the newly active tab. Focus only
 * moves when the strip already owns it — Ctrl+PageDown from the grid must not
 * pull keyboard focus out of the sheet.
 */
function syncFocusWithActiveTab(index: number) {
  if (!activeElementIsInStrip()) return;
  focusedTabIndex.value = index;
  focusTab(index);
}

/** Scroll only the strip: `nearest` on both axes keeps the page and grid still. */
function scrollTabIntoView(index: number) {
  if (index < 0) return;
  tabWrapperRefs.value.get(index)?.scrollIntoView?.({ block: "nearest", inline: "nearest" });
}

function onTabFocus(index: number) {
  focusedTabIndex.value = index;
  // Keyboard users reach the same preview; warm it without opening the popup.
  scheduleHoverPrefetch(index);
}

function onTabFocusOut(index: number) {
  // A focusin on the next tab arrives right after this focusout and restores
  // the roving stop; focus moving to the strip arrows leaves it cleared.
  if (focusedTabIndex.value === index) focusedTabIndex.value = null;
}

function onTabLeave() {
  cancelScheduledPrefetch();
  popupTimer = setTimeout(() => {
    hoveredTabIndex.value = null;
  }, 150);
}

function onTabClick(index: number) {
  props.controller.setActiveTabIndex(index);
}

/** Hidden tabs are labelled as such; normal tabs keep the default tooltip. */
function tabTitle(tab: XlsxWorkbookTab) {
  return tab.kind === "sheet" && tab.visibility === "hidden"
    ? m.value.xlsx.sheetTabs.hiddenTab
    : undefined;
}

function onTabDblClick(tab: XlsxWorkbookTab, index: number) {
  startRename(tab, index);
}

function setTabWrapperRef(index: number, el: unknown) {
  const element = el as HTMLElement | null;
  if (element) {
    tabWrapperRefs.value.set(index, element);
  } else {
    tabWrapperRefs.value.delete(index);
  }
}

function removeDragKeydownListener() {
  if (!dragKeydownTarget) return;
  dragKeydownTarget.removeEventListener("keydown", onTabDragKeydown, true);
  dragKeydownTarget = null;
}

function addDragKeydownListener() {
  removeDragKeydownListener();
  if (typeof document === "undefined" || typeof document.addEventListener !== "function") return;
  dragKeydownTarget = document;
  dragKeydownTarget.addEventListener("keydown", onTabDragKeydown, true);
}

/** Esc cancels a drag in progress, whether the key lands on the strip or not. */
function onTabDragKeydown(event: KeyboardEvent) {
  if (event.key !== "Escape" || draggingTabIndex.value === null) return;
  event.preventDefault?.();
  event.stopPropagation?.();
  cancelDrag();
}

function cancelDrag() {
  draggingTabIndex.value = null;
  dropSlot.value = null;
  removeDragKeydownListener();
}

function showMoveError(message: string, index: number) {
  moveError.value = message;
  const rect = tabWrapperRefs.value.get(index)?.getBoundingClientRect?.();
  moveErrorAnchor.value = rect ? { left: rect.left, top: rect.top } : null;
}

function onTabDragStart(index: number, event: DragEvent) {
  if (!canMoveSheets.value || renamingTabIndex.value === index) {
    event.preventDefault?.();
    return;
  }
  closeTabMenu();
  cancelDrag();
  moveError.value = null;
  moveErrorAnchor.value = null;
  draggingTabIndex.value = index;
  dropSlot.value = null;
  if (event.dataTransfer) {
    event.dataTransfer.effectAllowed = "move";
    try {
      event.dataTransfer.setData("text/plain", props.controller.tabs[index]?.name ?? "");
    } catch {
      // Some platform drag implementations refuse custom payloads; the move
      // itself does not depend on them.
    }
  }
  addDragKeydownListener();
}

function onTabDragOver(index: number, event: DragEvent) {
  if (draggingTabIndex.value === null) return;
  event.preventDefault?.();
  if (event.dataTransfer) event.dataTransfer.dropEffect = "move";
  const rect = (event.currentTarget as HTMLElement | null)?.getBoundingClientRect?.();
  const isAfter = rect
    ? event.clientX >= rect.left + rect.width / 2
    : false;
  dropSlot.value = isAfter ? index + 1 : index;
}

function onTabDrop(index: number, event: DragEvent) {
  if (draggingTabIndex.value === null) return;
  event.preventDefault?.();
  const fromIndex = draggingTabIndex.value;
  const insertionSlot = dropSlot.value ?? index;
  const toIndex = insertionSlot > fromIndex ? insertionSlot - 1 : insertionSlot;
  cancelDrag();
  if (toIndex === fromIndex) return;
  if (!props.controller.moveSheet(fromIndex, toIndex)) {
    showMoveError(m.value.xlsx.sheetTabs.moveFailed, toIndex);
    return;
  }
  // The strip re-renders at the new index; keep the moved tab visible once the
  // keyed wrappers have moved with it.
  void nextTick(() => {
    tabWrapperRefs.value.get(toIndex)?.scrollIntoView?.({ block: "nearest", inline: "nearest" });
  });
}

function renameErrorMessage(issue: XlsxSheetNameIssue) {
  const messages = m.value.xlsx;
  switch (issue) {
    case "empty":
      return messages.sheetTabs.renameEmpty;
    case "too-long":
      return messages.sheetTabs.renameTooLong;
    case "duplicate":
      return messages.sheetTabs.renameDuplicate;
    case "apostrophe":
      return messages.sheetTabs.renameApostrophe;
    case "invalid-characters":
    default:
      return messages.diagnostics.sheetNameInvalid;
  }
}

function closeRename() {
  renamingTabId.value = null;
  renameDraft.value = "";
  renameError.value = null;
  renameAnchor.value = null;
}

function refreshRenameAnchor() {
  const target = renameInputRef.value;
  const input = Array.isArray(target) ? target[0] : target;
  const rect = input?.getBoundingClientRect?.();
  renameAnchor.value = rect ? { left: rect.left, top: rect.top } : null;
}

function showRenameError(message: string) {
  renameError.value = message;
  refreshRenameAnchor();
}

function startRename(tab: XlsxWorkbookTab, index: number) {
  if (!canRenameSheets.value || tab.kind !== "sheet" || typeof tab.sheetIndex !== "number") {
    return;
  }
  closeTabMenu();
  props.controller.setActiveTabIndex(index);
  renamingTabId.value = tab.id;
  renameDraft.value = tab.name;
  renameError.value = null;
  void nextTick(() => {
    const target = renameInputRef.value;
    const input = Array.isArray(target) ? target[0] : target;
    input?.focus();
    input?.select();
    refreshRenameAnchor();
  });
}

function commitRename() {
  const index = renamingTabIndex.value;
  const tab = index >= 0 ? props.controller.tabs[index] : null;
  if (!tab || tab.kind !== "sheet" || typeof tab.sheetIndex !== "number") {
    closeRename();
    return;
  }
  const issue = props.controller.getSheetNameError(renameDraft.value, tab.sheetIndex);
  if (issue) {
    showRenameError(renameErrorMessage(issue));
    return;
  }
  if (!props.controller.renameSheet(tab.sheetIndex, renameDraft.value)) {
    showRenameError(m.value.xlsx.sheetTabs.renameFailed);
    return;
  }
  closeRename();
}

function cancelRename() {
  closeRename();
}

function onRenameInput(event: Event) {
  renameDraft.value = (event.target as HTMLInputElement | null)?.value ?? "";
  renameError.value = null;
}

function onRenameBlur() {
  // Blur commits, matching Excel; a rejected name keeps the editor open with
  // the reason visible instead of dropping the edit silently.
  if (renamingTabId.value === null) return;
  commitRename();
}

function onTabContextMenu(index: number, event: MouseEvent) {
  const tab = props.controller.tabs[index];
  if (!tab || tab.kind !== "sheet" || !canRenameSheets.value) {
    closeTabMenu();
    return;
  }
  event.preventDefault();
  props.controller.setActiveTabIndex(index);
  unhideMenuOpen.value = false;
  visibilityError.value = null;
  tabMenu.value = { tabId: tab.id, x: event.clientX, y: event.clientY };
}

function onTabKeydown(index: number, tab: XlsxWorkbookTab, event: KeyboardEvent) {
  // Keyboard users reach the same menu with the ContextMenu key or Shift+F10.
  if (event.key === "ContextMenu" || (event.shiftKey && event.key === "F10")) {
    if (!canRenameSheets.value || tab.kind !== "sheet") {
      return;
    }
    event.preventDefault();
    const rect = (event.currentTarget as HTMLElement | null)?.getBoundingClientRect?.();
    props.controller.setActiveTabIndex(index);
    unhideMenuOpen.value = false;
    visibilityError.value = null;
    tabMenu.value = { tabId: tab.id, x: rect?.left ?? 0, y: rect?.bottom ?? 0 };
    return;
  }

  // Arrow keys move the roving focus without changing the sheet; Enter/Space
  // commit the choice. Modifier combinations stay with the browser.
  if (event.altKey || event.ctrlKey || event.metaKey) return;
  const tabCount = props.controller.tabs.length;
  if (tabCount === 0) return;

  let targetIndex: number | null = null;
  switch (event.key) {
    case "ArrowLeft":
      targetIndex = index === 0 ? tabCount - 1 : index - 1;
      break;
    case "ArrowRight":
      targetIndex = index === tabCount - 1 ? 0 : index + 1;
      break;
    case "Home":
      targetIndex = 0;
      break;
    case "End":
      targetIndex = tabCount - 1;
      break;
    case "Enter":
    case " ":
    case "Spacebar":
    case "Space":
      event.preventDefault();
      focusedTabIndex.value = index;
      props.controller.setActiveTabIndex(index);
      focusTab(index);
      return;
    default:
      return;
  }

  event.preventDefault();
  focusedTabIndex.value = targetIndex;
  focusTab(targetIndex);
}

function closeTabMenu() {
  tabMenu.value = null;
  unhideMenuOpen.value = false;
}

/** Workbook sheet index for a tab, or null when the tab owns no worksheet. */
function resolveTabWorkbookSheetIndex(tab: XlsxWorkbookTab): number | null {
  if (typeof tab.workbookSheetIndex === "number") {
    return tab.workbookSheetIndex;
  }
  if (typeof tab.sheetIndex === "number") {
    return props.controller.sheets[tab.sheetIndex]?.workbookSheetIndex ?? null;
  }
  return null;
}

function showVisibilityError(message: string, anchor: { left: number; top: number } | null) {
  visibilityError.value = message;
  visibilityErrorAnchor.value = anchor;
}

function menuAnchor(): { left: number; top: number } | null {
  const menu = tabMenu.value;
  return menu ? { left: menu.x, top: menu.y } : null;
}

function hideFromMenu() {
  const menu = tabMenu.value;
  const index = menu ? props.controller.tabs.findIndex((tab) => tab.id === menu.tabId) : -1;
  const tab = index >= 0 ? props.controller.tabs[index] : null;
  const anchor = menuAnchor();
  closeTabMenu();
  if (!tab) return;
  const workbookSheetIndex = resolveTabWorkbookSheetIndex(tab);
  if (
    workbookSheetIndex === null
    || !props.controller.setSheetVisibility(workbookSheetIndex, "hidden")
  ) {
    showVisibilityError(m.value.xlsx.sheetTabs.hideFailed, anchor);
  }
}

function toggleUnhideMenu() {
  unhideMenuOpen.value = !unhideMenuOpen.value;
}

function unhideFromMenu(sheet: XlsxHiddenSheet) {
  const anchor = menuAnchor();
  closeTabMenu();
  if (!props.controller.setSheetVisibility(sheet.workbookSheetIndex, "visible")) {
    showVisibilityError(m.value.xlsx.sheetTabs.unhideFailed, anchor);
  }
}

function unhideTabFromStrip(tab: XlsxWorkbookTab) {
  const workbookSheetIndex = resolveTabWorkbookSheetIndex(tab);
  const index = props.controller.tabs.indexOf(tab);
  const rect = index >= 0 ? tabWrapperRefs.value.get(index)?.getBoundingClientRect?.() : null;
  if (
    workbookSheetIndex === null
    || !props.controller.setSheetVisibility(workbookSheetIndex, "visible")
  ) {
    showVisibilityError(
      m.value.xlsx.sheetTabs.unhideFailed,
      rect ? { left: rect.left, top: rect.top } : null,
    );
  }
}

function renameFromMenu() {
  const menu = tabMenu.value;
  closeTabMenu();
  if (!menu) return;
  const index = props.controller.tabs.findIndex((tab) => tab.id === menu.tabId);
  const tab = index >= 0 ? props.controller.tabs[index] : null;
  if (tab) startRename(tab, index);
}

function removeActiveSheet() {
  if (!canRemoveActiveSheet.value || isReadOnly.value) return;
  props.controller.removeActiveSheet();
}

function prevTab() {
  if (activeTabIndex.value > 0) {
    props.controller.setActiveTabIndex(activeTabIndex.value - 1);
  }
}

function nextTab() {
  if (activeTabIndex.value < props.controller.tabs.length - 1) {
    props.controller.setActiveTabIndex(activeTabIndex.value + 1);
  }
}

const tabsStyle = computed<CSSProperties>(() => ({
  alignItems: "center",
  backgroundColor: props.isDark ? "#27272a" : "#f4f4f5",
  borderTop: `1px solid ${props.isDark ? "#3f3f46" : "#e4e4e7"}`,
  display: "flex",
  flexShrink: "0",
  gap: "2px",
  overflow: "hidden",
  padding: "0 4px",
}));

function tabStyle(index: number): CSSProperties {
  const isActive = index === activeTabIndex.value;
  return {
    backgroundColor: isActive
      ? props.isDark ? "#3f3f46" : "#ffffff"
      : "transparent",
    border: isActive
      ? `1px solid ${props.isDark ? "#52525b" : "#d4d4d8"}`
      : "1px solid transparent",
    borderBottom: "none",
    borderRadius: "4px 4px 0 0",
    color: isActive
      ? "inherit"
      : props.isDark ? "#a1a1aa" : "#71717a",
    cursor: "pointer",
    fontSize: "12px",
    maxWidth: "120px",
    padding: "4px 10px",
    whiteSpace: "nowrap",
  };
}

function renameInputStyle(index: number): CSSProperties {
  return {
    ...tabStyle(index),
    cursor: "text",
    fontFamily: "inherit",
    minWidth: "72px",
    outline: `2px solid ${props.isDark ? "#60a5fa" : "#2563eb"}`,
    outlineOffset: "-1px",
  };
}

const renameErrorStyle = computed<CSSProperties>(() => ({
  backgroundColor: "#b91c1c",
  borderRadius: "4px",
  color: "#ffffff",
  fontSize: "11px",
  left: `${renameAnchor.value?.left ?? 0}px`,
  maxWidth: "260px",
  padding: "4px 8px",
  position: "fixed",
  top: `${renameAnchor.value?.top ?? 0}px`,
  transform: "translateY(calc(-100% - 8px))",
  whiteSpace: "normal",
  zIndex: 9998,
}));

const moveErrorStyle = computed<CSSProperties>(() => ({
  backgroundColor: "#b91c1c",
  borderRadius: "4px",
  color: "#ffffff",
  fontSize: "11px",
  left: `${moveErrorAnchor.value?.left ?? 0}px`,
  maxWidth: "260px",
  padding: "4px 8px",
  position: "fixed",
  top: `${moveErrorAnchor.value?.top ?? 0}px`,
  transform: "translateY(calc(-100% - 8px))",
  whiteSpace: "normal",
  zIndex: 9998,
}));

const visibilityErrorStyle = computed<CSSProperties>(() => ({
  backgroundColor: "#b91c1c",
  borderRadius: "4px",
  color: "#ffffff",
  fontSize: "11px",
  left: `${visibilityErrorAnchor.value?.left ?? 0}px`,
  maxWidth: "260px",
  padding: "4px 8px",
  position: "fixed",
  top: `${visibilityErrorAnchor.value?.top ?? 0}px`,
  transform: "translateY(calc(-100% - 8px))",
  whiteSpace: "normal",
  zIndex: 9998,
}));

const tabMenuStyle = computed<CSSProperties>(() => ({
  backgroundColor: props.isDark ? "#27272a" : "#ffffff",
  border: `1px solid ${props.isDark ? "#3f3f46" : "#e4e4e7"}`,
  borderRadius: "8px",
  boxShadow: "0 4px 16px rgba(0,0,0,0.16)",
  left: `${tabMenu.value?.x ?? 0}px`,
  padding: "4px",
  position: "fixed",
  top: `${tabMenu.value?.y ?? 0}px`,
  zIndex: 9999,
}));

// A switch can come from the strip, the prev/next arrows, the controller, or a
// restored workbook. Keep the active tab visible inside the strip and keep the
// roving tab stop on it without stealing focus from the grid.
watch(activeTabIndex, (index) => {
  void nextTick(() => {
    scrollTabIntoView(index);
    syncFocusWithActiveTab(index);
  });
}, { immediate: true });

// A sheet added or removed under the editor must not leave it pointing at a
// different tab, and a rebuilt tab strip closes a stale context menu.
watch(renamingTabIndex, (index) => {
  if (index < 0 && renamingTabId.value !== null) closeRename();
});
watch(() => props.controller.tabs, () => {
  closeTabMenu();
  cancelDrag();
});
// Losing edit rights (the host flips readOnly mid-edit) must not leave a live
// editor that can only fail on commit.
watch(canRenameSheets, (allowed) => {
  if (!allowed) {
    closeRename();
    closeTabMenu();
    cancelDrag();
  }
});

onUnmounted(() => {
  cancelScheduledPrefetch();
  cancelDrag();
  if (popupTimer) clearTimeout(popupTimer);
});
</script>

<style scoped>
.xlsx-sheettabs {
  user-select: none;
}

.xlsx-sheettabs__list {
  display: flex;
  flex: 1;
  gap: 2px;
  overflow-x: auto;
  scroll-behavior: smooth;
}

.xlsx-sheettabs__tab-wrapper {
  position: relative;
  flex-shrink: 0;
}

.xlsx-sheettabs__tab-wrapper--draggable {
  cursor: grab;
}

.xlsx-sheettabs__tab-wrapper--dragging {
  opacity: 0.45;
}

.xlsx-sheettabs__drop-indicator {
  background: #2563eb;
  border-radius: 1px;
  bottom: 2px;
  left: -2px;
  pointer-events: none;
  position: absolute;
  top: 2px;
  width: 2px;
  z-index: 2;
}

.xlsx-sheettabs__drop-indicator--after {
  left: auto;
  right: -2px;
}

.xlsx-sheettabs__tab {
  align-items: center;
  background: transparent;
  border: none;
  display: flex;
  font-family: inherit;
  gap: 4px;
  outline: none;
  overflow: hidden;
}

.xlsx-sheettabs__tab:hover {
  background: rgba(128, 128, 128, 0.1);
}

.xlsx-sheettabs__tab-name {
  overflow: hidden;
  text-overflow: ellipsis;
}

/* Hidden tabs stay in the strip only when the host asks for them; they are
   visibly different from a normal tab so "hide" is not mistaken for "delete". */
.xlsx-sheettabs__tab--hidden {
  font-style: italic;
  opacity: 0.55;
}

.xlsx-sheettabs__unhide-btn {
  align-items: center;
  background: transparent;
  border: 1px solid transparent;
  border-radius: 4px;
  color: inherit;
  cursor: pointer;
  display: flex;
  height: 22px;
  justify-content: center;
  opacity: 0.7;
  padding: 0;
  width: 22px;
}

.xlsx-sheettabs__unhide-btn:hover {
  background: rgba(128, 128, 128, 0.15);
  opacity: 1;
}

.xlsx-sheettabs__unhide-btn svg {
  fill: none;
  height: 14px;
  stroke: currentColor;
  stroke-linecap: round;
  stroke-linejoin: round;
  stroke-width: 1.8;
  width: 14px;
}

.xlsx-sheettabs__submenu {
  border-top: 1px solid rgba(128, 128, 128, 0.25);
  margin-top: 4px;
  padding-top: 4px;
}

.xlsx-sheettabs__menu-item--submenu {
  padding-left: 20px;
}

.xlsx-sheettabs__menu-empty {
  display: block;
  font-size: 11px;
  opacity: 0.6;
  padding: 4px 10px 4px 20px;
}

.xlsx-sheettabs__rename-input {
  box-sizing: border-box;
}

.xlsx-sheettabs__rename-error {
  background: #b91c1c;
  border-radius: 4px;
  color: #ffffff;
  font-size: 11px;
  max-width: 260px;
  padding: 4px 8px;
  white-space: normal;
  z-index: 9998;
}

.xlsx-sheettabs__menu-backdrop {
  inset: 0;
  position: fixed;
  z-index: 9998;
}

.xlsx-sheettabs__menu {
  min-width: 140px;
}

.xlsx-sheettabs__menu-item {
  background: transparent;
  border: none;
  border-radius: 4px;
  color: inherit;
  cursor: pointer;
  display: block;
  font-family: inherit;
  font-size: 12px;
  padding: 6px 10px;
  text-align: left;
  width: 100%;
}

.xlsx-sheettabs__menu-item:hover {
  background: rgba(128, 128, 128, 0.15);
}

.xlsx-sheettabs__menu-item:disabled {
  cursor: default;
  opacity: 0.4;
}

.xlsx-sheettabs__thumbnail {
  display: flex;
  flex-direction: column;
  gap: 4px;
  max-width: 212px;
}

.xlsx-sheettabs__thumbnail canvas {
  border-radius: 6px;
  display: block;
  height: auto;
  width: 100%;
}

.xlsx-sheettabs__add-btn,
.xlsx-sheettabs__remove-btn {
  align-items: center;
  background: transparent;
  border: 1px solid transparent;
  border-radius: 4px;
  color: inherit;
  cursor: pointer;
  display: flex;
  font-family: inherit;
  font-size: 16px;
  height: 28px;
  justify-content: center;
  width: 28px;
}

.xlsx-sheettabs__add-btn:hover,
.xlsx-sheettabs__remove-btn:hover:not(:disabled) {
  background: rgba(128, 128, 128, 0.15);
}

.xlsx-sheettabs__remove-btn:disabled {
  cursor: default;
  opacity: 0.3;
}

.xlsx-sheettabs__nav-btn {
  align-items: center;
  background: transparent;
  border: none;
  border-radius: 4px;
  color: inherit;
  cursor: pointer;
  display: flex;
  flex-shrink: 0;
  font-size: 16px;
  height: 28px;
  justify-content: center;
  width: 24px;
}

.xlsx-sheettabs__nav-btn:hover:not(:disabled) {
  background: rgba(128, 128, 128, 0.15);
}

.xlsx-sheettabs__nav-btn:disabled {
  cursor: default;
  opacity: 0.3;
}

.xlsx-sheettabs__nav-btn svg,
.xlsx-sheettabs__add-btn svg,
.xlsx-sheettabs__remove-btn svg {
  fill: none;
  height: 16px;
  stroke: currentColor;
  stroke-linecap: round;
  stroke-linejoin: round;
  stroke-width: 1.8;
  width: 16px;
}

@media (pointer: coarse) {
  .xlsx-sheettabs button {
    min-height: 44px;
    min-width: 44px;
  }
}
</style>

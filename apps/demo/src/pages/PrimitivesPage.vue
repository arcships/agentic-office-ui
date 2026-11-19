<script setup lang="ts">
import { ref, defineComponent } from "vue"
import {
  OfficeBadge,
  OfficeButton,
  OfficeCard,
  OfficeContextMenu,
  OfficeDialog,
  OfficeDropdownMenu,
  OfficeIconButton,
  OfficeInput,
  OfficePopover,
  OfficeScrollArea,
  OfficeSelect,
  OfficeSeparator,
  OfficeSheet,
  OfficeSkeleton,
  OfficeSpinner,
  OfficeTabs,
  OfficeTextarea,
  OfficeThemeProvider,
  OfficeToastProvider,
  OfficeToggle,
  OfficeTooltip,
  useOfficeToast,
  type OfficeMenuEntry,
  type OfficeSelectEntry,
  type OfficeTab,
  type OfficeTheme,
  type OfficeToastApi,
} from "@arcships/vue-ui"

const theme = ref<OfficeTheme>("light")
const density = ref<"comfortable" | "compact">("comfortable")
const inputValue = ref("Quarterly report")
const textareaValue = ref("Notes for the review…")
const selectValue = ref("docx")
const selectValues = ref("pdf")
const tabValue = ref("preview")
const pressed = ref(true)
const dialogOpen = ref(false)
const sheetOpen = ref(false)
const dropdownOpen = ref(false)
const lastAction = ref("—")

const tabs: OfficeTab[] = [
  { value: "preview", label: "Preview" },
  { value: "properties", label: "Properties" },
  { value: "history", label: "History", icon: "clock" },
]

const selectOptions: OfficeSelectEntry[] = [
  {
    type: "group",
    label: "Documents",
    options: [
      { value: "docx", label: "Word document", icon: "file-text" },
      { value: "xlsx", label: "Spreadsheet", icon: "table" },
    ],
  },
  { type: "separator" },
  {
    type: "group",
    label: "Other",
    options: [
      { value: "pdf", label: "PDF", icon: "file" },
      { value: "pptx", label: "Presentation", icon: "image", disabled: true },
    ],
  },
]

const menuItems: OfficeMenuEntry[] = [
  { label: "Copy", icon: "copy", shortcut: "⌘C" },
  { label: "Duplicate", icon: "plus", shortcut: "⌘D" },
  { type: "separator" },
  {
    type: "group",
    label: "View",
    items: [
      { label: "Show grid", type: "checkbox", checked: true },
      {
        type: "sub",
        label: "Zoom",
        icon: "search",
        items: [
          { label: "Fit width", id: "fit-width" },
          { label: "Actual size", id: "actual-size" },
        ],
      },
    ],
  },
  { label: "Delete", icon: "trash-2", variant: "destructive", disabled: true },
]

function handleMenuSelect(entry: OfficeMenuEntry) {
  if (entry.type === "separator") return
  lastAction.value = entry.label ?? entry.type ?? "item"
  if (entry.type === "checkbox") entry.checked = !entry.checked
}

const toastState = ref<OfficeToastApi | null>(null)

/** Runs under `OfficeToastProvider` so the toast manager can be injected. */
const ToastBinder = defineComponent({
  name: "ToastBinder",
  setup() {
    toastState.value = useOfficeToast()
    return () => null
  },
})

function showSuccess() {
  toastState.value?.success("Saved", "Document stored in the workspace.")
}
function showError() {
  toastState.value?.error("Export failed", "The worker returned an error.")
}
function showLoading() {
  const id = toastState.value?.loading("Preparing export", "Rendering pages…")
  if (id) {
    window.setTimeout(() => {
      toastState.value?.update(id, {
        title: "Export ready",
        description: "The PDF was downloaded.",
        status: "success",
        duration: 4000,
      })
    }, 900)
  }
}
</script>

<template>
  <OfficeThemeProvider :theme="theme" :density="density">
    <OfficeToastProvider :duration="5000">
      <ToastBinder />
      <section class="primitives-page" data-testid="primitives-gallery">
        <header class="primitives-header">
          <div>
            <p class="eyebrow">@arcships/vue-ui</p>
            <h2>Primitive gallery</h2>
            <p class="lead">
              Token-driven Reka UI primitives. Hover and keyboard-focus the
              controls to inspect the interaction states.
            </p>
          </div>
          <div class="theme-switch">
            <label>
              Theme
              <select v-model="theme" data-testid="gallery-theme">
                <option value="light">Light</option>
                <option value="dark">Dark</option>
                <option value="system">System</option>
              </select>
            </label>
            <label>
              Density
              <select v-model="density" data-testid="gallery-density">
                <option value="comfortable">Comfortable</option>
                <option value="compact">Compact</option>
              </select>
            </label>
          </div>
        </header>

        <section class="demo-section" data-testid="section-button">
          <h3>Button / IconButton</h3>
          <div class="row" data-testid="button-variants">
            <OfficeButton variant="default" data-testid="button-default">Default</OfficeButton>
            <OfficeButton variant="secondary">Secondary</OfficeButton>
            <OfficeButton variant="ghost">Ghost</OfficeButton>
            <OfficeButton variant="outline">Outline</OfficeButton>
            <OfficeButton variant="destructive">Destructive</OfficeButton>
            <OfficeButton variant="link">Link</OfficeButton>
          </div>
          <div class="row" data-testid="button-sizes">
            <OfficeButton size="xs">Extra small</OfficeButton>
            <OfficeButton size="sm">Small</OfficeButton>
            <OfficeButton size="default">Default</OfficeButton>
            <OfficeButton size="lg">Large</OfficeButton>
          </div>
          <div class="row" data-testid="button-states">
            <OfficeButton disabled data-testid="button-disabled">Disabled</OfficeButton>
            <OfficeButton loading data-testid="button-loading">Loading</OfficeButton>
            <OfficeButton as-child>
              <a href="#primitives-gallery" data-testid="button-as-child">Link slot</a>
            </OfficeButton>
            <OfficeIconButton label="More actions" icon="more-horizontal" />
            <OfficeIconButton label="Search" icon="search" variant="outline" size="icon-sm" />
            <OfficeIconButton label="Settings" icon="settings" active />
          </div>
        </section>

        <section class="demo-section" data-testid="section-input">
          <h3>Input / Textarea</h3>
          <div class="grid">
            <OfficeInput
              v-model="inputValue"
              placeholder="Document title"
              data-testid="input-default"
            />
            <OfficeInput
              v-model="inputValue"
              size="sm"
              placeholder="Small input"
            />
            <OfficeInput
              :model-value="'Invalid value'"
              invalid
              data-testid="input-invalid"
            />
            <OfficeInput :model-value="'Readonly'" readonly />
            <OfficeInput :model-value="'Disabled'" disabled data-testid="input-disabled" />
            <OfficeTextarea
              v-model="textareaValue"
              placeholder="Add a note"
              data-testid="textarea-default"
            />
          </div>
        </section>

        <section class="demo-section" data-testid="section-select">
          <h3>Select</h3>
          <div class="grid">
            <OfficeSelect
              id="gallery-select"
              v-model="selectValue"
              :options="selectOptions"
              placeholder="Choose a format"
              data-testid="select-default"
            />
            <OfficeSelect
              v-model="selectValues"
              :options="selectOptions"
              size="sm"
            />
            <OfficeSelect
              :model-value="'docx'"
              :options="selectOptions"
              disabled
              data-testid="select-disabled"
            />
            <OfficeSelect
              :model-value="''"
              :options="selectOptions"
              invalid
              placeholder="Invalid"
            />
          </div>
        </section>

        <section class="demo-section" data-testid="section-menu">
          <h3>DropdownMenu / ContextMenu</h3>
          <div class="row">
            <OfficeDropdownMenu
              v-model:open="dropdownOpen"
              :items="menuItems"
              label="Open menu"
              data-testid="dropdown-menu"
              @select="handleMenuSelect"
            />
            <OfficeDropdownMenu
              :items="menuItems.slice(0, 2)"
              label="Disabled trigger"
              disabled
            />
            <OfficeContextMenu :items="menuItems" @select="handleMenuSelect">
              <div class="context-target" data-testid="context-menu-target">
                Right-click this area
              </div>
            </OfficeContextMenu>
            <span class="action-log" data-testid="action-log">
              Last action: {{ lastAction }}
            </span>
          </div>
        </section>

        <section class="demo-section" data-testid="section-overlay">
          <h3>Dialog / Sheet / Popover / Tooltip</h3>
          <div class="row">
            <OfficeDialog
              v-model:open="dialogOpen"
              title="Rename document"
              description="Choose a new name for the shared document."
              data-testid="dialog"
            >
              <template #trigger>
                <OfficeButton data-testid="dialog-trigger">Open dialog</OfficeButton>
              </template>
              <OfficeInput :model-value="inputValue" aria-label="Document name" />
              <template #footer>
                <OfficeButton variant="ghost" @click="dialogOpen = false">Cancel</OfficeButton>
                <OfficeButton @click="dialogOpen = false">Save</OfficeButton>
              </template>
            </OfficeDialog>

            <OfficeSheet
              v-model:open="sheetOpen"
              title="Document details"
              description="Metadata shared with the workspace."
              side="right"
              data-testid="sheet"
            >
              <template #trigger>
                <OfficeButton variant="outline" data-testid="sheet-trigger">
                  Open sheet
                </OfficeButton>
              </template>
              <OfficeCard title="Quarterly report.docx" description="Updated just now">
                <p class="muted">Owner: Alex · 42 pages</p>
              </OfficeCard>
            </OfficeSheet>

            <OfficePopover label="Open popover" data-testid="popover">
              <p class="muted">
                Popovers use the same collision-aware positioning as menus.
              </p>
            </OfficePopover>

            <OfficeTooltip content="Copy to clipboard">
              <OfficeIconButton label="Copy to clipboard" icon="copy" data-testid="tooltip-trigger" />
            </OfficeTooltip>
          </div>
        </section>

        <section class="demo-section" data-testid="section-navigation">
          <h3>Tabs / Toggle / Badge</h3>
          <OfficeTabs v-model="tabValue" :tabs="tabs" data-testid="tabs">
            <template #panel-preview>
              <p class="muted">The document preview is rendered from the worker output.</p>
            </template>
            <template #panel-properties>
              <p class="muted">Title, author and revision metadata.</p>
            </template>
            <template #panel-history>
              <p class="muted">Recent edits and comments.</p>
            </template>
          </OfficeTabs>
          <div class="row row--spaced">
            <OfficeToggle v-model="pressed" label="Show grid" data-testid="toggle">
              Show grid
            </OfficeToggle>
            <OfficeBadge variant="default">Default</OfficeBadge>
            <OfficeBadge variant="secondary">Draft</OfficeBadge>
            <OfficeBadge variant="outline">Outline</OfficeBadge>
            <OfficeBadge variant="destructive">Blocked</OfficeBadge>
          </div>
        </section>

        <section class="demo-section" data-testid="section-display">
          <h3>Card / Separator / Skeleton / Spinner / ScrollArea</h3>
          <div class="display-grid">
            <OfficeCard
              title="Quarterly report"
              description="Shared with the finance team"
              data-testid="card"
            >
              <p class="muted">42 pages · 1.2 MB</p>
              <template #footer>
                <OfficeButton size="sm" variant="outline">Open</OfficeButton>
              </template>
            </OfficeCard>
            <div class="stack">
              <span class="muted">Above</span>
              <OfficeSeparator data-testid="separator" />
              <span class="muted">Below</span>
              <div class="row">
                <OfficeSpinner data-testid="spinner" />
                <OfficeSkeleton class="skeleton-line" data-testid="skeleton" />
              </div>
            </div>
            <OfficeScrollArea
              orientation="both"
              class="scroll-demo"
              data-testid="scroll-area"
            >
              <div class="scroll-content">
                <p v-for="index in 12" :key="index" class="muted">
                  Scrollable row {{ index }}
                </p>
              </div>
            </OfficeScrollArea>
          </div>
        </section>

        <section class="demo-section" data-testid="section-toast">
          <h3>Toast</h3>
          <div class="row">
            <OfficeButton variant="outline" @click="showSuccess">Show success</OfficeButton>
            <OfficeButton variant="outline" @click="showError">Show error</OfficeButton>
            <OfficeButton variant="outline" @click="showLoading">Show loading → success</OfficeButton>
          </div>
        </section>
      </section>
    </OfficeToastProvider>
  </OfficeThemeProvider>
</template>

<style scoped>
.primitives-page {
  display: flex;
  flex-direction: column;
  gap: 24px;
  padding: 24px;
  min-width: 0;
  background: var(--office-background);
  color: var(--office-foreground);
}

.primitives-header {
  display: flex;
  flex-wrap: wrap;
  align-items: flex-end;
  justify-content: space-between;
  gap: 16px;
}

.eyebrow {
  margin: 0 0 4px;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
  letter-spacing: 0.08em;
  text-transform: uppercase;
}

h2 {
  margin: 0;
  font-size: var(--office-text-xl);
  font-weight: 600;
}

.lead {
  margin: 6px 0 0;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-sm);
}

.theme-switch {
  display: flex;
  gap: 12px;
}

.theme-switch label {
  display: flex;
  flex-direction: column;
  gap: 4px;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
}

.demo-section {
  display: flex;
  flex-direction: column;
  gap: 14px;
  padding: 20px;
  border: 1px solid var(--office-border);
  border-radius: var(--office-radius-2xl);
  background: var(--office-card);
}

.demo-section h3 {
  margin: 0;
  font-size: var(--office-text-sm);
  font-weight: 600;
}

.row {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 10px;
}

.row--spaced {
  margin-top: 12px;
}

.grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 12px;
}

.display-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
  gap: 16px;
  align-items: start;
}

.stack {
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.context-target {
  display: inline-flex;
  align-items: center;
  min-height: 36px;
  padding: 0 12px;
  border: 1px dashed var(--office-input);
  border-radius: var(--office-radius-lg);
  color: var(--office-muted-foreground);
  font-size: var(--office-text-sm);
  user-select: none;
}

.action-log {
  color: var(--office-muted-foreground);
  font-size: var(--office-text-xs);
}

.muted {
  margin: 0;
  color: var(--office-muted-foreground);
  font-size: var(--office-text-sm);
}

.skeleton-line {
  inline-size: 160px;
  block-size: 14px;
}

.scroll-demo {
  block-size: 180px;
  border: 1px solid var(--office-border);
  border-radius: var(--office-radius-lg);
}

.scroll-content {
  display: flex;
  flex-direction: column;
  gap: 8px;
  padding: 12px;
  min-inline-size: 320px;
}

@media (max-width: 640px) {
  .primitives-page {
    padding: 12px;
  }
}
</style>

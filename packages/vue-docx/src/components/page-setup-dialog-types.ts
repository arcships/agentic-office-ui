// Shared prop/event types for the DOCX page-setup dialog.
//
// They live outside the SFC because `<script setup>` cannot contain ES module
// exports, and hosts that render the dialog themselves need the shapes.

import type {
  DocxPageMargins,
  DocxPageMarginsUpdate,
  DocxPageSize,
  DocxPageSizeUpdate,
} from "@arcships/docx-core"

export interface DocxPageSetupSectionOption {
  index: number
  /** One-based section number, the way Word labels it. */
  sectionNumber: number
}

export interface DocxPageSetupApplyPatch {
  margins?: DocxPageMarginsUpdate
  pageSize?: DocxPageSizeUpdate
}

export interface DocxPageSetupDialogProps {
  /** Section the dialog edits; `update:sectionIndex` reports a switch. */
  sectionIndex: number
  /** Sections offered in the picker. A single-section document hides it. */
  sections?: DocxPageSetupSectionOption[]
  pageSize: DocxPageSize
  margins: DocxPageMargins
  readOnly?: boolean
  isDark?: boolean
}

// editor-form-field.ts — form field operations (select/toggle/value/widget)
//
// Extracted from useDocxEditor.ts.

import type { DocModel, FormFieldRunNode } from "@arcships/docx-core"
import type { DocxFormFieldLocation } from "@arcships/docx-core"
import { cloneDocModel, getParagraphAtLocation } from "@arcships/docx-core"
import type { EditorCore } from "./editor-shared"
import { validateFormFieldValue } from "./form-field-value"

type ApplyModelChange = (
  updater: (current: DocModel) => DocModel,
  successStatus?: string,
  options?: { clearSelectedFormField?: boolean; pushHistory?: boolean },
) => boolean

function formFieldEditKey(location: DocxFormFieldLocation): string {
  return location.kind === "paragraph"
    ? `p:${location.nodeIndex}:${location.childIndex}`
    : `t:${location.tableIndex}:${location.rowIndex}:${location.cellIndex}:${location.paragraphIndex}:${location.childIndex}`
}

export function createEditorFormField(
  ctx: EditorCore,
  applyChange: ApplyModelChange,
) {
  const endFormFieldEdit = (location?: DocxFormFieldLocation): void => {
    if (
      !location ||
      ctx.activeFormFieldEditKey.value === formFieldEditKey(location)
    ) {
      ctx.activeFormFieldEditKey.value = undefined
    }
  }

  const selectFormField = (location?: DocxFormFieldLocation): void => {
    // Focusing another field (or leaving the editor) closes the previous
    // editing session so its next keystroke starts a new undo entry.
    if (
      ctx.activeFormFieldEditKey.value &&
      ctx.activeFormFieldEditKey.value !==
        (location ? formFieldEditKey(location) : undefined)
    ) {
      ctx.activeFormFieldEditKey.value = undefined
    }
    ctx.selectedFormFieldLocation.value = location
  }

  const toggleFormCheckbox = (location: DocxFormFieldLocation): void => {
    endFormFieldEdit()
    applyChange(
      (current) => {
        const nextModel = cloneDocModel(current)
        const { paragraph: para } = getParagraphAtLocation(nextModel, location)
        if (!para) return current
        const field = para.children[location.childIndex]
        if (!field || field.type !== "form-field" || field.fieldType !== "checkbox") return current
        const currentChecked =
          typeof field.checked === "boolean"
            ? field.checked
            : field.value === "1" || field.widget?.checkbox?.defaultChecked === true
        field.checked = !currentChecked
        // Keep the legacy value in sync: exporters and form-field consumers read
        // either representation, while the checkbox UI binds `checked`.
        field.value = field.checked ? "1" : "0"
        field.sourceXml = undefined
        para.sourceXml = undefined
        return nextModel
      },
      undefined,
      { clearSelectedFormField: false, pushHistory: true },
    )
  }

  const setFormFieldValue = (
    location: DocxFormFieldLocation,
    value: string,
  ): boolean => {
    const { paragraph } = getParagraphAtLocation(ctx.model.value, location)
    if (!paragraph) return false
    const currentField = paragraph.children[location.childIndex]
    if (!currentField || currentField.type !== "form-field") return false

    const validation = validateFormFieldValue(currentField, value)
    if (!validation.valid) return false

    const editKey = formFieldEditKey(location)
    const beginsSession = ctx.activeFormFieldEditKey.value !== editKey

    const changed = applyChange(
      (current) => {
        const nextModel = cloneDocModel(current)
        const { paragraph: para } = getParagraphAtLocation(nextModel, location)
        if (!para) return current
        const field = para.children[location.childIndex]
        if (!field || field.type !== "form-field") return current
        field.value = validation.value
        if (field.fieldType === "checkbox") {
          field.checked = validation.value === "1"
        }
        field.sourceXml = undefined
        para.sourceXml = undefined
        return nextModel
      },
      undefined,
      {
        // Keep the field selected so the inline widget keeps its focus ring
        // while the user types.
        clearSelectedFormField: false,
        // One undo entry per editing session: the first keystroke snapshots the
        // pre-edit model, subsequent keystrokes fold into the same entry.
        pushHistory: beginsSession,
      },
    )

    if (changed) ctx.activeFormFieldEditKey.value = editKey
    return true
  }

  const updateFormFieldWidget = (
    location: DocxFormFieldLocation,
    patch: Partial<NonNullable<FormFieldRunNode["widget"]>>
  ): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const { paragraph: para } = getParagraphAtLocation(nextModel, location)
      if (!para) return current
      const field = para.children[location.childIndex]
      if (!field || field.type !== "form-field") return current
      field.widget = { ...(field.widget ?? {}), ...patch }
      return nextModel
    })
  }

  return {
    selectFormField,
    toggleFormCheckbox,
    setFormFieldValue,
    endFormFieldEdit,
    updateFormFieldWidget,
  }
}

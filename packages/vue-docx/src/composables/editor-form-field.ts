// editor-form-field.ts — form field operations (select/toggle/value/widget)
//
// Extracted from useDocxEditor.ts.

import type { DocModel, FormFieldRunNode } from "@arcships/docx-core"
import type { DocxFormFieldLocation } from "@arcships/docx-core"
import { cloneDocModel, getParagraphAtLocation } from "@arcships/docx-core"
import type { EditorCore } from "./editor-shared"

export function createEditorFormField(
  ctx: EditorCore,
  applyChange: (updater: (current: DocModel) => DocModel, successStatus?: string) => void,
) {
  const selectFormField = (location?: DocxFormFieldLocation): void => {
    ctx.selectedFormFieldLocation.value = location
  }

  const toggleFormCheckbox = (location: DocxFormFieldLocation): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const { paragraph: para } = getParagraphAtLocation(nextModel, location)
      if (!para) return current
      const field = para.children[location.childIndex]
      if (!field || field.type !== "form-field" || field.fieldType !== "checkbox") return current
      // toggle: updating the field value
      const newValue = field.value === "1" ? "0" : "1"
      // Apply the value change
      field.value = newValue
      return nextModel
    })
  }

  const setFormFieldValue = (location: DocxFormFieldLocation, value: string): void => {
    applyChange((current) => {
      const nextModel = cloneDocModel(current)
      const { paragraph: para } = getParagraphAtLocation(nextModel, location)
      if (!para) return current
      const field = para.children[location.childIndex]
      if (!field || field.type !== "form-field") return current
      field.value = value
      return nextModel
    })
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
    updateFormFieldWidget,
  }
}

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
      if (field.fieldType === "checkbox") {
        field.checked = value === "1" || value.toLowerCase() === "true"
      }
      field.sourceXml = undefined
      para.sourceXml = undefined
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

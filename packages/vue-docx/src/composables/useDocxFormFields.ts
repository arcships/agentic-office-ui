// Vue composable: useDocxFormFields
// Migrated from upstream @extend-ai/react-docx, editor.tsx lines 31494-31540
//
// Exposes form field state and operations from the editor controller.

import {
  collectFormFieldsFromModel,
  type DocxEditorController,
  type DocxFormFieldLocation,
  type DocxFormFieldUpdate,
  type DocxSelectedFormField,
  type FormFieldRunNode,
} from "@arcships/docx-core"

export interface UseDocxFormFieldsResult {
  formFields: DocxSelectedFormField[]
  selectedFormField?: DocxSelectedFormField
  selectFormField: (location?: DocxFormFieldLocation) => void
  setFormFieldValue: (location: DocxFormFieldLocation, value: string) => boolean
  toggleFormCheckbox: (location: DocxFormFieldLocation) => void
  endFormFieldEdit: (location?: DocxFormFieldLocation) => void
  updateFormFieldWidget: (
    location: DocxFormFieldLocation,
    patch: Partial<NonNullable<FormFieldRunNode["widget"]>>
  ) => void
  updateSelectedFormFieldWidget: (
    patch: Partial<NonNullable<FormFieldRunNode["widget"]>>
  ) => void
  /** Value + configuration update in one undo entry (the properties dialog). */
  updateFormField: (
    location: DocxFormFieldLocation,
    patch: DocxFormFieldUpdate
  ) => boolean
  updateSelectedFormField: (patch: DocxFormFieldUpdate) => boolean
}

export function useDocxFormFields(
  editor: Pick<
    DocxEditorController,
    | "model"
    | "selectedFormField"
    | "selectFormField"
    | "setFormFieldValue"
    | "toggleFormCheckbox"
    | "endFormFieldEdit"
    | "updateFormFieldWidget"
    | "updateFormField"
  >
): UseDocxFormFieldsResult {
  const updateSelectedFormFieldWidget = (
    patch: Partial<NonNullable<FormFieldRunNode["widget"]>>
  ): void => {
    if (!editor.selectedFormField) return
    editor.updateFormFieldWidget(editor.selectedFormField.location, patch)
  }

  const updateSelectedFormField = (patch: DocxFormFieldUpdate): boolean => {
    if (!editor.selectedFormField) return false
    return editor.updateFormField(editor.selectedFormField.location, patch)
  }

  return {
    get formFields() { return collectFormFieldsFromModel(editor.model) },
    get selectedFormField() { return editor.selectedFormField },
    selectFormField: editor.selectFormField,
    setFormFieldValue: editor.setFormFieldValue,
    toggleFormCheckbox: editor.toggleFormCheckbox,
    endFormFieldEdit: editor.endFormFieldEdit,
    updateFormFieldWidget: editor.updateFormFieldWidget,
    updateSelectedFormFieldWidget,
    updateFormField: editor.updateFormField,
    updateSelectedFormField,
  }
}

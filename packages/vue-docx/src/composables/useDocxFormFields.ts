// Vue composable: useDocxFormFields
// Migrated from upstream @extend-ai/react-docx, editor.tsx lines 31494-31540
//
// Exposes form field state and operations from the editor controller.

import {
  collectFormFieldsFromModel,
  type DocxEditorController,
  type DocxFormFieldLocation,
  type DocxSelectedFormField,
  type FormFieldRunNode,
} from "@extend-ai/docx-core"

export interface UseDocxFormFieldsResult {
  formFields: DocxSelectedFormField[]
  selectedFormField?: DocxSelectedFormField
  selectFormField: (location?: DocxFormFieldLocation) => void
  setFormFieldValue: (location: DocxFormFieldLocation, value: string) => void
  toggleFormCheckbox: (location: DocxFormFieldLocation) => void
  updateFormFieldWidget: (
    location: DocxFormFieldLocation,
    patch: Partial<NonNullable<FormFieldRunNode["widget"]>>
  ) => void
  updateSelectedFormFieldWidget: (
    patch: Partial<NonNullable<FormFieldRunNode["widget"]>>
  ) => void
}

export function useDocxFormFields(
  editor: Pick<
    DocxEditorController,
    | "model"
    | "selectedFormField"
    | "selectFormField"
    | "setFormFieldValue"
    | "toggleFormCheckbox"
    | "updateFormFieldWidget"
  >
): UseDocxFormFieldsResult {
  const updateSelectedFormFieldWidget = (
    patch: Partial<NonNullable<FormFieldRunNode["widget"]>>
  ): void => {
    if (!editor.selectedFormField) return
    editor.updateFormFieldWidget(editor.selectedFormField.location, patch)
  }

  return {
    get formFields() { return collectFormFieldsFromModel(editor.model) },
    get selectedFormField() { return editor.selectedFormField },
    selectFormField: editor.selectFormField,
    setFormFieldValue: editor.setFormFieldValue,
    toggleFormCheckbox: editor.toggleFormCheckbox,
    updateFormFieldWidget: editor.updateFormFieldWidget,
    updateSelectedFormFieldWidget,
  }
}

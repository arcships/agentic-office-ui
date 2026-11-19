// editor-form-field.ts — form field operations (select/toggle/value/widget)
//
// Extracted from useDocxEditor.ts.

import type { DocModel, FormFieldRunNode } from "@arcships/docx-core"
import type { DocxFormFieldLocation, DocxFormFieldUpdate } from "@arcships/docx-core"
import { cloneDocModelShallow, cloneDocNode, getParagraphAtLocation } from "@arcships/docx-core"
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

/** Copy-on-write shell plus the paragraph/table that owns the form field. */
function cloneModelForFormFieldEdit(
  current: DocModel,
  location: DocxFormFieldLocation,
): DocModel {
  const next = cloneDocModelShallow(current)
  const nodeIndex =
    location.kind === "paragraph" ? location.nodeIndex : location.tableIndex
  const node = current.nodes[nodeIndex]
  if (node) {
    next.nodes[nodeIndex] = cloneDocNode(node)
  }
  return next
}

/** Placeholder and title-style text properties keep empty strings out of the model. */
function assignOptionalText(
  field: FormFieldRunNode,
  key: "title" | "tag" | "placeholder" | "checkedSymbol" | "uncheckedSymbol",
  value: string | undefined
): void {
  if (value === undefined) return
  const trimmed = value.trim()
  if (trimmed === "") delete field[key]
  else field[key] = trimmed
}

function mergeFormFieldWidget(
  current: FormFieldRunNode["widget"],
  patch: DocxFormFieldUpdate
): FormFieldRunNode["widget"] {
  if (patch.replaceWidget) return patch.widget
  if (!patch.widget) return current
  return {
    ...(current ?? {}),
    ...patch.widget,
    text:
      patch.widget.text === undefined
        ? current?.text
        : { ...(current?.text ?? {}), ...patch.widget.text },
    checkbox:
      patch.widget.checkbox === undefined
        ? current?.checkbox
        : { ...(current?.checkbox ?? {}), ...patch.widget.checkbox },
    dropdown:
      patch.widget.dropdown === undefined
        ? current?.dropdown
        : { ...(current?.dropdown ?? {}), ...patch.widget.dropdown },
  }
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
        const nextModel = cloneModelForFormFieldEdit(current, location)
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
        const nextModel = cloneModelForFormFieldEdit(current, location)
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
      const nextModel = cloneModelForFormFieldEdit(current, location)
      const { paragraph: para } = getParagraphAtLocation(nextModel, location)
      if (!para) return current
      const field = para.children[location.childIndex]
      if (!field || field.type !== "form-field") return current
      field.widget = { ...(field.widget ?? {}), ...patch }
      return nextModel
    })
  }

  /**
   * Applies a form-field configuration patch (value, options, title, widget...)
   * as one undo entry. The value is validated against the patched field, so an
   * option list and the value it holds are checked together.
   */
  const updateFormField = (
    location: DocxFormFieldLocation,
    patch: DocxFormFieldUpdate
  ): boolean => {
    const { paragraph } = getParagraphAtLocation(ctx.model.value, location)
    const currentField = paragraph?.children[location.childIndex]
    if (!currentField || currentField.type !== "form-field") {
      ctx.status.value = ctx.statusMessage("formFieldTargetUnavailable")
      return false
    }

    const nextFieldType = patch.fieldType ?? currentField.fieldType
    // Leaving the dropdown type drops the option list with it: keeping the old
    // entries would export a value list that the new field type cannot show.
    const nextOptions =
      patch.fieldType !== undefined && patch.fieldType !== "dropdown"
        ? patch.options
        : patch.options ?? currentField.options
    if (nextFieldType === "dropdown" && patch.value !== undefined) {
      const options = nextOptions ?? []
      // An empty dropdown is a valid "no choice yet" state while the dialog is
      // configuring the field; a required field still has to hold a value.
      if (options.length === 0 && patch.value !== "" && patch.required !== false) {
        ctx.status.value = ctx.statusMessage("formFieldOptionsRequired")
        return false
      }
    }

    const candidate: FormFieldRunNode = {
      ...currentField,
      fieldType: nextFieldType,
      options: nextOptions,
      widget: mergeFormFieldWidget(currentField.widget, patch),
    }
    if (patch.value !== undefined) candidate.value = patch.value
    if (patch.checked !== undefined) candidate.checked = patch.checked

    if (
      patch.value !== undefined &&
      (patch.value !== "" || candidate.required === true)
    ) {
      const validation = validateFormFieldValue(candidate, patch.value)
      if (!validation.valid) return false
      candidate.value = validation.value
    }

    applyChange(
      (current) => {
        const nextModel = cloneModelForFormFieldEdit(current, location)
        const { paragraph: para } = getParagraphAtLocation(nextModel, location)
        if (!para) return current
        const field = para.children[location.childIndex]
        if (!field || field.type !== "form-field") return current

        field.fieldType = candidate.fieldType
        if (candidate.options === undefined) delete field.options
        else field.options = candidate.options.map((option) => ({ ...option }))
        if (candidate.widget === undefined) delete field.widget
        else field.widget = { ...candidate.widget }
        if (candidate.value === undefined) delete field.value
        else field.value = candidate.value
        assignOptionalText(field, "title", patch.title)
        assignOptionalText(field, "tag", patch.tag)
        assignOptionalText(field, "placeholder", patch.placeholder)
        assignOptionalText(field, "checkedSymbol", patch.checkedSymbol)
        assignOptionalText(field, "uncheckedSymbol", patch.uncheckedSymbol)
        if (patch.required !== undefined) {
          if (patch.required) field.required = true
          else delete field.required
        }
        if (candidate.fieldType === "checkbox") {
          const checked =
            patch.checked ?? (candidate.value === "1" || candidate.value === "true")
          field.checked = checked
          field.value = checked ? "1" : "0"
        } else {
          delete field.checked
        }
        // The field is rebuilt from the model on export; keeping the verbatim
        // XML would write the old value/options back into the document.
        field.sourceXml = undefined
        para.sourceXml = undefined
        return nextModel
      },
      ctx.statusMessage("updatedFormField"),
      { clearSelectedFormField: false, pushHistory: true },
    )

    // A dialog save is a finished editing session: the next inline keystroke
    // opens its own undo entry instead of folding into the dialog's.
    endFormFieldEdit(location)
    return true
  }

  return {
    selectFormField,
    toggleFormCheckbox,
    setFormFieldValue,
    endFormFieldEdit,
    updateFormFieldWidget,
    updateFormField,
  }
}

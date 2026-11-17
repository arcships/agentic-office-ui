import { strToU8 } from "fflate";
import type {
  XlsxFormControl,
  XlsxFormControlInput,
  XlsxImageAnchor,
  XlsxImageMarker,
} from "../types";
import {
  type ArchiveEntries,
  CONTENT_TYPES_NS,
  CTRL_PROP_REL_TYPE,
  PKG_REL_NS,
  SPREADSHEET_NS,
  type WorkbookFormControlOrigin,
  getFirstChild,
  getFirstDescendant,
  getLocalElements,
  normalizeArchivePath,
  parseXml,
  readArchiveText,
  relativeArchivePath,
  relsPathForDocument,
  serializeXml,
} from "./image-parser";
import type { WorkbookImageAssets } from "./grid-render";

const XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const FORM_CONTROL_CONTENT_TYPE = "application/vnd.ms-excel.controlproperties+xml";

const OBJECT_TYPE_BY_KIND: Record<XlsxFormControl["kind"], string> = {
  button: "Button",
  checkbox: "CheckBox",
  dropdown: "DropDown",
  editbox: "EditBox",
  "group-box": "GroupBox",
  label: "Label",
  listbox: "ListBox",
  radio: "OptionButton",
  scrollbar: "ScrollBar",
  spinner: "SpinButton",
  unknown: "Unknown",
};

function createElement(document: XMLDocument, namespace: string, name: string) {
  return document.createElementNS(namespace, name);
}

function createMarker(document: XMLDocument, name: string, marker: XlsxImageMarker) {
  const element = createElement(document, XDR_NS, `xdr:${name}`);
  for (const [childName, value] of [
    ["col", marker.col],
    ["colOff", marker.colOffsetEmu],
    ["row", marker.row],
    ["rowOff", marker.rowOffsetEmu],
  ] as const) {
    const child = createElement(document, XDR_NS, `xdr:${childName}`);
    child.textContent = String(Math.max(0, Math.round(Number(value) || 0)));
    element.appendChild(child);
  }
  return element;
}

export function createAnchorElement(document: XMLDocument, anchor: XlsxImageAnchor) {
  const anchorElement = createElement(document, XDR_NS, "xdr:anchor");
  if (anchor.kind === "two-cell") {
    anchorElement.appendChild(createMarker(document, "from", anchor.from));
    anchorElement.appendChild(createMarker(document, "to", anchor.to));
    return anchorElement;
  }
  if (anchor.kind === "one-cell") {
    anchorElement.appendChild(createMarker(document, "from", anchor.from));
    const ext = createElement(document, XDR_NS, "xdr:ext");
    ext.setAttribute("cx", String(Math.max(0, Math.round(anchor.sizeEmu.cx))));
    ext.setAttribute("cy", String(Math.max(0, Math.round(anchor.sizeEmu.cy))));
    anchorElement.appendChild(ext);
    return anchorElement;
  }
  const pos = createElement(document, XDR_NS, "xdr:pos");
  pos.setAttribute("x", String(Math.max(0, Math.round(anchor.positionEmu.x))));
  pos.setAttribute("y", String(Math.max(0, Math.round(anchor.positionEmu.y))));
  const ext = createElement(document, XDR_NS, "xdr:ext");
  ext.setAttribute("cx", String(Math.max(0, Math.round(anchor.sizeEmu.cx))));
  ext.setAttribute("cy", String(Math.max(0, Math.round(anchor.sizeEmu.cy))));
  anchorElement.appendChild(pos);
  anchorElement.appendChild(ext);
  return anchorElement;
}

function findControlNode(document: XMLDocument, origin: WorkbookFormControlOrigin) {
  const controls = getLocalElements(document, "control");
  if (origin.shapeId !== null) {
    const match = controls.find((control) => Number(control.getAttribute("shapeId")) === origin.shapeId);
    if (match) return match;
  }
  return controls[origin.controlIndex] ?? null;
}

function ensureControlsContainer(document: XMLDocument) {
  const worksheet = document.documentElement;
  const existing = getFirstChild(worksheet, "controls");
  if (existing) return existing;
  const controls = createElement(document, SPREADSHEET_NS, "controls");
  const extLst = getFirstChild(worksheet, "extLst");
  worksheet.insertBefore(controls, extLst ?? null);
  return controls;
}

function updateControlAnchor(controlNode: Element, anchor: XlsxImageAnchor) {
  const controlPr = getFirstChild(controlNode, "controlPr") ?? controlNode;
  const existing = getFirstDescendant(controlPr, "anchor");
  const document = controlNode.ownerDocument as XMLDocument;
  const next = createAnchorElement(document, anchor);
  if (existing && existing.parentNode) {
    existing.parentNode.replaceChild(next, existing);
  } else {
    controlPr.appendChild(next);
  }
}

function objectTypeForKind(kind: XlsxFormControl["kind"]) {
  return OBJECT_TYPE_BY_KIND[kind] ?? "Unknown";
}

function buildCtrlPropDocument(kind: XlsxFormControl["kind"], input: Partial<XlsxFormControlInput>) {
  const document = parseXml("<formControlPr/>");
  if (!document) return null;
  const root = document.documentElement;
  root.setAttribute("xmlns", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
  root.setAttribute("objectType", objectTypeForKind(kind));
  if (input.checked !== undefined) root.setAttribute("checked", input.checked ? "1" : "0");
  if (input.linkedCell) root.setAttribute("fmlaLink", input.linkedCell);
  return document;
}

function ensureRelationshipsDocument(archive: ArchiveEntries, relsPath: string) {
  const existing = readArchiveText(archive, relsPath);
  if (existing) {
    const document = parseXml(existing);
    if (document) return document;
  }
  const document = parseXml("<Relationships/>");
  if (!document) return null;
  document.documentElement.setAttribute("xmlns", PKG_REL_NS);
  return document;
}

function relationshipIdForShape(shapeId: number) {
  return `rIdFormControl${shapeId}`;
}

function nextShapeId(controls: Element[]) {
  const ids = controls
    .map((control) => Number(control.getAttribute("shapeId")))
    .filter((value) => Number.isFinite(value) && value > 0);
  return (ids.length ? Math.max(...ids) : 1024) + 1;
}

function ensureContentTypeOverride(archive: ArchiveEntries, partName: string) {
  const xml = readArchiveText(archive, "[Content_Types].xml");
  const document = xml ? parseXml(xml) : parseXml("<Types/>");
  if (!document) return null;
  if (!document.documentElement.getAttribute("xmlns")) {
    document.documentElement.setAttribute("xmlns", CONTENT_TYPES_NS);
  }
  const normalizedPartName = normalizeArchivePath(partName);
  const overrides = getLocalElements(document, "Override");
  if (!overrides.some((node) => normalizeArchivePath(node.getAttribute("PartName") ?? "") === normalizedPartName)) {
    const override = document.createElementNS(CONTENT_TYPES_NS, "Override");
    override.setAttribute("PartName", `/${normalizedPartName}`);
    override.setAttribute("ContentType", FORM_CONTROL_CONTENT_TYPE);
    document.documentElement.appendChild(override);
  }
  return serializeXml(document);
}

function controlFromInput(
  workbookSheetIndex: number,
  shapeId: number,
  input: XlsxFormControlInput,
  zIndex: number,
): XlsxFormControl {
  return {
    anchor: input.anchor,
    checked: input.checked,
    editable: true,
    id: `form-control-${workbookSheetIndex}-shape-${shapeId}`,
    kind: input.kind ?? "checkbox",
    label: input.label,
    linkedCell: input.linkedCell,
    name: input.name ?? input.label,
    sheetIndex: workbookSheetIndex,
    shapeId,
    workbookSheetIndex,
    zIndex,
  };
}

export function addWorkbookFormControl(
  assets: WorkbookImageAssets,
  workbookSheetIndex: number,
  input: XlsxFormControlInput,
): XlsxFormControl | null {
  const sheetPath = assets.sheetPathsByWorkbookSheetIndex[workbookSheetIndex];
  if (!sheetPath) return null;
  const sheetXml = readArchiveText(assets.archive, sheetPath);
  const sheetDocument = sheetXml ? parseXml(sheetXml) : null;
  if (!sheetDocument) return null;

  const controls = getLocalElements(sheetDocument, "control");
  const shapeId = nextShapeId(controls);
  const relationshipId = relationshipIdForShape(shapeId);
  const ctrlPropPath = normalizeArchivePath(`xl/ctrlProps/formControl${shapeId}.xml`);
  const sheetRelsPath = relsPathForDocument(sheetPath);
  const relationshipsDocument = ensureRelationshipsDocument(assets.archive, sheetRelsPath);
  const ctrlPropDocument = buildCtrlPropDocument(input.kind ?? "checkbox", input);
  if (!relationshipsDocument || !ctrlPropDocument) return null;

  const relationship = relationshipsDocument.createElementNS(PKG_REL_NS, "Relationship");
  relationship.setAttribute("Id", relationshipId);
  relationship.setAttribute("Type", CTRL_PROP_REL_TYPE);
  relationship.setAttribute("Target", relativeArchivePath(sheetPath, ctrlPropPath));
  relationshipsDocument.documentElement.appendChild(relationship);

  const controlsContainer = ensureControlsContainer(sheetDocument);
  const controlNode = sheetDocument.createElementNS(SPREADSHEET_NS, "control");
  controlNode.setAttribute("shapeId", String(shapeId));
  if (input.name ?? input.label) controlNode.setAttribute("name", input.name ?? input.label ?? "");
  controlNode.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
  controlNode.setAttributeNS(
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "r:id",
    relationshipId,
  );
  const controlPr = sheetDocument.createElementNS(SPREADSHEET_NS, "controlPr");
  controlPr.appendChild(createAnchorElement(sheetDocument, input.anchor));
  controlNode.appendChild(controlPr);
  controlsContainer.appendChild(controlNode);

  assets.archive[ctrlPropPath] = strToU8(serializeXml(ctrlPropDocument));
  assets.archive[normalizeArchivePath(sheetRelsPath)] = strToU8(serializeXml(relationshipsDocument));
  assets.archive[normalizeArchivePath(sheetPath)] = strToU8(serializeXml(sheetDocument));
  const contentTypes = ensureContentTypeOverride(assets.archive, ctrlPropPath);
  if (contentTypes) assets.archive["[Content_Types].xml"] = strToU8(contentTypes);

  const existing = assets.formControlsByWorkbookSheetIndex[workbookSheetIndex] ?? [];
  const zIndex = input.zIndex ?? (existing.reduce((max, control) => Math.max(max, control.zIndex), 0) + 1);
  const control = controlFromInput(workbookSheetIndex, shapeId, input, zIndex);
  assets.formControlsByWorkbookSheetIndex[workbookSheetIndex] = [...existing, control]
    .sort((left, right) => left.zIndex - right.zIndex);
  assets.formControlOriginsById.set(control.id, {
    controlIndex: controls.length,
    ctrlPropPath,
    shapeId,
    sheetPath,
    workbookSheetIndex,
  });
  return control;
}

export function updateWorkbookFormControl(
  assets: WorkbookImageAssets,
  formControlId: string,
  patch: Partial<XlsxFormControlInput>,
): boolean {
  const origin = assets.formControlOriginsById.get(formControlId);
  if (!origin) return false;
  const sheetXml = readArchiveText(assets.archive, origin.sheetPath);
  const sheetDocument = sheetXml ? parseXml(sheetXml) : null;
  if (!sheetDocument) return false;
  const controlNode = findControlNode(sheetDocument, origin);
  if (!controlNode) return false;

  if (patch.name !== undefined || patch.label !== undefined) {
    controlNode.setAttribute("name", patch.name ?? patch.label ?? "");
  }
  if (patch.anchor) {
    updateControlAnchor(controlNode, patch.anchor);
  }
  if (origin.ctrlPropPath) {
    const ctrlPropXml = readArchiveText(assets.archive, origin.ctrlPropPath);
    const ctrlPropDocument = ctrlPropXml ? parseXml(ctrlPropXml) : null;
    if (ctrlPropDocument) {
      const root = ctrlPropDocument.documentElement;
      if (patch.kind) root.setAttribute("objectType", objectTypeForKind(patch.kind));
      if (patch.checked !== undefined) root.setAttribute("checked", patch.checked ? "1" : "0");
      if (patch.linkedCell !== undefined) {
        if (patch.linkedCell) root.setAttribute("fmlaLink", patch.linkedCell);
        else root.removeAttribute("fmlaLink");
      }
      assets.archive[normalizeArchivePath(origin.ctrlPropPath)] = strToU8(serializeXml(ctrlPropDocument));
    }
  }
  assets.archive[normalizeArchivePath(origin.sheetPath)] = strToU8(serializeXml(sheetDocument));

  assets.formControlsByWorkbookSheetIndex[origin.workbookSheetIndex] = (
    assets.formControlsByWorkbookSheetIndex[origin.workbookSheetIndex] ?? []
  ).map((control) => control.id === formControlId
    ? {
        ...control,
        ...(patch.anchor ? { anchor: patch.anchor } : {}),
        ...(patch.checked !== undefined ? { checked: patch.checked } : {}),
        ...(patch.kind ? { kind: patch.kind } : {}),
        ...(patch.label !== undefined ? { label: patch.label } : {}),
        ...(patch.linkedCell !== undefined ? { linkedCell: patch.linkedCell } : {}),
        ...(patch.name !== undefined ? { name: patch.name } : {}),
        ...(patch.zIndex !== undefined ? { zIndex: patch.zIndex } : {}),
      }
    : control
  ).sort((left, right) => left.zIndex - right.zIndex);
  return true;
}

export function deleteWorkbookFormControl(assets: WorkbookImageAssets, formControlId: string): boolean {
  const origin = assets.formControlOriginsById.get(formControlId);
  if (!origin) return false;
  const sheetXml = readArchiveText(assets.archive, origin.sheetPath);
  const sheetDocument = sheetXml ? parseXml(sheetXml) : null;
  if (!sheetDocument) return false;
  const controlNode = findControlNode(sheetDocument, origin);
  if (!controlNode) return false;
  controlNode.parentNode?.removeChild(controlNode);
  assets.archive[normalizeArchivePath(origin.sheetPath)] = strToU8(serializeXml(sheetDocument));

  if (origin.ctrlPropPath) {
    delete assets.archive[normalizeArchivePath(origin.ctrlPropPath)];
    const sheetRelsPath = relsPathForDocument(origin.sheetPath);
    const relsXml = readArchiveText(assets.archive, sheetRelsPath);
    const relsDocument = relsXml ? parseXml(relsXml) : null;
    if (relsDocument) {
      const relationshipId = origin.shapeId !== null ? relationshipIdForShape(origin.shapeId) : null;
      const relationship = getLocalElements(relsDocument, "Relationship").find((node) => (
        relationshipId !== null
          ? node.getAttribute("Id") === relationshipId
          : normalizeArchivePath(node.getAttribute("Target") ?? "") === normalizeArchivePath(
            relativeArchivePath(origin.sheetPath, origin.ctrlPropPath!),
          )
      ));
      relationship?.parentNode?.removeChild(relationship);
      assets.archive[normalizeArchivePath(sheetRelsPath)] = strToU8(serializeXml(relsDocument));
    }
    const contentTypesXml = readArchiveText(assets.archive, "[Content_Types].xml");
    const contentTypesDocument = contentTypesXml ? parseXml(contentTypesXml) : null;
    if (contentTypesDocument) {
      const override = getLocalElements(contentTypesDocument, "Override").find((node) => (
        normalizeArchivePath(node.getAttribute("PartName") ?? "") === normalizeArchivePath(origin.ctrlPropPath!)
      ));
      override?.parentNode?.removeChild(override);
      assets.archive["[Content_Types].xml"] = strToU8(serializeXml(contentTypesDocument));
    }
  }

  assets.formControlOriginsById.delete(formControlId);
  assets.formControlsByWorkbookSheetIndex[origin.workbookSheetIndex] = (
    assets.formControlsByWorkbookSheetIndex[origin.workbookSheetIndex] ?? []
  ).filter((control) => control.id !== formControlId);
  return true;
}

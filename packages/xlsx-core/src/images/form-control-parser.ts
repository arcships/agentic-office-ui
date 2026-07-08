import type {
  XlsxFormControl,
  XlsxShape,
} from "../types";
import {
  type ArchiveEntries,
  CTRL_PROP_REL_TYPE,
  type DrawingRectEmu,
  type RelationshipRecord,
  VML_DRAWING_REL_TYPE,
  getFirstChild,
  getFirstDescendant,
  getLocalElements,
  getRelationshipId,
  isElementNode,
  parseAnchor,
  parseXml,
  readArchiveText,
} from "./image-parser";
import { type WorkbookSheetState } from "./grid-render";
import { anchorToAbsoluteRect } from "./column-width";

// ── Shared text utilities ──────────────────────────────────────────────────

export function normalizeControlLabel(label: string | null | undefined) {
  if (!label) {
    return undefined;
  }

  const normalized = label
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return normalized.length > 0 ? normalized : undefined;
}

export function flattenShapeText(shape: XlsxShape) {
  const text = shape.paragraphs
    .flatMap((paragraph) => paragraph.runs.map((run) => run.text))
    .join(" ");
  return normalizeControlLabel(text);
}

// ── Form control types ─────────────────────────────────────────────────────

type ParsedSheetFormControl = {
  anchor: XlsxFormControl["anchor"] | null;
  controlRelationshipId: string | null;
  name?: string;
  shapeId: number | null;
};

type ParsedCtrlProp = {
  checked?: boolean;
  linkedCell?: string;
  objectType?: string;
};

type ParsedVmlFormControl = {
  checked?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  hidden: boolean;
  label?: string;
  linkedCell?: string;
  objectType?: string;
  shapeId: number | null;
  textAlign?: "center" | "left" | "right";
  textColor?: string;
  zIndex: number;
};

// ── Form control helpers ───────────────────────────────────────────────────

function parseSpreadsheetBooleanValue(value: string | null | undefined) {
  if (value === null || value === undefined) {
    return undefined;
  }

  const normalized = value.trim().toLowerCase();
  if (!normalized) {
    return true;
  }

  return !["0", "false", "none", "off", "unchecked"].includes(normalized);
}

function parseSpreadsheetBooleanNode(node: Element | null) {
  if (!node) {
    return undefined;
  }

  return parseSpreadsheetBooleanValue(node.getAttribute("val") ?? node.textContent);
}

function parseFormControlKind(rawType: string | null | undefined): XlsxFormControl["kind"] {
  const normalized = (rawType ?? "").trim().toLowerCase();
  switch (normalized) {
    case "button":
      return "button";
    case "checkbox":
      return "checkbox";
    case "drop":
      return "dropdown";
    case "editbox":
      return "editbox";
    case "gbox":
      return "group-box";
    case "label":
      return "label";
    case "list":
      return "listbox";
    case "radio":
      return "radio";
    case "scroll":
      return "scrollbar";
    case "spin":
      return "spinner";
    default:
      return "unknown";
  }
}

function parseFormControlShapeId(value: string | null | undefined) {
  const match = (value ?? "").match(/(\d+)(?!.*\d)/);
  if (!match) {
    return null;
  }

  const parsed = Number(match[1]);
  return Number.isFinite(parsed) ? parsed : null;
}

function parseCssDeclarationValue(styleText: string | null | undefined, property: string) {
  if (!styleText) {
    return null;
  }

  const pattern = new RegExp(`${property}\\s*:\\s*([^;]+)`, "i");
  const match = pattern.exec(styleText);
  return match?.[1]?.trim() ?? null;
}

function parseControlTextAlign(styleText: string | null | undefined): XlsxFormControl["textAlign"] {
  const value = parseCssDeclarationValue(styleText, "text-align")?.toLowerCase();
  if (value === "center" || value === "right") {
    return value;
  }
  return value === "left" ? "left" : undefined;
}

function parseVmlFontSizePt(value: string | null | undefined) {
  const parsed = Number(value ?? Number.NaN);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return undefined;
  }

  return parsed > 40 ? parsed / 20 : parsed;
}

function isPlaceholderFormControlName(name: string | null | undefined) {
  const normalized = normalizeControlLabel(name)?.toLowerCase();
  if (!normalized) {
    return false;
  }

  return /^(option button|group box|check box|drop down|dropdown|list box|edit box|scroll bar|spinner|spin button|button)\s+\d+$/.test(normalized);
}

function resolveNamedFormControlLabel(name: string | null | undefined) {
  return isPlaceholderFormControlName(name) ? undefined : normalizeControlLabel(name);
}

// ── Form control parsing ───────────────────────────────────────────────────

function parseSheetFormControlNodes(
  archive: ArchiveEntries,
  sheetPath: string
) {
  const sheetXml = readArchiveText(archive, sheetPath);
  if (!sheetXml) {
    return [] as ParsedSheetFormControl[];
  }

  const sheetDocument = parseXml(sheetXml);
  if (!sheetDocument) {
    return [] as ParsedSheetFormControl[];
  }

  return getLocalElements(sheetDocument, "control").map((controlNode) => ({
    anchor: parseAnchor(getFirstDescendant(controlNode, "anchor") ?? controlNode),
    controlRelationshipId: getRelationshipId(controlNode),
    name: controlNode.getAttribute("name") ?? undefined,
    shapeId: parseFormControlShapeId(controlNode.getAttribute("shapeId"))
  }));
}

function parseCtrlPropDocument(
  archive: ArchiveEntries,
  ctrlPropPath: string
) {
  const xml = readArchiveText(archive, ctrlPropPath);
  if (!xml) {
    return null;
  }

  const document = parseXml(xml);
  const root = document?.documentElement;
  if (!root) {
    return null;
  }

  return {
    checked: parseSpreadsheetBooleanValue(root.getAttribute("checked")),
    linkedCell: root.getAttribute("fmlaLink") ?? undefined,
    objectType: root.getAttribute("objectType") ?? undefined
  } satisfies ParsedCtrlProp;
}

function parseVmlFormControls(
  archive: ArchiveEntries,
  vmlDrawingPath: string
) {
  const xml = readArchiveText(archive, vmlDrawingPath);
  if (!xml) {
    return new Map<number, ParsedVmlFormControl>();
  }

  const document = parseXml(xml);
  if (!document) {
    return new Map<number, ParsedVmlFormControl>();
  }

  const controls = new Map<number, ParsedVmlFormControl>();
  for (const shapeNode of getLocalElements(document, "shape")) {
    const clientDataNode = getFirstChild(shapeNode, "ClientData");
    if (!clientDataNode) {
      continue;
    }

    const shapeId = parseFormControlShapeId(
      shapeNode.getAttributeNS("urn:schemas-microsoft-com:office:office", "spid")
      ?? shapeNode.getAttribute("o:spid")
      ?? shapeNode.getAttribute("spid")
      ?? shapeNode.getAttribute("id")
    );
    if (shapeId === null) {
      continue;
    }

    const styleText = shapeNode.getAttribute("style");
    const textboxNode = getFirstChild(shapeNode, "textbox");
    const fontNode = textboxNode ? getFirstDescendant(textboxNode, "font") : null;
    const textContainerNode = textboxNode ? getFirstDescendant(textboxNode, "div") : null;
    const label = normalizeControlLabel(textboxNode?.textContent);
    const zIndex = Number(parseCssDeclarationValue(styleText, "z-index") ?? Number.NaN);

    controls.set(shapeId, {
      checked: parseSpreadsheetBooleanNode(getFirstChild(clientDataNode, "Checked")),
      fontFamily: fontNode?.getAttribute("face") ?? undefined,
      fontSizePt: parseVmlFontSizePt(fontNode?.getAttribute("size")),
      hidden: (parseCssDeclarationValue(styleText, "visibility") ?? "").toLowerCase() === "hidden",
      label,
      linkedCell: normalizeControlLabel(getFirstChild(clientDataNode, "FmlaLink")?.textContent),
      objectType: clientDataNode.getAttribute("ObjectType") ?? undefined,
      shapeId,
      textAlign: parseControlTextAlign(textContainerNode?.getAttribute("style")),
      textColor: fontNode?.getAttribute("color") ?? undefined,
      zIndex: Number.isFinite(zIndex) ? zIndex : controls.size + 1
    });
  }

  return controls;
}

export function parseSheetFormControls(
  archive: ArchiveEntries,
  sheetPath: string,
  sheetRelationships: Map<string, RelationshipRecord>,
  workbookSheetIndex: number,
  zIndexBase: number
) {
  const controlNodes = parseSheetFormControlNodes(archive, sheetPath);
  if (controlNodes.length === 0) {
    return [] as XlsxFormControl[];
  }

  const legacyDrawingRelationship = [...sheetRelationships.values()].find(
    (relationship) => relationship.type === VML_DRAWING_REL_TYPE
  );
  const vmlControlsByShapeId = legacyDrawingRelationship
    ? parseVmlFormControls(archive, legacyDrawingRelationship.target)
    : new Map<number, ParsedVmlFormControl>();
  const parsedControls: XlsxFormControl[] = [];

  controlNodes.forEach((controlNode, index) => {
    if (!controlNode.anchor) {
      return;
    }

    const ctrlPropRelationship = controlNode.controlRelationshipId
      ? sheetRelationships.get(controlNode.controlRelationshipId) ?? null
      : null;
    const ctrlPropPath = ctrlPropRelationship?.type === CTRL_PROP_REL_TYPE
      ? ctrlPropRelationship.target
      : null;
    const ctrlProp = ctrlPropPath
      ? parseCtrlPropDocument(archive, ctrlPropPath)
      : null;
    const vmlControl = controlNode.shapeId !== null
      ? vmlControlsByShapeId.get(controlNode.shapeId) ?? null
      : null;
    const kind = parseFormControlKind(ctrlProp?.objectType ?? vmlControl?.objectType);

    parsedControls.push({
      anchor: controlNode.anchor,
      checked: ctrlProp?.checked ?? vmlControl?.checked,
      fontFamily: vmlControl?.fontFamily,
      fontSizePt: vmlControl?.fontSizePt,
      hidden: vmlControl?.hidden ?? false,
      id: `form-control-${workbookSheetIndex}-${index}`,
      kind,
      label: vmlControl?.label,
      linkedCell: ctrlProp?.linkedCell ?? vmlControl?.linkedCell,
      name: controlNode.name,
      sheetIndex: workbookSheetIndex,
      textAlign: vmlControl?.textAlign,
      textColor: vmlControl?.textColor,
      workbookSheetIndex,
      zIndex: zIndexBase + (vmlControl?.zIndex ?? index + 1)
    });
  });

  return parsedControls.sort((left, right) => left.zIndex - right.zIndex);
}

// ── Form control enrichment ─────────────────────────────────────────────────

function rectArea(rect: DrawingRectEmu) {
  return Math.max(0, rect.cx) * Math.max(0, rect.cy);
}

function rectIntersectionArea(left: DrawingRectEmu, right: DrawingRectEmu) {
  const overlapX = Math.max(0, Math.min(left.x + left.cx, right.x + right.cx) - Math.max(left.x, right.x));
  const overlapY = Math.max(0, Math.min(left.y + left.cy, right.y + right.cy) - Math.max(left.y, right.y));
  return overlapX * overlapY;
}

function rectCenterDistance(left: DrawingRectEmu, right: DrawingRectEmu) {
  const leftCenterX = left.x + left.cx / 2;
  const leftCenterY = left.y + left.cy / 2;
  const rightCenterX = right.x + right.cx / 2;
  const rightCenterY = right.y + right.cy / 2;
  return Math.hypot(leftCenterX - rightCenterX, leftCenterY - rightCenterY);
}

function findHiddenShapeControlMatch(
  control: XlsxFormControl,
  shapes: XlsxShape[],
  sheetState: WorkbookSheetState | null
): XlsxShape | null {
  const controlRect = anchorToAbsoluteRect(control.anchor, sheetState);
  const controlArea = Math.max(1, rectArea(controlRect));
  const placeholderName = isPlaceholderFormControlName(control.name);
  let bestMatch: XlsxShape | null = null;
  let bestScore = Number.NEGATIVE_INFINITY;

  shapes.forEach((shape) => {
    if (!shape.hidden) {
      return;
    }

    const shapeRect = anchorToAbsoluteRect(shape.anchor, sheetState);
    const shapeArea = Math.max(1, rectArea(shapeRect));
    const intersectionArea = rectIntersectionArea(controlRect, shapeRect);
    const overlapScore = intersectionArea / Math.min(controlArea, shapeArea);
    const distance = rectCenterDistance(controlRect, shapeRect);
    const maxDimension = Math.max(controlRect.cx, controlRect.cy, shapeRect.cx, shapeRect.cy, 1);
    const distanceScore = distance / maxDimension;
    const textLabel = flattenShapeText(shape);
    const sameName = normalizeControlLabel(shape.name)?.toLowerCase() === normalizeControlLabel(control.name)?.toLowerCase();
    let score = overlapScore * 4 - distanceScore;

    if (sameName) {
      score += 1;
    }

    if (control.kind === "group-box") {
      score += textLabel ? -3 : 0.5;
      score += shape.stroke?.none ? -1.5 : 0.75;
    } else {
      score += textLabel ? 2 : -1;
      score += shape.stroke?.none ? 0.2 : -0.5;
    }

    if (!placeholderName && textLabel && textLabel === resolveNamedFormControlLabel(control.name)) {
      score += 0.5;
    }

    if (score > bestScore) {
      bestMatch = shape;
      bestScore = score;
    }
  });

  return bestScore >= 0.25 ? bestMatch : null;
}

export function enrichFormControlsWithHiddenShapes(
  formControls: XlsxFormControl[],
  shapes: XlsxShape[],
  sheetState: WorkbookSheetState | null
) {
  return formControls.map((control) => {
    const matchedShape = findHiddenShapeControlMatch(control, shapes, sheetState);
    const matchedLabel = matchedShape ? flattenShapeText(matchedShape) : undefined;
    const fallbackLabel = resolveNamedFormControlLabel(control.name);
    let resolvedAnchor = control.anchor;
    if (matchedShape && (control.kind === "group-box" || matchedLabel || isPlaceholderFormControlName(control.name))) {
      resolvedAnchor = matchedShape.anchor;
    }

    return {
      ...control,
      anchor: resolvedAnchor,
      label: control.label ?? matchedLabel ?? fallbackLabel
    };
  });
}

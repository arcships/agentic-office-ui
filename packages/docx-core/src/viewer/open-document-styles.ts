// OpenDocument text styles: reads the ODF style model into the shared DOCX
// document-model style structures.
//
// ODF stores resolved formatting in `styles.xml` (named styles) and
// `content.xml` (automatic styles). Paragraph/list/table properties are
// hierarchical, so every lookup walks the parent-style chain and then merges the
// style's own properties. The document model has no ODF source markup, so the
// reader resolves the formatting once at import time.

import type {
  ParagraphAlignment,
  ParagraphBorderSet,
  ParagraphBorderStyle,
  ParagraphDirection,
  ParagraphIndent,
  ParagraphSpacing,
  ParagraphStyle,
  ParagraphStyleDefinition,
  TableBorderSet,
  TableBorderStyle,
  TableBoxSpacing,
  TableCellStyle,
  TableRowStyle,
  TableStyle,
  TextStyle,
} from "../engine/types"
import {
  attribute,
  childElements,
  decodeStyleName,
  descendantElements,
  firstChildElement,
  lengthToPt,
  lengthToTwips,
  localName,
  percentage,
} from "./open-document-xml"

type AttributeMap = Record<string, string>

export interface OdfStyleSheet {
  styles: Map<string, OdfRawStyle>
  defaults: Map<string, OdfRawStyle>
  listStyles: Map<string, Element>
  pageLayouts: Map<string, Element>
  masterPages: Map<string, Element>
  fontFaces: Map<string, string>
}

interface OdfRawStyle {
  name: string
  family: string
  parentName?: string
  displayName?: string
  defaultOutlineLevel?: number
  listStyleName?: string
  properties: Map<keyof OdfPropertyGroups, Element[]>
}

interface OdfPropertyGroups {
  paragraph: AttributeMap
  text: AttributeMap
  table: AttributeMap
  tableRow: AttributeMap
  tableColumn: AttributeMap
  tableCell: AttributeMap
  graphic: AttributeMap
  listLevel: AttributeMap
  listLabel: AttributeMap
}

export interface OdfResolvedStyle {
  paragraph: AttributeMap
  text: AttributeMap
  table: AttributeMap
  tableRow: AttributeMap
  tableColumn: AttributeMap
  tableCell: AttributeMap
  graphic: AttributeMap
  listStyleName?: string
  headingLevel?: number
  displayName?: string
  isDefault: boolean
}

export interface OdfParagraphStyleResult {
  style: ParagraphStyle
  definition: ParagraphStyleDefinition
  listStyleName?: string
  /** Paragraph-level text properties applied as the base run style. */
  runStyle?: TextStyle
}

export interface OdfListLevel {
  /** Zero-based level index. */
  level: number
  format: string
  text: string
  suffix: "tab" | "space" | "nothing"
  indent?: ParagraphIndent
  runStyle?: TextStyle
  bulletFontFamily?: string
  start?: number
}

export interface OdfListDefinition {
  name: string
  levels: OdfListLevel[]
}

export interface OdfPageLayout {
  sectionPropertiesXml: string
  pageWidthTwips: number
  pageHeightTwips: number
}

function emptyAttributeMap(): AttributeMap {
  return {}
}

function copyAttributes(element: Element): AttributeMap {
  const result: AttributeMap = {}
  for (let index = 0; index < element.attributes.length; index += 1) {
    const item = element.attributes.item(index)
    if (!item) continue
    result[localName(item)] = item.value
  }
  return result
}

function mergeAttributeMap(target: AttributeMap, source: AttributeMap): void {
  for (const [key, value] of Object.entries(source)) {
    if (value !== "") target[key] = value
  }
}

function emptyGroups(): OdfPropertyGroups {
  return {
    paragraph: emptyAttributeMap(),
    text: emptyAttributeMap(),
    table: emptyAttributeMap(),
    tableRow: emptyAttributeMap(),
    tableColumn: emptyAttributeMap(),
    tableCell: emptyAttributeMap(),
    graphic: emptyAttributeMap(),
    listLevel: emptyAttributeMap(),
    listLabel: emptyAttributeMap(),
  }
}

const GROUP_BY_ELEMENT: Readonly<Record<string, keyof OdfPropertyGroups>> = Object.freeze({
  "paragraph-properties": "paragraph",
  "text-properties": "text",
  "table-properties": "table",
  "table-row-properties": "tableRow",
  "table-column-properties": "tableColumn",
  "table-cell-properties": "tableCell",
  "graphic-properties": "graphic",
  "list-level-properties": "listLevel",
  "list-level-label-alignment": "listLabel",
})

function addStyleElement(style: OdfRawStyle, element: Element): void {
  const name = localName(element)
  const group = GROUP_BY_ELEMENT[name]
  if (!group) return
  const list = style.properties.get(group) ?? []
  list.push(element)
  style.properties.set(group, list)
}

function readRawStyle(element: Element, family: string): OdfRawStyle | undefined {
  const name = attribute(element, "name")
  if (!name) return undefined
  const style: OdfRawStyle = {
    name,
    family,
    parentName: attribute(element, "parent-style-name"),
    displayName: attribute(element, "display-name")
      ? decodeStyleName(attribute(element, "display-name") as string)
      : decodeStyleName(name),
    defaultOutlineLevel: numeric(attribute(element, "default-outline-level")),
    listStyleName: attribute(element, "list-style-name"),
    properties: new Map(),
  }
  for (const child of childElements(element)) addStyleElement(style, child)
  return style
}

function numeric(value: string | undefined): number | undefined {
  if (value === undefined) return undefined
  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : undefined
}

export function collectOdfStyles(documents: readonly Document[]): OdfStyleSheet {
  const sheet: OdfStyleSheet = {
    styles: new Map(),
    defaults: new Map(),
    listStyles: new Map(),
    pageLayouts: new Map(),
    masterPages: new Map(),
    fontFaces: new Map(),
  }
  for (const document of documents) {
    for (const element of descendantElements(document, "font-face")) {
      const name = attribute(element, "name")
      const family = attribute(element, "font-family")
      if (name && family) sheet.fontFaces.set(name, family.replace(/^["']|["']$/g, ""))
    }
    for (const element of descendantElements(document, "style")) {
      const family = attribute(element, "family")
      if (family) {
        const style = readRawStyle(element, family)
        if (style) sheet.styles.set(style.name, style)
      }
    }
    for (const element of descendantElements(document, "default-style")) {
      const family = attribute(element, "family")
      if (!family) continue
      const style = readRawStyle(element, family)
      if (style) sheet.defaults.set(family, style)
    }
    for (const element of descendantElements(document, "list-style")) {
      const name = attribute(element, "name")
      if (name && !sheet.listStyles.has(name)) sheet.listStyles.set(name, element)
    }
    for (const element of descendantElements(document, "page-layout")) {
      const name = attribute(element, "name")
      if (name && !sheet.pageLayouts.has(name)) sheet.pageLayouts.set(name, element)
    }
    for (const element of descendantElements(document, "master-page")) {
      const name = attribute(element, "name")
      if (name && !sheet.masterPages.has(name)) sheet.masterPages.set(name, element)
    }
  }
  return sheet
}

function applyRawStyle(sheet: OdfStyleSheet, raw: OdfRawStyle | undefined, groups: OdfPropertyGroups, seen: Set<string>): void {
  if (!raw || seen.has(raw.name)) return
  seen.add(raw.name)
  if (raw.parentName) {
    applyRawStyle(sheet, sheet.styles.get(raw.parentName), groups, seen)
  }
  for (const [group, elements] of raw.properties) {
    const target = groups[group]
    for (const element of elements) mergeAttributeMap(target, copyAttributes(element))
  }
}

export function resolveOdfStyle(
  sheet: OdfStyleSheet,
  name: string | undefined,
  family: "paragraph" | "text" | "table" | "table-cell" | "table-row" | "table-column" | "graphic",
): OdfResolvedStyle {
  const groups = emptyGroups()
  const explicit = name ? sheet.styles.get(name) : undefined
  const defaultStyle = sheet.defaults.get(family)
  if (defaultStyle) applyRawStyle(sheet, defaultStyle, groups, new Set())
  if (explicit) applyRawStyle(sheet, explicit, groups, new Set())

  let headingLevel = explicit ? headingLevelForStyle(sheet, explicit) : undefined
  if (headingLevel === undefined && explicit?.defaultOutlineLevel !== undefined) {
    headingLevel = explicit.defaultOutlineLevel
  }
  return {
    paragraph: groups.paragraph,
    text: groups.text,
    table: groups.table,
    tableRow: groups.tableRow,
    tableColumn: groups.tableColumn,
    tableCell: groups.tableCell,
    graphic: groups.graphic,
    listStyleName: explicit?.listStyleName,
    headingLevel,
    displayName: explicit?.displayName,
    isDefault: Boolean(defaultStyle && explicit === undefined),
  }
}

function headingLevelForStyle(sheet: OdfStyleSheet, style: OdfRawStyle): number | undefined {
  const seen = new Set<string>()
  let current: OdfRawStyle | undefined = style
  while (current && !seen.has(current.name)) {
    seen.add(current.name)
    const match = /^(?:heading|outline|h)[ _-]?(\d)$/i.exec(current.displayName ?? current.name)
    if (match) return Math.min(6, Math.max(1, Number(match[1])))
    if (current.defaultOutlineLevel !== undefined && current.defaultOutlineLevel > 0) {
      return Math.min(6, Math.max(1, current.defaultOutlineLevel))
    }
    current = current.parentName ? sheet.styles.get(current.parentName) : undefined
  }
  return undefined
}

function alignment(value: string | undefined): ParagraphAlignment | undefined {
  switch (value) {
    case "center": return "center"
    case "right": return "right"
    case "justify": return "justify"
    case "left": return "left"
    case "start": return "start"
    case "end": return "end"
    default: return undefined
  }
}

function direction(value: string | undefined): ParagraphDirection | undefined {
  if (!value) return undefined
  return /^rl/.test(value) || value === "tb-rl" ? "rtl" : undefined
}

function positiveTwips(value: string | undefined): number | undefined {
  const twips = lengthToTwips(value)
  return twips === undefined ? undefined : Math.max(0, twips)
}

function signedTwips(value: string | undefined): number | undefined {
  return lengthToTwips(value)
}

function paragraphSpacing(properties: AttributeMap): ParagraphSpacing | undefined {
  const before = positiveTwips(properties["margin-top"])
  const after = positiveTwips(properties["margin-bottom"])
  const lineHeight = properties["line-height"]
  const spacing: ParagraphSpacing = {}
  if (before !== undefined) spacing.beforeTwips = before
  if (after !== undefined) spacing.afterTwips = after
  const linePercentage = percentage(lineHeight)
  if (linePercentage !== undefined && linePercentage > 0) {
    // Word's automatic line rule stores 240ths of a line.
    spacing.lineTwips = Math.round((linePercentage / 100) * 240)
    spacing.lineRule = "auto"
  } else if (lineHeight !== undefined && lineHeight !== "normal") {
    const exact = lengthToTwips(lineHeight)
    if (exact !== undefined && exact > 0) {
      spacing.lineTwips = exact
      spacing.lineRule = "atLeast"
    }
  }
  return Object.keys(spacing).length > 0 ? spacing : undefined
}

function paragraphIndent(properties: AttributeMap): ParagraphIndent | undefined {
  const left = positiveTwips(properties["margin-left"])
  const right = positiveTwips(properties["margin-right"])
  const textIndent = signedTwips(properties["text-indent"])
  const indent: ParagraphIndent = {}
  if (left !== undefined) indent.leftTwips = left
  if (right !== undefined) indent.rightTwips = right
  if (textIndent !== undefined) {
    if (textIndent < 0) indent.hangingTwips = Math.abs(textIndent)
    else if (textIndent > 0) indent.firstLineTwips = textIndent
  }
  return Object.keys(indent).length > 0 ? indent : undefined
}

function borderStyle(value: string | undefined): ParagraphBorderStyle | TableBorderStyle | undefined {
  if (!value || value === "none" || value === "hidden") return undefined
  const parts = value.trim().split(/\s+/)
  if (parts.length < 2) return undefined
  const sizePt = lengthToPt(parts[0])
  const type = parts[1] === "solid" || parts[1] === "double" || parts[1] === "dotted"
    || parts[1] === "dashed" || parts[1] === "dash-dot" || parts[1] === "dot-dot-dash"
    || parts[1] === "groove" || parts[1] === "ridge" || parts[1] === "inset"
    || parts[1] === "outset"
    ? parts[1] === "dash-dot" ? "dotDash" : parts[1] === "dot-dot-dash" ? "dotDotDash" : parts[1]
    : parts[1]
  const color = parts.slice(2).find((part) => /^#|^rgb|^[a-z]+$/i.test(part))
  return {
    type,
    ...(color ? { color: color.startsWith("#") ? color : color.toLowerCase() } : {}),
    ...(sizePt !== undefined ? { sizeEighthPt: Math.max(1, Math.round(sizePt * 8)) } : {}),
  }
}

function borderSet(properties: AttributeMap): ParagraphBorderSet | undefined {
  const top = borderStyle(properties["border-top"] ?? properties["border"])
  const right = borderStyle(properties["border-right"] ?? properties["border"])
  const bottom = borderStyle(properties["border-bottom"] ?? properties["border"])
  const left = borderStyle(properties["border-left"] ?? properties["border"])
  if (!top && !right && !bottom && !left) return undefined
  return { top, right, bottom, left } as ParagraphBorderSet
}

function tableBorderSet(properties: AttributeMap): TableBorderSet | undefined {
  const top = borderStyle(properties["border-top"] ?? properties["border"])
  const right = borderStyle(properties["border-right"] ?? properties["border"])
  const bottom = borderStyle(properties["border-bottom"] ?? properties["border"])
  const left = borderStyle(properties["border-left"] ?? properties["border"])
  if (!top && !right && !bottom && !left) return undefined
  return { top, right, bottom, left } as TableBorderSet
}

function boxSpacing(properties: AttributeMap, prefix = "padding"): TableBoxSpacing | undefined {
  const all = lengthToTwips(properties[prefix])
  const top = lengthToTwips(properties[`${prefix}-top`]) ?? all
  const right = lengthToTwips(properties[`${prefix}-right`]) ?? all
  const bottom = lengthToTwips(properties[`${prefix}-bottom`]) ?? all
  const left = lengthToTwips(properties[`${prefix}-left`]) ?? all
  if (top === undefined && right === undefined && bottom === undefined && left === undefined) {
    return undefined
  }
  return { topTwips: top, rightTwips: right, bottomTwips: bottom, leftTwips: left }
}

function underlineStyle(value: string | undefined): TextStyle["underlineStyle"] {
  switch (value) {
    case "double": return "double"
    case "dotted": return "dotted"
    case "dash": return "dash"
    case "long-dash": return "dashLong"
    case "dot-dash": return "dotDash"
    case "dot-dot-dash": return "dotDotDash"
    case "wave": return "wave"
    case "bold": return "thick"
    case "solid":
    default:
      return value === undefined || value === "none" ? undefined : "single"
  }
}

function textStyleFromProperties(sheet: OdfStyleSheet, properties: AttributeMap, base?: TextStyle): TextStyle {
  const style: TextStyle = { ...(base ?? {}) }
  const fontName = properties["font-name"]
  const fontFamily = (fontName ? sheet.fontFaces.get(fontName) : undefined)
    ?? properties["font-family"]?.replace(/^["']|["']$/g, "")
  if (fontFamily) style.fontFamily = fontFamily
  const fontSizePt = lengthToPt(properties["font-size"])
  if (fontSizePt !== undefined) style.fontSizePt = fontSizePt

  const weight = properties["font-weight"]?.toLowerCase()
  if (weight === "bold" || (weight !== undefined && Number(weight) >= 600)) style.bold = true
  else if (weight === "normal" || (weight !== undefined && Number(weight) < 600)) style.bold = false
  const fontStyle = properties["font-style"]?.toLowerCase()
  if (fontStyle === "italic" || fontStyle === "oblique") style.italic = true
  else if (fontStyle === "normal") style.italic = false

  const underline = underlineStyle(properties["text-underline-style"])
  if (underline !== undefined) {
    style.underline = true
    style.underlineStyle = underline
  } else if (properties["text-underline-style"] === "none") {
    style.underline = false
    style.underlineStyle = undefined
  }
  const lineThrough = properties["text-line-through-style"]
  if (lineThrough !== undefined && lineThrough !== "none") {
    style.strike = true
    style.strikeStyle = lineThrough === "double" ? "double" : "single"
  } else if (lineThrough === "none") {
    style.strike = false
  }

  const color = properties["color"]
  if (color && color !== "font-color" && color !== "currentColor") style.color = color
  const background = properties["background-color"]
  if (background && background !== "transparent") style.backgroundColor = background

  const position = properties["text-position"]
  if (position?.startsWith("super")) style.verticalAlign = "superscript"
  else if (position?.startsWith("sub")) style.verticalAlign = "subscript"

  const variant = properties["font-variant"]
  if (variant === "small-caps") style.smallCaps = true
  const transform = properties["text-transform"]
  if (transform === "uppercase") style.caps = true

  const letterSpacing = lengthToTwips(properties["letter-spacing"])
  if (letterSpacing !== undefined) style.characterSpacingTwips = letterSpacing
  return style
}

export function resolveOdfTextStyle(
  sheet: OdfStyleSheet,
  name: string | undefined,
  base?: TextStyle,
): TextStyle | undefined {
  if (name === undefined && base === undefined) return undefined
  const resolved = resolveOdfStyle(sheet, name, "text")
  const style = textStyleFromProperties(sheet, resolved.text, base)
  return Object.keys(style).length > 0 ? style : undefined
}

export function resolveOdfParagraphStyle(
  sheet: OdfStyleSheet,
  name: string | undefined,
): OdfParagraphStyleResult {
  const resolved = resolveOdfStyle(sheet, name, "paragraph")
  const style: ParagraphStyle = {}
  const align = alignment(resolved.paragraph["text-align"])
  if (align) style.align = align
  const dir = direction(resolved.paragraph["writing-mode"])
  if (dir) style.direction = dir
  if (resolved.headingLevel !== undefined) style.headingLevel = resolved.headingLevel as ParagraphStyle["headingLevel"]
  if (name) style.styleId = name
  const styleName = resolved.displayName ?? (name ? decodeStyleName(name) : undefined)
  if (styleName) style.styleName = styleName
  const spacing = paragraphSpacing(resolved.paragraph)
  if (spacing) style.spacing = spacing
  const indent = paragraphIndent(resolved.paragraph)
  if (indent) style.indent = indent
  const background = resolved.paragraph["background-color"]
  if (background && background !== "transparent") style.backgroundColor = background
  const borders = borderSet(resolved.paragraph)
  if (borders) style.borders = borders
  if (resolved.paragraph["break-before"] === "page") style.pageBreakBefore = true
  if (resolved.paragraph["keep-with-next"] === "always") style.keepNext = true
  const resolvedRunStyle = textStyleFromProperties(sheet, resolved.text)
  const runStyle = Object.keys(resolvedRunStyle).length > 0 ? resolvedRunStyle : undefined

  const definition: ParagraphStyleDefinition = {
    id: name ?? "Standard",
    name: styleName ?? "Standard",
    ...(style.headingLevel !== undefined ? { headingLevel: style.headingLevel } : {}),
    ...(align ? { align } : {}),
    ...(dir ? { direction: dir } : {}),
    ...(spacing ? { spacing } : {}),
    ...(indent ? { indent } : {}),
    ...(style.backgroundColor ? { backgroundColor: style.backgroundColor } : {}),
    ...(borders ? { borders } : {}),
    ...(style.pageBreakBefore ? { pageBreakBefore: true } : {}),
    ...(style.keepNext ? { keepNext: true } : {}),
    ...(runStyle ? { runStyle } : {}),
    ...(resolved.isDefault ? { isDefault: true } : {}),
  }
  return {
    style,
    definition,
    listStyleName: resolved.listStyleName,
    ...(runStyle ? { runStyle } : {}),
  }
}

function listLevelFormat(value: string | undefined): string {
  switch (value) {
    case "a": return "lowerLetter"
    case "A": return "upperLetter"
    case "i": return "lowerRoman"
    case "I": return "upperRoman"
    case "01": return "decimalZero"
    case "1": return "decimal"
    case "":
    case undefined:
      return "none"
    default:
      return "decimal"
  }
}

function normalizeListTemplate(value: string | undefined): string | undefined {
  if (!value) return undefined
  // LibreOffice stores `%1%.`; OOXML templates use `%1.`.
  return value.replace(/%(\d+)%/g, "%$1")
}

function listLevelIndent(levelElement: Element): ParagraphIndent | undefined {
  const properties = firstChildElement(levelElement, "list-level-properties")
  const label = properties ? firstChildElement(properties, "list-level-label-alignment") : undefined
  if (!label) return undefined
  const left = positiveTwips(attribute(label, "margin-left"))
  const textIndent = signedTwips(attribute(label, "text-indent"))
  const indent: ParagraphIndent = {}
  if (left !== undefined) indent.leftTwips = left
  if (textIndent !== undefined) {
    if (textIndent < 0) indent.hangingTwips = Math.abs(textIndent)
    else if (textIndent > 0) indent.firstLineTwips = textIndent
  }
  return Object.keys(indent).length > 0 ? indent : undefined
}

export function resolveOdfListDefinition(
  sheet: OdfStyleSheet,
  name: string | undefined,
): OdfListDefinition | undefined {
  const element = name ? sheet.listStyles.get(name) : undefined
  if (!element) return undefined
  const levels: OdfListLevel[] = []
  for (const levelElement of childElements(element)) {
    const kind = localName(levelElement)
    if (kind !== "list-level-style-bullet" && kind !== "list-level-style-number") continue
    const levelNumber = numeric(attribute(levelElement, "level")) ?? levels.length + 1
    const level = Math.max(0, levelNumber - 1)
    const textProperties = firstChildElement(levelElement, "text-properties")
    const textStyle = textProperties
      ? textStyleFromProperties(sheet, copyAttributes(textProperties))
      : undefined
    const bullet = kind === "list-level-style-bullet"
    const numberFormat = listLevelFormat(attribute(levelElement, "num-format"))
    const bulletChar = attribute(levelElement, "bullet-char")?.replace(/[\uF0A7\uF0B7]/g, "•")
    const template = normalizeListTemplate(attribute(levelElement, "num-list-format"))
    const prefix = attribute(levelElement, "num-prefix") ?? ""
    const suffixText = attribute(levelElement, "num-suffix") ?? ""
    let text: string
    if (bullet) {
      text = bulletChar ?? "•"
    } else {
      let baseTemplate = template ?? "%1"
      if (template === undefined) {
        baseTemplate = suffixText ? `%1${suffixText}` : "%1"
      } else if (suffixText && !/[.)\]]$/.test(baseTemplate)) {
        baseTemplate += suffixText
      }
      text = `${prefix}${baseTemplate}`
    }
    const labelFollowedBy = (() => {
      const properties = firstChildElement(levelElement, "list-level-properties")
      const label = properties ? firstChildElement(properties, "list-level-label-alignment") : undefined
      return label ? attribute(label, "label-followed-by") : undefined
    })()
    const suffix: OdfListLevel["suffix"] = labelFollowedBy === "listtab"
      ? "tab"
      : labelFollowedBy === "nothing"
        ? "nothing"
        : "space"
    levels.push({
      level,
      format: bullet ? "bullet" : numberFormat,
      text,
      suffix,
      ...(listLevelIndent(levelElement) ? { indent: listLevelIndent(levelElement) } : {}),
      ...(textStyle && Object.keys(textStyle).length > 0 ? { runStyle: textStyle } : {}),
      ...(bullet && textStyle?.fontFamily ? { bulletFontFamily: textStyle.fontFamily } : {}),
      ...(numeric(attribute(levelElement, "start-value")) !== undefined
        ? { start: numeric(attribute(levelElement, "start-value")) }
        : {}),
    })
  }
  return levels.length > 0 ? { name: name ?? "List", levels } : undefined
}

export function resolveOdfTableStyle(sheet: OdfStyleSheet, name: string | undefined): TableStyle {
  const resolved = resolveOdfStyle(sheet, name, "table")
  const style: TableStyle = {}
  const width = lengthToTwips(resolved.table["width"])
  if (width !== undefined) style.widthTwips = width
  const indent = signedTwips(resolved.table["margin-left"])
  if (indent !== undefined) style.indentTwips = indent
  const cellSpacing = lengthToTwips(resolved.table["border-spacing"])
  if (cellSpacing !== undefined) style.cellSpacingTwips = cellSpacing
  style.layout = "fixed"
  return style
}

export function resolveOdfTableColumnWidth(sheet: OdfStyleSheet, name: string | undefined): number | undefined {
  const resolved = resolveOdfStyle(sheet, name, "table-column")
  return positiveTwips(resolved.tableColumn["column-width"])
}

export function resolveOdfTableRowStyle(sheet: OdfStyleSheet, name: string | undefined): TableRowStyle | undefined {
  const resolved = resolveOdfStyle(sheet, name, "table-row")
  const style: TableRowStyle = {}
  const height = positiveTwips(resolved.tableRow["row-height"] ?? resolved.tableRow["height"])
  if (height !== undefined) {
    style.heightTwips = height
    style.heightRule = "atLeast"
  }
  const background = resolved.tableRow["background-color"]
  if (background && background !== "transparent") style.backgroundColor = background
  return Object.keys(style).length > 0 ? style : undefined
}

export function resolveOdfTableCellStyle(sheet: OdfStyleSheet, name: string | undefined): TableCellStyle | undefined {
  const resolved = resolveOdfStyle(sheet, name, "table-cell")
  const style: TableCellStyle = {}
  const background = resolved.tableCell["background-color"]
  if (background && background !== "transparent") style.backgroundColor = background
  const verticalAlign = resolved.tableCell["vertical-align"]
  if (verticalAlign === "middle" || verticalAlign === "center") style.verticalAlign = "center"
  else if (verticalAlign === "bottom") style.verticalAlign = "bottom"
  else if (verticalAlign === "top") style.verticalAlign = "top"
  const margin = boxSpacing(resolved.tableCell)
  if (margin) style.marginTwips = margin
  const borders = tableBorderSet(resolved.tableCell)
  if (borders) style.borders = borders
  return Object.keys(style).length > 0 ? style : undefined
}

export function resolveOdfPageLayout(
  sheet: OdfStyleSheet,
  masterPageName: string | undefined,
): OdfPageLayout {
  const master = masterPageName ? sheet.masterPages.get(masterPageName) : sheet.masterPages.values().next().value as Element | undefined
  const layoutName = master ? attribute(master, "page-layout-name") : undefined
  const layout = layoutName ? sheet.pageLayouts.get(layoutName) : undefined
  const properties = layout ? firstChildElement(layout, "page-layout-properties") : undefined
  const values = properties ? copyAttributes(properties) : {}

  let width = positiveTwips(values["page-width"]) ?? 12240
  let height = positiveTwips(values["page-height"]) ?? 15840
  if (values["print-orientation"] === "landscape" && width < height) {
    ;[width, height] = [height, width]
  }
  const top = positiveTwips(values["margin-top"]) ?? 1440
  const right = positiveTwips(values["margin-right"]) ?? 1440
  const bottom = positiveTwips(values["margin-bottom"]) ?? 1440
  const left = positiveTwips(values["margin-left"]) ?? 1440
  const gutter = positiveTwips(values["margin-gutter"]) ?? 0
  const orientation = width > height ? ' w:orient="landscape"' : ""
  const sectionPropertiesXml =
    `<w:sectPr><w:pgSz w:w="${width}" w:h="${height}"${orientation}/>`
    + `<w:pgMar w:top="${top}" w:right="${right}" w:bottom="${bottom}" w:left="${left}"`
    + ` w:header="708" w:footer="708" w:gutter="${gutter}"/></w:sectPr>`
  return { sectionPropertiesXml, pageWidthTwips: width, pageHeightTwips: height }
}


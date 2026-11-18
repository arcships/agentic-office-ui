// clipboard-html.ts — HTML flavour of the DOCX clipboard.
//
// Copying a DOCX range writes `text/html` so Word and other editors receive
// bold/italic/underline/strike/colour/size/links/lists instead of flat text;
// `clipboard-html-parse.ts` maps the same tags and inline declarations back
// onto runs. Everything here is string and model work — no DOM APIs — so the
// same code runs in a Worker and in plain unit tests.
//
// Security: the clipboard is external input. Element content that can execute
// (script/style/iframe/object/embed/svg/…) is dropped by the parser, only
// whitelisted style declarations and colour/font-size/font-family shapes
// survive, and links are kept only when the caller's resolver (the Office URL
// policy in the viewer) accepts them.

import type {
  HeadingLevel,
  ParagraphAlignment,
  ParagraphStyle,
  TextStyle
} from "../engine/types"

export interface DocxClipboardRun {
  text: string
  style?: TextStyle
  /** Absolute or relative target already accepted by the caller's resolver. */
  link?: string
}

export interface DocxClipboardListInfo {
  type: "ordered" | "unordered"
  /** Zero-based OOXML list level. */
  level: number
}

export interface DocxClipboardParagraph {
  runs: DocxClipboardRun[]
  style?: ParagraphStyle
  list?: DocxClipboardListInfo
}

export interface DocxClipboardParseOptions {
  /**
   * Accepts or drops a link target. The Office URL policy
   * (`sanitizeOfficeLink`) is passed in by the viewer; core keeps no copy of
   * that policy. Without a resolver every link is dropped.
   */
  sanitizeLink?: (value: string) => string | undefined
  maxHtmlChars?: number
  maxParagraphs?: number
  maxRuns?: number
}

/** `text/html` flavour cap; a larger clipboard is ignored instead of parsed. */
export const MAX_DOCX_CLIPBOARD_HTML_CHARS = 512 * 1024
/** Upper bound on paragraphs taken from one clipboard payload. */
export const MAX_DOCX_CLIPBOARD_PARAGRAPHS = 2000
/** Upper bound on runs taken from one clipboard payload. */
export const MAX_DOCX_CLIPBOARD_RUNS = 20000

const HTML_ESCAPES: Record<string, string> = {
  amp: "&",
  apos: "'",
  bull: "•",
  copy: "©",
  gt: ">",
  hellip: "…",
  ldquo: "“",
  lsquo: "‘",
  lt: "<",
  mdash: "—",
  middot: "·",
  nbsp: "\u00a0",
  ndash: "–",
  quot: "\"",
  rdquo: "”",
  reg: "®",
  rsquo: "’",
  trade: "™"
}

const SAFE_COLOR_PATTERN =
  /^(?:#[0-9a-f]{3,8}|[a-z]{3,20}|rgba?\([\d\s.,%/]+\)|hsla?\([\d\s.,%/]+\))$/i
const SAFE_FONT_FAMILY_PATTERN = /^[A-Za-z0-9 _.-]{1,64}$/
const HEX_COLOR_PATTERN = /^[0-9a-f]{3,8}$/i

export const HEADING_FONT_SIZES_PT = [18, 16, 14, 13, 12, 11]
const UNDERLINE_STYLES = new Set<NonNullable<TextStyle["underlineStyle"]>>([
  "single",
  "words",
  "double",
  "thick",
  "dotted",
  "dottedHeavy",
  "dash",
  "dashedHeavy",
  "dashLong",
  "dashLongHeavy",
  "dotDash",
  "dashDotHeavy",
  "dotDotDash",
  "dashDotDotHeavy",
  "wave",
  "wavyHeavy",
  "wavyDouble"
])
const STRIKE_STYLES = new Set<NonNullable<TextStyle["strikeStyle"]>>(["single", "double"])

// ── text helpers ───────────────────────────────────────────────────

/** Decodes the entities Word and browsers put on the clipboard. */
export function decodeClipboardHtmlEntities(value: string): string {
  if (!value.includes("&")) return value
  return value.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z][a-zA-Z0-9]*);/g, (match, body: string) => {
    if (body.startsWith("#")) {
      const isHex = body[1] === "x" || body[1] === "X"
      const code = Number.parseInt(isHex ? body.slice(2) : body.slice(1), isHex ? 16 : 10)
      if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return match
      try {
        return String.fromCodePoint(code)
      } catch {
        return match
      }
    }
    return HTML_ESCAPES[body.toLowerCase()] ?? match
  })
}

/** Escapes text for the generated `text/html` payload. */
export function escapeClipboardHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
}

function escapeClipboardAttribute(value: string): string {
  return escapeClipboardHtml(value).replace(/'/g, "&#39;")
}

/** Collapses HTML whitespace the way a renderer would outside `<pre>`. */
export function collapseClipboardWhitespace(value: string): string {
  return value.replace(/[\t\n\r\f]+/g, " ").replace(/ {2,}/g, " ")
}

// ── style whitelist ────────────────────────────────────────────────

/** Keeps only colour-shaped values so a copied style cannot inject CSS. */
export function sanitizeClipboardColor(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined
  const trimmed = value.trim()
  if (!trimmed) return undefined
  const normalized = trimmed.toLowerCase()
  if (HEX_COLOR_PATTERN.test(normalized)) return `#${normalized}`
  if (!SAFE_COLOR_PATTERN.test(normalized)) return undefined
  return normalized
}

export function sanitizeClipboardFontFamily(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined
  const first = value.split(",")[0]?.trim().replace(/^["']|["']$/g, "")
  if (!first || !SAFE_FONT_FAMILY_PATTERN.test(first)) return undefined
  return first
}

function finiteNumber(value: unknown, min: number, max: number): number | undefined {
  if (typeof value !== "number" || !Number.isFinite(value)) return undefined
  if (value < min || value > max) return undefined
  return value
}

/**
 * Whitelists the run fields the clipboard represents.
 *
 * Exotic OOXML-only decorations (run borders, shadows, outlines, east-Asian
 * font slots) are intentionally dropped: they are not expressible in
 * `text/html`, and accepting arbitrary nested objects from the clipboard only
 * widens the input surface.
 */
export function sanitizeClipboardTextStyle(value: unknown): TextStyle | undefined {
  if (!value || typeof value !== "object") return undefined
  const source = value as Record<string, unknown>
  const style: TextStyle = {}
  if (typeof source.bold === "boolean") style.bold = source.bold
  if (typeof source.italic === "boolean") style.italic = source.italic
  if (typeof source.underline === "boolean") style.underline = source.underline
  if (typeof source.strike === "boolean") style.strike = source.strike
  if (typeof source.caps === "boolean") style.caps = source.caps
  if (typeof source.smallCaps === "boolean") style.smallCaps = source.smallCaps
  if (typeof source.underlineStyle === "string") {
    const underlineStyle = source.underlineStyle as NonNullable<TextStyle["underlineStyle"]>
    if (UNDERLINE_STYLES.has(underlineStyle)) style.underlineStyle = underlineStyle
  }
  if (typeof source.strikeStyle === "string") {
    const strikeStyle = source.strikeStyle as NonNullable<TextStyle["strikeStyle"]>
    if (STRIKE_STYLES.has(strikeStyle)) style.strikeStyle = strikeStyle
  }
  const color = sanitizeClipboardColor(source.color)
  if (color) style.color = color
  const highlight = sanitizeClipboardColor(source.highlight)
  if (highlight) style.highlight = highlight
  const backgroundColor = sanitizeClipboardColor(source.backgroundColor)
  if (backgroundColor) style.backgroundColor = backgroundColor
  const fontSizePt = finiteNumber(source.fontSizePt, 1, 400)
  if (fontSizePt !== undefined) style.fontSizePt = fontSizePt
  const fontFamily = sanitizeClipboardFontFamily(source.fontFamily)
  if (fontFamily) style.fontFamily = fontFamily
  if (source.verticalAlign === "superscript" || source.verticalAlign === "subscript") {
    style.verticalAlign = source.verticalAlign
  }
  const characterSpacingTwips = finiteNumber(source.characterSpacingTwips, -400, 4000)
  if (characterSpacingTwips !== undefined) style.characterSpacingTwips = characterSpacingTwips
  const characterScalePercent = finiteNumber(source.characterScalePercent, 1, 600)
  if (characterScalePercent !== undefined) style.characterScalePercent = characterScalePercent
  if (source.direction === "ltr" || source.direction === "rtl") {
    style.direction = source.direction
  }
  return Object.keys(style).length > 0 ? style : undefined
}

export function mergeStyles(base: TextStyle | undefined, delta: TextStyle | undefined): TextStyle | undefined {
  if (!delta) return base ? { ...base } : undefined
  if (!base) return { ...delta }
  return { ...base, ...delta }
}

// ── inline CSS ─────────────────────────────────────────────────────

export interface ParsedInlineStyle {
  style: TextStyle
  align?: ParagraphAlignment
  /** Word's `mso-list:l0 level1 lfo1` marker on a paragraph. */
  list?: DocxClipboardListInfo
  /** Word's `mso-list:Ignore` marker on the bullet/number run. */
  listIgnore?: boolean
}

function parseFontSizePt(value: string): number | undefined {
  const match = /^\s*(-?\d+(?:\.\d+)?)\s*(pt|px)\s*$/i.exec(value)
  if (!match) return undefined
  const amount = Number(match[1])
  if (!Number.isFinite(amount) || amount <= 0) return undefined
  const pt = match[2].toLowerCase() === "px" ? amount * 0.75 : amount
  return Math.max(1, Math.min(400, Math.round(pt * 10) / 10))
}

function fontWeightIsBold(value: string): boolean | undefined {
  const normalized = value.trim().toLowerCase()
  if (normalized === "normal" || normalized === "lighter") return false
  if (normalized === "bold" || normalized === "bolder") return true
  const numeric = Number.parseInt(normalized, 10)
  if (Number.isFinite(numeric)) return numeric >= 600
  return undefined
}

export function parseClipboardInlineStyle(value: string | undefined): ParsedInlineStyle {
  const result: ParsedInlineStyle = { style: {} }
  if (!value) return result

  for (const declaration of value.split(";")) {
    const separator = declaration.indexOf(":")
    if (separator <= 0) continue
    const property = declaration.slice(0, separator).trim().toLowerCase()
    const raw = declaration.slice(separator + 1).trim()
    if (!property || !raw) continue

    switch (property) {
      case "font-weight": {
        const bold = fontWeightIsBold(raw)
        if (bold !== undefined) result.style.bold = bold
        break
      }
      case "font-style": {
        const normalized = raw.toLowerCase()
        if (normalized === "italic" || normalized === "oblique") result.style.italic = true
        else if (normalized === "normal") result.style.italic = false
        break
      }
      case "text-decoration":
      case "text-decoration-line": {
        const normalized = raw.toLowerCase()
        if (normalized.includes("none")) {
          result.style.underline = false
          result.style.strike = false
        } else {
          if (normalized.includes("underline")) result.style.underline = true
          if (normalized.includes("line-through")) result.style.strike = true
        }
        break
      }
      case "color": {
        const color = sanitizeClipboardColor(raw)
        if (color) result.style.color = color
        break
      }
      case "background":
      case "background-color": {
        const color = sanitizeClipboardColor(raw)
        if (color) result.style.backgroundColor = color
        break
      }
      case "font-size": {
        const fontSizePt = parseFontSizePt(raw)
        if (fontSizePt !== undefined) result.style.fontSizePt = fontSizePt
        break
      }
      case "font-family": {
        const fontFamily = sanitizeClipboardFontFamily(raw)
        if (fontFamily) result.style.fontFamily = fontFamily
        break
      }
      case "vertical-align": {
        const normalized = raw.toLowerCase()
        if (normalized === "super") result.style.verticalAlign = "superscript"
        else if (normalized === "sub") result.style.verticalAlign = "subscript"
        break
      }
      case "text-transform": {
        if (raw.trim().toLowerCase() === "uppercase") result.style.caps = true
        break
      }
      case "font-variant": {
        if (raw.trim().toLowerCase().includes("small-caps")) result.style.smallCaps = true
        break
      }
      case "text-align": {
        const normalized = raw.trim().toLowerCase()
        if (
          normalized === "left" ||
          normalized === "right" ||
          normalized === "center" ||
          normalized === "justify"
        ) {
          result.align = normalized
        }
        break
      }
      case "mso-list": {
        const normalized = raw.trim().toLowerCase()
        if (normalized.includes("ignore")) {
          result.listIgnore = true
          break
        }
        const levelMatch = /level(\d+)/.exec(normalized)
        const level = levelMatch ? Math.max(0, Math.min(8, Number(levelMatch[1]) - 1)) : 0
        // `mso-list` names a numbering definition but not its format; Word's
        // HTML bullet lists are the common case, so treat it as a bullet list.
        result.list = { type: "unordered", level }
        break
      }
      default:
        break
    }
  }

  return result
}

/** Serializes a run style into portable inline CSS for `text/html`. */
export function clipboardRunCss(style?: TextStyle): string {
  if (!style) return ""
  const declarations: string[] = []
  if (style.bold) declarations.push("font-weight:bold")
  if (style.italic) declarations.push("font-style:italic")
  const decorations: string[] = []
  if (style.underline) decorations.push("underline")
  if (style.strike) decorations.push("line-through")
  if (decorations.length > 0) declarations.push(`text-decoration:${decorations.join(" ")}`)
  if (style.caps) declarations.push("text-transform:uppercase")
  if (style.smallCaps) declarations.push("font-variant:small-caps")
  const color = sanitizeClipboardColor(style.color)
  if (color) declarations.push(`color:${color}`)
  const background = sanitizeClipboardColor(style.backgroundColor ?? style.highlight)
  if (background) declarations.push(`background-color:${background}`)
  if (typeof style.fontSizePt === "number" && Number.isFinite(style.fontSizePt)) {
    declarations.push(`font-size:${style.fontSizePt}pt`)
  }
  const fontFamily = sanitizeClipboardFontFamily(style.fontFamily)
  if (fontFamily) declarations.push(`font-family:${escapeClipboardAttribute(fontFamily)}`)
  if (style.verticalAlign === "superscript") declarations.push("vertical-align:super")
  if (style.verticalAlign === "subscript") declarations.push("vertical-align:sub")
  return declarations.join(";")
}

function paragraphOpenTag(paragraph: DocxClipboardParagraph): string {
  const headingLevel = paragraph.style?.headingLevel
  if (paragraph.list) return "li"
  if (headingLevel && headingLevel >= 1 && headingLevel <= 6) return `h${headingLevel}`
  return "p"
}

function paragraphCss(paragraph: DocxClipboardParagraph): string {
  const declarations: string[] = ["margin:0"]
  const align = paragraph.style?.align
  if (align === "left" || align === "right" || align === "center" || align === "justify") {
    declarations.push(`text-align:${align}`)
  }
  const headingLevel = paragraph.style?.headingLevel
  if (paragraph.list) {
    const level = Math.max(0, Math.min(8, paragraph.list.level ?? 0))
    if (level > 0) declarations.push(`margin-left:${level * 24}px`)
  } else if (headingLevel && headingLevel >= 1 && headingLevel <= 6) {
    // The model derives heading size from the run; emit the same default so a
    // heading pasted into Word keeps its weight even without a style sheet.
    const size = HEADING_FONT_SIZES_PT[headingLevel - 1]
    declarations.push(`font-size:${size}pt`, "font-weight:bold")
  }
  return declarations.join(";")
}

// ── HTML serialization ─────────────────────────────────────────────

/** Plain text of clipboard paragraphs, one line per paragraph. */
export function clipboardTextFromParagraphs(
  paragraphs: readonly DocxClipboardParagraph[]
): string {
  return paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(""))
    .join("\n")
}

/** Splits plain text into clipboard paragraphs. */
export function clipboardParagraphsFromText(text: string): DocxClipboardParagraph[] {
  const normalized = text.replace(/\r\n?/g, "\n")
  if (!normalized) return []
  return normalized.split("\n").map((line) => ({
    runs: line ? [{ text: line }] : []
  }))
}

function renderClipboardRuns(
  runs: readonly DocxClipboardRun[],
  sanitizeLink?: (value: string) => string | undefined
): string {
  const parts: string[] = []
  for (const run of runs) {
    if (!run.text) continue
    // A link copied out of the document still has to pass the policy: the
    // clipboard is read by apps that will happily follow a `javascript:` URL.
    const link = run.link && sanitizeLink ? sanitizeLink(run.link) : run.link
    const segments = run.text.split("\n")
    const rendered: string[] = []
    segments.forEach((segment, index) => {
      if (index > 0) rendered.push("<br>")
      if (!segment) return
      const css = clipboardRunCss(run.style)
      if (link) {
        rendered.push(
          `<a href="${escapeClipboardAttribute(link)}"${
            css ? ` style="${escapeClipboardAttribute(css)}"` : ""
          }>${escapeClipboardHtml(segment)}</a>`
        )
        return
      }
      if (!css) {
        rendered.push(escapeClipboardHtml(segment))
        return
      }
      rendered.push(`<span style="${escapeClipboardAttribute(css)}">${escapeClipboardHtml(segment)}</span>`)
    })
    parts.push(rendered.join(""))
  }
  return parts.join("")
}

function listTag(type: DocxClipboardListInfo["type"]): string {
  return type === "ordered" ? "ol" : "ul"
}

/**
 * Renders clipboard paragraphs to the `text/html` flavour.
 *
 * Consecutive list paragraphs share one `<ul>`/`<ol>` so Word reconstructs the
 * list instead of a set of unrelated bullets.
 */
export function renderDocxClipboardHtml(
  paragraphs: readonly DocxClipboardParagraph[],
  options?: {
    maxChars?: number
    /** Drops links the destination app would otherwise follow. */
    sanitizeLink?: (value: string) => string | undefined
  }
): string {
  const maxChars = Math.max(1, options?.maxChars ?? MAX_DOCX_CLIPBOARD_HTML_CHARS)
  const parts: string[] = []
  let length = 0
  let openList: DocxClipboardListInfo["type"] | undefined

  const push = (chunk: string): boolean => {
    if (length + chunk.length > maxChars) return false
    parts.push(chunk)
    length += chunk.length
    return true
  }

  for (const paragraph of paragraphs) {
    const listType = paragraph.list?.type
    let chunk = ""
    if (listType) {
      if (openList !== listType) {
        if (openList) chunk += `</${listTag(openList)}>`
        chunk += `<${listTag(listType)}>`
        openList = listType
      }
    } else if (openList) {
      chunk += `</${listTag(openList)}>`
      openList = undefined
    }

    const tag = paragraphOpenTag(paragraph)
    const css = paragraphCss(paragraph)
    // A paragraph is emitted whole or not at all: a half-written open tag would
    // leave the `text/html` flavour unbalanced at the cap.
    chunk += `<${tag}${css ? ` style="${escapeClipboardAttribute(css)}"` : ""}>`
    chunk += renderClipboardRuns(paragraph.runs, options?.sanitizeLink)
    chunk += `</${tag}>`
    if (!push(chunk)) break
  }

  if (openList) push(`</${listTag(openList)}>`)
  return parts.join("")
}


export function sanitizeClipboardListInfo(value: unknown): DocxClipboardListInfo | undefined {
  if (!value || typeof value !== "object") return undefined
  const source = value as Record<string, unknown>
  if (source.type !== "ordered" && source.type !== "unordered") return undefined
  const level = finiteNumber(source.level, 0, 8) ?? 0
  return { type: source.type, level: Math.round(level) }
}

function sameStyle(left?: TextStyle, right?: TextStyle): boolean {
  return JSON.stringify(left ?? null) === JSON.stringify(right ?? null)
}

export function mergeAdjacentRuns(runs: readonly DocxClipboardRun[]): DocxClipboardRun[] {
  const merged: DocxClipboardRun[] = []
  for (const run of runs) {
    if (!run.text) continue
    const previous = merged[merged.length - 1]
    if (previous && previous.link === run.link && sameStyle(previous.style, run.style)) {
      previous.text += run.text
      continue
    }
    const copy: DocxClipboardRun = { text: run.text }
    if (run.style) copy.style = run.style
    if (run.link) copy.link = run.link
    merged.push(copy)
  }
  return merged
}

export function trimRunEdges(runs: DocxClipboardRun[]): DocxClipboardRun[] {
  while (runs.length > 0) {
    const first = runs[0]
    first.text = first.text.replace(/^[\s\u00a0]+/, "")
    if (first.text) break
    runs.shift()
  }
  while (runs.length > 0) {
    const last = runs[runs.length - 1]
    last.text = last.text.replace(/[\s\u00a0]+$/, "")
    if (last.text) break
    runs.pop()
  }
  return runs
}

/**
 * Sanitizes blocks that arrived as structured clipboard data.
 *
 * The same rules apply to `text/html` parsing, so a payload coming from our own
 * copy and a payload coming from Word cannot smuggle in unsupported styles or
 * unapproved links.
 */
export function sanitizeClipboardParagraphs(
  value: unknown,
  options?: DocxClipboardParseOptions
): DocxClipboardParagraph[] {
  if (!Array.isArray(value)) return []
  const maxParagraphs = Math.max(1, options?.maxParagraphs ?? MAX_DOCX_CLIPBOARD_PARAGRAPHS)
  const maxRuns = Math.max(1, options?.maxRuns ?? MAX_DOCX_CLIPBOARD_RUNS)
  const paragraphs: DocxClipboardParagraph[] = []
  let runCount = 0

  for (const item of value) {
    if (paragraphs.length >= maxParagraphs) break
    if (!item || typeof item !== "object") continue
    const source = item as Record<string, unknown>
    if (!Array.isArray(source.runs)) continue

    const runs: DocxClipboardRun[] = []
    for (const run of source.runs) {
      if (runCount >= maxRuns) break
      if (!run || typeof run !== "object") continue
      const runSource = run as Record<string, unknown>
      if (typeof runSource.text !== "string" || !runSource.text) continue
      const link =
        typeof runSource.link === "string" && options?.sanitizeLink
          ? options.sanitizeLink(runSource.link)
          : undefined
      const clipboardRun: DocxClipboardRun = { text: runSource.text }
      const style = sanitizeClipboardTextStyle(runSource.style)
      if (style) clipboardRun.style = style
      if (link) clipboardRun.link = link
      runs.push(clipboardRun)
      runCount += 1
    }

    const list = sanitizeClipboardListInfo(source.list)
    const paragraph: DocxClipboardParagraph = { runs: mergeAdjacentRuns(runs) }
    if (list) paragraph.list = list
    const styleSource = source.style
    if (styleSource && typeof styleSource === "object") {
      const style: ParagraphStyle = {}
      const sourceStyle = styleSource as Record<string, unknown>
      const align = sourceStyle.align
      if (align === "left" || align === "right" || align === "center" || align === "justify") {
        style.align = align
      }
      const headingLevel = finiteNumber(sourceStyle.headingLevel, 1, 6)
      if (headingLevel !== undefined) style.headingLevel = Math.round(headingLevel) as HeadingLevel
      if (Object.keys(style).length > 0) paragraph.style = style
    }
    paragraphs.push(paragraph)
  }

  return paragraphs
}


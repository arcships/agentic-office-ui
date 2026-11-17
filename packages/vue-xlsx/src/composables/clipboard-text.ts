// Clipboard primitives with no imports: TSV quoting/parsing plus the
// ClipboardItem -> writeText write ladder used when copying a range.
//
// TSV is only a lossless interchange format when a field holding a tab, a line
// break or a quote is quoted the way Excel writes it. Writing the raw display
// value made a cell with an embedded newline copy into several cells, and
// splitting pasted text on "\t"/"\n" mis-split any quoted external TSV, so the
// writer and the parser below have to agree on the same quoting rules.

const FIELD_SEPARATOR = "\t";
const ROW_SEPARATOR = "\r\n";
const QUOTE = '"';
const QUOTE_PATTERN = /[\t\n\r"]/;

/** Quotes a single TSV field when its text would otherwise split the grid. */
export function escapeTsvField(value: string): string {
  if (!QUOTE_PATTERN.test(value)) return value;
  return `${QUOTE}${value.replace(/"/g, `${QUOTE}${QUOTE}`)}${QUOTE}`;
}

/** Serializes a grid of display values into the `text/plain` clipboard flavour. */
export function serializeClipboardText(rows: readonly (readonly string[])[]): string {
  return rows
    .map((row) => row.map(escapeTsvField).join(FIELD_SEPARATOR))
    .join(ROW_SEPARATOR);
}

/** Parses the `text/plain` clipboard flavour, honouring quoted TSV fields. */
export function parseClipboardText(text: string): string[][] {
  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;

  const endField = () => {
    row.push(field);
    field = "";
  };
  const endRow = () => {
    endField();
    rows.push(row);
    row = [];
  };

  for (let index = 0; index < normalized.length; index += 1) {
    const character = normalized[index];
    if (quoted) {
      if (character !== QUOTE) {
        field += character;
      } else if (normalized[index + 1] === QUOTE) {
        field += QUOTE;
        index += 1;
      } else {
        quoted = false;
      }
      continue;
    }
    // A quote only opens a field at its start; elsewhere it is literal text.
    if (character === QUOTE && field === "") {
      quoted = true;
    } else if (character === FIELD_SEPARATOR) {
      endField();
    } else if (character === "\n") {
      endRow();
    } else {
      field += character;
    }
  }
  endRow();

  // A trailing separator terminates the last row instead of opening an empty one.
  if (normalized.endsWith("\n") && rows.length > 1) {
    const lastRow = rows[rows.length - 1];
    if (lastRow.length === 1 && lastRow[0] === "") rows.pop();
  }

  return rows;
}

export interface ClipboardWriteTarget {
  write?(items: unknown[]): Promise<void>;
  writeText?(text: string): Promise<void>;
}

export interface ClipboardWriteTargets {
  /** Custom MIME type carrying the structured range payload. */
  mimeType: string;
  clipboard?: ClipboardWriteTarget | null;
  /** ClipboardItem-compatible constructor; absent in browsers without ClipboardItem. */
  ClipboardItem?: (new (items: Record<string, Blob>) => unknown) | null;
}

export interface ClipboardWritePayload {
  html: string;
  structured: string;
  text: string;
}

/**
 * Writes a copied range to the clipboard, returning whether anything reached it.
 *
 * `write()` rejects as a whole when a single type is unsupported (Firefox
 * rejects custom MIME types) or when the permission is denied, so a rejected
 * composite write still falls back to the plain-text flavour. Callers get a
 * boolean instead of a rejection so a denied copy can be reported to the user.
 */
export async function writeClipboardPayload(
  payload: ClipboardWritePayload,
  targets: ClipboardWriteTargets,
): Promise<boolean> {
  const { clipboard } = targets;
  if (!clipboard) return false;

  if (targets.ClipboardItem && clipboard.write) {
    try {
      const item = new targets.ClipboardItem({
        [targets.mimeType]: new Blob([payload.structured], { type: targets.mimeType }),
        "text/html": new Blob([payload.html], { type: "text/html" }),
        "text/plain": new Blob([payload.text], { type: "text/plain" }),
      });
      await clipboard.write([item]);
      return true;
    } catch {
      // Fall through to the text flavour; the composite item is all-or-nothing.
    }
  }

  if (!clipboard.writeText) return false;
  try {
    await clipboard.writeText(payload.text);
    return true;
  } catch {
    return false;
  }
}

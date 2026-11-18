// Word segmentation for the document block model.
//
// The upstream parser segments with Unicode UAX #29 and keeps every
// non-whitespace segment, including punctuation and symbols. This port keeps the
// same "no dictionary" contract but additionally splits CJK runs per code point,
// because a whole Chinese sentence is not a useful citation/highlight unit.

import type { DocumentWord, DocxSourceRef } from "./types";

export interface SegmentedWord {
  text: string;
  /** Inclusive UTF-16 offset into the segmented text. */
  start: number;
  /** Exclusive UTF-16 offset into the segmented text. */
  end: number;
}

const CJK_PATTERN = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u;
const WHITESPACE_PATTERN = /\s/u;

/** True when the code point starts a new word on its own (CJK scripts). */
export function isCjkWordCharacter(value: string): boolean {
  return CJK_PATTERN.test(value);
}

/**
 * Splits text into offsets for word-level highlighting.
 *
 * - whitespace separates words and belongs to none;
 * - a CJK code point is its own word;
 * - combining marks stay with the preceding word;
 * - every other non-whitespace code point belongs to the current word, so
 *   punctuation and symbols are retained instead of dropped.
 */
export function segmentDocumentWords(text: string): SegmentedWord[] {
  const words: SegmentedWord[] = [];
  let wordStart = -1;
  let wordEnd = -1;
  let previousWasCjk = false;

  const flush = (): void => {
    if (wordStart < 0) return;
    words.push({ text: text.slice(wordStart, wordEnd), start: wordStart, end: wordEnd });
    wordStart = -1;
    wordEnd = -1;
    previousWasCjk = false;
  };

  for (let offset = 0; offset < text.length; ) {
    const codePoint = text.codePointAt(offset);
    if (codePoint === undefined) break;
    const value = String.fromCodePoint(codePoint);
    const nextOffset = offset + value.length;

    if (WHITESPACE_PATTERN.test(value)) {
      flush();
      offset = nextOffset;
      continue;
    }

    const combining = /^\p{M}/u.test(value);
    const cjk = isCjkWordCharacter(value);

    if (combining && wordStart >= 0 && !previousWasCjk) {
      wordEnd = nextOffset;
      offset = nextOffset;
      continue;
    }

    if (cjk) {
      flush();
      words.push({ text: value, start: offset, end: nextOffset });
      previousWasCjk = true;
      offset = nextOffset;
      continue;
    }

    if (wordStart < 0) {
      wordStart = offset;
      wordEnd = nextOffset;
    } else {
      wordEnd = nextOffset;
    }
    previousWasCjk = false;
    offset = nextOffset;
  }

  flush();
  return words;
}

/**
 * Builds the word objects for one paragraph block.
 *
 * Word ids end in `/word/{start}-{end}` using offsets in the block content, as
 * in the upstream contract, and each word carries the paragraph source
 * reference narrowed to its UTF-16 range.
 */
export function buildDocumentWords(
  blockId: string,
  content: string,
  ref: DocxSourceRef,
): DocumentWord[] {
  return segmentDocumentWords(content).map((word) => ({
    object: "word",
    id: `${blockId}/word/${word.start}-${word.end}`,
    blockId,
    content: word.text,
    sourceRefs: [
      {
        ...ref,
        start: word.start,
        end: word.end,
      },
    ],
  }));
}

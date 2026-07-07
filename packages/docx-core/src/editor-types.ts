// Editor selection and text range types extracted from editor.tsx.
// These types are shared between editor state (core/state.ts) and
// the editor component, which lives in vue-docx.

export type DocxEditorSelection =
  | {
      kind: "paragraph";
      nodeIndex: number;
    }
  | {
      kind: "table-cell";
      tableIndex: number;
      rowIndex: number;
      cellIndex: number;
    };

interface DocxTableCellLocation {
  tableIndex: number;
  rowIndex: number;
  cellIndex: number;
}

interface DocxTableCellSelectionRange {
  tableIndex: number;
  anchorRowIndex: number;
  anchorCellIndex: number;
  focusRowIndex: number;
  focusCellIndex: number;
}

interface ParagraphLocationInBody {
  kind: "paragraph";
  nodeIndex: number;
}

interface ParagraphLocationInCell {
  kind: "table-cell";
  tableIndex: number;
  rowIndex: number;
  cellIndex: number;
  paragraphIndex: number;
}

export type DocxTextRangeLocation =
  | ParagraphLocationInBody
  | ParagraphLocationInCell;

export interface DocxTextRangeBoundary {
  location: DocxTextRangeLocation;
  offset: number;
}

export interface DocxTextRange {
  start: DocxTextRangeBoundary;
  end: DocxTextRangeBoundary;
}

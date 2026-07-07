// Default blank-document model factory.
// Upstream editor.tsx: lines 4546-4644.

import type { DocModel } from "../../engine/types";
import { cloneDocModel } from "../../engine/clone";

export const defaultStarterModel: DocModel = {
  nodes: [
    {
      type: "paragraph",
      children: [{ type: "text", text: "" }],
    },
  ],
  metadata: {
    sourceParts: 1,
    warnings: [],
    headerSections: [],
    footerSections: [],
    paragraphStyles: [
      {
        id: "Normal",
        name: "Body",
        isDefault: true,
        // Word's blank-document default body typeface. Without this the body
        // default fell through to a heading font / Times New Roman and did not
        // match the "Calibri" shown in the toolbar.
        runStyle: { fontFamily: "Calibri", fontSizePt: 11 },
      },
      {
        id: "Heading1",
        name: "Heading 1",
        headingLevel: 1,
        isPrimary: true,
        runStyle: {
          fontFamily: "Calibri Light",
          fontSizePt: 16,
          bold: true,
          color: "#2f5496",
        },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        headingLevel: 2,
        isPrimary: true,
        runStyle: {
          fontFamily: "Calibri Light",
          fontSizePt: 13,
          bold: true,
          color: "#2f5496",
        },
      },
      {
        id: "Heading3",
        name: "Heading 3",
        headingLevel: 3,
        isPrimary: true,
        runStyle: {
          fontFamily: "Calibri",
          fontSizePt: 12,
          bold: true,
          color: "#1f3763",
        },
      },
      {
        id: "Heading4",
        name: "Heading 4",
        headingLevel: 4,
        runStyle: {
          fontFamily: "Calibri",
          fontSizePt: 11,
          bold: true,
          color: "#1f3763",
        },
      },
      {
        id: "Heading5",
        name: "Heading 5",
        headingLevel: 5,
        runStyle: {
          fontFamily: "Calibri",
          fontSizePt: 11,
          bold: true,
          color: "#1f3763",
        },
      },
      {
        id: "Heading6",
        name: "Heading 6",
        headingLevel: 6,
        runStyle: {
          fontFamily: "Calibri",
          fontSizePt: 11,
          bold: true,
          color: "#1f3763",
        },
      },
    ],
    defaultParagraphStyleId: "Normal",
  },
};

export function createBlankDocumentModel(): DocModel {
  return cloneDocModel(defaultStarterModel);
}

import type {
  DocumentNoteDefinition,
  FooterSection,
  FormFieldRunNode,
  HeaderSection,
  HeadingLevel,
  ImageRunNode,
  NumberingDefinitionSet,
  NumberingLevelDefinition,
  ParagraphAlignment,
  ParagraphBorderSet,
  ParagraphBorderStyle,
  ParagraphIndent,
  ParagraphNode,
  ParagraphStyleDefinition,
  TableBorderSet,
  TableBorderStyle,
  TableCellContentNode,
  TableCellStyle,
  TableNode,
  TableRowStyle,
  TextRunNode,
  DocModel,
} from "../../engine/types";
// TODO: replace with local ooxml package type once migrated into docx-core.
import type { DocumentLayoutMetrics } from "../../viewer/section-layout";
import { twipsToPixels } from "../../viewer/section-layout";
import {
  collectTableExplicitPageBreakInfo,
  collectTopLevelExplicitPageBreakStartNodeIndexes,
} from "../../viewer/pagination-breaks";
import {
  reconcilePageCountCandidateToTargetCountByScalingHeight,
  resolveMeasuredBodyFooterOverlapLatchState,
  shouldAllowStoredPageCountReduction,
} from "../../viewer/page-count-reconciliation";
import {
  splitParagraphChildrenAtTextOffsets,
  updateParagraphText,
  updateTableCellParagraphText,
  updateTableCellParagraphTextRecursive,
  updateTableCellText,
} from "../editor-ops";


export {
  collectTableExplicitPageBreakInfo,
  collectTopLevelExplicitPageBreakStartNodeIndexes,
  reconcilePageCountCandidateToTargetCountByScalingHeight,
  resolveMeasuredBodyFooterOverlapLatchState,
  shouldAllowStoredPageCountReduction,
  splitParagraphChildrenAtTextOffsets,
  updateParagraphText,
  updateTableCellParagraphText,
  updateTableCellParagraphTextRecursive,
  updateTableCellText,
  twipsToPixels,
};

export type {
  DocumentNoteDefinition,
  FooterSection,
  FormFieldRunNode,
  HeaderSection,
  HeadingLevel,
  ImageRunNode,
  NumberingDefinitionSet,
  NumberingLevelDefinition,
  ParagraphAlignment,
  ParagraphBorderSet,
  ParagraphBorderStyle,
  ParagraphIndent,
  ParagraphNode,
  ParagraphStyleDefinition,
  TableBorderSet,
  TableBorderStyle,
  TableCellContentNode,
  TableCellStyle,
  TableNode,
  TableRowStyle,
  TextRunNode,
  DocModel,
  DocumentLayoutMetrics,
};

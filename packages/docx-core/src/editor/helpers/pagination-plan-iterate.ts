import type {
  DocModel,
  ParagraphNode,
  TableNode,
  TableRowStyle,
  TableCellStyle,
  TableCellContentNode,
  ParagraphAlignment,
  ParagraphBorderSet,
  ParagraphBorderStyle,
  ParagraphIndent,
  ParagraphStyleDefinition,
  TableBorderSet,
  TableBorderStyle,
  TextRunNode,
  ImageRunNode,
  HeadingLevel,
  DocumentNoteDefinition,
  HeaderSection,
  FooterSection,
  FormFieldRunNode,
  NumberingDefinitionSet,
  NumberingLevelDefinition,
} from "../../engine/types";
import type { DocumentLayoutMetrics } from "../../viewer/section-layout";
import { twipsToPixels } from "../../viewer/section-layout";

export type CSSProperties = Record<string, string | number | undefined>;

// TODO: locate upstream pagination-plan-iterate implementation in the split plan docs.
export interface PaginationPlanIterateResult {
  model: DocModel;
  layout: DocumentLayoutMetrics;
  measuredHeightPx: number;
  pageCount: number;
}

export interface PaginationPlanIterateParams {
  model: DocModel;
  layout: DocumentLayoutMetrics;
  initialPageCount?: number;
}

export function iteratePaginationPlan(
  params: PaginationPlanIterateParams
): PaginationPlanIterateResult {
  const pageCount = Math.max(1, Math.round(params.initialPageCount ?? 1));
  const measuredHeightPx = twipsToPixels(pageCount * 240) ?? pageCount * 16;

  return {
    model: params.model,
    layout: params.layout,
    measuredHeightPx,
    pageCount,
  };
}

export type PaginationPlanIterateParagraph = ParagraphNode;
export type PaginationPlanIterateTable = TableNode;
export type PaginationPlanIterateTableRowStyle = TableRowStyle;
export type PaginationPlanIterateTableCellStyle = TableCellStyle;
export type PaginationPlanIterateTableCellContentNode = TableCellContentNode;
export type PaginationPlanIterateParagraphAlignment = ParagraphAlignment;
export type PaginationPlanIterateParagraphBorderSet = ParagraphBorderSet;
export type PaginationPlanIterateParagraphBorderStyle = ParagraphBorderStyle;
export type PaginationPlanIterateParagraphIndent = ParagraphIndent;
export type PaginationPlanIterateParagraphStyleDefinition = ParagraphStyleDefinition;
export type PaginationPlanIterateTableBorderSet = TableBorderSet;
export type PaginationPlanIterateTableBorderStyle = TableBorderStyle;
export type PaginationPlanIterateTextRunNode = TextRunNode;
export type PaginationPlanIterateImageRunNode = ImageRunNode;
export type PaginationPlanIterateHeadingLevel = HeadingLevel;
export type PaginationPlanIterateDocumentNoteDefinition = DocumentNoteDefinition;
export type PaginationPlanIterateHeaderSection = HeaderSection;
export type PaginationPlanIterateFooterSection = FooterSection;
export type PaginationPlanIterateFormFieldRunNode = FormFieldRunNode;
export type PaginationPlanIterateNumberingDefinitionSet = NumberingDefinitionSet;
export type PaginationPlanIterateNumberingLevelDefinition = NumberingLevelDefinition;

export const paginationPlanIterateStyles: CSSProperties = {
  display: "block",
  position: "relative",
};

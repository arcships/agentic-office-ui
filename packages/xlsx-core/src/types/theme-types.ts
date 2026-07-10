/** Pure workbook theme data shared by core color helpers and runtime types. */
export interface XlsxThemePalette {
  colorsByIndex: Record<number, string>;
  majorLatinFont?: string;
  minorLatinFont?: string;
}

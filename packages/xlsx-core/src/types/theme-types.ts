/** Pure workbook theme data shared by core color helpers and runtime types. */
export interface XlsxThemePalette {
  colorsByIndex: Record<number, string>;
  /**
   * Effective legacy indexed colour palette, keyed by OOXML `indexed` value.
   * `buildThemePalette()` fills the ECMA-376 defaults and applies the
   * workbook's `<colors><indexedColors>` overrides from `xl/styles.xml`.
   */
  indexedColors?: Record<number, string>;
  majorLatinFont?: string;
  minorLatinFont?: string;
}

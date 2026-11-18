// xlsx-file-types.ts — Accepted workbook container policy shared by the file
// picker, drag-and-drop import and the controller loader.
//
// The list mirrors what the XLSX loader can actually parse: OOXML workbooks
// and templates, legacy binary `.xls`/`.xlsb`, and delimited `.csv` text.

export const XLSX_FILE_EXTENSIONS = [
  ".xlsx",
  ".xls",
  ".xlsm",
  ".xlsb",
  ".xltx",
  ".xltm",
  ".csv",
] as const

export const XLSX_FILE_MIME_TYPES = [
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel",
  "application/vnd.ms-excel.sheet.macroEnabled.12",
  "application/vnd.ms-excel.sheet.binary.macroEnabled.12",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
  "application/vnd.ms-excel.template.macroEnabled.12",
  "text/csv",
] as const

/** Value for an `<input type="file">` accept attribute. */
export const XLSX_FILE_ACCEPT = [...XLSX_FILE_EXTENSIONS, ...XLSX_FILE_MIME_TYPES].join(",")

/** Lower-case extension including the leading dot, or undefined. */
export function xlsxFileExtension(fileName: string): string | undefined {
  const match = /(\.[^.\\/]+)$/.exec(fileName.trim())
  return match?.[1]?.toLowerCase()
}

export function isSupportedXlsxFileName(fileName: string): boolean {
  const extension = xlsxFileExtension(fileName)
  return extension !== undefined &&
    (XLSX_FILE_EXTENSIONS as readonly string[]).includes(extension)
}

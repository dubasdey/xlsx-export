/**
 * Supported Excel cell values.
 *
 * - string: Text values
 * - number: Numeric values
 * - boolean: Boolean values
 * - Date: Date values serialized as ISO strings
 * - null / undefined: Empty cells
 */
export type ExcelValue = string | number | boolean | Date | null | undefined;

/**
 * A single row of Excel data where object keys map to column headers.
 */
export type ExcelRow = Record<string, ExcelValue>;

/**
 * Export options for Excel generation.
 */
export interface ExcelOptions {
  /** Optional sheet name for the exported workbook. */
  sheetName?: string;

  /** Optional file name for the download. */
  filename?: string;

  /** Optional explicit column order. */
  columns?: string[];
}

/**
 * Internal ZIP entry representation used by the XLSX builder.
 */
export interface ZipEntry {
  name: string;
  data: Uint8Array;
}

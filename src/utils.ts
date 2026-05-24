import { ExcelRow } from './types';

/**
 * Determine the column order for the worksheet.
 *
 * If explicit columns are provided, they are returned as-is.
 * Otherwise the function collects keys from the row data in insertion order.
 */
export function normalizeColumns(rows: ExcelRow[], explicitColumns?: string[]): string[] {
  if (explicitColumns && explicitColumns.length > 0) {
    return explicitColumns;
  }

  const order: string[] = [];
  const seen = new Set<string>();

  rows.forEach((row) => {
    Object.keys(row).forEach((key) => {
      if (!seen.has(key)) {
        seen.add(key);
        order.push(key);
      }
    });
  });

  return order;
}

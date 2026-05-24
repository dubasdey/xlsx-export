import { ExcelOptions, ExcelRow } from './types';
import { downloadXlsx, generateXlsx } from './generator';

/**
 * Service that exposes Excel export helpers for Angular applications.
 */
export class ExcelExportService {
  /**
   * Generates an XLSX Blob from the provided row data.
   *
   * @param rows Array of row objects where keys become column headers.
   * @param options Export options including worksheet name and file name.
   */
  generate(rows: ExcelRow[], options: ExcelOptions = {}): Blob {
    return generateXlsx(rows, options);
  }

  /**
   * Triggers the browser download of the generated XLSX file.
   *
   * @param rows Array of row objects where keys become column headers.
   * @param options Export options including worksheet name and file name.
   */
  save(rows: ExcelRow[], options: ExcelOptions = {}): void {
    downloadXlsx(rows, options);
  }
}

import { ExcelOptions, ExcelRow } from './types';
import {
  buildAppPropertiesXml,
  buildContentTypesXml,
  buildCorePropertiesXml,
  buildRelsRootXml,
  buildStylesXml,
  buildWorkbookXml,
  buildWorkbookRelsXml,
  buildWorksheetXml,
} from './xml';
import { normalizeColumns } from './utils';
import { buildZip, encodeUtf8 } from './zip';

/**
 * Generates an XLSX file blob from row data.
 *
 * @param rows Spreadsheet rows where keys map to column headers.
 * @param options Optional export settings for file name and sheet name.
 */
export function generateXlsx(rows: ExcelRow[], options: ExcelOptions = {}): Blob {
  const sheetName = options.sheetName ?? 'Sheet1';
  const columns = normalizeColumns(rows, options.columns);
  const zipEntries = [
    { name: '[Content_Types].xml', data: encodeUtf8(buildContentTypesXml()) },
    { name: '_rels/.rels', data: encodeUtf8(buildRelsRootXml()) },
    { name: 'xl/workbook.xml', data: encodeUtf8(buildWorkbookXml(sheetName)) },
    { name: 'xl/_rels/workbook.xml.rels', data: encodeUtf8(buildWorkbookRelsXml()) },
    { name: 'xl/worksheets/sheet1.xml', data: encodeUtf8(buildWorksheetXml(columns, rows)) },
    { name: 'xl/styles.xml', data: encodeUtf8(buildStylesXml()) },
    { name: 'docProps/core.xml', data: encodeUtf8(buildCorePropertiesXml()) },
    { name: 'docProps/app.xml', data: encodeUtf8(buildAppPropertiesXml(sheetName)) },
  ];
  const zip = buildZip(zipEntries);
  return new Blob([zip.buffer as ArrayBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

/**
 * Creates a browser download for the generated XLSX blob.
 *
 * @param rows Spreadsheet rows where keys map to column headers.
 * @param options Optional export settings for file name and sheet name.
 */
export function downloadXlsx(rows: ExcelRow[], options: ExcelOptions = {}): void {
  const filename = options.filename ?? 'export.xlsx';
  const blob = generateXlsx(rows, options);
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');

  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

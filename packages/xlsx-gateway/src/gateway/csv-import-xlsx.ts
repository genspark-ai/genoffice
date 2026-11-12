/**
 * The archive half of the CSV import: a delimited text becomes an xlsx.
 *
 * Desktop only. The parser lives in csv-import, which the sheets renderer
 * imports as well (Data › From Text/CSV), and a browser bundle cannot carry
 * a node stream — so the part that writes the archive stays here, where the
 * node APIs are.
 */

import { Readable } from 'node:stream'

import JSZip from 'jszip'

import {
  csvGridShape,
  csvRows,
  escapeXml,
  resolveImportDelimiter,
  worksheetXmlChunks,
  type CsvGridShape,
  type CsvImportOptions,
} from './csv-import'
import { DEFAULT_THEME_XML } from './xlsx-default-theme'
import { MINIMAL_STYLESHEET_XML } from './xlsx-default-styles'
import { validateSheetName } from './xlsx-sheets'

export async function csvToXlsxBuffer(
  csvText: string,
  sheetName = 'Sheet1',
  options?: CsvImportOptions,
): Promise<Buffer> {
  const delimiter = resolveImportDelimiter(csvText)
  const shape = csvGridShape(csvText, delimiter, options)
  if (shape.rows === 0) throw new Error('The CSV file has no data rows.')
  return xlsxBufferFromGrid(csvRows(csvText, delimiter), shape, sheetName, options)
}

/**
 * Open-path conversion. `delimiter` pins the split for callers that already
 * know the format (a .tsv); leaving it unset lets the sniffer and its
 * prose-shatter guard decide, which is right for a bare .csv but unreliable
 * for a tab-delimited file whose fields hold enough commas to out-count tabs.
 */
export async function csvToXlsxBufferForOpen(
  csvText: string,
  sheetName = 'Sheet1',
  delimiter?: string,
  options?: CsvImportOptions,
): Promise<{ buffer: Buffer; empty: boolean }> {
  const resolved = delimiter ?? resolveImportDelimiter(csvText)
  const shape = csvGridShape(csvText, resolved, options)
  const buffer = await xlsxBufferFromGrid(csvRows(csvText, resolved), shape, sheetName, options)
  return { buffer, empty: shape.rows === 0 }
}

/**
 * xlsx from the app's OWN comma-serialized sheet grid (AI create_document):
 * unlike the import path above, the delimiter is fixed to comma — cell text
 * may legitimately hold more semicolons/tabs than commas (csvField quotes
 * neither), and sniffing would then split the wrong columns — and an
 * all-empty grid becomes a valid blank workbook instead of an import error.
 */
export async function sheetCsvToXlsxBuffer(
  csvText: string,
  sheetName = 'Sheet1',
  options?: CsvImportOptions,
): Promise<Buffer> {
  const shape = csvGridShape(csvText, ',', options)
  return xlsxBufferFromGrid(csvRows(csvText, ','), shape, sheetName, options)
}

/** minimal empty workbook: the backing file for a "new blank spreadsheet" tab */
export async function blankXlsxBuffer(sheetName = 'Sheet1'): Promise<Buffer> {
  const shape: CsvGridShape = { rows: 0, columns: 1 }
  return xlsxBufferFromGrid([], shape, sheetName)
}

/**
 * The workbook for a grid, with the worksheet streamed into the archive: the
 * XML never exists as one string, and the archive takes it as it comes.
 */
async function xlsxBufferFromGrid(
  rows: Iterable<readonly string[]>,
  shape: CsvGridShape,
  sheetName: string,
  options?: CsvImportOptions,
): Promise<Buffer> {
  validateSheetName(sheetName)
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<sheets><sheet name="${escapeXml(sheetName)}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>',
  )
  zip.file('xl/theme/theme1.xml', DEFAULT_THEME_XML)
  zip.file('xl/styles.xml', MINIMAL_STYLESHEET_XML)
  zip.file('xl/worksheets/sheet1.xml', worksheetXmlStream(rows, shape, options))
  return zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    // the worksheet is a stream: without this the archive would collect it in
    // memory before compressing it, which is the cost the import cannot afford
    streamFiles: true,
  })
}

/**
 * The worksheet as a byte stream. The chunks go in as UTF-8, the encoding the
 * worksheet declares: the archive turns a string into UTF-16 bytes, which
 * would leave the XML unreadable.
 */
function worksheetXmlStream(
  rows: Iterable<readonly string[]>,
  shape: CsvGridShape,
  options?: CsvImportOptions,
): Readable {
  return Readable.from(
    (function* (): Generator<Buffer> {
      // The archive reads the worksheet as it compresses it, so the rows
      // arrive at the compressor's pace: the row count is the progress, and an
      // abort lands while the compression is still running (the error the
      // generator throws ends the archive).
      const chunks = worksheetXmlChunks(rows, shape, {
        onRow: (rowsWritten) => {
          options?.onProgress?.({
            phase: 'write',
            fraction: Math.min(rowsWritten / Math.max(shape.rows, 1), 1),
          })
        },
      })
      for (const chunk of chunks) {
        options?.signal?.throwIfAborted()
        yield Buffer.from(chunk, 'utf8')
      }
    })(),
  )
}

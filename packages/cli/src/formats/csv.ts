import {
  csvToXlsxBuffer,
  decodeCsvBuffer,
  parseCsv,
  sniffDelimiter,
} from '@genoffice/xlsx-gateway/gateway/csv-import'

export interface CsvInfo {
  rows: number
  columns: number
  delimiter: string
}

export function csvInfo(bytes: Uint8Array): CsvInfo {
  const text = decodeCsvBuffer(bytes)
  const delimiter = sniffDelimiter(text)
  const rows = parseCsv(text, delimiter)
  return {
    rows: rows.length,
    columns: rows.reduce((max, r) => Math.max(max, r.length), 0),
    delimiter,
  }
}

export async function csvToXlsx(bytes: Uint8Array, sheetName: string): Promise<Uint8Array> {
  const buf = await csvToXlsxBuffer(decodeCsvBuffer(bytes), sheetName)
  return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength)
}

/** Excel sheet name from a file stem: forbidden characters become `_`, edge apostrophes go, 31 chars max. */
export function sheetNameFromStem(stem: string): string {
  const name = stem
    .replace(/[\\/?*[\]:]/g, '_')
    .replace(/^'+|'+$/g, '')
    .slice(0, 31)
    .replace(/'+$/, '')
  return name || 'Sheet1'
}

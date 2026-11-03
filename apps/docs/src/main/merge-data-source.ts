import { readFile } from 'node:fs/promises'
import { basename, extname } from 'node:path'
import { mergeDataFromRows } from '@genoffice/docx-engine'
import { decodeCsvBuffer, parseCsv } from '@genoffice/xlsx-gateway/gateway/csv-import'
import { readSheetNames, readSheetRows } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { MergeDataSource, MergeDataSourceFile } from '../shared/ipc'

/**
 * The data source of a mail merge: the local spreadsheet the records come
 * from. The file is read here, in the main process, because that is where the
 * file system is; the renderer only ever sees the records.
 *
 * The parsing itself belongs to xlsx-gateway (the CSV reader Sheets opens a
 * .csv with, the workbook reader), and the record shape to the docx engine —
 * so a list that opens correctly in Sheets opens correctly here.
 * genoffice#2097.
 */

/** the file extensions a data source may have */
export const MERGE_DATA_SOURCE_EXTENSIONS = ['xlsx', 'csv'] as const

/**
 * The file a recipient list comes from, and the sheets it offers.
 *
 * The sheets are listed before one of them is read: a workbook holds several,
 * and the user picks the one the merge draws on. A csv has no sheets.
 */
export async function readMergeDataSourceFile(path: string): Promise<MergeDataSourceFile> {
  const extension = extname(path).toLowerCase()
  const sheets = extension === '.xlsx' ? await readSheetNames(await readFile(path)) : []
  return { path, name: basename(path), sheets }
}

/**
 * Read a local .xlsx or .csv as a mail merge data source.
 *
 * The chosen sheet of a workbook (the first one by default) and the first row
 * of a CSV name the fields; the rows below them are the records.
 */
export async function readMergeDataSource(
  path: string,
  sheetName?: string,
): Promise<MergeDataSource> {
  const extension = extname(path).toLowerCase()
  const bytes = await readFile(path)
  let rows: string[][] | null = null
  let sheet = ''
  if (extension === '.csv') {
    rows = parseCsv(decodeCsvBuffer(bytes))
  } else if (extension === '.xlsx') {
    const names = await readSheetNames(bytes)
    sheet = sheetName && names.includes(sheetName) ? sheetName : (names[0] ?? '')
    rows = await readSheetRows(bytes, sheet || undefined)
  }
  if (!rows) throw new Error(`unsupported data source: ${extension || path}`)
  const data = mergeDataFromRows(namedHeaders(rows))
  return {
    path,
    name: basename(path),
    sheet,
    fields: [...data.fields],
    records: [...data.records],
  }
}

/**
 * A header cell that is empty names no field, and two cells with the same
 * name would fight over one value: the second becomes "Name2", the way Word
 * names them.
 */
function namedHeaders(rows: readonly (readonly string[])[]): readonly (readonly string[])[] {
  const header = rows[0]
  if (!header) return []
  const seen = new Map<string, number>()
  const names = header.map((raw, index) => {
    const name = raw.trim() || `Column${index + 1}`
    const count = seen.get(name) ?? 0
    seen.set(name, count + 1)
    return count === 0 ? name : `${name}${count + 1}`
  })
  return [names, ...rows.slice(1)]
}

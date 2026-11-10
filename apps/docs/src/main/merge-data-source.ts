import { readFile } from 'node:fs/promises'
import { basename, extname } from 'node:path'
import { mergeDataFromRows, type MergeRecord } from '@genoffice/docx-engine'
import { decodeCsvBuffer, parseCsv } from '@genoffice/xlsx-gateway/gateway/csv-import'
import { readSheetRows } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'

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

/** A data source the mail merge reads: the field names and the records. */
export interface MergeDataSource {
  /** absolute path of the file the records came from */
  path: string
  /** the file's name, for the UI */
  name: string
  /** the column names, in file order */
  fields: string[]
  /** one entry per record */
  records: MergeRecord[]
}

/** the file extensions a data source may have */
export const MERGE_DATA_SOURCE_EXTENSIONS = ['xlsx', 'csv'] as const

/**
 * Read a local .xlsx or .csv as a mail merge data source.
 *
 * The first sheet of a workbook and the first row of a CSV name the fields;
 * the rows below them are the records.
 */
export async function readMergeDataSource(path: string): Promise<MergeDataSource> {
  const extension = extname(path).toLowerCase()
  const bytes = await readFile(path)
  const rows =
    extension === '.csv'
      ? parseCsv(decodeCsvBuffer(bytes))
      : extension === '.xlsx'
        ? await readSheetRows(bytes)
        : null
  if (!rows) throw new Error(`unsupported data source: ${extension || path}`)
  const data = mergeDataFromRows(namedHeaders(rows))
  return {
    path,
    name: basename(path),
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

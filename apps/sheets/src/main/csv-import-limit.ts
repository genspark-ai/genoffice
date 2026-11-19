/**
 * The ceiling a delimited import (a .csv/.tsv opened as a workbook) may
 * reach. The shell owns the setting and pushes the user's choice; the core's
 * default is what an import gets until it does.
 */

import { CSV_IMPORT_DEFAULT_LIMIT_MB } from '@genoffice/xlsx-gateway/gateway/csv-limits'

let limitBytes = CSV_IMPORT_DEFAULT_LIMIT_MB * 1024 * 1024

/** the ceiling a delimited import may reach, in bytes */
export function delimitedImportLimit(): number {
  return limitBytes
}

/**
 * the user's ceiling for a delimited import, in bytes. A value that is not a
 * whole number of bytes is refused: the ceiling is what stands between a text
 * file and the memory it takes to convert it, so it never falls back to
 * "no limit" by accident.
 */
export function setDelimitedImportLimit(bytes: number): void {
  if (!Number.isInteger(bytes) || bytes <= 0) return
  limitBytes = bytes
}

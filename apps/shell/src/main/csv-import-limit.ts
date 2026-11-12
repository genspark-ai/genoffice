/**
 * The ceiling a delimited import (a .csv/.tsv opened as a workbook) may
 * reach, in MB. The setting lives in userData/app-settings.json; the editors
 * take the value through setDelimitedImportLimit.
 */

import { CSV_IMPORT_DEFAULT_LIMIT_MB, CSV_IMPORT_LIMIT_CHOICES_MB } from '../shared/home-api'

/**
 * The ceiling from the stored settings: the user's choice when it is one of
 * the offered steps, the default otherwise. A value the setting does not offer
 * is not honoured — the choices are what the editors were told about.
 */
export function csvImportLimitMb(settings: Record<string, unknown>): number {
  const saved = settings.csvImportLimitMb
  const mb = typeof saved === 'number' ? saved : CSV_IMPORT_DEFAULT_LIMIT_MB
  return CSV_IMPORT_LIMIT_CHOICES_MB.includes(mb as (typeof CSV_IMPORT_LIMIT_CHOICES_MB)[number])
    ? mb
    : CSV_IMPORT_DEFAULT_LIMIT_MB
}

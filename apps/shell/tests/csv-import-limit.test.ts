import { describe, expect, it } from 'vitest'

import { csvImportLimitMb } from '../src/main/csv-import-limit'
import { CSV_IMPORT_DEFAULT_LIMIT_MB, CSV_IMPORT_LIMIT_CHOICES_MB } from '../src/shared/home-api'

/**
 * The ceiling a delimited import (a .csv/.tsv opened as a workbook) may
 * reach: the user's choice from the settings, the default when there is none.
 */

describe('csvImportLimitMb', () => {
  it('offers the default among the steps the setting presents', () => {
    expect(CSV_IMPORT_LIMIT_CHOICES_MB).toContain(CSV_IMPORT_DEFAULT_LIMIT_MB)
  })

  it('falls back to the default when the setting says nothing', () => {
    expect(csvImportLimitMb({})).toBe(CSV_IMPORT_DEFAULT_LIMIT_MB)
  })

  it('honours a step the setting offers', () => {
    for (const mb of CSV_IMPORT_LIMIT_CHOICES_MB) {
      expect(csvImportLimitMb({ csvImportLimitMb: mb })).toBe(mb)
    }
  })

  it('refuses a value the setting does not offer', () => {
    expect(csvImportLimitMb({ csvImportLimitMb: 0 })).toBe(CSV_IMPORT_DEFAULT_LIMIT_MB)
    expect(csvImportLimitMb({ csvImportLimitMb: -1 })).toBe(CSV_IMPORT_DEFAULT_LIMIT_MB)
    expect(csvImportLimitMb({ csvImportLimitMb: 12.5 })).toBe(CSV_IMPORT_DEFAULT_LIMIT_MB)
    expect(csvImportLimitMb({ csvImportLimitMb: '512' })).toBe(CSV_IMPORT_DEFAULT_LIMIT_MB)
  })
})

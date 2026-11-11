import { describe, expect, it } from 'vitest'

import { CSV_IMPORT_DEFAULT_LIMIT_MB } from '@genoffice/xlsx-gateway/gateway/csv-limits'

import { delimitedImportLimit, setDelimitedImportLimit } from '../src/main/csv-import-limit'

/**
 * The ceiling a delimited import may reach (a .csv/.tsv opened as a
 * workbook). The shell owns the setting and pushes the user's choice here.
 */

describe('the delimited import ceiling', () => {
  it('starts at the core default', () => {
    expect(delimitedImportLimit()).toBe(CSV_IMPORT_DEFAULT_LIMIT_MB * 1024 * 1024)
  })

  it('takes the ceiling the shell pushes', () => {
    setDelimitedImportLimit(512 * 1024 * 1024)
    expect(delimitedImportLimit()).toBe(512 * 1024 * 1024)
    setDelimitedImportLimit(CSV_IMPORT_DEFAULT_LIMIT_MB * 1024 * 1024)
  })

  it('refuses a ceiling that is not a whole number of bytes', () => {
    const before = delimitedImportLimit()
    setDelimitedImportLimit(0)
    expect(delimitedImportLimit()).toBe(before)
    setDelimitedImportLimit(-1)
    expect(delimitedImportLimit()).toBe(before)
    setDelimitedImportLimit(1.5)
    expect(delimitedImportLimit()).toBe(before)
  })
})

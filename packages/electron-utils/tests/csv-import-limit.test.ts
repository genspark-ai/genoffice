import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import {
  CSV_IMPORT_MAX_BYTES_KEY,
  DEFAULT_CSV_IMPORT_MAX_BYTES,
  MAX_CSV_IMPORT_MAX_BYTES,
  MIN_CSV_IMPORT_MAX_BYTES,
  csvImportAllowed,
  formatCsvImportBytes,
  readCsvImportMaxBytes,
  sanitizeCsvImportMaxBytes,
  writeCsvImportMaxBytes,
} from '../src/csv-import-limit'

const settingsFile = (content: string): string => {
  const path = join(mkdtempSync(join(tmpdir(), 'csv-import-limit-')), 'app-settings.json')
  writeFileSync(path, content)
  return path
}

describe('sanitizeCsvImportMaxBytes', () => {
  it('keeps a reasonable value', () => {
    expect(sanitizeCsvImportMaxBytes(64 * 1024 * 1024)).toBe(64 * 1024 * 1024)
  })

  it('defaults on anything unusable', () => {
    for (const value of [undefined, null, '128MB', NaN, Infinity, {}]) {
      expect(sanitizeCsvImportMaxBytes(value), String(value)).toBe(DEFAULT_CSV_IMPORT_MAX_BYTES)
    }
  })

  it('clamps to the 1 MiB … 512 MiB range', () => {
    expect(sanitizeCsvImportMaxBytes(0)).toBe(MIN_CSV_IMPORT_MAX_BYTES)
    expect(sanitizeCsvImportMaxBytes(1024)).toBe(MIN_CSV_IMPORT_MAX_BYTES)
    expect(sanitizeCsvImportMaxBytes(10 * 1024 * 1024 * 1024)).toBe(MAX_CSV_IMPORT_MAX_BYTES)
  })

  it('truncates fractions', () => {
    expect(sanitizeCsvImportMaxBytes(32 * 1024 * 1024 + 0.9)).toBe(32 * 1024 * 1024)
  })
})

describe('readCsvImportMaxBytes', () => {
  it('reads the configured cap', () => {
    expect(
      readCsvImportMaxBytes(settingsFile(JSON.stringify({ [CSV_IMPORT_MAX_BYTES_KEY]: 40 * 1024 * 1024 }))),
    ).toBe(40 * 1024 * 1024)
  })

  it('falls back to the default for a missing key, malformed file or missing file', () => {
    expect(readCsvImportMaxBytes(settingsFile('{}'))).toBe(DEFAULT_CSV_IMPORT_MAX_BYTES)
    expect(readCsvImportMaxBytes(settingsFile('not json'))).toBe(DEFAULT_CSV_IMPORT_MAX_BYTES)
    expect(readCsvImportMaxBytes('/nonexistent/app-settings.json')).toBe(DEFAULT_CSV_IMPORT_MAX_BYTES)
  })
})

describe('writeCsvImportMaxBytes', () => {
  it('persists the clamped value without dropping other settings', () => {
    const path = settingsFile('{"language":"en"}')
    const written = writeCsvImportMaxBytes(path, 10 * 1024 * 1024 * 1024)
    expect(written).toBe(MAX_CSV_IMPORT_MAX_BYTES)
    const saved = JSON.parse(readFileSync(path, 'utf8')) as Record<string, unknown>
    expect(saved.language).toBe('en')
    expect(readCsvImportMaxBytes(path)).toBe(MAX_CSV_IMPORT_MAX_BYTES)
  })
})

describe('csvImportAllowed', () => {
  it('allows files past the old 32 MiB ceiling under the new default', () => {
    const oldCeiling = 32 * 1024 * 1024
    expect(csvImportAllowed(oldCeiling + 1)).toBe(true)
    expect(csvImportAllowed(DEFAULT_CSV_IMPORT_MAX_BYTES)).toBe(true)
  })

  it('rejects only above the cap, and honors a configured cap', () => {
    expect(csvImportAllowed(DEFAULT_CSV_IMPORT_MAX_BYTES + 1)).toBe(false)
    expect(csvImportAllowed(100 * 1024 * 1024, 64 * 1024 * 1024)).toBe(false)
    expect(csvImportAllowed(40 * 1024 * 1024, 64 * 1024 * 1024)).toBe(true)
  })
})

describe('formatCsvImportBytes', () => {
  it('rounds to whole MB', () => {
    expect(formatCsvImportBytes(128 * 1024 * 1024)).toBe('128 MB')
    expect(formatCsvImportBytes(33_554_432)).toBe('32 MB')
    expect(formatCsvImportBytes(1.4 * 1024 * 1024)).toBe('1 MB')
  })
})

import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  readLegacyImportPrompt,
  readLegacyImportSkip,
  writeLegacyImportPrompt,
  writeLegacyImportSkip,
} from '../src/legacy-import-prefs'

let dir: string
let settingsPath: string

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'legacy-import-'))
  settingsPath = join(dir, 'app-settings.json')
})
afterEach(() => rmSync(dir, { recursive: true, force: true }))

describe('legacy import skip preference', () => {
  it('defaults to asking when the file is missing or corrupt', () => {
    expect(readLegacyImportSkip(settingsPath, 'word97')).toBe(false)
    expect(readLegacyImportPrompt(settingsPath)).toBe(true)
    writeFileSync(settingsPath, '{not json')
    expect(readLegacyImportSkip(settingsPath, 'word97')).toBe(false)
    expect(readLegacyImportPrompt(settingsPath)).toBe(true)
  })

  it('remembers one format without touching the others or unrelated keys', () => {
    writeFileSync(settingsPath, JSON.stringify({ authorName: 'A' }))
    writeLegacyImportSkip(settingsPath, 'word97', true)
    expect(readLegacyImportSkip(settingsPath, 'word97')).toBe(true)
    expect(readLegacyImportSkip(settingsPath, 'hwp')).toBe(false)
    expect(readLegacyImportPrompt(settingsPath)).toBe(false)
    expect(JSON.parse(readFileSync(settingsPath, 'utf8')).authorName).toBe('A')
  })

  it('settings switch clears or sets every format', () => {
    writeLegacyImportSkip(settingsPath, 'ods', true)
    writeLegacyImportPrompt(settingsPath, true)
    expect(readLegacyImportSkip(settingsPath, 'ods')).toBe(false)
    expect(readLegacyImportPrompt(settingsPath)).toBe(true)
    writeLegacyImportPrompt(settingsPath, false)
    for (const kind of ['word97', 'odt', 'hwp', 'ods', 'odp'] as const)
      expect(readLegacyImportSkip(settingsPath, kind)).toBe(true)
  })

  it('ignores non-boolean junk under the key', () => {
    writeFileSync(
      settingsPath,
      JSON.stringify({ legacyImportSkipConfirm: { word97: 'yes', odt: 1 } }),
    )
    expect(readLegacyImportSkip(settingsPath, 'word97')).toBe(false)
    expect(readLegacyImportPrompt(settingsPath)).toBe(true)
  })
})

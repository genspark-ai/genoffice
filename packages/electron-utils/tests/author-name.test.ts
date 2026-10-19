import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import {
  AUTHOR_NAME_MAX,
  readAuthorNameSetting,
  resolveAuthorName,
  sanitizeAuthorName,
  systemAuthorName,
} from '../src/author-name'

const settingsFile = (content: string): string => {
  const path = join(mkdtempSync(join(tmpdir(), 'author-name-')), 'app-settings.json')
  writeFileSync(path, content)
  return path
}

describe('sanitizeAuthorName', () => {
  it('trims, collapses whitespace and caps the length', () => {
    expect(sanitizeAuthorName('  Ada   Lovelace \n')).toBe('Ada Lovelace')
    expect(sanitizeAuthorName('x'.repeat(AUTHOR_NAME_MAX + 5))).toHaveLength(AUTHOR_NAME_MAX)
  })

  it('rejects non-strings and blank values', () => {
    expect(sanitizeAuthorName(42)).toBe('')
    expect(sanitizeAuthorName('   ')).toBe('')
  })
})

describe('readAuthorNameSetting', () => {
  it('reads the stored name', () => {
    expect(readAuthorNameSetting(settingsFile('{"authorName":" Ada "}'))).toBe('Ada')
  })

  it('is empty for a missing key, a malformed file or a missing file', () => {
    expect(readAuthorNameSetting(settingsFile('{}'))).toBe('')
    expect(readAuthorNameSetting(settingsFile('[1]'))).toBe('')
    expect(readAuthorNameSetting(settingsFile('not json'))).toBe('')
    expect(readAuthorNameSetting('/nonexistent/app-settings.json')).toBe('')
  })
})

describe('resolveAuthorName', () => {
  it('prefers the stored name and falls back to the OS account', () => {
    expect(resolveAuthorName(settingsFile('{"authorName":"Ada"}'))).toBe('Ada')
    expect(resolveAuthorName(settingsFile('{"authorName":"  "}'))).toBe(systemAuthorName())
    expect(systemAuthorName().length).toBeGreaterThan(0)
  })
})

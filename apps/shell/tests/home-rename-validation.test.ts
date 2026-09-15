import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { isSameFile, isValidRawRenameName, isValidRenameName } from '../src/main/rename-validation'

describe('home rename validation', () => {
  it('rejects every Windows-illegal name character with the localized gate', () => {
    for (const bad of ['\\', '/', ':', '*', '?', '"', '<', '>', '|', 'a\0b', 'a\x01b']) {
      expect(isValidRenameName(`report${bad}.pdf`)).toBe(false)
    }
    expect(isValidRenameName('')).toBe(false)
    expect(isValidRenameName('quarterly report (final).pdf')).toBe(true)
    expect(isValidRenameName('ski⛷report.pdf')).toBe(true)
  })

  it('treats only the same inode as the same file', () => {
    const dir = mkdtempSync(join(tmpdir(), 'rename-'))
    const a = join(dir, 'Report.pdf')
    const b = join(dir, 'other.pdf')
    writeFileSync(a, 'a')
    writeFileSync(b, 'b')
    expect(isSameFile(a, a)).toBe(true)
    expect(isSameFile(a, b)).toBe(false)
    expect(isSameFile(a, join(dir, 'missing.pdf'))).toBe(false)
  })

  it('rejects Windows reserved names, trailing dots, and overlong names', () => {
    for (const bad of [
      'CON',
      'con.pdf',
      'PRN.docx',
      'aux',
      'NUL.txt',
      'COM1',
      'com9.pdf',
      'LPT1',
      'lpt9.xlsx',
    ]) {
      expect(isValidRenameName(bad)).toBe(false)
    }
    expect(isValidRenameName('report.')).toBe(false)
    expect(isValidRenameName('report ')).toBe(false)
    expect(isValidRenameName('.')).toBe(false)
    expect(isValidRenameName('a'.repeat(256))).toBe(false)
    expect(isValidRenameName('com10.pdf')).toBe(true)
    expect(isValidRenameName('my CON file.pdf')).toBe(true)
    expect(isValidRenameName('a'.repeat(255))).toBe(true)
  })

  it('rejects raw names with surrounding whitespace instead of silently trimming', () => {
    // The IPC handler validates the raw name: trimming "report " to
    // "report" first made the trailing-space gate unreachable.
    for (const raw of ['report ', ' report', '  report  ', 'report\t', '\nreport.pdf']) {
      expect(isValidRawRenameName(raw)).toBe(false)
    }
    expect(isValidRawRenameName('quarterly report (final).pdf')).toBe(true)
    expect(isValidRawRenameName('report.')).toBe(false)
    expect(isValidRawRenameName('')).toBe(false)
  })
})

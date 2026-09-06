import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { isSameFile, isValidRenameName } from '../src/main/rename-validation'

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
})

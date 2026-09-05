import { describe, expect, it } from 'vitest'

import { isCaseOnlyRename, isValidRenameName } from '../src/main/rename-validation'

describe('home rename validation', () => {
  it('rejects every Windows-illegal name character with the localized gate', () => {
    for (const bad of ['\\', '/', ':', '*', '?', '"', '<', '>', '|', 'a\0b', 'a\x01b']) {
      expect(isValidRenameName(`report${bad}.pdf`)).toBe(false)
    }
    expect(isValidRenameName('')).toBe(false)
    expect(isValidRenameName('quarterly report (final).pdf')).toBe(true)
    expect(isValidRenameName('ski⛷report.pdf')).toBe(true)
  })

  it('recognizes case-only renames so the exists gate can be skipped', () => {
    expect(isCaseOnlyRename('/d/Report.pdf', '/d/report.pdf')).toBe(true)
    expect(isCaseOnlyRename('/d/report.pdf', '/d/report.pdf')).toBe(false)
    expect(isCaseOnlyRename('/d/report.pdf', '/d/other.pdf')).toBe(false)
    expect(isCaseOnlyRename('/d/report.pdf', '/e/report.pdf')).toBe(false)
  })
})

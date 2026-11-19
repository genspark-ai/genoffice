import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { bareFileName, uniquePathIn } from '../src/main/file-name'

describe('bareFileName', () => {
  it('keeps plain names untouched', () => {
    expect(bareFileName('report.docx')).toBe('report.docx')
    expect(bareFileName('my notes v2.md')).toBe('my notes v2.md')
  })

  it('reduces separator-carrying suggestions to their last segment', () => {
    expect(bareFileName('sub/dir/report.docx')).toBe('report.docx')
    expect(bareFileName('../../etc/passwd')).toBe('passwd')
    expect(bareFileName('/absolute/path/x.xlsx')).toBe('x.xlsx')
    expect(bareFileName('C:\\Users\\x\\y.docx')).toBe('y.docx')
  })

  it('returns an empty string for navigation segments and pure separators', () => {
    expect(bareFileName('..')).toBe('')
    expect(bareFileName('.')).toBe('')
    expect(bareFileName('/')).toBe('')
    expect(bareFileName('\\')).toBe('')
  })
})

describe('uniquePathIn', () => {
  it('numbers a taken name and gives a suggestion that names nothing a name of its own', () => {
    const dir = mkdtempSync(join(tmpdir(), 'unique-path-'))
    try {
      expect(uniquePathIn(dir, 'report.docx')).toBe(join(dir, 'report.docx'))
      writeFileSync(join(dir, 'report.docx'), '')
      expect(uniquePathIn(dir, 'sub/report.docx')).toBe(join(dir, 'report-2.docx'))
      // a trailing separator or a navigation segment leaves no name: never the directory itself
      for (const empty of ['report/', 'sub/dir/', '..', '/'])
        expect(uniquePathIn(dir, empty)).toBe(join(dir, 'untitled'))
      writeFileSync(join(dir, 'untitled'), '')
      expect(uniquePathIn(dir, 'report/')).toBe(join(dir, 'untitled-2'))
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })
})

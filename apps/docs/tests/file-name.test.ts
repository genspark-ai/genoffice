import { describe, expect, it } from 'vitest'
import { bareFileName } from '../src/main/file-name'

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

import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { uniqueGeneratedPdfPath } from '../src/main/generated-output'

describe('uniqueGeneratedPdfPath', () => {
  it('keeps generated PDFs inside the configured directory', () => {
    expect(uniqueGeneratedPdfPath('/save', '../report-merged.pdf', () => false)).toBe(
      join('/save', 'report-merged.pdf'),
    )
  })

  it('adds a PDF extension and skips existing names', () => {
    const occupied = new Set([join('/save', 'report.pdf'), join('/save', 'report-2.pdf')])
    expect(uniqueGeneratedPdfPath('/save', 'report', (path) => occupied.has(path))).toBe(
      join('/save', 'report-3.pdf'),
    )
  })

  it('sanitizes characters that are invalid in file names', () => {
    expect(uniqueGeneratedPdfPath('/save', 'a:b?.pdf', () => false)).toBe(join('/save', 'a_b_.pdf'))
  })

  it('keeps the stem of a suggested name with a colon (win32 basename drive-strip)', () => {
    // path.win32.basename treats a colon as the drive separator, so the raw
    // name must have its colons replaced before basename ever sees it:
    // `a:b?.pdf` would otherwise come out as `b_.pdf` on Windows.
    expect(uniqueGeneratedPdfPath('/save', 'Report: Q1.pdf', () => false)).toBe(
      join('/save', 'Report_ Q1.pdf'),
    )
  })

  it('still reduces a relative full path suggestion to its file name', () => {
    expect(uniqueGeneratedPdfPath('/save', '../report-merged.pdf', () => false)).toBe(
      join('/save', 'report-merged.pdf'),
    )
  })

  it('collapses a win32 drive-path suggestion to its file name on Windows', () => {
    // basename is platform-flavored: win32 splits on backslashes and would
    // have eaten the stem had the colon not been replaced first; on posix
    // the same string is one file name and the sanitizer flattens it.
    const winPath = 'C:\\users\\report.pdf'
    const expected = process.platform === 'win32' ? 'report.pdf' : 'C__users_report.pdf'
    expect(uniqueGeneratedPdfPath('/save', winPath, () => false)).toBe(join('/save', expected))
  })
})

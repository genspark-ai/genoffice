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
})

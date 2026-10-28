import { describe, expect, it } from 'vitest'

import { autoFileNameFromBody, sanitizeFileBaseName } from '../src/renderer/file-actions'

describe('sanitizeFileBaseName', () => {
  it('keeps an ordinary heading as the file base', () => {
    expect(sanitizeFileBaseName('Quarterly report')).toBe('Quarterly report')
  })

  it('declines Windows reserved names so the save keeps its Untitled name', () => {
    for (const reserved of ['CON', 'prn', 'AUX', 'nul', 'COM1', 'com9', 'LPT1', 'lpt9']) {
      expect(sanitizeFileBaseName(reserved)).toBeNull()
    }
    expect(sanitizeFileBaseName('My CON file')).toBe('My CON file')
    expect(sanitizeFileBaseName('com10')).toBe('com10')
  })
})

const paragraph = (text: string) => ({
  type: 'docParagraph',
  content: [{ type: 'text', text }],
})
const heading = (text: string) => ({
  type: 'docHeading',
  content: [{ type: 'text', text }],
})

describe('autoFileNameFromBody', () => {
  it('names the file after the first heading when one exists', () => {
    expect(autoFileNameFromBody([heading('Quarterly report'), paragraph('Numbers follow.')])).toBe(
      'Quarterly report.docx',
    )
  })

  it('falls back to the first paragraph for prose documents without a heading', () => {
    expect(autoFileNameFromBody([paragraph('The quarterly numbers missed every target.')])).toBe(
      'The quarterly number.docx',
    )
  })

  it('caps the prose fallback at the first 20 characters', () => {
    const prose = 'abcdefghijklmnopqrstuvwxyz'
    expect(autoFileNameFromBody([paragraph(prose)])).toBe('abcdefghijklmnopqrst.docx')
  })

  it('skips a leading empty paragraph instead of naming the file after it', () => {
    expect(autoFileNameFromBody([paragraph(''), paragraph('Real opening line')])).toBe(
      'Real opening line.docx',
    )
  })

  it('sanitizes the prose fallback like a heading base', () => {
    expect(autoFileNameFromBody([paragraph('Q1/Q2: "final?"')])).toBe('Q1 Q2 final.docx')
  })

  it('keeps the Untitled name for a document with no usable text', () => {
    expect(autoFileNameFromBody([])).toBeNull()
    expect(autoFileNameFromBody([paragraph('')])).toBeNull()
    // a heading that sanitizes away entirely (reserved name) must not block the
    // paragraph fallback
    expect(autoFileNameFromBody([heading('CON'), paragraph('Fallback prose')])).toBe(
      'Fallback prose.docx',
    )
  })
})

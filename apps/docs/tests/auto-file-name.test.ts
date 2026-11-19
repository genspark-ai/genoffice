import { describe, expect, it } from 'vitest'

import { sanitizeFileBaseName } from '../src/renderer/file-actions'

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

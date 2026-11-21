/** A page-number instruction typed for a header or footer becomes the live field, not a cached number. */
import { describe, expect, it } from 'vitest'
import { liveHfFieldOf } from '../src/renderer/editor/hf-editor'

describe('live header and footer fields', () => {
  it('names the live field for a bare page instruction, whatever its case or spacing', () => {
    expect(liveHfFieldOf('PAGE')).toBe('PAGE')
    expect(liveHfFieldOf(' numpages ')).toBe('NUMPAGES')
    expect(liveHfFieldOf('SectionPages')).toBe('SECTIONPAGES')
  })

  it('leaves other fields, and page fields with switches, to the cached form', () => {
    expect(liveHfFieldOf('DATE')).toBeNull()
    expect(liveHfFieldOf('PAGE \\* roman')).toBeNull()
    expect(liveHfFieldOf('FILENAME')).toBeNull()
    expect(liveHfFieldOf('')).toBeNull()
  })
})

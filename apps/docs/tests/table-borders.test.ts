import { describe, expect, it } from 'vitest'

import { tableBordersCss } from '../src/renderer/editor/extensions'

describe('tableBordersCss', () => {
  it('emits edge variables for declared sides and none otherwise', () => {
    const css = tableBordersCss({ top: { style: 'single', szEighths: 4, color: '000000' } })
    expect(css).toContain('--doc-b-t:1px solid #000000')
    expect(css).toContain('--doc-b-b:none')
    expect(css).toContain('--doc-b-h:none')
    expect(css).toContain('--doc-b-v:none')
  })

  it('emits insideH/insideV variables for table-level inner borders', () => {
    const css = tableBordersCss({
      insideH: { style: 'single', szEighths: 4, color: 'FF0000' },
      insideV: { style: 'single', szEighths: 12, color: '0000FF' },
    })
    expect(css).toContain('--doc-b-h:1px solid #FF0000')
    expect(css).toContain('--doc-b-v:2px solid #0000FF')
  })

  it('returns no declarations for a null attr', () => {
    expect(tableBordersCss(null)).toEqual([])
  })
})

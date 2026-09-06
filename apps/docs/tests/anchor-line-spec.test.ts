import { describe, expect, it } from 'vitest'
import { anchorLineSpec } from '../src/renderer/editor/extensions'

describe('anchorLineSpec', () => {
  it('lays the line out with the paragraph format but leaves the page break on the picture', () => {
    const [tag, attrs] = anchorLineSpec({
      pageBreakBefore: true,
      keepNext: true,
      spaceBefore: 240,
      styleId: 'Caption',
    }) as [string, Record<string, string>, unknown]
    expect(tag).toBe('p')
    expect(attrs.class).toContain('doc-anchor-line')
    expect(attrs.class).not.toContain('page-break-before')
    expect(attrs['data-page-break-label']).toBeUndefined()
    expect(attrs['data-para']).toBeUndefined()
    expect(attrs.style).toContain('margin-top')
  })
})

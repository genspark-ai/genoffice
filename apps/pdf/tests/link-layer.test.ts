import { describe, expect, it } from 'vitest'
import { collectPageLinks, MAX_PAGE_LINKS } from '../src/renderer/LinkLayer'

describe('collectPageLinks', () => {
  it('keeps links with finite rects and targets', () => {
    const links = collectPageLinks([
      { subtype: 'Link', rect: [0, 0, 10, 10], url: 'https://a.com' },
      { subtype: 'Link', rect: [0, 0, 10, 10], dest: [0] },
      { subtype: 'Widget', rect: [0, 0, 10, 10], url: 'https://b.com' },
      { subtype: 'Link', rect: [0, 0, 10, 10] },
    ])
    expect(links).toHaveLength(2)
  })

  it('drops non-finite rects and caps per-page count', () => {
    expect(
      collectPageLinks([{ subtype: 'Link', rect: [0, 0, NaN, 10], url: 'https://a.com' }]),
    ).toHaveLength(0)
    const many = Array.from({ length: MAX_PAGE_LINKS + 100 }, (_, i) => ({
      subtype: 'Link',
      rect: [0, i, 10, i + 1] as [number, number, number, number],
      url: 'https://a.com',
    }))
    expect(collectPageLinks(many)).toHaveLength(MAX_PAGE_LINKS)
  })
})

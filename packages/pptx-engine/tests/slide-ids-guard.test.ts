/** nextSlideId stays fast and bounded on hostile id attributes. */
import { describe, it, expect } from 'vitest'
import { nextSlideId } from '../src/slide-ids'

function presWithIds(ids: string): string {
  return `<p:sldIdLst>${ids}</p:sldIdLst>`
}

describe('nextSlideId hostile inputs', () => {
  it('ignores out-of-range ids instead of forcing the ceiling scan', () => {
    const pres = presWithIds(
      '<p:sldId id="256" r:id="rId2"/>' + '<p:sldId id="99999999999" r:id="rId3"/>',
    )
    expect(nextSlideId(pres)).toBe(257)
  })

  it('ignores below-minimum ids when allocating', () => {
    const pres = presWithIds('<p:sldId id="5" r:id="rId2"/>' + '<p:sldId id="256" r:id="rId3"/>')
    expect(nextSlideId(pres)).toBe(257)
  })

  it('allocates max+1 without spreading huge id sets', () => {
    const count = 200000
    let ids = ''
    for (let i = 0; i < count; i++) ids += `<p:sldId id="${256 + i}" r:id="rId${i}"/>`
    expect(nextSlideId(presWithIds(ids))).toBe(256 + count)
  })

  it('ignores single-quoted out-of-range ids as well', () => {
    const pres = presWithIds(
      "<p:sldId id='99999999999' r:id='rId2'/>" + '<p:sldId id="300" r:id="rId3"/>',
    )
    expect(nextSlideId(pres)).toBe(301)
  })
})

import { describe, expect, it } from 'vitest'
import { borderMergeFlags, sameBorderGroup } from '../src/renderer/editor/para-border-merge'

const red = JSON.stringify({ b: { color: 'C00000', szPt: 1 } })

describe('sameBorderGroup', () => {
  it('identical borders + lines + shading merge', () => {
    expect(
      sameBorderGroup({ borders: 'b', borderLines: red }, { borders: 'b', borderLines: red }),
    ).toBe(true)
  })

  it('borderLines compare by value, not key order / formatting', () => {
    const reordered = JSON.stringify({ b: { szPt: 1, color: 'C00000' } })
    expect(
      sameBorderGroup({ borders: 'b', borderLines: red }, { borders: 'b', borderLines: reordered }),
    ).toBe(true)
  })

  it('any difference splits the group', () => {
    const blue = JSON.stringify({ b: { color: '0000FF', szPt: 1 } })
    expect(
      sameBorderGroup({ borders: 'b', borderLines: red }, { borders: 'tb', borderLines: red }),
    ).toBe(false)
    expect(
      sameBorderGroup({ borders: 'b', borderLines: red }, { borders: 'b', borderLines: blue }),
    ).toBe(false)
    expect(
      sameBorderGroup(
        { borders: 'b', borderLines: red },
        { borders: 'b', borderLines: red, shadingFill: 'EEEEEE' },
      ),
    ).toBe(false)
  })

  it('border-less paragraphs never merge', () => {
    expect(sameBorderGroup({}, {})).toBe(false)
    expect(sameBorderGroup({ borders: 'b' }, {})).toBe(false)
  })
})

describe('borderMergeFlags', () => {
  const bordered = { borders: 'b', borderLines: red }

  it('inner boundaries of a run suppress bottom above / top below', () => {
    expect(borderMergeFlags([bordered, bordered, bordered])).toEqual([
      { suppressTop: false, suppressBottom: true },
      { suppressTop: true, suppressBottom: true },
      { suppressTop: true, suppressBottom: false },
    ])
  })

  it('a non-matching block (table, different style) breaks the group', () => {
    expect(borderMergeFlags([bordered, {}, bordered])).toEqual([
      { suppressTop: false, suppressBottom: false },
      { suppressTop: false, suppressBottom: false },
      { suppressTop: false, suppressBottom: false },
    ])
  })

  it('single bordered paragraph keeps both edges', () => {
    expect(borderMergeFlags([bordered])).toEqual([{ suppressTop: false, suppressBottom: false }])
  })
})

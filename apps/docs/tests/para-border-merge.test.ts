import { describe, expect, it } from 'vitest'
import {
  affectsBorderGroup,
  borderMergeFlags,
  sameBorderGroup,
  type StyleBorderInfo,
  type StyleBorderLookup,
} from '../src/renderer/editor/para-border-merge'

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

describe('style-level borders (Word merges same-style bordered paragraphs)', () => {
  const h2: StyleBorderInfo = {
    borderSides: { t: { none: true, spacePt: 30 }, b: { szPt: 0.75, spacePt: 3 } },
  }
  const lookup: StyleBorderLookup = (id) => (id === 'H2' ? h2 : undefined)
  const styled = { styleId: 'H2' }

  it('two paragraphs of a bordered style form one group', () => {
    expect(sameBorderGroup(styled, styled, lookup)).toBe(true)
    expect(borderMergeFlags([styled, styled, { styleId: 'Normal' }], lookup)).toEqual([
      { suppressTop: false, suppressBottom: true },
      { suppressTop: true, suppressBottom: false },
      { suppressTop: false, suppressBottom: false },
    ])
  })

  it('a non-paragraph block never joins the group, even when Normal carries borders', () => {
    const normal: StyleBorderLookup = () => h2
    expect(borderMergeFlags([{}, { break: true }, {}], normal)).toEqual([
      { suppressTop: false, suppressBottom: false },
      { suppressTop: false, suppressBottom: false },
      { suppressTop: false, suppressBottom: false },
    ])
    expect(borderMergeFlags([{}, {}], normal)[0].suppressBottom).toBe(true)
  })

  it('without the style table only direct attrs count', () => {
    expect(sameBorderGroup(styled, styled)).toBe(false)
  })

  it('a direct side overrides the style side per side', () => {
    expect(sameBorderGroup(styled, { ...styled, borderReset: 't' }, lookup)).toBe(false)
    expect(
      sameBorderGroup(styled, { ...styled, borderReset: 't', borderPad: '{"t":30}' }, lookup),
    ).toBe(true)
    expect(
      sameBorderGroup(
        styled,
        { ...styled, borders: 'b', borderLines: '{"b":{"szPt":0.75,"spacePt":3}}' },
        lookup,
      ),
    ).toBe(true)
  })

  it('style fill vs cleared fill split the group', () => {
    const filled: StyleBorderLookup = () => ({ ...h2, shadingFill: 'EEEEEE' })
    expect(sameBorderGroup(styled, { ...styled, shadingClear: true }, filled)).toBe(false)
    expect(sameBorderGroup(styled, { ...styled, shadingFill: 'EEEEEE' }, filled)).toBe(true)
  })
})

describe('affectsBorderGroup', () => {
  it('pad-only and style-bordered paragraphs force a recompute', () => {
    expect(affectsBorderGroup({ borderPad: '{"t":30}' })).toBe(true)
    expect(affectsBorderGroup({ borderReset: 'tb' })).toBe(true)
    expect(affectsBorderGroup({ shadingClear: true })).toBe(true)
    const lookup: StyleBorderLookup = () => ({ borderSides: { t: { none: true, spacePt: 30 } } })
    expect(affectsBorderGroup({ styleId: 'H2' }, lookup)).toBe(true)
    expect(affectsBorderGroup({ styleId: 'H2' })).toBe(false)
    expect(affectsBorderGroup({})).toBe(false)
  })
})

import { describe, expect, it } from 'vitest'
import { wrapGroupCss, type WrapPicture } from '../src/renderer/editor/wrap-group'

// prod corpus 20261020 sample 004: two right-wrapped pictures of one paragraph,
// the upper one 51pt past the column edge; CSS float placement pushed it left
const upper: WrapPicture = {
  side: 'right',
  x: 522.8,
  y: 52.1,
  w: 190.2,
  h: 126.4,
  dist: [12, 0, 0, 0],
}
const lower: WrapPicture = {
  side: 'right',
  x: 524.3,
  y: 217,
  w: 186.6,
  h: 139.9,
  dist: [12, 0, 0, 0],
}

describe('wrapGroupCss', () => {
  it('leaves a single picture per side to the plain float', () => {
    expect(wrapGroupCss([lower]).styles).toEqual([null])
    expect(wrapGroupCss([lower, { ...upper, side: 'left' }]).styles).toEqual([null, null])
  })

  it('the first picture floats with the union box; later ones overlay at their offsets', () => {
    const { styles } = wrapGroupCss([lower, upper])
    const lead = styles[0]!
    expect(lead).toContain('margin-top:52.1px')
    expect(lead).toContain('padding-top:164.9px')
    expect(lead).toContain('padding-bottom:0.0px')
    // the margin box starts at the union's text-side edge, the picture keeps its x
    expect(lead).toContain('margin-left:13.5px')
    expect(lead).toContain('margin-right:calc(100% - 710.9px)')
    // stair profile in margin-box coordinates (the reference box of a bare
    // shape-outside; margin-top lies inside it): the exclusion starts at the
    // upper rectangle's top, not at the paragraph top
    expect(lead).toMatch(
      /shape-outside:polygon\(0\.0px 52\.1px, 0\.0px 178\.5px, 202\.2px 178\.5px, 202\.2px 217\.0px, 1\.5px 217\.0px, 1\.5px 356\.9px, 202\.2px 356\.9px, 202\.2px 52\.1px\)/,
    )
    const firstY = /polygon\([\d.]+px ([\d.]+)px/.exec(lead)![1]
    expect(firstY).toBe('52.1')
    expect(lead).toContain(`margin-top:${firstY}px`)
    expect(styles[1]).toBe('position:absolute;float:none;margin:0;left:522.8px;top:52.1px')
  })

  it('prefers an unturned lead; a turned lead rotates about its own picture', () => {
    const turned = wrapGroupCss([{ ...lower, turned: true }, upper]).styles
    expect(turned[1]).toContain('padding-bottom:178.4px')
    expect(turned[0]).toBe('position:absolute;float:none;margin:0;left:524.3px;top:217.0px')
    const both = wrapGroupCss([
      { ...lower, turned: true },
      { ...upper, turned: true },
    ]).styles
    expect(both[0]).toContain('transform-origin:50% 234.9px')
  })

  it('a left-side group measures from the column start', () => {
    const a: WrapPicture = { side: 'left', x: 0, y: 0, w: 100, h: 50, dist: [0, 12, 0, 0] }
    const b: WrapPicture = { side: 'left', x: 20, y: 80, w: 100, h: 50, dist: [0, 12, 0, 0] }
    const lead = wrapGroupCss([a, b]).styles[0]!
    expect(lead).toContain('margin-left:0.0px')
    expect(lead).toContain('margin-right:32.0px')
    expect(lead).toContain('padding-bottom:80.0px')
    expect(lead).toContain(
      'polygon(112.0px 0.0px, 112.0px 50.0px, 0.0px 50.0px, 0.0px 80.0px, 132.0px 80.0px, 132.0px 130.0px, 0.0px 130.0px, 0.0px 0.0px)',
    )
  })
})

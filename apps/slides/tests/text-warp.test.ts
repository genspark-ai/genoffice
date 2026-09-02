import { describe, expect, it } from 'vitest'
import { warpGlyphs } from '../src/renderer/text-warp'
import type { GlyphDraw } from '../src/renderer/konva-adapter'

const glyph = (text: string, x: number): GlyphDraw => ({
  text,
  x,
  y: 0,
  fontSize: 20,
  fontFamily: 'Arial',
  fill: '#000',
  fontStyle: 'normal',
  textDecoration: '',
})
const measure = (text: string) => [...text].length * 10

describe('warpGlyphs', () => {
  it('unknown presets return null (caller keeps the straight layout)', () => {
    expect(warpGlyphs([glyph('Hi', 0)], 100, 40, { prst: 'textNoShape' }, measure)).toBeNull()
    expect(warpGlyphs([glyph('Hi', 0)], 100, 40, { prst: 'textRingInside' }, measure)).toBeNull()
  })

  it('wave1 lifts the quarter-way character and drops the three-quarter one', () => {
    const out = warpGlyphs([glyph('abcd', 0)], 200, 40, { prst: 'textWave1' }, measure)!
    // repacked at 10px/char, centers at u=1/8,3/8,5/8,7/8: sin>0 → +y for u<0.5
    expect(out[1]!.y).toBeGreaterThan(out[2]!.y)
    expect(out[0]!.rotation).toBe(0) // waves stay upright
  })

  it('arch rotates characters along the tangent and centers the block', () => {
    const out = warpGlyphs([glyph('abcd', 0)], 200, 40, { prst: 'textArchUp' }, measure)!
    expect(out[0]!.rotation!).toBeLessThan(0) // rising into the arch
    expect(out[3]!.rotation!).toBeGreaterThan(0)
    const centers = out.map((g) => g.x)
    expect((centers[0]! + centers[3]!) / 2).toBeCloseTo(100, 0) // centered in the box
  })

  it('triangleInverted squeezes mid-box characters via scaleY', () => {
    const out = warpGlyphs([glyph('abcd', 0)], 200, 40, { prst: 'textTriangleInverted' }, measure)!
    expect(out[1]!.scaleY!).toBeLessThan(out[0]!.scaleY!)
  })

  it('splits multi-char runs and skips whitespace glyphs', () => {
    const out = warpGlyphs([glyph('a b', 0)], 100, 40, { prst: 'textWave1' }, measure)!
    expect(out.map((g) => g.text)).toEqual(['a', 'b'])
    expect(out[1]!.x).toBeGreaterThan(out[0]!.x + 10) // the swallowed space keeps its advance
  })
})

import { describe, it, expect } from 'vitest'
import { resolveColorNode } from '../src/color'

describe('resolveColorNode a:prstClr', () => {
  it('resolves preset names', () => {
    expect(resolveColorNode({ 'a:prstClr': { '@_val': 'red' } }, undefined)).toBe('#FF0000')
    expect(resolveColorNode({ 'a:prstClr': { '@_val': 'cornflowerBlue' } }, undefined)).toBe(
      '#6495ED',
    )
    expect(resolveColorNode({ 'a:prstClr': { '@_val': 'dkGray' } }, undefined)).toBe('#A9A9A9')
  })

  it('applies lumOff/alpha modifiers (Aspose watermark: translucent pink)', () => {
    const node = {
      'a:prstClr': {
        '@_val': 'red',
        'a:lumOff': { '@_val': '30000' },
        'a:alpha': { '@_val': '40000' },
      },
    }
    // lumOff in HSL luminance: red L 0.5 → 0.8 with saturation kept = #FF9999
    // (matches Office's published tint table, e.g. 4472C4 lumMod40/lumOff60 → B4C7E7)
    expect(resolveColorNode(node, undefined)).toBe('#FF999966')
  })

  it('unknown preset name resolves to undefined', () => {
    expect(resolveColorNode({ 'a:prstClr': { '@_val': 'nope' } }, undefined)).toBeUndefined()
  })
})

describe('HSL modifiers', () => {
  it('hueOff rotates hue (SmartArt colorful cycle)', () => {
    // red + 120° hue offset (val in 1/60000 deg) → green, saturation/luminance kept
    const node = {
      'a:srgbClr': { '@_val': 'FF0000', 'a:hueOff': { '@_val': String(120 * 60000) } },
    }
    expect(resolveColorNode(node, undefined)).toBe('#00FF00')
  })

  it('satOff shifts saturation', () => {
    // 50%-saturation red minus 25% saturation
    const node = {
      'a:srgbClr': { '@_val': 'BF4040', 'a:satOff': { '@_val': '-25000' } },
    }
    expect(resolveColorNode(node, undefined)).toBe('#9F6060')
  })

  it('lumMod/lumOff operate on HSL luminance with saturation kept', () => {
    // Per-channel ±1/255: the HSL round-trip rounds differently across platforms
    const expectHexClose = (got: string | undefined, want: string) => {
      expect(got).toBeTruthy()
      for (let i = 1; i < 7; i += 2) {
        const g = parseInt(got!.slice(i, i + 2), 16)
        const w = parseInt(want.slice(i, i + 2), 16)
        expect(Math.abs(g - w)).toBeLessThanOrEqual(1)
      }
    }
    // Office tint table: 4472C4 "Lighter 60%" (lumMod 40 + lumOff 60) = B4C7E7
    const lighter = {
      'a:srgbClr': {
        '@_val': '4472C4',
        'a:lumMod': { '@_val': '40000' },
        'a:lumOff': { '@_val': '60000' },
      },
    }
    expectHexClose(resolveColorNode(lighter, undefined), '#B4C7E7')
    // Office tint table: 4472C4 "Darker 25%" (lumMod 75) = 2F5496
    const darker = {
      'a:srgbClr': { '@_val': '4472C4', 'a:lumMod': { '@_val': '75000' } },
    }
    expectHexClose(resolveColorNode(darker, undefined), '#2F5496')
    // Dark saturated accent + lighten stays vivid (prod_011 org chart: 7F0000 → #FF6666)
    const vivid = {
      'a:srgbClr': {
        '@_val': '7F0000',
        'a:lumMod': { '@_val': '40000' },
        'a:lumOff': { '@_val': '60000' },
      },
    }
    expectHexClose(resolveColorNode(vivid, undefined), '#FF6666')
  })
})

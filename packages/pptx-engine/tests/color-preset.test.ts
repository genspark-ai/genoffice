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
    expect(resolveColorNode(node, undefined)).toBe('#FF4D4D66')
  })

  it('unknown preset name resolves to undefined', () => {
    expect(resolveColorNode({ 'a:prstClr': { '@_val': 'nope' } }, undefined)).toBeUndefined()
  })
})

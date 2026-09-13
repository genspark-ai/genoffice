import { describe, it, expect } from 'vitest'
import { resolveColorNode } from '../src/color'
import { sysColorHex } from '../src/theme'

describe('a:sysClr resolution', () => {
  it('window/windowText ignore a stale lastClr (PowerPoint paints the live system colors)', () => {
    // A lecture deck: lt1 = sysClr window lastClr=C0C0C0, PowerPoint background is white
    expect(sysColorHex('window', 'C0C0C0')).toBe('#FFFFFF')
    expect(sysColorHex('windowText', '808080')).toBe('#000000')
    expect(
      resolveColorNode({ 'a:sysClr': { '@_val': 'window', '@_lastClr': 'C0C0C0' } }, undefined),
    ).toBe('#FFFFFF')
  })

  it('other system colors keep the cached lastClr, modifiers still apply', () => {
    expect(sysColorHex('highlight', '3399FF')).toBe('#3399FF')
    expect(sysColorHex('highlight', undefined)).toBe('#000000')
    expect(
      resolveColorNode(
        {
          'a:sysClr': {
            '@_val': 'windowText',
            '@_lastClr': '000000',
            'a:alpha': { '@_val': '50000' },
          },
        },
        undefined,
      ),
    ).toBe('#00000080')
  })
})

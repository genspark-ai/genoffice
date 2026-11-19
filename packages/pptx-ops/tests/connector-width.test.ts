import { describe, expect, it } from 'vitest'
import { resolveConnectorWidthEmu } from '../src/ops/arrange-ops'

/**
 * Regression for addConnector line width: a huge but finite widthPt
 * (e.g. 1e308) overflows Math.round(w * EMU_PER_PT) to Infinity, and the
 * old `widthEmu <= 0` check let Infinity through into the OOXML.
 */
describe('resolveConnectorWidthEmu', () => {
  it('accepts normal widths', () => {
    expect(resolveConnectorWidthEmu({})).toBe(12700)
    expect(resolveConnectorWidthEmu({ widthPt: 1 })).toBe(12700)
    expect(resolveConnectorWidthEmu({ widthEmu: 9525 })).toBe(9525)
  })

  it('rejects zero, negative, and overflowing widths', () => {
    expect(() => resolveConnectorWidthEmu({ widthPt: 0 })).toThrow(/finite number > 0/)
    expect(() => resolveConnectorWidthEmu({ widthPt: -2 })).toThrow(/finite number > 0/)
    expect(() => resolveConnectorWidthEmu({ widthPt: 1e308 })).toThrow(/finite number > 0/)
    expect(() => resolveConnectorWidthEmu({ widthEmu: 0 })).toThrow(/finite number > 0/)
  })
})

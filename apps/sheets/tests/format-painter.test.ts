import { FormatPainterStatus } from '@univerjs/sheets-ui'
import { describe, expect, it } from 'vitest'

import { formatPainterClickAction, formatPainterTurnedOff } from '../src/renderer/format-painter'

const { OFF, ONCE, INFINITE } = FormatPainterStatus

describe('formatPainterClickAction', () => {
  it('arms the one-shot painter on a single click while off', () => {
    expect(formatPainterClickAction(OFF, 1)).toBe('once')
  })

  it('cancels an active painter on a single click', () => {
    expect(formatPainterClickAction(ONCE, 1)).toBe('cancel')
    expect(formatPainterClickAction(INFINITE, 1)).toBe('cancel')
  })

  it('escalates the first click of a double-click into the locked mode', () => {
    expect(
      formatPainterClickAction(formatPainterClickAction(OFF, 1) === 'once' ? ONCE : OFF, 2),
    ).toBe('lock')
  })

  it('ignores the second click of a double-click that cancelled the painter', () => {
    expect(formatPainterClickAction(OFF, 2)).toBe('ignore')
  })

  it('does not re-arm a locked painter on a repeated double-click', () => {
    expect(formatPainterClickAction(INFINITE, 2)).toBe('ignore')
  })
})

describe('formatPainterTurnedOff', () => {
  it('reports the transition out of an active painter from either mode', () => {
    expect(formatPainterTurnedOff(ONCE, OFF)).toBe(true)
    expect(formatPainterTurnedOff(INFINITE, OFF)).toBe(true)
  })

  it('stays quiet while arming, escalating or already off', () => {
    expect(formatPainterTurnedOff(OFF, ONCE)).toBe(false)
    expect(formatPainterTurnedOff(ONCE, INFINITE)).toBe(false)
    expect(formatPainterTurnedOff(OFF, OFF)).toBe(false)
  })
})

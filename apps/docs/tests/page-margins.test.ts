import { describe, expect, it } from 'vitest'
import {
  cmFromTwips,
  marginsFitPage,
  twipsFromCmInput,
} from '../src/renderer/components/MarginDialog'

describe('page margins', () => {
  it('preserves the original twips when a rounded display value is unchanged', () => {
    const moderateMargin = 1080
    const displayed = String(cmFromTwips(moderateMargin))

    expect(displayed).toBe('1.91')
    expect(twipsFromCmInput(displayed, moderateMargin)).toBe(moderateMargin)
  })

  it('converts an edited centimeter value to twips', () => {
    expect(twipsFromCmInput('2', 1080)).toBe(1134)
  })

  it('rejects margins that do not leave the minimum page body', () => {
    expect(marginsFitPage({ top: 720, right: 5000, bottom: 720, left: 5000 }, 10319, 14572)).toBe(
      false,
    )
    expect(marginsFitPage({ top: 720, right: 720, bottom: 720, left: 720 }, 10319, 14572)).toBe(
      true,
    )
  })
})

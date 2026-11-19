import { describe, expect, it } from 'vitest'
import {
  paperCss,
  seedDraftColors,
  textEraseKey,
  textEraseProbe,
} from '../src/renderer/text-edit-preview'
import type { TextDraft } from '../src/renderer/text-edit-preview'

const draft = (paper?: string): TextDraft => ({
  origIdx: 0,
  rect: [10, 20, 200, 40],
  oldText: 'Heading',
  fontSize: 24,
  value: 'Heading',
  paper,
})

describe('seedDraftColors', () => {
  it('keeps default ink for white text over the paper fallback', () => {
    const d = seedDraftColors(draft(), { reason: null, baseColor: [250, 250, 250] })
    expect(d.seedInk).toBeUndefined()
  })

  it('seeds white ink when the sampled page behind the run is dark', () => {
    const d = seedDraftColors(draft('#0b1220'), { reason: null, baseColor: [250, 250, 250] })
    expect(d.seedInk).toBe('#fafafa')
  })

  it('keeps default ink when the ink matches the sampled page', () => {
    const d = seedDraftColors(draft('#0b1220'), { reason: null, baseColor: [12, 18, 34] })
    expect(d.seedInk).toBeUndefined()
  })
})

describe('paperCss / erase keys', () => {
  it('goes transparent once erased, else the sampled color, else nothing', () => {
    expect(paperCss(true, '#0b1220')).toEqual({ background: 'transparent' })
    expect(paperCss(false, '#0b1220')).toEqual({ background: '#0b1220' })
    expect(paperCss(false, undefined)).toEqual({})
  })

  it('keys a draft and the pending edit it becomes to the same run', () => {
    const a = textEraseProbe(2, [10, 20, 200, 40], 'Heading', 24)
    expect(a.newText).toBe('')
    expect(textEraseKey(a)).toBe(
      textEraseKey({
        pageIndex: 2,
        rect: [10, 20, 200, 40],
        oldText: 'Heading',
        newText: 'New',
        fontSize: 30,
      }),
    )
    expect(textEraseKey(a)).not.toBe(textEraseKey({ ...a, oldText: 'Other' }))
  })
})

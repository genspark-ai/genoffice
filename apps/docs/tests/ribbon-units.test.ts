import { describe, expect, it } from 'vitest'
import { paperSizeCaption } from '../src/renderer/components/ribbon-layout-tab'

describe('paperSizeCaption', () => {
  it('matches the legacy hardcoded captions with the cm unit', () => {
    expect(paperSizeCaption(11906, 16838, 'cm')).toBe('21 × 29.7 cm')
    expect(paperSizeCaption(12240, 15840, 'cm')).toBe('21.59 × 27.94 cm')
    expect(paperSizeCaption(12240, 20160, 'cm')).toBe('21.59 × 35.56 cm')
    expect(paperSizeCaption(10319, 14572, 'cm')).toBe('18.2 × 25.7 cm')
  })

  it('honors the translated unit', () => {
    expect(paperSizeCaption(11906, 16838, '\u5398\u7c73')).toBe('21 × 29.7 \u5398\u7c73')
  })
})

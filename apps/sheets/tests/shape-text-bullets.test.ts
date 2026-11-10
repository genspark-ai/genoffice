import { describe, expect, it } from 'vitest'

import {
  shapeBulletWidth,
  shapeParagraphIndentStyle,
  shapeParagraphMarkers,
  type ShapeParagraph,
} from '../src/renderer/shape-text-bullets'

const numbered = (text: string, scheme = 'arabicPeriod', startAt?: number): ShapeParagraph => ({
  marginLeft: 18,
  indent: -18,
  bulletScheme: scheme,
  ...(startAt === undefined ? {} : { bulletStartAt: startAt }),
  runs: [{ text }],
})
const plain = (text: string): ShapeParagraph => ({ runs: [{ text }] })

describe('shapeParagraphMarkers', () => {
  it('numbers consecutive buAutoNum paragraphs and restarts after plain text', () => {
    const markers = shapeParagraphMarkers([
      numbered('a'),
      numbered('b'),
      plain('note'),
      numbered('c'),
    ])
    expect(markers).toEqual(['1.', '2.', undefined, '1.'])
  })

  it('honours startAt and the scheme, and skips empty paragraphs without resetting', () => {
    const markers = shapeParagraphMarkers([
      numbered('a', 'arabicParenR', 3),
      { ...numbered(' ', 'arabicParenR', 3) },
      numbered('b', 'arabicParenR', 3),
      numbered('c', 'alphaLcPeriod'),
      numbered('d', 'romanLcPeriod'),
      numbered('e', 'romanLcPeriod'),
    ])
    expect(markers).toEqual(['3)', undefined, '4)', 'a.', 'i.', 'ii.'])
  })

  it('falls back to arabic for unknown schemes and passes buChar through', () => {
    expect(shapeParagraphMarkers([numbered('a', 'somethingNew')])).toEqual(['1.'])
    expect(shapeParagraphMarkers([{ bulletChar: '-', runs: [{ text: 'x' }] }, plain('')])).toEqual([
      '-',
      undefined,
    ])
  })
})

describe('hanging indent', () => {
  it('pads wrapped lines to marL and pulls the first line back by indent', () => {
    expect(shapeParagraphIndentStyle(numbered('a'))).toEqual({
      paddingLeft: 'calc(var(--shape-px, 1px) * 24)',
      textIndent: 'calc(var(--shape-px, 1px) * -24)',
    })
    expect(shapeBulletWidth(numbered('a'))).toBe('calc(var(--shape-px, 1px) * 24)')
  })

  it('clamps a hang deeper than marL and leaves non-hanging bullets inline', () => {
    const deep: ShapeParagraph = { marginLeft: 9, indent: -18, runs: [{ text: 'a' }] }
    expect(shapeParagraphIndentStyle(deep).textIndent).toBe('calc(var(--shape-px, 1px) * -12)')
    expect(shapeBulletWidth({ marginLeft: 10, bulletChar: '-', runs: [] })).toBeUndefined()
    expect(shapeParagraphIndentStyle(plain('a'))).toEqual({})
  })
})

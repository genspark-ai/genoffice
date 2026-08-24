import { describe, expect, it } from 'vitest'

import {
  bidiVisualOrder,
  classifyBidiGlyph,
  logicalGlyphContent,
  reorderRichCellSkeleton,
  resolveBidiLevels,
} from '../src/renderer/rich-text-bidi-fix'

const AR_PRICE = 'السعر' // Arabic "price"
const AR_HEH = 'هـ' // heh + tatweel
const HE_SHALOM = 'שלום'
const AR_INDIC = '١٤٤٦' // Arabic-Indic 1446

const reverse = (text: string): string => [...text].reverse().join('')

// Univer's ArabicHandler emits one glyph per contiguous Arabic chunk with the
// characters already reversed; everything else is one glyph per character.
function univerGlyphs(words: string[]): { content: string; width: number; left: number }[] {
  let left = 0
  return words.map((content) => {
    const glyph = { content, width: content.length * 10, left }
    left += glyph.width
    return glyph
  })
}

function visualContents(glyphs: { content: string; left: number }[]): string[] {
  return [...glyphs].sort((a, b) => a.left - b.left).map((g) => g.content)
}

function makeSkeleton(
  id: string,
  dataStream: string,
  lines: { paragraphIndex?: number; glyphs: { content: string; width: number; left: number }[] }[],
) {
  return {
    getViewModel: () => ({
      getDataModel: () => ({ getSnapshot: () => ({ id, body: { dataStream } }) }),
    }),
    getSkeletonData: () => ({
      pages: [
        {
          sections: [
            {
              columns: [
                {
                  lines: lines.map((line) => ({
                    paragraphIndex: line.paragraphIndex,
                    divides: [{ glyphGroup: line.glyphs }],
                  })),
                },
              ],
            },
          ],
        },
      ],
    }),
  }
}

describe('logicalGlyphContent', () => {
  it('un-reverses multi-char Arabic chunks (ArabicHandler order)', () => {
    expect(logicalGlyphContent(reverse(AR_PRICE))).toBe(AR_PRICE)
  })

  it('leaves single chars, Hebrew and Latin untouched', () => {
    expect(logicalGlyphContent('ع')).toBe('ع')
    expect(logicalGlyphContent(HE_SHALOM)).toBe(HE_SHALOM)
    expect(logicalGlyphContent('abc')).toBe('abc')
    expect(logicalGlyphContent(`${AR_PRICE} `)).toBe(`${AR_PRICE} `)
  })
})

describe('classifyBidiGlyph', () => {
  it('classifies scripts and digits', () => {
    expect(classifyBidiGlyph(AR_PRICE)).toBe('AL')
    expect(classifyBidiGlyph(HE_SHALOM.slice(0, 1))).toBe('R')
    expect(classifyBidiGlyph('a')).toBe('L')
    expect(classifyBidiGlyph('7')).toBe('EN')
    expect(classifyBidiGlyph(AR_INDIC)).toBe('AN')
    expect(classifyBidiGlyph(' ')).toBe('WS')
    expect(classifyBidiGlyph('(')).toBe('ON')
  })

  it('lets a strong letter win over leading digits in a mixed chunk', () => {
    expect(classifyBidiGlyph(`${AR_INDIC}${AR_PRICE}`)).toBe('AL')
  })

  it('honors directional marks', () => {
    expect(classifyBidiGlyph('‏')).toBe('R')
    expect(classifyBidiGlyph('‎')).toBe('L')
  })
})

describe('resolveBidiLevels + bidiVisualOrder', () => {
  it('keeps pure LTR text in order', () => {
    const levels = resolveBidiLevels(['L', 'WS', 'EN', 'WS', 'L'], false)
    expect(bidiVisualOrder(levels)).toEqual([0, 1, 2, 3, 4])
  })

  it('reorders an RTL paragraph with embedded European digits', () => {
    // AL WS EN EN WS AL — digits stay LTR inside the reversed line
    const levels = resolveBidiLevels(['AL', 'WS', 'EN', 'EN', 'WS', 'AL'], true)
    expect(bidiVisualOrder(levels)).toEqual([5, 4, 2, 3, 1, 0])
  })

  it('swaps adjacent Arabic words inside an LTR paragraph', () => {
    const levels = resolveBidiLevels(['L', 'WS', 'AL', 'WS', 'AL', 'WS', 'L'], false)
    expect(bidiVisualOrder(levels)).toEqual([0, 1, 4, 3, 2, 5, 6])
  })

  it('resolves neutrals between RTL letters to RTL', () => {
    const levels = resolveBidiLevels(['R', 'ON', 'R'], true)
    expect(levels).toEqual([1, 1, 1])
  })

  // W2 seeds from sos (R in an RTL paragraph, never AL): matches python-bidi
  // for both scripts.
  it('keeps EN digits and their percent sign together in a Hebrew paragraph', () => {
    // "shalom 10%" — % joins the digits (W5 on EN), all level 2
    expect(resolveBidiLevels(['R', 'WS', 'EN', 'EN', 'ET'], true)).toEqual([1, 1, 2, 2, 2])
    // leading digits with no strong letter before them stay EN too
    expect(resolveBidiLevels(['EN', 'EN', 'ET', 'WS', 'R'], true)).toEqual([2, 2, 2, 1, 1])
  })

  it('turns EN after an Arabic letter into AN, detaching the percent sign', () => {
    // "<AR> 10%" — W2 makes the digits AN, so W5 no longer captures the ET
    expect(resolveBidiLevels(['AL', 'WS', 'EN', 'EN', 'ET'], true)).toEqual([1, 1, 2, 2, 1])
  })

  it('keeps an NBSP-separated number as one run in an RTL paragraph', () => {
    // NBSP is CS, not WS — "1<NBSP>234" must not split (W4), per python-bidi
    expect(classifyBidiGlyph('\u00A0')).toBe('CS')
    expect(resolveBidiLevels(['R', 'WS', 'EN', 'CS', 'EN', 'EN', 'EN'], true)).toEqual([
      1, 1, 2, 2, 2, 2, 2,
    ])
  })
})

describe('reorderRichCellSkeleton', () => {
  it('restores Arabic glyph content and seats an RTL line in visual order', () => {
    const glyphs = univerGlyphs([reverse(AR_PRICE), ' ', '1', '4', '4', '6', ' ', reverse(AR_HEH)])
    const dataStream = `${AR_PRICE} 1446 ${AR_HEH}\r\n`
    const skeleton = makeSkeleton('rich-cell', dataStream, [
      { paragraphIndex: dataStream.length - 2, glyphs },
    ])
    reorderRichCellSkeleton(skeleton)
    expect(visualContents(glyphs)).toEqual([AR_HEH, ' ', '1', '4', '4', '6', ' ', AR_PRICE])
    expect(glyphs[0]?.content).toBe(AR_PRICE)
    // widths travel with their glyphs, so run styling stays attached
    expect(glyphs.reduce((sum, g) => Math.max(sum, g.left + g.width), 0)).toBe(130)
  })

  it('mirrors paired brackets on RTL runs', () => {
    const he = [...HE_SHALOM]
    const glyphs = univerGlyphs([...he, ' ', '(', ...he, ')', ' ', '1', '2', '3'])
    const dataStream = `${HE_SHALOM} (${HE_SHALOM}) 123\r\n`
    const skeleton = makeSkeleton('rich-cell', dataStream, [
      { paragraphIndex: dataStream.length - 2, glyphs },
    ])
    reorderRichCellSkeleton(skeleton)
    const visual = visualContents(glyphs)
    expect(visual.slice(0, 4)).toEqual(['1', '2', '3', ' '])
    expect(visual[4]).toBe('(')
    expect(visual[9]).toBe(')')
    expect(visual.slice(5, 9)).toEqual([...HE_SHALOM].reverse())
    expect(visual.slice(11)).toEqual([...HE_SHALOM].reverse())
  })

  it('is idempotent across repeated calculate passes', () => {
    const glyphs = univerGlyphs([reverse(AR_PRICE), ' ', reverse(AR_HEH)])
    const dataStream = `${AR_PRICE} ${AR_HEH}\r\n`
    const skeleton = makeSkeleton('rich-cell', dataStream, [
      { paragraphIndex: dataStream.length - 2, glyphs },
    ])
    reorderRichCellSkeleton(skeleton)
    const after = glyphs.map((g) => ({ ...g }))
    reorderRichCellSkeleton(skeleton)
    expect(glyphs).toEqual(after)
  })

  it('reprocesses a line whose glyphs were rebuilt by a relayout', () => {
    const line: {
      paragraphIndex?: number
      glyphs: { content: string; width: number; left: number }[]
    } = { glyphs: univerGlyphs([reverse(AR_PRICE), ' ', reverse(AR_HEH)]) }
    const dataStream = `${AR_PRICE} ${AR_HEH}\r\n`
    line.paragraphIndex = dataStream.length - 2
    const skeleton = makeSkeleton('rich-cell', dataStream, [line])
    reorderRichCellSkeleton(skeleton)
    // simulate an incremental relayout handing out fresh reversed glyphs
    line.glyphs = univerGlyphs([reverse(AR_PRICE), ' ', reverse(AR_HEH)])
    reorderRichCellSkeleton(skeleton)
    expect(visualContents(line.glyphs)).toEqual([AR_HEH, ' ', AR_PRICE])
  })

  it('ignores non rich-cell documents (the in-cell editor)', () => {
    const glyphs = univerGlyphs([reverse(AR_PRICE), ' ', reverse(AR_HEH)])
    const before = glyphs.map((g) => ({ ...g }))
    const skeleton = makeSkeleton('__INTERNAL_EDITOR__DOCS_NORMAL', `${AR_PRICE} ${AR_HEH}\r\n`, [
      { glyphs },
    ])
    reorderRichCellSkeleton(skeleton)
    expect(glyphs).toEqual(before)
  })

  it('leaves LTR-only rich documents untouched', () => {
    const glyphs = univerGlyphs(['T', 'o', 't', 'a', 'l', ' ', '1', '4'])
    const before = glyphs.map((g) => ({ ...g }))
    const skeleton = makeSkeleton('rich-cell', 'Total 14\r\n', [{ paragraphIndex: 8, glyphs }])
    reorderRichCellSkeleton(skeleton)
    expect(glyphs).toEqual(before)
  })

  it('reorders an RTL run with trailing digits inside an LTR paragraph', () => {
    // UAX#9: the digits become AN after the Arabic word and sit LEFT of it
    // ("Pr 10 <AR>"), verified against python-bidi.
    const glyphs = univerGlyphs(['P', 'r', ' ', reverse(AR_PRICE), ' ', '1', '0'])
    const dataStream = `Pr ${AR_PRICE} 10\r\n`
    const skeleton = makeSkeleton('rich-cell', dataStream, [
      { paragraphIndex: dataStream.length - 2, glyphs },
    ])
    reorderRichCellSkeleton(skeleton)
    expect(visualContents(glyphs)).toEqual(['P', 'r', ' ', '1', '0', ' ', AR_PRICE])
  })
})

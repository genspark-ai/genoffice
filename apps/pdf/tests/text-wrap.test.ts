// @vitest-environment jsdom
import { beforeAll, describe, expect, it, vi } from 'vitest'
import {
  joinBlockLines,
  mapLineRangeToBlock,
  spliceBlockText,
  wrapText,
} from '../src/renderer/text-wrap'

// jsdom has no canvas: stub measureText with a fixed per-char width (0.5×font size
// for ASCII, 1× for others — the CJK em-square convention)
beforeAll(() => {
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({
    font: '',
    measureText: (s: string) => ({
      width: [...s].reduce((w, ch) => w + (ch.charCodeAt(0) < 128 ? 50 : 100), 0),
    }),
  } as unknown as CanvasRenderingContext2D)
})

const FAMILY = 'stub'

describe('wrapText', () => {
  it('wraps Latin at word boundaries', () => {
    // 10pt font → ASCII char is 5pt; 100pt width fits 20 chars
    expect(wrapText('aaaa bbbb cccc dddd eeee', 100, 10, FAMILY)).toEqual([
      'aaaa bbbb cccc dddd',
      'eeee',
    ])
  })

  it('breaks CJK after any character', () => {
    // CJK char is 10pt at 10pt font; width 35 fits 3 per line
    expect(wrapText('一二三四五六七', 35, 10, FAMILY)).toEqual(['一二三', '四五六', '七'])
  })

  it('never starts a line with closing punctuation (kinsoku)', () => {
    // Width fits exactly 3 CJK chars; the 。 would land at line start —
    // the preceding char moves down with it
    const lines = wrapText('一二三。四五', 30, 10, FAMILY)
    expect(lines.every((l) => !'，。'.includes(l[0]!))).toBe(true)
    expect(lines.join('')).toBe('一二三。四五')
  })

  it('never ends a line with an opening bracket', () => {
    const lines = wrapText('一二（三四五', 30, 10, FAMILY)
    expect(lines.every((l) => !l.endsWith('（'))).toBe(true)
    expect(lines.join('')).toBe('一二（三四五')
  })

  it('hard-breaks a single unit wider than the block', () => {
    const lines = wrapText('aaaaaaaaaaaaaaaaaaaaaaaa', 50, 10, FAMILY)
    expect(lines.length).toBeGreaterThan(1)
    expect(lines.join('')).toBe('aaaaaaaaaaaaaaaaaaaaaaaa')
  })

  it('returns one line when everything fits', () => {
    expect(wrapText('short', 1000, 10, FAMILY)).toEqual(['short'])
  })
})

describe('joinBlockLines', () => {
  it('joins Latin lines with a space', () => {
    expect(joinBlockLines(['the quick brown', 'fox jumps'])).toBe('the quick brown fox jumps')
  })

  it('joins CJK lines without a space', () => {
    expect(joinBlockLines(['段落聚类以基线', '分桶为第一步'])).toBe('段落聚类以基线分桶为第一步')
  })

  it('keeps a space at mixed boundaries', () => {
    expect(joinBlockLines(['uses pdf.js', '提取文本'])).toBe('uses pdf.js 提取文本')
  })
})

describe('spliceBlockText', () => {
  it('folds one line edit into the paragraph', () => {
    expect(
      spliceBlockText('alpha beta gamma delta', [{ oldText: 'gamma', newText: 'gamma123' }]),
    ).toBe('alpha beta gamma123 delta')
  })

  it('matches whitespace-insensitively (line joins differ in synthesized spaces)', () => {
    // The block joins with a space, the DOM line grouping produced none
    expect(
      spliceBlockText('the quick brown fox jumps', [
        { oldText: 'quick  brown fox', newText: 'lazy dog' },
      ]),
    ).toBe('the lazy dog jumps')
  })

  it('folds several non-overlapping edits by document order', () => {
    expect(
      spliceBlockText('abcdefghij', [
        { oldText: 'hi', newText: '8_9' },
        { oldText: 'bc', newText: '2_3' },
      ]),
    ).toBe('a2_3defg8_9j')
  })

  it('uses the hint to pick the right occurrence of a repeated line', () => {
    expect(spliceBlockText('samesame', [{ oldText: 'same', newText: 'X', hint: 4 }])).toBe('sameX')
  })

  it('returns null when an oldText is not in the block', () => {
    expect(spliceBlockText('abc', [{ oldText: 'de', newText: 'x' }])).toBeNull()
  })

  it('returns null when two edits overlap', () => {
    expect(
      spliceBlockText('abcde', [
        { oldText: 'abc', newText: 'a' },
        { oldText: 'cd', newText: 'b' },
      ]),
    ).toBeNull()
  })

  it('unifies radical-block codepoints on both sides (pdf.js cmap quirk)', () => {
    // pdf.js extracted Kangxi radicals (U+2F83/U+2F17 for U+81EA/U+5341); the
    // block text and the folded result must carry the real unified ideographs
    // PDFium extracts and fonts can draw
    expect(joinBlockLines(['\u2f83\u5efa\u5b89\u2f17\u4e09\u5e74'])).toBe(
      '\u81ea\u5efa\u5b89\u5341\u4e09\u5e74',
    )
    expect(
      spliceBlockText('\u81ea\u5efa\u5b89\u5341\u4e09\u5e74\u81f3\u5609\u5e73\u4e94\u5e74', [
        {
          oldText: '\u2f83\u5efa\u5b89\u2f17\u4e09\u5e74',
          newText: '\u2f83\u5efa\u5b89\u2f17\u4e09\u5e74X',
        },
      ]),
    ).toBe('\u81ea\u5efa\u5b89\u5341\u4e09\u5e74X\u81f3\u5609\u5e73\u4e94\u5e74')
  })

  it('unifies Radicals Supplement codepoints that NFKC cannot decompose', () => {
    // U+2EE6 (C-simplified bird) and U+2ECB (C-simplified cart) have no NFKC
    // decomposition; only the shared RADICAL_EQUIV table folds them to the
    // unified ideographs (U+9E1F, U+8F66) PDFium extracts
    expect(joinBlockLines(['\u83dc\u2ee6\u4e0e\u2ecb\u8f86'])).toBe(
      '\u83dc\u9e1f\u4e0e\u8f66\u8f86',
    )
    expect(
      spliceBlockText('\u83dc\u9e1f\u4e0e\u8f66\u8f86', [
        { oldText: '\u83dc\u2ee6', newText: '\u83dc\u2ee6X' },
      ]),
    ).toBe('\u83dc\u9e1fX\u4e0e\u8f66\u8f86')
  })

  it('folds CJK Strokes even when no radical-block char is present', () => {
    // U+31CF (stroke) folds to U+4E40; the gate must not require U+2E80-2FDF
    expect(joinBlockLines(['\u31cfX'])).toBe('\u4e40X')
    expect(spliceBlockText('\u4e40X rest', [{ oldText: '\u31cfX', newText: 'Y' }])).toBe('Y rest')
  })

  it('aligns an unfolded radical against an already-folded astral ideograph', () => {
    // U+2E87 folds to U+20628 (supplementary plane, two UTF-16 units): the
    // haystack carries the surrogates as separate units, the needle folds
    // one radical into both — unit-level entries keep the ranks aligned
    expect(spliceBlockText('\u{20628}\u5b50', [{ oldText: '\u2e87', newText: 'X' }])).toBe(
      'X\u5b50',
    )
  })

  it('an edit spanning a visual line break still folds (crossing spans)', () => {
    // oldText mixes the tail of line 1 and the head of line 2, carrying the
    // join space the block text does not have at that position
    expect(
      spliceBlockText('headtail nexttext rest', [
        { oldText: 'tail  next', newText: 'tail next123' },
      ]),
    ).toBe('headtail next123text rest')
  })

  it('reports where each newText landed via outRanges (input-edit order)', () => {
    const out: [number, number][] = []
    const folded = spliceBlockText(
      'abcdefghij',
      [
        { oldText: 'hi', newText: '8_9' },
        { oldText: 'bc', newText: '2_3' },
      ],
      out,
    )
    expect(folded).toBe('a2_3defg8_9j')
    // outRanges follows the input order even though the splice emits by position
    expect(out).toEqual([
      [8, 11],
      [1, 4],
    ])
    expect(folded!.slice(...out[0]!)).toBe('8_9')
    expect(folded!.slice(...out[1]!)).toBe('2_3')
  })

  it('outRanges covers the radical-folded newText the splice inserted', () => {
    const out: [number, number][] = []
    const folded = spliceBlockText(
      '\u81ea\u5efa\u5b89',
      [{ oldText: '\u2f83', newText: '\u2f83X' }],
      out,
    )
    expect(folded).toBe('\u81eaX\u5efa\u5b89')
    expect(folded!.slice(...out[0]!)).toBe('\u81eaX')
  })
})

describe('mapLineRangeToBlock', () => {
  it('maps a mid-line range to block offsets', () => {
    // block = two joined lines; select "bravo" on line 1
    const block = 'alpha bravo charlie delta'
    expect(mapLineRangeToBlock(block, 'alpha bravo charlie', 6, 11)).toEqual([6, 11])
  })

  it('ignores whitespace differences between DOM line and joined text', () => {
    // DOM line synthesizes a double space the block text does not have
    const block = 'alpha bravo charlie'
    expect(mapLineRangeToBlock(block, 'alpha  bravo', 7, 12)).toEqual([6, 11])
  })

  it('locates a later line inside the block', () => {
    const block = '\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b'
    // line 2 of the paragraph: chars 5..8; select its middle two
    expect(mapLineRangeToBlock(block, '\u4e94\u516d\u4e03\u516b', 1, 3)).toEqual([5, 7])
  })

  it('folds radical variants like the splice matcher', () => {
    // DOM line carries U+2EE6 where the block text has the unified U+9E1F
    expect(mapLineRangeToBlock('\u83dc\u9e1f\u4e0e', '\u83dc\u2ee6\u4e0e', 1, 2)).toEqual([1, 2])
  })

  it('maps across an astral fold without losing unit alignment', () => {
    // line has the radical U+2E87 (1 unit), block has its fold U+20628 (2 units)
    expect(mapLineRangeToBlock('\u{20628}\u5b50', '\u2e87\u5b50', 0, 1)).toEqual([0, 2])
    expect(mapLineRangeToBlock('\u{20628}\u5b50', '\u2e87\u5b50', 1, 2)).toEqual([2, 3])
  })

  it('collapses a whitespace-only range to a caret', () => {
    expect(mapLineRangeToBlock('alpha bravo', 'alpha bravo', 5, 6)).toEqual([6, 6])
  })

  it('returns null when the line is not in the block', () => {
    expect(mapLineRangeToBlock('alpha bravo', 'zulu yankee', 0, 4)).toBeNull()
  })
})

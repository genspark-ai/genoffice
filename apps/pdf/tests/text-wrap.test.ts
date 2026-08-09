// @vitest-environment jsdom
import { beforeAll, describe, expect, it, vi } from 'vitest'
import { joinBlockLines, wrapText } from '../src/renderer/text-wrap'

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

import { describe, expect, it } from 'vitest'
import type { PageEntry } from '../src/renderer/search'
import { groupPageBlocks } from '../src/renderer/text-block'

interface Frag {
  t: string
  x: number
  y: number
  w: number
  h: number
  rot?: boolean
}

const page = (frags: Frag[]): PageEntry => {
  let text = ''
  const items = frags.map((f) => {
    const start = text.length
    text += f.t
    return {
      start,
      end: text.length,
      x: f.x,
      y: f.y,
      w: f.w,
      h: f.h,
      ...(f.rot ? { rot: true as const } : {}),
    }
  })
  return { text, lower: text.toLowerCase(), items }
}

/** One 200pt-wide body line at the given baseline */
const bodyLine = (t: string, y: number, x = 50): Frag => ({ t, x, y, w: 200, h: 12 })

describe('groupPageBlocks', () => {
  it('returns nothing for empty or whitespace-only pages', () => {
    expect(groupPageBlocks(page([]))).toEqual([])
    expect(groupPageBlocks(page([{ t: '   ', x: 50, y: 700, w: 20, h: 12 }]))).toEqual([])
  })

  it('merges consecutive same-width lines into one paragraph', () => {
    const blocks = groupPageBlocks(
      page([bodyLine('First line', 700), bodyLine('second line', 686), bodyLine('third.', 672)]),
    )
    expect(blocks).toHaveLength(1)
    const b = blocks[0]!
    expect(b.align).toBe('left')
    expect(b.lines.map((l) => l.text)).toEqual(['First line', 'second line', 'third.'])
    expect(b.lineHeight).toBe(14)
    expect(b.fontSize).toBe(12)
    expect(b.rect[0]).toBe(50)
    expect(b.rect[1]).toBeCloseTo(669.6) // last baseline 672 minus 0.2× descent allowance
    expect(b.rect[2]).toBe(250)
    expect(b.rect[3]).toBe(712)
  })

  it('splits paragraphs at vertical gaps beyond 1.5× the leading', () => {
    const blocks = groupPageBlocks(
      page([
        bodyLine('para one a', 700),
        bodyLine('para one b', 686),
        bodyLine('para two a', 640),
        bodyLine('para two b', 626),
      ]),
    )
    expect(blocks.map((b) => b.lines.length)).toEqual([2, 2])
  })

  it('keeps a heading out of the body paragraph (font-size mismatch)', () => {
    const blocks = groupPageBlocks(
      page([
        { t: 'Heading', x: 50, y: 716, w: 120, h: 18 },
        bodyLine('body one', 700),
        bodyLine('body two', 686),
      ]),
    )
    expect(blocks).toHaveLength(2)
    expect(blocks.find((b) => b.lines.length === 1)!.lines[0]!.text).toBe('Heading')
  })

  it('keeps interleaved columns apart (gutter split + per-column attachment)', () => {
    const blocks = groupPageBlocks(
      page([
        { t: 'left one', x: 50, y: 700, w: 200, h: 12 },
        { t: 'right one', x: 300, y: 700, w: 200, h: 12 },
        { t: 'left two', x: 50, y: 686, w: 200, h: 12 },
        { t: 'right two', x: 300, y: 686, w: 200, h: 12 },
      ]),
    )
    expect(blocks).toHaveLength(2)
    const texts = blocks.map((b) => b.lines.map((l) => l.text))
    expect(texts).toContainEqual(['left one', 'left two'])
    expect(texts).toContainEqual(['right one', 'right two'])
  })

  it('merges CJK body text with generous (2×) leading', () => {
    const blocks = groupPageBlocks(
      page([
        { t: '由一百多座小岛与桥梁编织而成', x: 50, y: 700, w: 200, h: 10 },
        { t: '没有车马只有贡多拉悠悠划过', x: 50, y: 680, w: 200, h: 10 },
        { t: '威尼斯把慢写进了它的地图里', x: 50, y: 660, w: 200, h: 10 },
      ]),
    )
    expect(blocks).toHaveLength(1)
    expect(blocks[0]!.lineHeight).toBe(20)
  })

  it('re-splits a stray same-size line glued by the generous bootstrap', () => {
    // A standalone line 22pt above a paragraph whose own leading is 12pt:
    // bootstrap merges it, the outlier pass must split it back off
    const blocks = groupPageBlocks(
      page([
        { t: 'standalone lead-in', x: 50, y: 700, w: 200, h: 10 },
        { t: 'body line one here', x: 50, y: 678, w: 200, h: 10 },
        { t: 'body line two here', x: 50, y: 666, w: 200, h: 10 },
        { t: 'body line three xx', x: 50, y: 654, w: 200, h: 10 },
      ]),
    )
    expect(blocks.map((b) => b.lines.length).sort()).toEqual([1, 3])
  })

  it('re-splits a stray line glued onto a two-line paragraph (two leads)', () => {
    // Leads are [22, 12]: a plain median would pick the stray 22 itself and approve
    // it — the lower median (12) must expose it as an outlier
    const blocks = groupPageBlocks(
      page([
        { t: 'standalone lead-in', x: 50, y: 700, w: 200, h: 10 },
        { t: 'body line one here', x: 50, y: 678, w: 200, h: 10 },
        { t: 'body line two here', x: 50, y: 666, w: 200, h: 10 },
      ]),
    )
    expect(blocks.map((b) => b.lines.length).sort()).toEqual([1, 2])
  })

  it('re-splits a two-line group whose lone gap exceeds any designed leading', () => {
    // Gap 24pt on a 10pt face: within the 2.5× bootstrap but beyond the ~2× cap of
    // real body leading, so this is two blocks, not one paragraph
    const blocks = groupPageBlocks(
      page([
        { t: 'stray title line', x: 50, y: 700, w: 200, h: 10 },
        { t: 'first body line', x: 50, y: 676, w: 200, h: 10 },
      ]),
    )
    expect(blocks.map((b) => b.lines.length)).toEqual([1, 1])
  })

  it('keeps a two-line CJK paragraph with designed 2× leading together', () => {
    const blocks = groupPageBlocks(
      page([
        { t: '由一百多座小岛与桥梁编织而成', x: 50, y: 700, w: 200, h: 10 },
        { t: '没有车马只有贡多拉悠悠划过桥', x: 50, y: 680, w: 200, h: 10 },
      ]),
    )
    expect(blocks).toHaveLength(1)
    expect(blocks[0]!.lines).toHaveLength(2)
  })

  it('splits adjacent labeled paragraphs at a paragraph-final short line', () => {
    // Footer-box style: two labeled paragraphs, same leading and font size; the
    // first ends in a line far shorter than the shared extent, so the next line
    // starts a new paragraph even though leading alone cannot tell them apart
    const blocks = groupPageBlocks(
      page([
        {
          t: '\u6570\u636e\u6765\u6e90\uff1a\u5e74\u62a5\u53ca\u5b63\u5ea6\u4e1a\u7ee9\u516c\u544a\u62ab\u9732\u6587\u4ef6',
          x: 50,
          y: 700,
          w: 400,
          h: 10,
        },
        {
          t: '\u516c\u5f00\u5e02\u573a\u884c\u60c5\u6570\u636e\u3002',
          x: 50,
          y: 681,
          w: 90,
          h: 10,
        },
        {
          t: '\u514d\u8d23\u58f0\u660e\uff1a\u672c\u62a5\u544a\u4ec5\u4f9b\u7814\u7a76\u53c2\u8003\u4e0d\u6784\u6210\u6295\u8d44\u5efa\u8bae',
          x: 50,
          y: 662,
          w: 400,
          h: 10,
        },
      ]),
    )
    expect(blocks.map((b) => b.lines.length)).toEqual([2, 1])
  })

  it('keeps a ragged wrap that the next word explains (no paragraph-end split)', () => {
    // Line 1 ends 60pt short, but the next line starts with a 10-char word
    // (~72pt at 12pt): the wrap is forced, not a paragraph boundary
    const blocks = groupPageBlocks(
      page([
        { t: 'ragged text before a', x: 50, y: 700, w: 160, h: 12 },
        { t: 'wideworddd continues here', x: 50, y: 686, w: 220, h: 12 },
        { t: 'and ends.', x: 50, y: 672, w: 70, h: 12 },
      ]),
    )
    expect(blocks).toHaveLength(1)
  })

  it('tolerates a first-line indent', () => {
    const blocks = groupPageBlocks(
      page([{ t: 'indented', x: 62, y: 700, w: 188, h: 12 }, bodyLine('flush left line', 686)]),
    )
    expect(blocks).toHaveLength(1)
  })

  it('merges per-glyph CJK runs into one line without synthesized spaces', () => {
    const glyphs: Frag[] = Array.from({ length: 6 }, (_, i) => ({
      t: String.fromCharCode(0x4e00 + i),
      x: 50 + i * 12,
      y: 700,
      w: 12,
      h: 12,
    }))
    const blocks = groupPageBlocks(
      page([...glyphs, { t: '\u7b2c\u4e8c\u884c\u7b2c\u4e8c\u884c', x: 50, y: 686, w: 72, h: 12 }]),
    )
    expect(blocks).toHaveLength(1)
    expect(blocks[0]!.lines[0]!.text).toBe('一丁丂七丄丅')
  })

  it('synthesizes a space at visible gaps between runs on one baseline', () => {
    const blocks = groupPageBlocks(
      page([
        { t: 'Hello', x: 50, y: 700, w: 40, h: 12 },
        { t: 'world', x: 95, y: 700, w: 40, h: 12 },
      ]),
    )
    expect(blocks[0]!.lines[0]!.text).toBe('Hello world')
  })

  it('ignores rotated runs', () => {
    const blocks = groupPageBlocks(
      page([bodyLine('normal', 700), { t: 'rotated', x: 400, y: 400, w: 60, h: 12, rot: true }]),
    )
    expect(blocks).toHaveLength(1)
    expect(blocks[0]!.lines[0]!.text).toBe('normal')
  })

  it('detects centered blocks by their common midline', () => {
    const blocks = groupPageBlocks(
      page([
        { t: 'a centered heading', x: 100, y: 700, w: 100, h: 12 },
        { t: 'short', x: 125, y: 686, w: 50, h: 12 },
        { t: 'medium line here', x: 110, y: 672, w: 80, h: 12 },
      ]),
    )
    expect(blocks).toHaveLength(1)
    expect(blocks[0]!.align).toBe('center')
  })

  it('detects right-aligned blocks by their common right edge', () => {
    const blocks = groupPageBlocks(
      page([
        { t: 'flush to the right edge', x: 150, y: 700, w: 100, h: 12 },
        { t: 'short one', x: 200, y: 686, w: 50, h: 12 },
        { t: 'medium right', x: 170, y: 672, w: 80, h: 12 },
      ]),
    )
    expect(blocks).toHaveLength(1)
    expect(blocks[0]!.align).toBe('right')
  })

  it('detects centered lines of similar width (shared left edge must not win)', () => {
    // All left edges agree within the tolerance, but the centers agree better:
    // a common multi-line centered heading, not a left-aligned paragraph
    const blocks = groupPageBlocks(
      page([
        { t: 'centered heading one', x: 100, y: 700, w: 100, h: 12 },
        { t: 'centered heading 2', x: 104, y: 686, w: 92, h: 12 },
        { t: 'centered heading iii', x: 102, y: 672, w: 96, h: 12 },
      ]),
    )
    expect(blocks).toHaveLength(1)
    expect(blocks[0]!.align).toBe('center')
  })

  it('detects right-aligned lines of similar width (shared left edge must not win)', () => {
    const blocks = groupPageBlocks(
      page([
        { t: 'right aligned line', x: 150, y: 700, w: 100, h: 12 },
        { t: 'right aligned 2', x: 154, y: 686, w: 96, h: 12 },
        { t: 'right aligned iii', x: 152, y: 672, w: 98, h: 12 },
      ]),
    )
    expect(blocks).toHaveLength(1)
    expect(blocks[0]!.align).toBe('right')
  })

  it('detects a two-line centered block whose last line is longest', () => {
    const blocks = groupPageBlocks(
      page([
        { t: 'short title', x: 125, y: 700, w: 50, h: 12 },
        { t: 'a longer centered second line', x: 100, y: 686, w: 100, h: 12 },
      ]),
    )
    expect(blocks).toHaveLength(1)
    expect(blocks[0]!.align).toBe('center')
  })

  it('detects a two-line right-aligned block', () => {
    const blocks = groupPageBlocks(
      page([
        { t: 'flush right longer', x: 150, y: 700, w: 100, h: 12 },
        { t: 'short', x: 200, y: 686, w: 50, h: 12 },
      ]),
    )
    expect(blocks).toHaveLength(1)
    expect(blocks[0]!.align).toBe('right')
  })

  it('keeps a two-line indented paragraph left-aligned', () => {
    const blocks = groupPageBlocks(
      page([
        { t: 'indented first line xx', x: 62, y: 700, w: 188, h: 12 },
        { t: 'flush second line yyyy', x: 50, y: 686, w: 200, h: 12 },
      ]),
    )
    expect(blocks).toHaveLength(1)
    expect(blocks[0]!.align).toBe('left')
  })

  it('keeps a hanging-indent paragraph left-aligned', () => {
    // First line flush, body lines inset one em, right edges flush: bibliography style
    const blocks = groupPageBlocks(
      page([
        { t: 'Author, A. (2026). The', x: 50, y: 700, w: 200, h: 12 },
        { t: 'title of the work meq', x: 62, y: 686, w: 188, h: 12 },
        { t: 'publisher, place.', x: 62, y: 672, w: 140, h: 12 },
      ]),
    )
    expect(blocks).toHaveLength(1)
    expect(blocks[0]!.align).toBe('left')
  })

  it('keeps a two-line hanging indent left-aligned', () => {
    const blocks = groupPageBlocks(
      page([
        { t: 'Author, B. (2026). Any', x: 50, y: 700, w: 200, h: 12 },
        { t: 'inset second line meq', x: 62, y: 686, w: 188, h: 12 },
      ]),
    )
    expect(blocks).toHaveLength(1)
    expect(blocks[0]!.align).toBe('left')
  })

  it('single-line blocks carry no alignment evidence', () => {
    const blocks = groupPageBlocks(page([bodyLine('just one line', 700)]))
    expect(blocks[0]!.align).toBe('left')
  })

  it('does not bridge blocks whose font sizes drift within one line (dominant size wins)', () => {
    const blocks = groupPageBlocks(
      page([
        { t: 'mostly twelve ', x: 50, y: 700, w: 180, h: 12 },
        { t: 'X', x: 232, y: 700, w: 18, h: 18 },
        bodyLine('next line of the para', 686),
      ]),
    )
    expect(blocks).toHaveLength(1)
    expect(blocks[0]!.fontSize).toBe(12)
  })
})

import { describe, expect, it } from 'vitest'
import { foldMap, norm, spliceIntoEngine } from '../src/main/text-edit'
import { chainLayers } from '../src/shared/x-layers'

describe('chainLayers', () => {
  const glyphs = (starts: number[], w = 10) => starts.map((x) => ({ left: x, right: x + w }))
  it('keeps a plain line in one chain', () => {
    expect(chainLayers(glyphs([0, 12, 24, 36]))).toEqual([0, 0, 0, 0])
  })
  it('separates a rebuilt run stacked over a per-glyph line', () => {
    // The realistic stack: rebuilt lines are one object (one long span), the original
    // CJK line underneath is per-glyph — the long span conflicts with every glyph
    const rects = [
      { left: 0, right: 10 }, // glyph 1
      { left: 0, right: 200 }, // rebuilt line drawn over the whole row
      { left: 12, right: 22 }, // glyph 2
      { left: 24, right: 34 }, // glyph 3
    ]
    expect(chainLayers(rects)).toEqual([0, 1, 0, 0])
  })
  it('keeps glyphs after the shorter overlay with their own line', () => {
    // Overlay covers only the first half of the row; later glyphs still chain to the
    // glyph line, not to the overlay's tail
    const rects = [
      { left: 0, right: 10 }, // glyph 1
      { left: 0, right: 30 }, // short overlay across glyphs 1-2
      { left: 12, right: 22 }, // glyph 2
      { left: 24, right: 34 }, // glyph 3 (overlaps overlay's tail end)
      { left: 36, right: 46 }, // glyph 4 (past the overlay)
    ]
    expect(chainLayers(rects)).toEqual([0, 1, 0, 0, 0])
  })
  it('splits exact overlays', () => {
    expect(chainLayers(glyphs([0, 0, 12, 12]))).toEqual([0, 1, 0, 1])
  })
  it('does not let a short overlay ending in a word gap steal the next word', () => {
    // Overlay ends at 23, one px before the next word — nearer than the previous word's
    // tail (gap 2). Within the tie margin the many-span text chain must win.
    const rects = [
      { left: 0, right: 10 }, // word 1
      { left: 0, right: 23 }, // short overlay ending mid-gap
      { left: 12, right: 22 }, // word 2
      { left: 24, right: 34 }, // word 3 — overlay tail gap 1, word chain gap 2
    ]
    expect(chainLayers(rects, 5)).toEqual([0, 1, 0, 0])
  })
})

describe('norm', () => {
  it('folds Kangxi radicals to unified ideographs', () => {
    // pdf.js extracts U+2F00-block radicals from some fonts' cmaps; pdfium extracts
    // the unified chars — both sides must land on the same key
    expect(norm('⼆〇⼆五')).toBe(norm('二〇二五'))
  })
  it('stays whitespace-insensitive', () => {
    expect(norm('a b\tc\n')).toBe('abc')
  })
})

describe('foldMap', () => {
  it('maps folded units back to source indices', () => {
    const { units, idx } = foldMap('a⼆ b')
    expect(units).toEqual(['a', '二', 'b'])
    expect(idx).toEqual([0, 1, 3, 4])
  })
  it('collapses whitespace runs to one unit when keeping spaces', () => {
    const { units } = foldMap('a  b', true)
    expect(units).toEqual(['a', ' ', 'b'])
  })
})

describe('spliceIntoEngine', () => {
  it('keeps engine codepoints on unedited text', () => {
    // Renderer-side oldText carries the radical variant; the engine text must survive
    expect(spliceIntoEngine('签署日期为', '签署⽇期为', '签署⽇期为好')).toBe('签署日期为好')
  })
  it('applies mid-string replacements', () => {
    expect(
      spliceIntoEngine('The quick brown fox', 'The quick brown fox', 'The quick red fox'),
    ).toBe('The quick red fox')
  })
  it('keeps engine spacing when the renderer synthesized spaces', () => {
    expect(spliceIntoEngine('签署日期', '签署 日期', '签署 好期')).toBe('签署好期')
  })
  it('supports pure insertion of a space', () => {
    expect(spliceIntoEngine('helloworld', 'helloworld', 'hello world')).toBe('hello world')
  })
  it('supports deletion', () => {
    expect(spliceIntoEngine('abcdef', 'abcdef', 'abef')).toBe('abef')
  })
  it('returns the engine text unchanged for a no-op edit', () => {
    expect(spliceIntoEngine('日期', '⽇期', '⽇期')).toBe('日期')
  })
  it('joins typed text directly onto a retained word with no stray space', () => {
    expect(spliceIntoEngine('Refont me please', 'Refont me please', 'Refonted text')).toBe(
      'Refonted text',
    )
  })
  it('keeps single engine spaces around a replaced word', () => {
    expect(
      spliceIntoEngine('The quick brown fox', 'The quick brown fox', 'The quick red fox'),
    ).toBe('The quick red fox')
  })
})

describe('spliceIntoEngine (multi-line block reflow)', () => {
  // Engine text is the join of per-line objects with no separator; oldText is the
  // renderer's joinBlockLines (space at Latin boundaries); newText is the reflow
  // with '\n' between the new lines. The '\n's must survive the splice, or the
  // rebuilt run collapses into one overflowing line at save.
  it('keeps reflow line breaks on unedited Latin prefix lines', () => {
    expect(
      spliceIntoEngine(
        'The quick brownfox jumps', // two engine objects: 'The quick brown' + 'fox jumps'
        'The quick brown fox jumps',
        'The quick brown fox\nleaps',
      ),
    ).toBe('The quick brown fox\nleaps')
  })
  it('keeps line breaks when nothing changed (pure reflow)', () => {
    expect(spliceIntoEngine('helloworld', 'hello world', 'hello\nworld')).toBe('hello\nworld')
  })
  it('keeps engine codepoints on fully retained lines', () => {
    // Engine has the unified ideograph, renderer extracted the Kangxi radical: the
    // untouched first line must come back with the engine's own chars
    expect(spliceIntoEngine('签署日期为今天', '签署⽇期为今天', '签署⽇期为\n今天好')).toBe(
      '签署日期为\n今天好',
    )
  })
  it('rebuilds suffix lines from the engine after an edit at the start', () => {
    expect(spliceIntoEngine('大家好威尼斯真美', '大家好威尼斯真美', '嗨大家好\n威尼斯真美')).toBe(
      '嗨大家好\n威尼斯真美',
    )
  })
  it('restores the space where a reflow moved a word across an engine line join', () => {
    // 'brown' and 'fox' were split by the original visual line break (engine join has
    // no char between them); the reflow now puts them on one line — a space must appear
    expect(
      spliceIntoEngine(
        'The quick brownfox jumps high', // 'The quick brown' + 'fox jumps high'
        'The quick brown fox jumps high',
        'The quick brown fox\njumps higher',
      ),
    ).toBe('The quick brown fox\njumps higher')
  })
})

import { describe, it, expect } from 'vitest'
import { HeuristicMetrics, OpentypeMetrics, type OpentypeFontLike } from '../src/metrics'
import { layoutText } from '../src/text-layout'
import { makeViewport } from '../src/coords'
import type { TextBody } from '@genoffice/pptx-engine'

const vp = makeViewport({ cx: 9525 * 1000, cy: 9525 * 1000 }, 1000) // scale 1

function body(partial: Partial<TextBody> & Pick<TextBody, 'paragraphs'>): TextBody {
  return {
    anchor: 'top',
    insets: { l: 0, t: 0, r: 0, b: 0 },
    autofit: 'none',
    wrap: true,
    ...partial,
  }
}

describe('2.3 metrics', () => {
  const m = new HeuristicMetrics()

  it('CJK advances ~1em, latin narrower', () => {
    const style = { fontFamily: 'Arial', fontSizePx: 100, bold: false, italic: false }
    const cjk = m.measure('中文', style)
    const latin = m.measure('ab', style)
    expect(cjk).toBeCloseTo(200, 0) // 2 * 1em * 100px
    expect(latin).toBeLessThan(cjk)
  })

  it('opentype metrics falls back to heuristic when glyph missing (CJK in latin font)', () => {
    const latinOnly: OpentypeFontLike = {
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      getAdvanceWidth: () => 0, // opentype gives 0/notdef width for missing glyphs
      charToGlyphIndex: (ch) => (ch.codePointAt(0)! < 0x2e80 ? 1 : 0),
    }
    const om = new OpentypeMetrics(() => latinOnly)
    const style = { fontFamily: 'Latin', fontSizePx: 100, bold: false, italic: false }
    // CJK → falls back to the heuristic (≈1em per char) instead of 0
    expect(om.measure('中文', style)).toBeCloseTo(200, 0)
  })

  it('opentype metrics uses real font when available, else fallback', () => {
    const fakeFont: OpentypeFontLike = {
      unitsPerEm: 1000,
      ascender: 800,
      descender: -200,
      getAdvanceWidth: (t, s) => t.length * s * 0.5,
    }
    const om = new OpentypeMetrics((st) => (st.fontFamily === 'Real' ? fakeFont : undefined))
    const real = { fontFamily: 'Real', fontSizePx: 100, bold: false, italic: false }
    expect(om.measure('abcd', real)).toBeCloseTo(200, 4) // 4 * 100 * 0.5
    const met = om.metrics(real)
    expect(met.ascent).toBeCloseTo(80, 4)
    expect(met.descent).toBeCloseTo(20, 4)
    expect(met.lineHeight).toBeCloseTo(100, 4)
    // hhea lineGap becomes external leading (advances lines, outside the line box)
    const om2 = new OpentypeMetrics(() => ({ ...fakeFont, lineGap: 100 }))
    const met2 = om2.metrics(real)
    expect(met2.lineHeight).toBeCloseTo(100, 4)
    expect(met2.externalLeading).toBeCloseTo(10, 4)
    // Unknown fonts fall back to the heuristic (no throw)
    const unknown = { fontFamily: 'Nope', fontSizePx: 100, bold: false, italic: false }
    expect(om.measure('ab', unknown)).toBeGreaterThan(0)
  })
})

describe('2.3 text layout', () => {
  it('single short line stays one line', () => {
    const layout = layoutText({
      body: body({ paragraphs: [{ runs: [{ text: 'Hi', fontSize: 18 }] }] }),
      boxWidthPx: 400,
      boxHeightPx: 200,
      metrics: new HeuristicMetrics(),
      vp,
    })
    expect(layout.lines.length).toBe(1)
    expect(layout.lines[0]!.runs.map((r) => r.text).join('')).toBe('Hi')
  })

  it('wraps long text into multiple lines at word boundaries', () => {
    const long = 'word '.repeat(40).trim()
    const layout = layoutText({
      body: body({ paragraphs: [{ runs: [{ text: long, fontSize: 18 }] }] }),
      boxWidthPx: 200,
      boxHeightPx: 2000,
      metrics: new HeuristicMetrics(),
      vp,
    })
    expect(layout.lines.length).toBeGreaterThan(1)
    // No line is wider than the available width
    for (const ln of layout.lines) {
      const w = ln.runs.reduce((a, r) => a + r.widthPx, 0)
      expect(w).toBeLessThanOrEqual(200 + 1)
    }
  })

  it('empty paragraph with a textless marker run keeps its style (line height + readback)', () => {
    const marked = layoutText({
      body: body({ paragraphs: [{ runs: [{ text: '', fontSize: 32 }] }] }),
      boxWidthPx: 400,
      boxHeightPx: 200,
      metrics: new HeuristicMetrics(),
      vp,
    })
    const plain = layoutText({
      body: body({ paragraphs: [{ runs: [] }] }),
      boxWidthPx: 400,
      boxHeightPx: 200,
      metrics: new HeuristicMetrics(),
      vp,
    })
    // The marker styles the empty line: taller than the default-styled empty paragraph,
    // and exposed as a zero-width run so the editor/ribbon can read the format back
    expect(marked.lines[0]!.height).toBeGreaterThan(plain.lines[0]!.height)
    expect(plain.lines[0]!.runs.length).toBe(0)
    const run = marked.lines[0]!.runs[0]!
    expect(run.text).toBe('')
    expect(run.widthPx).toBe(0)
    expect(run.fontSizePx).toBeCloseTo((32 * 96) / 72, 1)
  })

  it('wrap=false keeps single line even if overflowing', () => {
    const long = 'word '.repeat(40).trim()
    const layout = layoutText({
      body: body({ wrap: false, paragraphs: [{ runs: [{ text: long, fontSize: 18 }] }] }),
      boxWidthPx: 200,
      boxHeightPx: 2000,
      metrics: new HeuristicMetrics(),
      vp,
    })
    expect(layout.lines.length).toBe(1)
  })

  it('CJK wraps per-character', () => {
    const cjk = '中文换行测试内容比较长需要折行的一段文字'.repeat(3)
    const layout = layoutText({
      body: body({ paragraphs: [{ runs: [{ text: cjk, fontSize: 24 }] }] }),
      boxWidthPx: 150,
      boxHeightPx: 3000,
      metrics: new HeuristicMetrics(),
      vp,
    })
    expect(layout.lines.length).toBeGreaterThan(1)
  })

  it('kinsoku: a closing mark never starts a line (its predecessor is pulled down)', () => {
    // 4 chars fit per line (4 × 24px = 96px); without kinsoku 、would head line 2
    const layout = layoutText({
      body: body({ paragraphs: [{ runs: [{ text: 'ああああ、いい', fontSize: 24 }] }] }),
      boxWidthPx: 96,
      boxHeightPx: 3000,
      metrics: new HeuristicMetrics(),
      vp,
    })
    const lines = layout.lines.map((l) => l.runs.map((r) => r.text).join(''))
    expect(lines[0]).toBe('あああ')
    expect(lines[1]!.startsWith('あ、')).toBe(true)
  })

  it('kinsoku: chained closing marks pull the whole tail down', () => {
    const layout = layoutText({
      body: body({ paragraphs: [{ runs: [{ text: 'あああ。」いい', fontSize: 24 }] }] }),
      boxWidthPx: 96,
      boxHeightPx: 3000,
      metrics: new HeuristicMetrics(),
      vp,
    })
    const lines = layout.lines.map((l) => l.runs.map((r) => r.text).join(''))
    expect(lines[0]).toBe('ああ')
    expect(lines[1]!.startsWith('あ。」')).toBe(true)
  })

  it('kinsoku: an opening bracket never ends a line', () => {
    const layout = layoutText({
      body: body({ paragraphs: [{ runs: [{ text: 'あああ「いいい', fontSize: 24 }] }] }),
      boxWidthPx: 96,
      boxHeightPx: 3000,
      metrics: new HeuristicMetrics(),
      vp,
    })
    const lines = layout.lines.map((l) => l.runs.map((r) => r.text).join(''))
    expect(lines[0]).toBe('あああ')
    expect(lines[1]!.startsWith('「い')).toBe(true)
  })

  it('autofit=shrink reduces font scale when content overflows height', () => {
    const many = Array.from({ length: 30 }, (_, i) => ({
      runs: [{ text: `Line ${i} of text`, fontSize: 40 }],
    }))
    const layout = layoutText({
      body: body({ autofit: 'shrink', paragraphs: many }),
      boxWidthPx: 400,
      boxHeightPx: 200,
      metrics: new HeuristicMetrics(),
      vp,
    })
    expect(layout.fontScale).toBeLessThan(1)
    // Step floor 25%: with very heavy content, stop at 25% and allow overflow
    expect(layout.fontScale).toBeGreaterThanOrEqual(0.25)
  })

  it('"\\n" sentinel (from <a:br/>) forces a line break even without wrap', () => {
    const layout = layoutText({
      body: body({ wrap: false, paragraphs: [{ runs: [{ text: 'A\nB', fontSize: 18 }] }] }),
      boxWidthPx: 400,
      boxHeightPx: 200,
      metrics: new HeuristicMetrics(),
      vp,
    })
    expect(layout.lines.length).toBe(2)
    expect(layout.lines[0]!.runs.map((r) => r.text).join('')).toBe('A')
    expect(layout.lines[1]!.runs.map((r) => r.text).join('')).toBe('B')
  })

  it('lineHeight % and spaceBefore/spaceAfter affect vertical metrics', () => {
    const single = layoutText({
      body: body({ paragraphs: [{ runs: [{ text: 'x', fontSize: 20 }] }] }),
      boxWidthPx: 400,
      boxHeightPx: 400,
      metrics: new HeuristicMetrics(),
      vp,
    })
    const spaced = layoutText({
      body: body({
        paragraphs: [
          { runs: [{ text: 'x', fontSize: 20 }], lineHeight: 200, spaceBefore: 12, spaceAfter: 6 },
        ],
      }),
      boxWidthPx: 400,
      boxHeightPx: 400,
      metrics: new HeuristicMetrics(),
      vp,
    })
    expect(spaced.contentHeight).toBeGreaterThan(single.contentHeight)
    // PowerPoint ignores space-before on the frame's first paragraph (0047 measured);
    // the taller line comes from lineHeight/spaceAfter only
    expect(spaced.lines[0]!.top).toBe(0)
  })

  it('space-before applies from the second paragraph on, never the first', () => {
    const layout = layoutText({
      body: body({
        paragraphs: [
          { runs: [{ text: 'a', fontSize: 20 }], spaceBefore: 12 },
          { runs: [{ text: 'b', fontSize: 20 }], spaceBefore: 12 },
        ],
      }),
      boxWidthPx: 400,
      boxHeightPx: 400,
      metrics: new HeuristicMetrics(),
      vp,
    })
    expect(layout.lines[0]!.top).toBe(0)
    expect(layout.lines[1]!.top - (layout.lines[0]!.top + layout.lines[0]!.height)).toBeCloseTo(
      (12 * 96) / 72,
      1,
    )
  })

  it('align=center centers each line horizontally (baked into run x)', () => {
    const m = new HeuristicMetrics()
    const layout = layoutText({
      body: body({ paragraphs: [{ runs: [{ text: '中文', fontSize: 18 }], align: 'center' }] }),
      boxWidthPx: 400,
      boxHeightPx: 100,
      metrics: m,
      vp,
    })
    const runs = layout.lines[0]!.runs
    const lineW = runs.reduce((a, r) => a + r.widthPx, 0)
    expect(runs[0]!.x).toBeCloseTo((400 - lineW) / 2, 1)
  })

  it('align=right flushes line to the right edge', () => {
    const layout = layoutText({
      body: body({ paragraphs: [{ runs: [{ text: 'ab', fontSize: 18 }], align: 'right' }] }),
      boxWidthPx: 400,
      boxHeightPx: 100,
      metrics: new HeuristicMetrics(),
      vp,
    })
    const runs = layout.lines[0]!.runs
    const lineW = runs.reduce((a, r) => a + r.widthPx, 0)
    expect(runs[runs.length - 1]!.x + runs[runs.length - 1]!.widthPx).toBeCloseTo(400, 1)
    expect(runs[0]!.x).toBeCloseTo(400 - lineW, 1)
  })

  it('anchor=middle bakes vertical offset into line top/baseline', () => {
    const layout = layoutText({
      body: body({ anchor: 'middle', paragraphs: [{ runs: [{ text: 'Hi', fontSize: 18 }] }] }),
      boxWidthPx: 400,
      boxHeightPx: 200,
      metrics: new HeuristicMetrics(),
      vp,
    })
    const line = layout.lines[0]!
    const expectedTop = (200 - layout.contentHeight) / 2
    expect(line.top).toBeCloseTo(expectedTop, 1)
    expect(line.runs[0]!.baselineY).toBeGreaterThan(expectedTop)
  })

  it('anchor=top keeps lines at top (no offset)', () => {
    const layout = layoutText({
      body: body({ anchor: 'top', paragraphs: [{ runs: [{ text: 'Hi', fontSize: 18 }] }] }),
      boxWidthPx: 400,
      boxHeightPx: 200,
      metrics: new HeuristicMetrics(),
      vp,
    })
    expect(layout.lines[0]!.top).toBe(0)
  })

  it('preserves vertical anchor and multi-run color/bold', () => {
    const layout = layoutText({
      body: body({
        anchor: 'middle',
        paragraphs: [
          {
            runs: [
              { text: 'A', fontSize: 18, bold: true, color: '#FF0000' },
              { text: 'B', fontSize: 18, italic: true, color: '#00FF00' },
            ],
          },
        ],
      }),
      boxWidthPx: 400,
      boxHeightPx: 200,
      metrics: new HeuristicMetrics(),
      vp,
    })
    expect(layout.anchor).toBe('middle')
    const runs = layout.lines[0]!.runs
    expect(runs.find((r) => r.text === 'A')?.bold).toBe(true)
    expect(runs.find((r) => r.text === 'A')?.color).toBe('#FF0000')
    expect(runs.find((r) => r.text === 'B')?.italic).toBe(true)
  })
})

describe('bullets and indentation (bullet / marL / indent)', () => {
  const m = new HeuristicMetrics()
  const run = (text: string) => ({ text, fontSize: 10, color: '#111111' })

  it('buChar: first line prefixed with bullet glyph, body starts at marL (hanging indent)', () => {
    const lay = layoutText({
      body: body({
        paragraphs: [
          {
            runs: [run('item one')],
            bullet: { type: 'char', char: '•' },
            marL: 127000, // ≈13.3px
            indent: -127000,
          },
        ],
      }),
      boxWidthPx: 400,
      boxHeightPx: 100,
      metrics: m,
      vp,
    })
    const runs = lay.lines[0]!.runs
    expect(runs[0]!.text).toBe('•')
    expect(runs[0]!.x).toBeCloseTo(0, 5) // marL+indent = 0
    const textStart = runs[1]!
    expect(textStart.x).toBeCloseTo(127000 / 9525, 1) // body line start = marL
  })

  it('buAutoNum: consecutive paragraphs numbered 1. 2., reset after a non-numbered paragraph', () => {
    const paragraphs = [
      { runs: [run('a')], bullet: { type: 'number' as const }, marL: 127000, indent: -127000 },
      { runs: [run('b')], bullet: { type: 'number' as const }, marL: 127000, indent: -127000 },
      { runs: [run('plain')] },
      { runs: [run('c')], bullet: { type: 'number' as const }, marL: 127000, indent: -127000 },
    ]
    const lay = layoutText({
      body: body({ paragraphs }),
      boxWidthPx: 400,
      boxHeightPx: 200,
      metrics: m,
      vp,
    })
    expect(lay.lines[0]!.runs[0]!.text).toBe('1.')
    expect(lay.lines[1]!.runs[0]!.text).toBe('2.')
    expect(lay.lines[2]!.runs[0]!.text).toBe('plain')
    expect(lay.lines[3]!.runs[0]!.text).toBe('1.') // reset
  })

  it('buSzPct/buClr: bullet glyph drawn scaled and in its own color', () => {
    const lay = layoutText({
      body: body({
        paragraphs: [
          {
            runs: [run('item')],
            bullet: { type: 'char', char: '•', sizePct: 75, color: '#C00000' },
            marL: 127000,
            indent: -127000,
          },
        ],
      }),
      boxWidthPx: 400,
      boxHeightPx: 100,
      metrics: m,
      vp,
    })
    const runs = lay.lines[0]!.runs
    expect(runs[0]!.text).toBe('•')
    expect(runs[0]!.color).toBe('#C00000')
    expect(runs[0]!.fontSizePx).toBeCloseTo(runs[1]!.fontSizePx * 0.75, 5)
  })

  it('buNone/no bullet: no symbol added; positive indent applies to the first line only', () => {
    const lay = layoutText({
      body: body({
        paragraphs: [
          {
            runs: [run('中文正文换行测试中文正文换行测试中文正文换行测试')],
            bullet: { type: 'none' },
            marL: 95250, // 10px
            indent: 47625, // first line +5px
          },
        ],
      }),
      boxWidthPx: 120,
      boxHeightPx: 200,
      metrics: m,
      vp,
    })
    expect(lay.lines.length).toBeGreaterThan(1)
    expect(lay.lines[0]!.runs[0]!.text).not.toBe('•')
    expect(lay.lines[0]!.runs[0]!.x).toBeCloseTo(15, 0) // marL + indent
    expect(lay.lines[1]!.runs[0]!.x).toBeCloseTo(10, 0) // marL
  })
})

describe('text outline (rPr a:ln) passthrough to GlyphRun', () => {
  const m = new HeuristicMetrics()

  it('outline color + EMU width converted to px; runs without outline omit the field', () => {
    const t = layoutText({
      body: body({
        paragraphs: [
          {
            runs: [
              {
                text: 'Art',
                fontFamily: 'Arial',
                fontSize: 40,
                outline: { color: '#FF0000', widthEmu: 12700 },
              },
              { text: 'plain', fontFamily: 'Arial', fontSize: 40 },
            ],
          },
        ],
      }),
      boxWidthPx: 10000,
      boxHeightPx: 1000,
      metrics: m,
      vp,
    })
    const runs = t.lines[0]!.runs
    expect(runs[0]!.outline?.color).toBe('#FF0000')
    // 12700 EMU = 1pt, scale=1 -> 96/72 px
    expect(runs[0]!.outline?.widthPx).toBeCloseTo(96 / 72, 3)
    expect(runs[1]!.outline).toBeUndefined()
  })
})

describe('2.3 letter-spacing (rPr spc) and substitute family for missing fonts', () => {
  const m = new HeuristicMetrics()

  it('letterSpacing counts into run width and the next run x (appended after each character)', () => {
    // Two runs: "abcd" with 2pt letter spacing + a following "x"; the latter's x = the former's width
    const t = layoutText({
      body: body({
        paragraphs: [
          {
            runs: [
              { text: 'abcd', fontFamily: 'Arial', fontSize: 75, letterSpacing: 2 },
              { text: 'x', fontFamily: 'Arial', fontSize: 75 },
            ],
          },
        ],
      }),
      boxWidthPx: 10000,
      boxHeightPx: 1000,
      metrics: m,
      vp,
    })
    const runs = t.lines[0]!.runs
    expect(runs).toHaveLength(2)
    const plain = m.measure('abcd', {
      fontFamily: 'Arial',
      fontSizePx: 100,
      bold: false,
      italic: false,
    })
    const lsPx = (2 * 100) / 75 // 2pt → px (at scale=1, 1pt=4/3px)
    expect(runs[0]!.letterSpacingPx).toBeCloseTo(lsPx, 3)
    expect(runs[0]!.widthPx).toBeCloseTo(plain + 4 * lsPx, 3)
    expect(runs[1]!.x).toBeCloseTo(runs[0]!.x + runs[0]!.widthPx, 3)
    expect(runs[1]!.letterSpacingPx).toBeUndefined()
  })

  it('displayFamily substitute name written into GlyphRun.fontFamily (measurement and drawing share the source)', () => {
    const withSub = {
      metrics: (s: any) => m.metrics(s),
      measure: (t: string, s: any) => m.measure(t, s),
      displayFamily: (s: any) => (s.fontFamily === 'Playfair Display' ? 'Georgia' : s.fontFamily),
    }
    const t = layoutText({
      body: body({
        paragraphs: [{ runs: [{ text: 'Hi', fontFamily: 'Playfair Display', fontSize: 20 }] }],
      }),
      boxWidthPx: 10000,
      boxHeightPx: 1000,
      metrics: withSub,
      vp,
    })
    expect(t.lines[0]!.runs[0]!.fontFamily).toBe('Georgia')
  })
})

describe('spcPct paragraph spacing + justify alignment', () => {
  // Heuristic: 18pt → 24px, line height 1.2em = 28.8px
  it('spaceBeforePct=100 adds one single line height before the paragraph', () => {
    const layout = layoutText({
      body: body({
        paragraphs: [
          { runs: [{ text: 'a', fontSize: 18 }] },
          { runs: [{ text: 'b', fontSize: 18 }], spaceBeforePct: 100 },
        ],
      }),
      boxWidthPx: 400,
      boxHeightPx: 400,
      metrics: new HeuristicMetrics(),
      vp,
    })
    expect(layout.lines[1]!.top).toBeCloseTo(28.8 + 28.8, 1)
  })

  it('spaceAfterPct=50 appends half a single line height', () => {
    const layout = layoutText({
      body: body({
        paragraphs: [
          { runs: [{ text: 'a', fontSize: 18 }], spaceAfterPct: 50 },
          { runs: [{ text: 'b', fontSize: 18 }] },
        ],
      }),
      boxWidthPx: 400,
      boxHeightPx: 400,
      metrics: new HeuristicMetrics(),
      vp,
    })
    expect(layout.lines[1]!.top).toBeCloseTo(28.8 + 14.4, 1)
  })

  it('justify: wrapped lines fill the available width, the last line of a paragraph stays left-aligned', () => {
    const long = 'word '.repeat(20).trim()
    const layout = layoutText({
      body: body({ paragraphs: [{ runs: [{ text: long, fontSize: 18 }], align: 'justify' }] }),
      boxWidthPx: 200,
      boxHeightPx: 2000,
      metrics: new HeuristicMetrics(),
      vp,
    })
    expect(layout.lines.length).toBeGreaterThan(1)
    for (let i = 0; i < layout.lines.length; i++) {
      const ln = layout.lines[i]!
      const last = ln.runs[ln.runs.length - 1]!
      const per = last.justifyExtraPx ?? 0
      // Last glyph's right edge = x + width - the extra trailing letter-space
      const rightEdge = last.x + last.widthPx - per * (per ? 1 : 0)
      if (i < layout.lines.length - 1) {
        expect(per).toBeGreaterThan(0)
        expect(rightEdge).toBeCloseTo(200, 0)
      } else {
        expect(per).toBe(0)
        expect(rightEdge).toBeLessThan(200)
      }
    }
  })

  it('justify does not affect lines ended by a soft break', () => {
    const layout = layoutText({
      body: body({
        paragraphs: [
          {
            runs: [
              { text: 'short', fontSize: 18 },
              { text: '\n', fontSize: 18 },
              { text: 'more text here', fontSize: 18 },
            ],
            align: 'justify',
          },
        ],
      }),
      boxWidthPx: 400,
      boxHeightPx: 400,
      metrics: new HeuristicMetrics(),
      vp,
    })
    // Soft-wrapped lines are not justified
    expect(layout.lines[0]!.runs.every((r) => !r.justifyExtraPx)).toBe(true)
  })
})

describe('shrink discrete steps + superscript/subscript baseline', () => {
  it('shrink scale lands on a PowerPoint step (7.5% increments), lnSpcReduction follows', () => {
    const many = Array.from({ length: 8 }, (_, i) => ({
      runs: [{ text: `Row ${i}`, fontSize: 20 }],
    }))
    // 8 lines × 20pt (26.7px×1.2=32px) = 256px, box height 210 → needs ≤82% → step 0.775/0.1
    const layout = layoutText({
      body: body({ autofit: 'shrink', paragraphs: many }),
      boxWidthPx: 400,
      boxHeightPx: 210,
      metrics: new HeuristicMetrics(),
      vp,
    })
    const STEPS = [0.925, 0.85, 0.775, 0.7, 0.625, 0.55, 0.475, 0.4, 0.325, 0.25]
    expect(STEPS.some((s) => Math.abs(s - layout.fontScale) < 1e-9)).toBe(true)
    expect(layout.contentHeight).toBeLessThanOrEqual(210)
    // Step pairing: below 85% comes with line-spacing reduction
    if (layout.fontScale <= 0.85) expect(layout.lnSpcReduction).toBeGreaterThan(0)
  })

  it('existing scale caps the ceiling: only steps smaller than it are used', () => {
    const many = Array.from({ length: 20 }, () => ({
      runs: [{ text: '很长的一行文字内容', fontSize: 20 }],
    }))
    const layout = layoutText({
      body: body({ autofit: 'shrink', fontScale: 0.7, lnSpcReduction: 0.2, paragraphs: many }),
      boxWidthPx: 400,
      boxHeightPx: 120,
      metrics: new HeuristicMetrics(),
      vp,
    })
    expect(layout.fontScale).toBeLessThan(0.7)
    // Existing 20% line-spacing reduction: stepping down further never regresses to a smaller reduction
    expect(layout.lnSpcReduction).toBe(0.2)
  })

  it('superscript baseline=30 → baseline rises 30% of font size, subscript goes the other way', () => {
    const layout = layoutText({
      body: body({
        paragraphs: [
          {
            runs: [
              { text: 'x', fontSize: 18 },
              { text: '2', fontSize: 18, baseline: 30 },
              { text: 'n', fontSize: 18, baseline: -25 },
            ],
          },
        ],
      }),
      boxWidthPx: 400,
      boxHeightPx: 200,
      metrics: new HeuristicMetrics(),
      vp,
    })
    const [base, sup, sub] = layout.lines[0]!.runs
    // 18pt = 24px; raised by 24×0.3 = 7.2px
    expect(base!.baselineY - sup!.baselineY).toBeCloseTo(7.2, 1)
    expect(sub!.baselineY - base!.baselineY).toBeCloseTo(6, 1)
  })
})

describe('vertical text (bodyPr vert) column layout', () => {
  const m = new HeuristicMetrics()
  const layoutV = (b: TextBody, w = 200, h = 300) =>
    layoutText({ body: b, boxWidthPx: w, boxHeightPx: h, metrics: m, vp })

  it('eaVert: two CJK paragraphs → two columns, right-to-left, chars flow downward within a column, all inside the box', () => {
    const layout = layoutV(
      body({
        vert: 'eaVert',
        paragraphs: [
          { runs: [{ text: '縦書き', fontSize: 18 }] },
          { runs: [{ text: '二列目', fontSize: 18 }] },
        ],
      }),
    )
    expect(layout.vert).toBe('eaVert')
    expect(layout.lines).toHaveLength(2)
    const [c1, c2] = layout.lines
    // One run per char per column; the second column (second paragraph) has a smaller x than the first
    expect(c1!.runs.map((r) => r.text)).toEqual(['縦', '書', 'き'])
    expect(c2!.runs[0]!.x).toBeLessThan(c1!.runs[0]!.x)
    // anchor=top → the first column starts flush with the right edge
    const right = Math.max(...c1!.runs.map((r) => r.x + r.widthPx))
    expect(right).toBeLessThanOrEqual(200)
    expect(right).toBeGreaterThan(200 - 30)
    // baselineY strictly increases within a column
    for (const col of layout.lines) {
      for (let i = 1; i < col.runs.length; i++) {
        expect(col.runs[i]!.baselineY).toBeGreaterThan(col.runs[i - 1]!.baselineY)
      }
      for (const r of col.runs) {
        expect(r.x).toBeGreaterThanOrEqual(0)
        expect(r.x + r.widthPx).toBeLessThanOrEqual(200)
        expect(r.baselineY).toBeLessThanOrEqual(300)
      }
    }
    // Paragraph identity preserved (the editor splits paragraphs by paraStart)
    expect(c1!.paraStart).toBe(true)
    expect(c2!.paraStart).toBe(true)
  })

  it('eaVert: overflowing text wraps to new columns, continuation columns have paraStart=false and stay right-to-left', () => {
    const layout = layoutV(
      body({
        vert: 'eaVert',
        paragraphs: [{ runs: [{ text: '縦書きテスト'.repeat(5), fontSize: 18 }] }],
      }),
      300,
      100,
    )
    expect(layout.lines.length).toBeGreaterThan(1)
    expect(layout.lines[0]!.paraStart).toBe(true)
    expect(layout.lines.slice(1).every((l) => l.paraStart === false)).toBe(true)
    for (const col of layout.lines) expect(col.height).toBeLessThanOrEqual(100)
    const xs = layout.lines.map((l) => l.runs[0]!.x)
    for (let i = 1; i < xs.length; i++) expect(xs[i]!).toBeLessThan(xs[i - 1]!)
    // Text is lossless (reassembled column by column)
    expect(layout.lines.flatMap((l) => l.runs.map((r) => r.text)).join('')).toBe(
      '縦書きテスト'.repeat(5),
    )
  })

  it('vert/vert270/wordArtVert degrade to the same column layout as eaVert (vert flag preserved)', () => {
    const mk = (vert: TextBody['vert']) =>
      layoutV(body({ vert, paragraphs: [{ runs: [{ text: '縦書', fontSize: 18 }] }] }))
    const coords = (l: ReturnType<typeof mk>) =>
      l.lines.map((ln) => ln.runs.map((r) => [r.text, r.x, r.baselineY]))
    const ea = mk('eaVert')
    for (const v of ['vert', 'vert270', 'wordArtVert'] as const) {
      const l = mk(v)
      expect(l.vert).toBe(v)
      expect(coords(l)).toEqual(coords(ea))
    }
  })

  it('anchor acts along the text flow: bottom hugs the left, middle centers', () => {
    const paras = [{ runs: [{ text: '縦', fontSize: 18 }] }]
    const left = (a: TextBody['anchor']) =>
      Math.min(
        ...layoutV(body({ vert: 'eaVert', anchor: a, paragraphs: paras })).lines[0]!.runs.map(
          (r) => r.x,
        ),
      )
    const top = left('top')
    const mid = left('middle')
    const bot = left('bottom')
    expect(bot).toBeLessThan(mid)
    expect(mid).toBeLessThan(top)
    expect(bot).toBeGreaterThanOrEqual(0)
    expect(bot).toBeLessThan(6) // column width 28.8, glyph width 24 → centered offset 2.4
  })
})

describe('vertical layout latin word rotation', () => {
  it('in eaVert latin words rotate 90° as whole words, CJK stays upright char by char', async () => {
    const { layoutText } = await import('../src/text-layout')
    const { HeuristicMetrics } = await import('../src/metrics')
    const { makeViewport } = await import('../src/coords')
    const vp = makeViewport({ cx: 9525 * 1000, cy: 9525 * 1000 }, 1000) // scale 1
    const out = layoutText({
      body: {
        vert: 'eaVert',
        paragraphs: [{ runs: [{ text: '縦書きABC組版' }] }],
      } as never,
      boxWidthPx: 300,
      boxHeightPx: 200,
      metrics: new HeuristicMetrics(),
      vp,
    })
    const runs = out.lines.flatMap((l) => l.runs).filter((r) => !r.isBullet)
    const rotated = runs.filter((r) => r.rotate90)
    const upright = runs.filter((r) => !r.rotate90)
    expect(rotated.length).toBe(1)
    expect(rotated[0]!.text).toBe('ABC')
    // CJK stays upright char by char
    expect(upright.map((r) => r.text).join('')).toBe('縦書き組版')
    // Rotated word anchors stay within the column (not outside the content area)
    expect(rotated[0]!.x).toBeGreaterThan(0)
    expect(rotated[0]!.x).toBeLessThanOrEqual(300)
  })
})

describe('text highlight propagation', () => {
  it('run highlight lands on every glyph run split from it, others stay clean', () => {
    const layout = layoutText({
      body: body({
        paragraphs: [
          {
            runs: [
              { text: 'marked text', fontSize: 18, highlight: '#FF0000' },
              { text: ' plain', fontSize: 18 },
            ],
          },
        ],
      }),
      boxWidthPx: 400,
      boxHeightPx: 200,
      metrics: new HeuristicMetrics(),
      vp,
    })
    const runs = layout.lines.flatMap((l) => l.runs)
    const marked = runs.filter((r) => r.highlight)
    expect(marked.length).toBeGreaterThan(0)
    expect(marked.every((r) => r.highlight === '#FF0000')).toBe(true)
    expect(marked.map((r) => r.text).join('')).toBe('marked text')
    expect(
      runs
        .filter((r) => !r.highlight)
        .map((r) => r.text)
        .join(''),
    ).toBe(' plain')
  })
})

describe('table cell edge spacing (trimEdgeSpacing)', () => {
  const paras = [
    { runs: [{ text: 'first', fontSize: 12 }], spaceBefore: 10, spaceAfter: 10 },
    { runs: [{ text: 'last', fontSize: 12 }], spaceBefore: 10, spaceAfter: 10 },
  ]
  it('drops the first space-before and last space-after only', () => {
    const plain = layoutText({
      body: body({ paragraphs: paras }),
      boxWidthPx: 400,
      boxHeightPx: 200,
      metrics: new HeuristicMetrics(),
      vp,
    })
    const trimmed = layoutText({
      body: body({ paragraphs: paras }),
      boxWidthPx: 400,
      boxHeightPx: 200,
      metrics: new HeuristicMetrics(),
      vp,
      trimEdgeSpacing: true,
    })
    // First-paragraph space-before is dropped in every body (PowerPoint semantics);
    // trimEdgeSpacing additionally drops the last space-after (pt → px at scale 1: ×96/72)
    expect(plain.contentHeight - trimmed.contentHeight).toBeCloseTo((10 * 96) / 72, 1)
    // inner spacing (after-first + before-last) is kept
    expect(trimmed.lines[1]!.top - (trimmed.lines[0]!.top + trimmed.lines[0]!.height)).toBeCloseTo(
      (20 * 96) / 72,
      1,
    )
    // first line starts at the very top
    expect(trimmed.lines[0]!.top).toBe(0)
  })
})

describe('autofit ignores trailing blank paragraphs', () => {
  it('keeps the stored scale when only trailing blanks overflow', () => {
    // 3 content lines fit the box; 4 trailing blanks push contentHeight past it
    const paragraphs = [
      { runs: [{ text: 'one', fontSize: 18 }] },
      { runs: [{ text: 'two', fontSize: 18 }] },
      { runs: [{ text: 'three', fontSize: 18 }] },
      { runs: [] },
      { runs: [] },
      { runs: [] },
      { runs: [] },
    ]
    const layout = layoutText({
      body: body({ paragraphs, autofit: 'shrink' }),
      boxWidthPx: 400,
      boxHeightPx: 100,
      metrics: new HeuristicMetrics(),
      vp,
    })
    expect(layout.fontScale).toBe(1)
  })
  it('still shrinks when real content overflows', () => {
    const paragraphs = Array.from({ length: 12 }, () => ({
      runs: [{ text: 'line of text', fontSize: 18 }],
    }))
    const layout = layoutText({
      body: body({ paragraphs, autofit: 'shrink' }),
      boxWidthPx: 400,
      boxHeightPx: 100,
      metrics: new HeuristicMetrics(),
      vp,
    })
    expect(layout.fontScale).toBeLessThan(1)
  })
})

describe('buAutoNum startAt', () => {
  const m = new HeuristicMetrics()
  const style = { fontFamily: 'Arial', fontSizePx: 20, bold: false, italic: false }
  const para = (text: string, startAt?: number) => ({
    runs: [{ text, ...style, fontSize: 20 }],
    bullet: { type: 'number' as const, ...(startAt != null ? { startAt } : {}) },
    marL: 457200,
    indent: -457200,
  })

  it('starts the sequence at startAt and continues from there', () => {
    const layout = layoutText({
      body: body({ paragraphs: [para('first', 3), para('second')] as any }),
      boxWidthPx: 800,
      boxHeightPx: 200,
      metrics: m,
      vp,
    })
    const bullets = layout.lines
      .map((l) => l.runs.find((r: any) => r.isBullet))
      .filter(Boolean)
      .map((r: any) => r.text)
    expect(bullets).toEqual(['3.', '4.'])
  })

  it('defaults to 1 without startAt', () => {
    const layout = layoutText({
      body: body({ paragraphs: [para('only')] as any }),
      boxWidthPx: 800,
      boxHeightPx: 200,
      metrics: m,
      vp,
    })
    const b = layout.lines[0]!.runs.find((r: any) => r.isBullet) as any
    expect(b.text).toBe('1.')
  })
})

describe('lnSpc > 100%: excess spacing sits above the glyphs', () => {
  const m = new HeuristicMetrics()
  const style = { fontFamily: 'Arial', fontSizePx: 20, bold: false, italic: false }
  const para = (text: string, lineHeight?: number) => ({
    runs: [{ text, ...style, fontSize: 20 }],
    ...(lineHeight != null ? { lineHeight } : {}),
  })

  it('150% spacing pushes the first baseline down by the half-line excess', () => {
    const single = layoutText({
      body: body({ paragraphs: [para('watermark')] as any }),
      boxWidthPx: 800,
      boxHeightPx: 400,
      metrics: m,
      vp,
    })
    const spaced = layoutText({
      body: body({ paragraphs: [para('watermark', 150)] as any }),
      boxWidthPx: 800,
      boxHeightPx: 400,
      metrics: m,
      vp,
    })
    // PowerPoint model (48pt probe): explicit pct spacing places the baseline at
    // 0.7333 x the line box, ignoring font metrics; single spacing bottom-anchors
    // (box - descent). Heuristic metrics: descent 0.2em.
    const b1 = single.lines[0]!.runs[0]!.baselineY
    const b2 = spaced.lines[0]!.runs[0]!.baselineY
    expect(single.lines[0]!.height).toBeCloseTo(spaced.lines[0]!.height / 1.5, 1)
    expect(b1).toBeCloseTo(single.lines[0]!.height - 0.2 * (single.lines[0]!.height / 1.2), 1)
    expect(b2).toBeCloseTo(spaced.lines[0]!.height * (0.88 / 1.2), 1)
  })

  it('sub-100% spacing shrinks the box and lifts the baseline with it (ink may poke above)', () => {
    const tight = layoutText({
      body: body({ paragraphs: [para('text', 70)] as any }),
      boxWidthPx: 800,
      boxHeightPx: 400,
      metrics: m,
      vp,
    })
    const single = layoutText({
      body: body({ paragraphs: [para('text')] as any }),
      boxWidthPx: 800,
      boxHeightPx: 400,
      metrics: m,
      vp,
    })
    expect(tight.lines[0]!.height).toBeCloseTo(single.lines[0]!.height * 0.7, 1)
    // generic law: baseline = 0.7333 x box — above the single-spacing baseline
    expect(tight.lines[0]!.runs[0]!.baselineY).toBeCloseTo(tight.lines[0]!.height * (0.88 / 1.2), 1)
    expect(tight.lines[0]!.runs[0]!.baselineY).toBeLessThan(single.lines[0]!.runs[0]!.baselineY)
  })
})

describe('bodyPr numCol columns', () => {
  it('flows lines into the next column when the box height is exceeded', () => {
    const paras = Array.from({ length: 8 }, (_, i) => ({
      runs: [{ text: `p${i}`, fontSize: 20 }],
    }))
    const layout = layoutText({
      body: { ...body({ paragraphs: paras }), numCol: 2, spcCol: 0 },
      boxWidthPx: 400,
      boxHeightPx: 4 * 20 * (96 / 72) * 1.2 + 2, // fits 4 lines per column
      metrics: new HeuristicMetrics(),
      vp,
    })
    const xs = layout.lines.map((l) => l.runs[0]!.x)
    // First 4 lines in column 1 (x < 200), rest shifted a full column stride right
    expect(xs.slice(0, 4).every((x) => x < 200)).toBe(true)
    expect(xs.slice(4).every((x) => x >= 200)).toBe(true)
    // Column 2 restarts at the top, baselines rebased with the line tops
    expect(layout.lines[4]!.top).toBe(0)
    expect(layout.lines[4]!.runs[0]!.baselineY).toBeLessThan(layout.lines[3]!.runs[0]!.baselineY)
    expect(layout.contentHeight).toBeCloseTo(layout.lines[3]!.top + layout.lines[3]!.height, 1)
  })

  it('single column behavior unchanged when numCol absent', () => {
    const layout = layoutText({
      body: body({ paragraphs: [{ runs: [{ text: 'x', fontSize: 20 }] }] }),
      boxWidthPx: 400,
      boxHeightPx: 100,
      metrics: new HeuristicMetrics(),
      vp,
    })
    expect(layout.lines.length).toBe(1)
  })
})

describe('WordArt effect passthrough', () => {
  const m = new HeuristicMetrics()
  const run = {
    text: 'Word',
    fontSize: 20,
    color: '#111111',
    gradient: {
      stops: [
        { pos: 0, color: '#FF0000' },
        { pos: 1, color: '#0000FF' },
      ],
      angle: 5400000,
    },
    glow: { color: '#ED7D31', radius: 53100 },
    reflection: true,
  }

  it('gradient/glow/reflection reach the glyph runs (angle in degrees, glow in px)', () => {
    const layout = layoutText({
      body: body({ paragraphs: [{ runs: [run] }] }),
      boxWidthPx: 400,
      boxHeightPx: 100,
      metrics: m,
      vp,
    })
    const g = layout.lines[0]!.runs[0]!
    expect(g.gradient?.angleDeg).toBe(90)
    expect(g.gradient?.stops).toHaveLength(2)
    expect(g.glow?.color).toBe('#ED7D31')
    expect(g.glow?.blurPx).toBeGreaterThan(0)
    expect(g.reflection).toBe(true)
  })

  it('bodyPr extrusion projects to a screen offset (lat 30° → downward)', () => {
    const layout = layoutText({
      body: body({
        paragraphs: [{ runs: [{ text: 'W', fontSize: 20, color: '#111111' }] }],
        extrusion3d: { color: '#C0504D', depthEmu: 127000, latDeg: 30, lonDeg: 0 },
      }),
      boxWidthPx: 400,
      boxHeightPx: 100,
      metrics: m,
      vp,
    })
    expect(layout.extrusion?.color).toBe('#C0504D')
    expect(layout.extrusion?.dx).toBeCloseTo(0, 5)
    // PowerPoint renders the depth below the glyphs for a positive camera latitude
    expect(layout.extrusion!.dy).toBeGreaterThan(0)
  })
})

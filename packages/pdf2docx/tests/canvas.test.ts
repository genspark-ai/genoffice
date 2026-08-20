/** P19 page classifier: flow | canvas | scan (conservative — canvas needs high confidence). */
import { describe, expect, it } from 'vitest'
import {
  classifyPage,
  classifyPages,
  computeCanvasPrior,
  isSlideProducer,
  isSlideSizedPage,
} from '../src/analyze/canvas'
import type { ImageBlock, IrPage, Line, Span, TextBlock } from '../src/ir'

function span(text: string, over: Partial<Span> = {}): Span {
  return {
    text,
    box: { x0: 72, y0: 690, x1: 300, y1: 700 },
    fontSize: 12,
    fontFamily: 'Helvetica',
    bold: false,
    italic: false,
    color: '000000',
    dir: 'ltr',
    script: 'latin',
    ...over,
  }
}

function lineAt(text: string, x0: number, top: number, fontSize = 12, width = 200): Line {
  const s = span(text, { fontSize, box: { x0, y0: top - fontSize, x1: x0 + width, y1: top } })
  return { spans: [s], box: s.box, baseline: top - fontSize * 0.8, endsWithHyphen: false }
}

function blockOf(lines: Line[], over: Partial<TextBlock> = {}): TextBlock {
  return {
    kind: 'text',
    lines,
    box: {
      x0: Math.min(...lines.map((l) => l.box.x0)),
      y0: Math.min(...lines.map((l) => l.box.y0)),
      x1: Math.max(...lines.map((l) => l.box.x1)),
      y1: Math.max(...lines.map((l) => l.box.y1)),
    },
    align: 'left',
    firstLineIndentPt: 0,
    dir: 'ltr',
    ...over,
  }
}

function page(blocks: IrPage['blocks'], over: Partial<IrPage> = {}): IrPage {
  return {
    index: 0,
    widthPt: 960,
    heightPt: 540,
    rotation: 0,
    blocks,
    degraded: false,
    scanned: false,
    hasStructTree: false,
    ...over,
  }
}

/** typical slide: big sparse display text islands over a colored wash */
function slidePage(index = 0): IrPage {
  return page(
    [
      blockOf([lineAt('Quarterly Review', 60, 500, 36, 500)]),
      blockOf([lineAt('Growth is strong', 60, 350, 20, 300)]),
      blockOf([lineAt('Costs are flat', 520, 350, 20, 300)]),
      blockOf([lineAt('appendix note', 60, 60, 14, 200)]),
    ],
    { index, bgColor: '1a2b3c' },
  )
}

/** dense report page that happens to use slide-sized paper (AI-generated) */
function denseSlideSizedPage(index = 0): IrPage {
  const lines: Line[] = []
  for (let i = 0; i < 34; i++) {
    lines.push(
      lineAt(
        '这是一段密集的正文内容,讨论市场的现状与未来的发展趋势,涉及供给侧与需求侧的多重因素分析。',
        60,
        510 - i * 14,
        11,
        840,
      ),
    )
  }
  return page([blockOf(lines)], { index })
}

const SLIDE_META = { producer: 'Microsoft® PowerPoint® 2016' }

describe('doc-level priors', () => {
  it('recognizes slide producers (PowerPoint/Keynote/Slides/Impress/WPS 演示)', () => {
    expect(isSlideProducer({ producer: 'Microsoft® PowerPoint® 2010' })).toBe(true)
    expect(isSlideProducer({ creator: 'PowerPoint' })).toBe(true)
    expect(isSlideProducer({ creator: 'Keynote' })).toBe(true)
    expect(isSlideProducer({ creator: 'WPS 演示' })).toBe(true)
    expect(isSlideProducer({ producer: 'LibreOffice Impress' })).toBe(true)
    expect(isSlideProducer({ producer: 'LibreOffice 26.2', creator: 'Writer' })).toBe(false)
    expect(isSlideProducer({ producer: 'pypdf' })).toBe(false)
    expect(isSlideProducer({})).toBe(false)
  })

  it('recognizes slide page geometry (16:9 / 4:3 landscape), not A4/letter landscape', () => {
    expect(isSlideSizedPage(960, 540)).toBe(true)
    expect(isSlideSizedPage(720, 405)).toBe(true)
    expect(isSlideSizedPage(720, 540)).toBe(true)
    expect(isSlideSizedPage(1440, 810)).toBe(true)
    expect(isSlideSizedPage(842, 595)).toBe(false) // A4 landscape
    expect(isSlideSizedPage(792, 612)).toBe(false) // letter landscape
    expect(isSlideSizedPage(595, 842)).toBe(false) // portrait
  })

  it('computes a strong prior for a PowerPoint deck and none for a dense pypdf report', () => {
    const deck = computeCanvasPrior([slidePage(0), slidePage(1)], SLIDE_META)
    expect(deck.slideProducer).toBe(true)
    expect(deck.pointsNeeded).toBe(2)
    const report = computeCanvasPrior([denseSlideSizedPage(0), denseSlideSizedPage(1)], {
      producer: 'pypdf',
    })
    expect(report.slideProducer).toBe(false)
    expect(report.pointsNeeded).toBeGreaterThanOrEqual(4)
  })
})

describe('classifyPage', () => {
  it('classifies a sparse display slide as canvas under a strong prior', () => {
    const prior = computeCanvasPrior([slidePage()], SLIDE_META)
    expect(classifyPage(slidePage(), prior)).toBe('canvas')
  })

  it('keeps scan/degraded pages as scan', () => {
    const prior = computeCanvasPrior([slidePage()], SLIDE_META)
    expect(classifyPage(page([], { scanned: true }), prior)).toBe('scan')
    expect(classifyPage(page([], { degraded: true }), prior)).toBe('scan')
  })

  it('never marks portrait or non-slide-sized pages canvas, even with a strong prior', () => {
    const prior = computeCanvasPrior([slidePage()], SLIDE_META)
    const portrait = page([blockOf([lineAt('Title', 60, 700, 36, 400)])], {
      widthPt: 612,
      heightPt: 792,
      bgColor: '1a2b3c',
    })
    expect(classifyPage(portrait, prior)).toBe('flow')
    const a4Land = { ...slidePage(), widthPt: 842, heightPt: 595 }
    expect(classifyPage(a4Land, prior)).toBe('flow')
  })

  it('keeps dense text pages flow even on slide-sized paper with a strong prior', () => {
    const prior = computeCanvasPrior([denseSlideSizedPage()], SLIDE_META)
    expect(classifyPage(denseSlideSizedPage(), prior)).toBe('flow')
  })

  it('a dense slide-sized report doc (no slide producer) stays all-flow', () => {
    const pages = [denseSlideSizedPage(0), denseSlideSizedPage(1), denseSlideSizedPage(2)]
    classifyPages(pages, { producer: 'pypdf' })
    expect(pages.every((p) => p.canvas === undefined)).toBe(true)
  })

  it('a sparse-display deck without metadata still classifies canvas via geometry prior', () => {
    const pages = [slidePage(0), slidePage(1), slidePage(2)]
    classifyPages(pages, {})
    expect(pages.every((p) => p.canvas === true)).toBe(true)
  })

  it('background coverage counts behind-floats and panels', () => {
    const wallpaper: ImageBlock = {
      kind: 'image',
      box: { x0: 0, y0: 0, x1: 960, y1: 540 },
      data: new Uint8Array(8),
      mime: 'image/png',
      pixelWidth: 2,
      pixelHeight: 2,
      float: { wrap: 'behind', xOffsetPt: 0 },
    }
    const p = page([
      wallpaper,
      blockOf([lineAt('Big cover title', 100, 400, 40, 600)]),
      blockOf([lineAt('subtitle line', 100, 200, 18, 400)]),
      blockOf([lineAt('date footer', 100, 60, 12, 200)]),
    ])
    const prior = computeCanvasPrior([p], {})
    expect(classifyPage(p, prior)).toBe('canvas')
  })

  it('cross-page sentence continuity suppresses the geometry prior', () => {
    // two slide-sized pages whose joint reads like one flowing paragraph
    const mk = (index: number): IrPage => {
      const lines: Line[] = []
      for (let i = 0; i < 6; i++) {
        lines.push(
          lineAt(
            'a body line of running prose that fills the column width here',
            60,
            500 - i * 16,
            12,
            840,
          ),
        )
      }
      return page([blockOf(lines)], { index })
    }
    const pages = [mk(0), mk(1), mk(2), mk(3)]
    const prior = computeCanvasPrior(pages, {})
    expect(prior.continuityShare).toBeGreaterThan(0.5)
    expect(prior.pointsNeeded).toBeGreaterThanOrEqual(4)
  })
})

// ── newsletter admission (P20 C) ──

/** A4-landscape bulletin page: regions + table + a cross-section backtrack */
function newsletterPage(index = 0, backtrack = true): IrPage {
  const left = blockOf([
    lineAt('Officers and roles listed for the year, with plenty of text here', 40, 550, 11, 300),
  ])
  const leftB = blockOf([
    lineAt('President-elect, secretary, treasurer and the full committee row', 40, 520, 11, 300),
  ])
  const right = blockOf([
    lineAt('Rotary Club of Clare weekly bulletin masthead and headline text!!', 460, 545, 16, 320),
  ])
  // the second section's first block starts far ABOVE section 1's bottom
  const lower = blockOf([
    lineAt(
      'Club meetings每 Monday 6.00pm at the restaurant, apologies to the',
      460,
      backtrack ? 500 : 280,
      11,
      320,
    ),
  ])
  const lower2 = blockOf([
    lineAt(
      'invocation and toast text for the week goes here in the region.',
      460,
      backtrack ? 470 : 250,
      11,
      320,
    ),
  ])
  const table = {
    kind: 'table',
    box: { x0: 40, y0: 100, x1: 400, y1: 480 },
    rows: [],
    colWidthsPt: [180, 180],
  } as unknown as IrPage['blocks'][number]
  const blocks: IrPage['blocks'] = [left, leftB, table, right, lower, lower2]
  const p = page(blocks, { index, widthPt: 842, heightPt: 595 })
  // strong: a tall LEFT region section followed by the RIGHT region whose
  // first block starts near the page top — the cross-section chain runs
  // backwards by ~455pt. weak: the sections stack top-to-bottom instead.
  p.sections = backtrack
    ? [
        {
          box: { x0: 40, y0: 90, x1: 400, y1: 560 },
          columns: [
            { box: { x0: 40, y0: 90, x1: 220, y1: 560 }, blocks: [left, leftB] },
            { box: { x0: 230, y0: 90, x1: 400, y1: 560 }, blocks: [table] },
          ],
          gutterWidthsPt: [10],
          dir: 'ltr',
        },
        {
          box: { x0: 460, y0: 90, x1: 800, y1: 550 },
          columns: [
            { box: { x0: 460, y0: 90, x1: 640, y1: 550 }, blocks: [right, lower] },
            { box: { x0: 650, y0: 90, x1: 800, y1: 550 }, blocks: [lower2] },
          ],
          gutterWidthsPt: [10],
          dir: 'ltr',
        },
      ]
    : [
        {
          box: { x0: 40, y0: 505, x1: 800, y1: 560 },
          columns: [
            { box: { x0: 40, y0: 505, x1: 400, y1: 560 }, blocks: [left, leftB] },
            { box: { x0: 460, y0: 505, x1: 800, y1: 560 }, blocks: [right] },
          ],
          gutterWidthsPt: [60],
          dir: 'ltr',
        },
        {
          box: { x0: 40, y0: 90, x1: 800, y1: 500 },
          columns: [
            { box: { x0: 460, y0: 90, x1: 800, y1: 500 }, blocks: [lower, lower2] },
            { box: { x0: 40, y0: 90, x1: 400, y1: 500 }, blocks: [table] },
          ],
          gutterWidthsPt: [60],
          dir: 'ltr',
        },
      ]
  return p
}

describe('newsletter admission (P20 C)', () => {
  it('detects the strong page: regions + table + huge cross-section backtrack', async () => {
    const { isNewsletterStrongPage } = await import('../src/analyze/canvas')
    expect(isNewsletterStrongPage(newsletterPage(0, true))).toBe(true)
    expect(isNewsletterStrongPage(newsletterPage(0, false))).toBe(false)
  })

  it('a strong page unlocks the sibling region page; both go canvas', () => {
    const strong = newsletterPage(0, true)
    const weak = newsletterPage(1, false)
    classifyPages([strong, weak], {})
    expect(strong.canvas).toBe(true)
    expect(weak.canvas).toBe(true)
  })

  it('without a strong page nothing changes', () => {
    const weakA = newsletterPage(0, false)
    const weakB = newsletterPage(1, false)
    classifyPages([weakA, weakB], {})
    expect(weakA.canvas).toBeUndefined()
    expect(weakB.canvas).toBeUndefined()
  })

  it('a page without tables never reads as newsletter-strong', async () => {
    const { isNewsletterStrongPage } = await import('../src/analyze/canvas')
    const p = newsletterPage(0, true)
    p.blocks = p.blocks.filter((b) => b.kind !== 'table')
    for (const s of p.sections!)
      for (const c of s.columns) {
        c.blocks = c.blocks.filter((b) => b.kind !== 'table')
      }
    expect(isNewsletterStrongPage(p)).toBe(false)
  })

  it('one strong page lost in a long ordinary document stays flow', () => {
    const pages: IrPage[] = [newsletterPage(0, true)]
    for (let i = 1; i < 8; i++) {
      const lines: Line[] = []
      for (let k = 0; k < 20; k++)
        lines.push(
          lineAt('ordinary dense report body text on portrait paper', 60, 700 - k * 14, 11, 460),
        )
      pages.push(page([blockOf(lines)], { index: i, widthPt: 595, heightPt: 842 }))
    }
    classifyPages(pages, {})
    expect(pages.every((p) => p.canvas !== true)).toBe(true)
  })
})

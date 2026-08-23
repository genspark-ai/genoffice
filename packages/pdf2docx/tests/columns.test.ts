/** XY-Cut section/column detection unit tests: hand-built chars, no wasm. */
import { describe, expect, it } from 'vitest'
import { analyzePage } from '../src/analyze'
import { detectSections, type SectionElement } from '../src/analyze/columns'
import { clusterCombiningMarks, groupIntoLines } from '../src/analyze/lines'
import { splitIntoUnits } from '../src/analyze/units'
import type { ExtractedPage } from '../src/extract'
import type { PdfChar } from '../src/ir'
import { mkChar, mkText } from './helpers/chars'

const elementsOf = (chars: PdfChar[]): SectionElement[] =>
  splitIntoUnits(groupIntoLines(clusterCombiningMarks(chars))).map((u) => ({
    box: u.box,
    unit: u,
  }))

/** n rows of two side-by-side sentence columns (x 72.. / 320..) */
function twoColumnChars(rows = 6, leftX = 72, rightX = 320, topY = 700): PdfChar[] {
  const left = [
    'The quick brown fox jumps ov',
    'lazy dog while a cat watches',
    'from the windowsill and then',
    'sunlight drifts over the roo',
    'pages of an open book turnin',
    'slowly in the afternoon bree',
  ]
  const right = [
    'Meanwhile across the road th',
    'baker sets out fresh loaves,',
    'steam rising from the crusts',
    'as customers queue politely,',
    'coins ready in their pockets',
    'waiting for the door to open',
  ]
  const chars: PdfChar[] = []
  for (let i = 0; i < rows; i++) {
    const y = topY - i * 14
    chars.push(...mkText(left[i % left.length]!, leftX, { y }).chars)
    chars.push(...mkText(right[i % right.length]!, rightX, { y }).chars)
  }
  return chars
}

function extractedPage(chars: PdfChar[], over: Partial<ExtractedPage> = {}): ExtractedPage {
  return {
    index: 0,
    widthPt: 612,
    heightPt: 792,
    rotation: 0,
    chars,
    images: [],
    paths: [],
    degraded: false,
    scanned: false,
    hasStructTree: false,
    vectorRegions: [],
    badUnicodeRatio: 0,
    ...over,
  }
}

describe('splitIntoUnits: displaced whitespace (P20)', () => {
  it('drops a trailing NBSP drawn at the line left edge (Word-export artifact)', () => {
    // "Contact" at x=100.., plus an NBSP whose glyph sits at x=100.5 —
    // x-sorting would land it between C and o and mint "C ontact"
    const chars = [
      ...mkText('Contact', 100, { fontSize: 11 }).chars,
      mkChar(' ', 100.5, { fontSize: 11, width: 2.7 }),
    ]
    const units = splitIntoUnits(groupIntoLines(clusterCombiningMarks(chars)))
    expect(units).toHaveLength(1)
    expect(units[0]!.chars.map((c) => c.text).join('')).toBe('Contact')
  })

  it('keeps a real interior space that sits in its own gap', () => {
    const chars = [...mkText('ab', 100).chars, mkChar(' ', 110), ...mkText('cd', 116).chars]
    const units = splitIntoUnits(groupIntoLines(clusterCombiningMarks(chars)))
    expect(units).toHaveLength(1)
    expect(units[0]!.chars.map((c) => c.text).join('')).toBe('ab cd')
  })
})

describe('detectSections: table-anchored columns (P16 E)', () => {
  it('a lone table block anchors its column beside prose', () => {
    // slide: table (one aggregate block) left, prose lines right
    const proseChars = [
      ...mkText('The revenue expenses and net', 500, { y: 690 }).chars,
      ...mkText('income of the venture over t', 500, { y: 670 }).chars,
      ...mkText('the next five years will gro', 500, { y: 650 }).chars,
      ...mkText('and focus on the key drivers', 500, { y: 630 }).chars,
    ]
    const table: SectionElement = {
      box: { x0: 49, y0: 560, x1: 449, y1: 700 },
      block: {
        kind: 'table',
        box: { x0: 49, y0: 560, x1: 449, y1: 700 },
        colWidthsPt: [100, 100, 100, 100],
        rows: [],
        confidence: 0.8,
      },
    }
    const sections = detectSections([...elementsOf(proseChars), table], 540, 960)
    expect(sections).toHaveLength(1)
    expect(sections[0]!.columns).toHaveLength(2)
    expect(sections[0]!.columns[0]!.elements.some((e) => e.block?.kind === 'table')).toBe(true)
  })
})

describe('detectSections', () => {
  it('a plain paragraph page is one single-column section', () => {
    const chars = [
      ...mkText('The quick brown fox jumps over the lazy dog', 72, { y: 700 }).chars,
      ...mkText('while the cat watches from a warm windowsill', 72, { y: 686 }).chars,
      ...mkText('in the afternoon sunshine of early October.', 72, { y: 672 }).chars,
    ]
    const sections = detectSections(elementsOf(chars), 792)
    expect(sections).toHaveLength(1)
    expect(sections[0]!.columns).toHaveLength(1)
  })

  it('a two-column page yields one section with two columns and a gutter', () => {
    const sections = detectSections(elementsOf(twoColumnChars()), 792)
    expect(sections).toHaveLength(1)
    const s = sections[0]!
    expect(s.columns).toHaveLength(2)
    expect(s.columns[0]!.elements).toHaveLength(6)
    expect(s.columns[1]!.elements).toHaveLength(6)
    expect(s.gutters).toHaveLength(1)
    expect(s.columns[0]!.box.x1).toBeLessThan(s.columns[1]!.box.x0)
    expect(s.dir).toBe('ltr')
  })

  it('supports three columns', () => {
    const texts = ['alpha beta gamma dl', 'delta epsilon zeta t', 'ta theta iota kappa']
    const chars: PdfChar[] = []
    for (let i = 0; i < 5; i++) {
      const y = 700 - i * 14
      chars.push(...mkText(texts[i % 3]!, 60, { y }).chars)
      chars.push(...mkText(texts[(i + 1) % 3]!, 240, { y }).chars)
      chars.push(...mkText(texts[(i + 2) % 3]!, 420, { y }).chars)
    }
    const sections = detectSections(elementsOf(chars), 792)
    expect(sections).toHaveLength(1)
    expect(sections[0]!.columns).toHaveLength(3)
    expect(sections[0]!.gutters).toHaveLength(2)
  })

  it('columns may have unequal widths', () => {
    const chars: PdfChar[] = []
    for (let i = 0; i < 4; i++) {
      const y = 700 - i * 14
      chars.push(...mkText('short col text here', 72, { y }).chars) // ~95pt wide
      chars.push(...mkText('this right column is much wider than the left one', 220, { y }).chars)
    }
    const sections = detectSections(elementsOf(chars), 792)
    expect(sections[0]!.columns).toHaveLength(2)
    const [a, b] = sections[0]!.columns
    expect(b!.box.x1 - b!.box.x0).toBeGreaterThan(1.5 * (a!.box.x1 - a!.box.x0))
  })

  it('splits a full-width title above a two-column body into two sections', () => {
    const title = mkText('Annual Report Overview Twenty Twenty Six', 160, { y: 730 }).chars
    const sections = detectSections(elementsOf([...title, ...twoColumnChars()]), 792)
    expect(sections).toHaveLength(2)
    expect(sections[0]!.columns).toHaveLength(1)
    expect(sections[1]!.columns).toHaveLength(2)
  })

  it('closes a column section when full-width text resumes below', () => {
    const chars = [
      ...twoColumnChars(),
      ...mkText('This closing paragraph spans the entire page width below both columns.', 72, {
        y: 600,
      }).chars,
    ]
    const sections = detectSections(elementsOf(chars), 792)
    expect(sections).toHaveLength(2)
    expect(sections[0]!.columns).toHaveLength(2)
    expect(sections[1]!.columns).toHaveLength(1)
  })

  it('degrades a two-element pair (author names) to a single column', () => {
    const chars = [
      ...mkText('Alice Cooper', 100, { y: 700 }).chars,
      ...mkText('Bob Dylan', 350, { y: 700 }).chars,
      ...mkText('University A', 100, { y: 686 }).chars,
      ...mkText('University B', 350, { y: 686 }).chars,
    ]
    const sections = detectSections(elementsOf(chars), 792)
    expect(sections.every((s) => s.columns.length === 1)).toBe(true)
  })

  // ── P14 A: slide card layouts (PowerPoint→PDF implicit two-column) ──

  /** two side-by-side cards: number badge + title on one baseline, body below */
  function cardChars(closingLineX?: number): PdfChar[] {
    const chars: PdfChar[] = [
      ...mkText('01', 72, { y: 700 }).chars,
      ...mkText('card one title text here', 110, { y: 700 }).chars,
      ...mkText('body line one for card one', 110, { y: 686 }).chars,
      ...mkText('body line two for card one', 110, { y: 672 }).chars,
      ...mkText('02', 320, { y: 700 }).chars,
      ...mkText('card two title text here', 358, { y: 700 }).chars,
      ...mkText('body line one for card two', 358, { y: 686 }).chars,
      ...mkText('body line two for card two', 358, { y: 672 }).chars,
    ]
    if (closingLineX !== undefined) {
      chars.push(
        ...mkText(
          'a closing paragraph spanning the full page width below the cards',
          closingLineX,
          {
            y: 640,
          },
        ).chars,
      )
    }
    return chars
  }

  it('same-baseline cards split into two columns despite number-badge sliver columns (P14 A)', () => {
    const sections = detectSections(elementsOf(cardChars()), 540, 960)
    expect(sections).toHaveLength(1)
    const s = sections[0]!
    expect(s.columns).toHaveLength(2)
    // badge + title + 2 body lines per card; the badge's pseudo-gutter is pruned
    expect(s.columns[0]!.elements).toHaveLength(4)
    expect(s.columns[1]!.elements).toHaveLength(4)
  })

  it('a full-width closer starting right of a badge pseudo-gutter still closes the section (P14 A)', () => {
    // x=112 rides the [~82,110] badge gutter while killing the real one
    const sections = detectSections(elementsOf(cardChars(112)), 540, 960)
    expect(sections).toHaveLength(2)
    expect(sections[0]!.columns).toHaveLength(2)
    expect(sections[1]!.columns).toHaveLength(1)
  })

  it('a tall gutterless title/quote stack closes before side-by-side columns open (P14 A)', () => {
    const stack = [
      ...mkText('a big standalone title line', 72, { y: 760 }).chars,
      ...mkText('quote line one below the title text', 72, { y: 742 }).chars,
      ...mkText('quote line two continues the block', 72, { y: 724 }).chars,
    ]
    const sections = detectSections(elementsOf([...stack, ...cardChars()]), 540, 960)
    expect(sections).toHaveLength(2)
    expect(sections[0]!.columns).toHaveLength(1)
    expect(sections[0]!.columns[0]!.elements).toHaveLength(3)
    expect(sections[1]!.columns).toHaveLength(2)
  })

  // ── P15 B: stats row + photo caption set apart by a vast gutter ──

  /** three stat blocks (number + two label lines) and an optional far-right caption pair */
  function statsChars(captionX?: number): PdfChar[] {
    const stats: PdfChar[] = []
    for (const [i, x] of [72, 240, 408].entries()) {
      stats.push(...mkText(`${i + 8}0+`, x, { y: 700, fontSize: 30 }).chars)
      stats.push(...mkText('YEARS OF HISTORY', x, { y: 668 }).chars)
      stats.push(...mkText('since the last century', x, { y: 654 }).chars)
    }
    if (captionX !== undefined) {
      stats.push(...mkText('Sheng Jian Bao', captionX, { y: 668 }).chars)
      stats.push(...mkText('the iconic breakfast', captionX, { y: 654 }).chars)
    }
    return stats
  }

  it('keeps a 2-element caption column set apart by a vast gutter (P15 B)', () => {
    // caption at x=790: ~280pt clear of the last stat block (vast vs 24pt slide gutterMin)
    const sections = detectSections(elementsOf(statsChars(790)), 540, 960)
    expect(sections).toHaveLength(1)
    const s = sections[0]!
    expect(s.columns).toHaveLength(4)
    expect(s.columns[3]!.elements).toHaveLength(2)
    // stat labels must not absorb the caption text
    expect(s.columns[2]!.elements).toHaveLength(3)
  })

  it('still merges a 2-element column behind an ordinary gutter', () => {
    // caption at x=545: ~34pt gap — an ordinary column gap, not a deliberate set-apart
    const sections = detectSections(elementsOf(statsChars(545)), 540, 960)
    expect(sections).toHaveLength(1)
    const s = sections[0]!
    expect(s.columns).toHaveLength(3)
    expect(s.columns[2]!.elements).toHaveLength(5)
  })

  it('an RTL two-column page reports rtl direction', () => {
    const chars: PdfChar[] = []
    for (let i = 0; i < 4; i++) {
      const y = 700 - i * 14
      chars.push(...mkText('مرحبا بالعالم مرحبا بالعالم', 60, { y }).chars)
      chars.push(...mkText('שלום עולם שלום עולם שלום', 340, { y }).chars)
    }
    const sections = detectSections(elementsOf(chars), 792)
    expect(sections).toHaveLength(1)
    expect(sections[0]!.columns).toHaveLength(2)
    expect(sections[0]!.dir).toBe('rtl')
  })
})

describe('analyzePage: multi-column reading order', () => {
  it('reads the left column before the right column', () => {
    const page = analyzePage(extractedPage(twoColumnChars()))
    expect(page.sections).toHaveLength(1)
    const section = page.sections![0]!
    expect(section.columns).toHaveLength(2)
    const textOf = (col: number): string =>
      section.columns[col]!.blocks.flatMap((b) => (b.kind === 'text' ? b.lines : []))
        .flatMap((l) => l.spans.map((s) => s.text))
        .join(' ')
    expect(textOf(0)).toContain('quick brown fox')
    expect(textOf(1)).toContain('across the road')
    // flattened page.blocks: all left-column content precedes right-column content
    const flat = page.blocks
      .flatMap((b) => (b.kind === 'text' ? b.lines : []))
      .flatMap((l) => l.spans.map((s) => s.text))
      .join(' ')
    expect(flat.indexOf('quick brown fox')).toBeLessThan(flat.indexOf('across the road'))
  })

  it('reads an RTL page right column first', () => {
    const chars: PdfChar[] = []
    for (let i = 0; i < 4; i++) {
      const y = 700 - i * 14
      // left column carries a Latin marker word, right column pure Arabic
      chars.push(...mkText('west column text here now', 60, { y }).chars)
      chars.push(...mkText('مرحبا بالعالم مرحبا بالعالم', 340, { y }).chars)
    }
    const page = analyzePage(extractedPage(chars))
    const section = page.sections![0]!
    expect(section.dir).toBe('rtl')
    expect(section.columns).toHaveLength(2)
    // reading order: right (Arabic) column first
    expect(section.columns[0]!.box.x0).toBeGreaterThan(section.columns[1]!.box.x0)
  })

  it('keeps a plain single-column page on the P1/P2 path (one section, one column)', () => {
    const chars = [
      // centered against the mirrored body (bodyLeft 72 → mirrored right 540)
      ...mkText('Sample Document', 239, { y: 720, fontSize: 18 }).chars,
      ...mkText('The quick brown fox jumps over the lazy dog today.', 72, { y: 660 }).chars,
      ...mkText('It then naps in the warm afternoon sun for hours.', 72, { y: 646 }).chars,
    ]
    const page = analyzePage(extractedPage(chars))
    expect(page.sections).toHaveLength(1)
    expect(page.sections![0]!.columns).toHaveLength(1)
    // the mirrored-body centering inference still applies (P2 behavior)
    const title = page.blocks.find(
      (b) => b.kind === 'text' && b.lines[0]!.spans[0]!.text.includes('Sample'),
    )
    expect(title && title.kind === 'text' && title.align).toBe('center')
  })
})

describe('mergeTwinSections / leader index pages (P22 D)', () => {
  it('re-joins a torn two-column run (same split, adjacent sections)', () => {
    // 12 two-column body rows with a mid-page full-width-ish line that
    // kills the gutter for one slab and tears the run in two
    const chars: PdfChar[] = []
    for (let i = 0; i < 6; i++) {
      const y = 720 - i * 14
      chars.push(...mkText('left column body text here', 60, { y }).chars)
      chars.push(...mkText('right column body text too', 340, { y }).chars)
    }
    // the tear: one line whose x-range crosses the gutter
    chars.push(
      ...mkText('a heading line crossing the middle of the page width', 60, { y: 630 }).chars,
    )
    for (let i = 0; i < 6; i++) {
      const y = 610 - i * 14
      chars.push(...mkText('left column body text more', 60, { y }).chars)
      chars.push(...mkText('right column body text more', 340, { y }).chars)
    }
    const page = analyzePage(extractedPage(chars))
    // without merging this shatters into [2col, 1col, 2col] or worse; the
    // crossing line's own section may remain, but the two 2-col runs that
    // share the same gutter must not BOTH survive as separate sections
    const twoCol = page.sections!.filter((s) => s.columns.length === 2)
    expect(twoCol.length).toBeLessThanOrEqual(1)
  })

  it('keeps differently-split multi-column sections apart', () => {
    const chars: PdfChar[] = []
    for (let i = 0; i < 4; i++) {
      const y = 720 - i * 14
      chars.push(...mkText('left one text block here', 60, { y }).chars)
      chars.push(...mkText('right one text block here', 340, { y }).chars)
    }
    for (let i = 0; i < 4; i++) {
      const y = 640 - i * 14
      chars.push(...mkText('narrow', 60, { y }).chars)
      chars.push(
        ...mkText('a wide right column starting far left of the other gutter', 160, { y }).chars,
      )
    }
    const page = analyzePage(extractedPage(chars))
    const twoCol = page.sections!.filter((s) => s.columns.length === 2)
    // different gutter x-ranges → no merge
    expect(twoCol.length).toBe(2)
  })

  it('forces a dot-leader index page into one single-column section', () => {
    const chars: PdfChar[] = []
    for (let i = 0; i < 12; i++) {
      const y = 700 - i * 12
      chars.push(...mkText('Entry name', 60, { y }).chars)
      for (let x = 130; x <= 220; x += 6) chars.push(mkChar('.', x, { y, width: 2 }))
      chars.push(...mkText(String(10 + i), 226, { y }).chars)
      // second index column on the same row
      chars.push(...mkText('Other entry', 320, { y }).chars)
      for (let x = 400; x <= 490; x += 6) chars.push(mkChar('.', x, { y, width: 2 }))
      chars.push(...mkText(String(30 + i), 496, { y }).chars)
    }
    const page = analyzePage(extractedPage(chars))
    expect(page.sections!.every((s) => s.columns.length === 1)).toBe(true)
  })
})

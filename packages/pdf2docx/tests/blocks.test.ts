import { describe, expect, it } from 'vitest'
import { groupIntoBlocks } from '../src/analyze'
import type { Line, Span } from '../src/ir'
import type { Rect } from '../src/geometry'

/** hand-built IR line: box + one span covering it */
function mkLine(x0: number, x1: number, topY: number, fontSize = 10, text = 'x'): Line {
  const box: Rect = { x0, x1, y0: topY - fontSize, y1: topY }
  const span: Span = {
    text,
    box,
    fontSize,
    fontFamily: 'Helvetica',
    bold: false,
    italic: false,
    color: '000000',
    dir: 'ltr',
    script: 'latin',
  }
  return { spans: [span], box, baseline: topY - fontSize * 0.8, endsWithHyphen: false }
}

// body: 72..540, line height 10. Tight leading: 12pt pitch → 2pt box gap.
const LEFT = 72
const RIGHT = 540

describe('paragraph clustering', () => {
  it('keeps tightly-leaded full lines in one block', () => {
    const lines = [
      mkLine(LEFT, RIGHT, 700),
      mkLine(LEFT, RIGHT, 688),
      mkLine(LEFT, RIGHT - 100, 676), // short last line
    ]
    expect(groupIntoBlocks(lines)).toHaveLength(1)
  })

  it('splits on a vertical gap beyond the median-based tolerance', () => {
    const lines = [
      mkLine(LEFT, RIGHT, 700),
      mkLine(LEFT, RIGHT, 688), // gap 2
      mkLine(LEFT, RIGHT, 676), // gap 2
      mkLine(LEFT, RIGHT, 656), // gap 10 > 1.5×2 and > 0.45×10
      mkLine(LEFT, RIGHT, 644),
    ]
    const blocks = groupIntoBlocks(lines)
    expect(blocks).toHaveLength(2)
    expect(blocks[0]!.lines).toHaveLength(3)
    expect(blocks[1]!.lines).toHaveLength(2)
  })

  it('never joins across gaps of several line heights, even when they are the median', () => {
    // sparse leftovers around consumed form/table regions: uniform ~70pt gaps
    // make the median gap ≈ every gap — the absolute cap must still split
    const lines = [
      mkLine(LEFT, RIGHT, 700),
      mkLine(LEFT, RIGHT, 620),
      mkLine(LEFT, RIGHT, 540),
      mkLine(LEFT, RIGHT, 460),
    ]
    expect(groupIntoBlocks(lines)).toHaveLength(4)
  })

  it('splits when a line at the body edge is followed by an indented line', () => {
    const lines = [
      mkLine(LEFT, RIGHT, 700),
      mkLine(LEFT, RIGHT, 688),
      mkLine(LEFT + 20, RIGHT, 676), // 2-em indent after a flush line
      mkLine(LEFT, RIGHT, 664),
    ]
    const blocks = groupIntoBlocks(lines)
    expect(blocks).toHaveLength(2)
    expect(blocks[1]!.lines).toHaveLength(2)
    expect(blocks[1]!.firstLineIndentPt).toBeCloseTo(20)
  })

  it('splits after a short left-anchored line (paragraph end)', () => {
    const lines = [
      mkLine(LEFT, RIGHT, 700),
      mkLine(LEFT, RIGHT - 200, 688), // ends 200pt short of the right edge
      mkLine(LEFT, RIGHT, 676),
      mkLine(LEFT, RIGHT, 664),
    ]
    const blocks = groupIntoBlocks(lines)
    expect(blocks).toHaveLength(2)
    expect(blocks[0]!.lines).toHaveLength(2)
  })

  it('does not apply the short-line rule to centered lines', () => {
    // a centered title stack: varying lefts → not left-anchored → no split
    const lines = [mkLine(200, 412, 700, 10), mkLine(230, 382, 688, 10)]
    expect(groupIntoBlocks(lines)).toHaveLength(1)
  })
})

describe('alignment inference', () => {
  it('multi-line left-aligned block (ragged right)', () => {
    const lines = [
      mkLine(LEFT, RIGHT, 700),
      mkLine(LEFT, RIGHT - 60, 688),
      mkLine(LEFT, RIGHT - 25, 676),
    ]
    const blocks = groupIntoBlocks(lines)
    expect(blocks[0]!.align).toBe('left')
  })

  it('multi-line centered block', () => {
    const ctx = { bodyLeft: LEFT, bodyRight: RIGHT }
    const lines = [
      mkLine(206, 406, 700), // center 306
      mkLine(246, 366, 688),
      mkLine(226, 386, 676),
    ]
    const blocks = groupIntoBlocks(lines, ctx)
    expect(blocks[0]!.align).toBe('center')
  })

  it('multi-line right-aligned block (ragged left)', () => {
    const ctx = { bodyLeft: LEFT, bodyRight: RIGHT }
    const lines = [mkLine(300, RIGHT, 700), mkLine(340, RIGHT, 688), mkLine(320, RIGHT, 676)]
    const blocks = groupIntoBlocks(lines, ctx)
    expect(blocks[0]!.align).toBe('right')
  })

  it('justified block: non-final lines share a flush right edge (P16 G)', () => {
    const lines = [
      mkLine(LEFT, RIGHT, 700),
      mkLine(LEFT, RIGHT - 1, 688),
      mkLine(LEFT, RIGHT + 0.5, 676),
      mkLine(LEFT, RIGHT - 1.5, 664),
      mkLine(LEFT, RIGHT - 180, 652), // short last line
    ]
    const blocks = groupIntoBlocks(lines)
    expect(blocks[0]!.align).toBe('justify')
  })

  it('ragged full-width prose stays left-aligned', () => {
    const lines = [
      mkLine(LEFT, RIGHT, 700),
      mkLine(LEFT, RIGHT - 14, 688),
      mkLine(LEFT, RIGHT - 5, 676),
      mkLine(LEFT, RIGHT - 9, 664),
      mkLine(LEFT, RIGHT - 120, 652),
    ]
    const blocks = groupIntoBlocks(lines)
    expect(blocks[0]!.align).toBe('left')
  })

  it('short blocks never claim justify (insufficient evidence)', () => {
    const lines = [
      mkLine(LEFT, RIGHT, 700),
      mkLine(LEFT, RIGHT, 688),
      mkLine(LEFT, RIGHT - 200, 676),
    ]
    expect(groupIntoBlocks(lines)[0]!.align).toBe('left')
  })

  it('single centered line against the page body', () => {
    const ctx = { bodyLeft: LEFT, bodyRight: RIGHT }
    const title = [mkLine(256, 356, 700, 14, 'Title')] // centered at 306
    expect(groupIntoBlocks(title, ctx)[0]!.align).toBe('center')
  })

  it('single right-anchored line against the page body', () => {
    const ctx = { bodyLeft: LEFT, bodyRight: RIGHT }
    const line = [mkLine(440, RIGHT, 700)]
    expect(groupIntoBlocks(line, ctx)[0]!.align).toBe('right')
  })

  it('detects a first-line indent in a left-aligned paragraph', () => {
    const lines = [
      mkLine(LEFT + 21, RIGHT, 700), // ~2-em first-line indent
      mkLine(LEFT, RIGHT, 688),
      mkLine(LEFT, RIGHT - 40, 676),
    ]
    const blocks = groupIntoBlocks(lines, { bodyLeft: LEFT, bodyRight: RIGHT })
    expect(blocks).toHaveLength(1)
    expect(blocks[0]!.align).toBe('left')
    expect(blocks[0]!.firstLineIndentPt).toBeCloseTo(21)
  })
})

describe('intra-paragraph hard breaks (P7)', () => {
  const ctx = { bodyLeft: LEFT, bodyRight: RIGHT }
  /** a full-width body line elsewhere on the page anchors the wrap edge */
  const anchor = mkLine(LEFT, RIGHT, 600)

  it('marks the break in a two-line spaced-out title (both lines far short of the edge)', () => {
    const lines = [
      mkLine(LEFT, LEFT + 85, 700, 10, 'T E C H N I C A L'),
      mkLine(LEFT, LEFT + 58, 688, 10, 'R E V I E W'),
      anchor,
    ]
    const blocks = groupIntoBlocks(lines, ctx)
    expect(blocks[0]!.lines).toHaveLength(2)
    expect(blocks[0]!.lines[0]!.hardBreakBefore).toBeUndefined()
    expect(blocks[0]!.lines[1]!.hardBreakBefore).toBe(true)
  })

  it('does not mark ordinary auto-wrapped lines (full lines, short last line)', () => {
    const lines = [
      mkLine(LEFT, RIGHT, 700, 10, 'full line of text here'),
      mkLine(LEFT, RIGHT - 2, 688, 10, 'another full line'),
      mkLine(LEFT, RIGHT - 120, 676, 10, 'short tail'),
    ]
    const blocks = groupIntoBlocks(lines, ctx)
    expect(blocks).toHaveLength(1)
    expect(blocks[0]!.lines.every((l) => !l.hardBreakBefore)).toBe(true)
  })

  it('does not mark a ragged-right wrap whose leftover cannot hold the next word', () => {
    const lines = [
      mkLine(LEFT, RIGHT - 30, 700, 10, 'ragged wrap line'), // 30pt leftover
      mkLine(LEFT, LEFT + 45, 688, 10, 'wideword'), // one 45pt word — did not fit
      anchor,
    ]
    const blocks = groupIntoBlocks(lines, ctx)
    expect(blocks[0]!.lines[1]!.hardBreakBefore).toBeUndefined()
  })

  it('marks breaks in a centered multi-line title stack', () => {
    const mid = (LEFT + RIGHT) / 2
    const lines = [
      mkLine(mid - 80, mid + 80, 700, 10, 'A Longer Centered Title'),
      mkLine(mid - 30, mid + 30, 688, 10, 'Short Row'),
      mkLine(mid - 60, mid + 60, 676, 10, 'And A Third'),
      anchor,
    ]
    const blocks = groupIntoBlocks(lines, ctx)
    expect(blocks[0]!.lines).toHaveLength(3)
    expect(blocks[0]!.align).toBe('center')
    expect(blocks[0]!.lines[1]!.hardBreakBefore).toBe(true)
    expect(blocks[0]!.lines[2]!.hardBreakBefore).toBe(true)
  })

  it('marks breaks in an open-leaded slide zone stack far from the frame edge (P11 b)', () => {
    // list zone at x 300..420 inside a 72..540 frame: the whole-frame ratio
    // guard never fires there; open leading (36pt pitch at 20pt type) plus
    // first-word room in the zone's own extent marks the breaks
    const lines = [
      mkLine(300, 360, 700, 20, 'item one'),
      mkLine(300, 370, 664, 20, 'item three'),
      mkLine(300, 420, 628, 20, 'item two longer'),
      mkLine(LEFT, RIGHT, 500, 10, 'full-width anchor line far below'),
    ]
    const blocks = groupIntoBlocks(lines, ctx)
    expect(blocks[0]!.lines).toHaveLength(3)
    expect(blocks[0]!.lines[1]!.hardBreakBefore).toBe(true)
    expect(blocks[0]!.lines[2]!.hardBreakBefore).toBe(true)
  })

  it('does not zone-break dense-leaded prose in a narrow zone (P11 b)', () => {
    // same zone geometry but single-spaced (24pt pitch at 20pt type)
    const lines = [
      mkLine(300, 360, 700, 20, 'item one'),
      mkLine(300, 420, 676, 20, 'item two longer'),
      anchor,
    ]
    const blocks = groupIntoBlocks(lines, ctx)
    expect(blocks[0]!.lines).toHaveLength(2)
    expect(blocks[0]!.lines[1]!.hardBreakBefore).toBeUndefined()
  })

  it('pins authored breaks in a display-heading stack even at the wrap edge (P12 B)', () => {
    // 54pt two-line cover title whose first line runs flush to the wrap edge:
    // leftover ≈ 0 defeats the leftover gates, but display type keeps its
    // authored boundary (substitute fonts would re-wrap it in an ugly spot)
    const lines = [
      mkLine(LEFT + 200, RIGHT, 700, 54, '打造明星产品'),
      mkLine(LEFT, RIGHT - 30, 635, 54, '集成产品开发（IPD）'),
      mkLine(LEFT, RIGHT, 400, 10, 'full-width anchor line far below'),
    ]
    const blocks = groupIntoBlocks(lines, ctx)
    expect(blocks[0]!.lines).toHaveLength(2)
    expect(blocks[0]!.lines[1]!.hardBreakBefore).toBe(true)
  })

  it('does not pin big-print prose taller than the heading line cap (P12 B)', () => {
    const lines = [
      mkLine(LEFT, RIGHT, 700, 26, 'big print prose wrapping on'),
      mkLine(LEFT, RIGHT - 2, 668, 26, 'and on across several full'),
      mkLine(LEFT, RIGHT - 5, 636, 26, 'lines of large accessible'),
      mkLine(LEFT, LEFT + 90, 604, 26, 'body text'),
    ]
    const blocks = groupIntoBlocks(lines, ctx)
    expect(blocks[0]!.lines.length).toBeGreaterThan(3)
    expect(blocks[0]!.lines.every((l) => !l.hardBreakBefore)).toBe(true)
  })

  it('never marks after a hyphenated line (hyphenation is a soft wrap)', () => {
    // 25pt leftover (> tolerance), next first word ~13pt (fits) — only the
    // hyphen guard keeps this soft
    const short = mkLine(LEFT, RIGHT - 25, 700, 10, 'ends in a hyphen-')
    const next = mkLine(LEFT, RIGHT, 688, 10, 'a bcdefghijklmnopqrstuvwxyz and more')
    const blocks = groupIntoBlocks([{ ...short, endsWithHyphen: true }, next], ctx)
    expect(blocks).toHaveLength(1)
    expect(blocks[0]!.lines[1]!.hardBreakBefore).toBeUndefined()
  })
})

describe('run-on list-item split (P8 C)', () => {
  it('splits short left-aligned indented entries into their own paragraphs', () => {
    // marker-less bullet list (vector-drawn bullets): four tightly-leaded
    // entries sharing a shallow indent, each stopping short of the group's
    // own right extent although the next entry's first word had room there
    const intro = mkLine(LEFT, RIGHT, 706, 10, 'the intro line:')
    const items = [
      mkLine(LEFT + 20, LEFT + 300, 686, 10, 'align the goals with human values'),
      mkLine(LEFT + 20, LEFT + 290, 674, 10, 'see inside the network internals'),
      mkLine(LEFT + 20, LEFT + 390, 662, 10, 'train the model on explicit principles'),
      mkLine(LEFT + 20, LEFT + 330, 650, 10, 'find and mitigate frontier risks'),
    ]
    const blocks = groupIntoBlocks([intro, ...items])
    expect(blocks).toHaveLength(5)
    expect(blocks.slice(1).map((b) => b.lines)).toEqual(items.map((l) => [l]))
  })

  it('keeps a narrow paragraph whole when its lines wrap greedily at its own width', () => {
    // a paragraph narrower than the page column (page body widened by an
    // unrelated wide title): its lines end short of the PAGE column but
    // flush with each other — the next word never fit the leftover
    const title = mkLine(LEFT + 130, RIGHT, 710, 14, 'A Wide Unrelated Title')
    const para = [
      mkLine(LEFT, LEFT + 230, 686, 10, 'narrow column prose wraps here'),
      mkLine(LEFT, LEFT + 233, 674, 10, 'and keeps filling its own edge'),
      mkLine(LEFT, LEFT + 110, 662, 10, 'short last line.'),
    ]
    const blocks = groupIntoBlocks([title, ...para])
    expect(blocks).toHaveLength(2)
    expect(blocks[1]!.lines).toHaveLength(3)
  })

  it('keeps full-width prose whole (natural wraps never split)', () => {
    const lines = [
      mkLine(LEFT, RIGHT, 700, 10, 'full width body text line one that'),
      mkLine(LEFT, RIGHT - 4, 688, 10, 'continues to the column edge and'),
      mkLine(LEFT, RIGHT - 180, 676, 10, 'ends on a short final line.'),
    ]
    expect(groupIntoBlocks(lines)).toHaveLength(1)
  })
})

describe('text-level hyphenation detection (P21 B)', () => {
  const line = (text: string, topY: number, x1 = LEFT + 200): Line =>
    mkLine(LEFT, x1, topY, 10, text)

  it('joins a hyphen-ended line onto a lowercase Latin continuation', () => {
    const blocks = groupIntoBlocks([line('what can be ex-', 700), line('pected here', 688)])
    expect(blocks[0]!.lines[0]!.endsWithHyphen).toBe(true)
  })

  it('joins a broken compound too — indistinguishable from hyphenation', () => {
    const blocks = groupIntoBlocks([line('it is self-', 700), line('contained', 688)])
    expect(blocks[0]!.lines[0]!.endsWithHyphen).toBe(true)
  })

  it('keeps the hyphen when the next line starts uppercase or a digit', () => {
    const upper = groupIntoBlocks([line('the Smith-', 700), line('Jones theorem', 688)])
    expect(upper[0]!.lines[0]!.endsWithHyphen).toBe(false)
    const digit = groupIntoBlocks([line('during 2012-', 700), line('2013 the rate', 688)])
    expect(digit[0]!.lines[0]!.endsWithHyphen).toBe(false)
  })

  it('does not apply to CJK-adjacent hyphens', () => {
    const blocks = groupIntoBlocks([
      line('\u89c1\u7b2c\u4e00\u7ae0-', 700),
      line('\u7eed\u8868\u5185\u5bb9\u5728\u6b64', 688),
    ])
    expect(blocks[0]!.lines[0]!.endsWithHyphen).toBe(false)
  })

  it('always joins a soft hyphen (U+00AD), even before uppercase', () => {
    const blocks = groupIntoBlocks([line('Vertrags\u00ad', 700), line('Partner sind', 688)])
    expect(blocks[0]!.lines[0]!.endsWithHyphen).toBe(true)
  })
})

describe('short-line verse runs (P21 C)', () => {
  // verse: left-aligned tight stack, ragged rights, next word always had room.
  // The contract: no two verse lines flow into one re-wrappable paragraph —
  // either the P8 C split makes each line its own paragraph, or the verse
  // pass pins hard breaks at every boundary.
  it('never lets a ragged short-line stack rejoin into flowing text', () => {
    const lines = [
      mkLine(LEFT, LEFT + 150, 700, 10, 'Wer reitet so spat durch Nacht'),
      mkLine(LEFT, LEFT + 120, 688, 10, 'Es ist der Vater mit'),
      mkLine(LEFT, LEFT + 170, 676, 10, 'Er hat den Knaben wohl in dem Arm'),
      mkLine(LEFT, LEFT + 130, 664, 10, 'Er fasst ihn sicher'),
    ]
    const blocks = groupIntoBlocks(lines)
    const joinable = blocks.some((b) => b.lines.some((l, i) => i > 0 && l.hardBreakBefore !== true))
    expect(joinable).toBe(false)
  })

  it('marks punctuation-terminated stacks even when rights nearly agree', () => {
    const lines = [
      mkLine(LEFT, LEFT + 230, 700, 10, 'Ich liebe dich, mich reizt deine Gestalt;'),
      mkLine(LEFT, LEFT + 239, 688, 10, 'Und bist du nicht willig, so brauch ich Gewalt.'),
      mkLine(LEFT, LEFT + 214, 676, 10, 'Mein Vater, mein Vater, jetzt fasst er mich an!'),
      mkLine(LEFT, LEFT + 170, 664, 10, 'Erlkonig hat mir ein Leids getan!'),
    ]
    const blocks = groupIntoBlocks(lines)
    const flat = blocks.flatMap((b) => b.lines)
    expect(flat.slice(1).every((l) => l.hardBreakBefore === true)).toBe(true)
  })

  it('leaves greedily wrapped prose alone (saturated lines, no leftover room)', () => {
    const lines = [
      mkLine(LEFT, RIGHT, 700, 10, 'full prose line that reaches the wrap edge here'),
      mkLine(LEFT, RIGHT - 4, 688, 10, 'another nearly full prose line of the paragraph'),
      mkLine(LEFT, RIGHT - 2, 676, 10, 'and one more just as full as the previous ones'),
      mkLine(LEFT, LEFT + 120, 664, 10, 'short final line.'),
    ]
    const blocks = groupIntoBlocks(lines)
    const flat = blocks.flatMap((b) => b.lines)
    expect(flat.some((l) => l.hardBreakBefore)).toBe(false)
  })

  it('does not chain across a hyphenated wrap', () => {
    const lines = [
      mkLine(LEFT, LEFT + 150, 700, 10, 'a short line here'),
      mkLine(LEFT, LEFT + 148, 688, 10, 'that ends with hyphen-'),
      mkLine(LEFT, LEFT + 152, 676, 10, 'ated continuation text'),
    ]
    const blocks = groupIntoBlocks(lines)
    const flat = blocks.flatMap((b) => b.lines)
    expect(flat.some((l) => l.hardBreakBefore)).toBe(false)
  })
})

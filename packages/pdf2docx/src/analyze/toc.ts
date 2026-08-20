/**
 * TOC-entry detection (P6). Dot-leader lines ("Heading ......... 42") read
 * naturally as one text line, but leaving them as plain paragraphs loses
 * their structure: the literal dots neither survive font substitution nor
 * mark the line as a table-of-contents entry. Detected lines become their
 * own blocks carrying `tocEntry`; the rebuild layer swaps the dots for a
 * right-aligned dot-leader tab and a TOC pStyle.
 *
 * Only unambiguous lines convert: a non-empty title, at least four leader
 * dots, and a trailing arabic/roman page number.
 */
import type { Line, PdfChar, Span, TextBlock } from '../ir'
import { analyzeChars } from './chars'
import { firstStrongDir } from './rtl'
import type { LineUnit } from './units'
import { clusterUnitRows } from './units'

/** title ... leader dots ... page number */
const TOC_LINE_RE = /^(.*?[^\s.])\s*([.·]\s*){4,}\s*([0-9]+|[ivxlcdm]+)\s*$/i
/** each this many points of extra indent nests the entry one level deeper */
const LEVEL_INDENT_PT = 14
const MAX_TOC_LEVEL = 9

// ── leaderless entries ("ACKNOWLEDGEMENTS        iv") — unit-row based ──
/** a page number unit: bare arabic (≤3 digits) or roman page number */
const PAGENUM_UNIT_RE = /^(?:[0-9]{1,3}|[ivxlcdm]{1,6})$/i
/** at least this many consecutive rows make a leaderless TOC run */
const TOC_RUN_MIN_ROWS = 3
/** page-number right edges must line up within this many ems */
const PAGENUM_ALIGN_TOL_EMS = 1.5
/** rows further apart than this many row heights break the run */
const TOC_ROW_GAP_MAX_RATIO = 2.5

// ── dot-leader index pages (P22 D) ──
/** a run of at least this many dot glyphs (spaces between allowed) is a leader */
const LEADER_RUN_MIN_DOTS = 4
/** pages with at least this many leader lines… */
export const LEADER_PAGE_MIN_LINES = 10
/** …making up at least this share of the page's lines are index pages */
export const LEADER_PAGE_MIN_SHARE = 0.3

/**
 * The char sequence contains a dot-leader ENTRY: a run of dots followed by a
 * number ("entry .... 28"). The trailing number is load-bearing — dotted
 * fill-in lines on forms ("Name: ......") must not read as index entries.
 */
export function hasDotLeaderRun(chars: readonly PdfChar[]): boolean {
  let run = 0
  let armed = false
  for (const c of chars) {
    if (c.text === '.' || c.text === '·') {
      run++
      if (run >= LEADER_RUN_MIN_DOTS) armed = true
    } else if (c.text === ' ' || c.code === 0x20) {
      continue
    } else if (armed && c.text >= '0' && c.text <= '9') {
      return true
    } else {
      run = 0
      armed = false
    }
  }
  return false
}

/** split one line's spans at a text offset (span boundaries preserved) */
function spansBefore(spans: readonly Span[], end: number): Span[] {
  const out: Span[] = []
  let at = 0
  for (const span of spans) {
    const take = Math.min(span.text.length, Math.max(0, end - at))
    if (take > 0) {
      const text = span.text.slice(0, take).replace(/\s+$/u, '')
      if (text.length > 0) out.push({ ...span, text })
    }
    at += span.text.length
    if (at >= end) break
  }
  return out
}

/** a detected entry: title spans only; dots and number move to `tocEntry` */
function tocBlockOf(line: Line, src: TextBlock, level: number, pageNumber: string): TextBlock {
  const text = line.spans.map((s) => s.text).join('')
  const m = TOC_LINE_RE.exec(text)!
  const titleEnd = m[1]!.length
  const spans = spansBefore(line.spans, titleEnd)
  const titled: Line = { ...line, spans: spans.length > 0 ? spans : line.spans }
  return {
    kind: 'text',
    lines: [titled],
    box: line.box,
    align: 'left',
    firstLineIndentPt: 0,
    dir: firstStrongDir([titled]),
    tocEntry: { level, pageNumber },
  }
}

const romanRank = (s: string): number => {
  const one: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 }
  let total = 0
  const t = s.toLowerCase()
  for (let i = 0; i < t.length; i++) {
    const cur = one[t[i]!] ?? 0
    const next = one[t[i + 1] ?? ''] ?? 0
    total += cur < next ? -cur : cur
  }
  return total
}

interface TocRowCand {
  units: LineUnit[]
  box: LineUnit['box']
  /** trailing page-number unit */
  num: LineUnit
  roman: boolean
  value: number
}

export interface DetectedTocRows {
  blocks: TextBlock[]
  /** units that stay in the regular flow */
  remainingUnits: LineUnit[]
}

/**
 * Leaderless TOC entries: runs of ≥3 consecutive rows whose LAST unit is a
 * bare page number, right edges aligned, values not decreasing (roman front
 * matter may precede arabic chapters). Detected rows leave the flow as
 * TOC-entry blocks; anything else stays put — miss rather than misfire.
 */
export function detectTocRows(units: readonly LineUnit[]): DetectedTocRows {
  const rows = clusterUnitRows(units)
  const cands: Array<TocRowCand | null> = rows.map((row) => {
    if (row.units.length < 2) return null
    const num = row.units[row.units.length - 1]!
    const text = num.chars
      .map((c) => c.text)
      .join('')
      .trim()
    if (!PAGENUM_UNIT_RE.test(text)) return null
    const roman = /^[ivxlcdm]+$/i.test(text) && !/^\d+$/.test(text)
    return {
      units: row.units,
      box: row.box,
      num,
      roman,
      value: roman ? romanRank(text) : parseInt(text, 10),
    }
  })

  const blocks: TextBlock[] = []
  const consumed = new Set<LineUnit>()
  const acceptRun = (run: TocRowCand[]): void => {
    if (run.length < TOC_RUN_MIN_ROWS) return
    const em = run[0]!.num.fontSize || 12
    const rights = run.map((c) => c.num.box.x1)
    if (Math.max(...rights) - Math.min(...rights) > PAGENUM_ALIGN_TOL_EMS * em) return
    // roman front matter first, then arabic; each part non-decreasing
    let sawArabic = false
    let prev = -1
    for (const c of run) {
      if (c.roman && sawArabic) return
      if (!c.roman && !sawArabic) {
        sawArabic = true
        prev = -1
      }
      if (c.value < prev) return
      prev = c.value
    }
    const minX0 = Math.min(...run.map((c) => c.box.x0))
    for (const c of run) {
      const titleChars = c.units.slice(0, -1).flatMap((u) => u.chars)
      const lines = analyzeChars(titleChars)
      if (lines.length === 0) continue
      const level = Math.min(
        MAX_TOC_LEVEL,
        1 + Math.max(0, Math.floor((c.box.x0 - minX0) / LEVEL_INDENT_PT)),
      )
      blocks.push({
        kind: 'text',
        lines,
        box: c.box,
        align: 'left',
        firstLineIndentPt: 0,
        dir: firstStrongDir(lines),
        tocEntry: {
          level,
          pageNumber: c.num.chars
            .map((ch) => ch.text)
            .join('')
            .trim(),
        },
      })
      for (const u of c.units) consumed.add(u)
    }
  }
  let run: TocRowCand[] = []
  const close = (): void => {
    acceptRun(run)
    run = []
  }
  for (let i = 0; i < rows.length; i++) {
    const cand = cands[i]
    if (!cand) {
      close()
      continue
    }
    const prev = run[run.length - 1]
    if (prev) {
      const gap = prev.box.y0 - cand.box.y1
      if (gap > TOC_ROW_GAP_MAX_RATIO * Math.max(prev.box.y1 - prev.box.y0, 1)) close()
    }
    run.push(cand)
  }
  close()
  return { blocks, remainingUnits: units.filter((u) => !consumed.has(u)) }
}

/**
 * Split dot-leader lines out of a column's blocks into TOC-entry blocks.
 * Non-matching lines pass through grouped as before.
 */
export function detectTocBlocks(blocks: TextBlock[]): TextBlock[] {
  const matches = new Map<Line, { title: string; pageNumber: string }>()
  for (const block of blocks) {
    if (block.list || block.dir === 'rtl') continue
    for (const line of block.lines) {
      const m = TOC_LINE_RE.exec(line.spans.map((s) => s.text).join(''))
      if (m) matches.set(line, { title: m[1]!, pageNumber: m[3]! })
    }
  }
  if (matches.size === 0) return blocks

  // nesting from indentation relative to the leftmost entry
  const minX0 = Math.min(...[...matches.keys()].map((l) => l.box.x0))
  const levelOf = (line: Line): number =>
    Math.min(MAX_TOC_LEVEL, 1 + Math.max(0, Math.floor((line.box.x0 - minX0) / LEVEL_INDENT_PT)))

  const out: TextBlock[] = []
  for (const block of blocks) {
    if (!block.lines.some((l) => matches.has(l))) {
      out.push(block)
      continue
    }
    let plain: Line[] = []
    const flush = (): void => {
      if (plain.length > 0) {
        out.push({
          ...block,
          lines: plain,
          box: plain
            .map((l) => l.box)
            .reduce((a, b) => ({
              x0: Math.min(a.x0, b.x0),
              y0: Math.min(a.y0, b.y0),
              x1: Math.max(a.x1, b.x1),
              y1: Math.max(a.y1, b.y1),
            })),
        })
      }
      plain = []
    }
    for (const line of block.lines) {
      const m = matches.get(line)
      if (m) {
        flush()
        out.push(tocBlockOf(line, block, levelOf(line), m.pageNumber))
      } else {
        plain.push(line)
      }
    }
    flush()
  }
  return out
}

/** Paragraph clustering over the per-page search-index items (PDF user space).
    Lines first (baseline bucketing, split at column gutters), then paragraphs
    (leading + x-overlap + font-size heuristics). Deliberately conservative:
    splitting a paragraph in two is recoverable (line-level editing still works),
    merging across columns is not. */

import type { PageEntry } from './search'

export interface BlockLine {
  /** [x1, y1, x2, y2] in PDF user space (y-up) */
  rect: [number, number, number, number]
  text: string
  /** Baseline y */
  y: number
  fontSize: number
}

export interface TextBlock {
  rect: [number, number, number, number]
  /** Width-weighted dominant font size across the block's lines */
  fontSize: number
  /** Dominant baseline-to-baseline distance (fontSize × 1.2 for single-line blocks) */
  lineHeight: number
  /** How the block's lines are set against its extent. Justified text reads as
      'left' (a greedy re-wrap fills lines nearly full anyway); single-line blocks
      are 'left' — one line carries no alignment evidence. */
  align: 'left' | 'center' | 'right'
  lines: BlockLine[]
}

interface Frag {
  x: number
  y: number
  w: number
  h: number
  text: string
}

const median = (nums: number[]): number => {
  const s = [...nums].sort((a, b) => a - b)
  return s[Math.floor(s.length / 2)]!
}

/** Lower median (even counts take the smaller middle value): the outlier re-split
    compares each lead against this baseline, and a single glued-on gap must not be
    able to raise it — with two leads the plain median IS the stray gap, which would
    make it approve itself. */
const lowMedian = (nums: number[]): number => {
  const s = [...nums].sort((a, b) => a - b)
  return s[Math.ceil(s.length / 2) - 1]!
}

/** Width-weighted dominant value, bucketed to 0.5 units so one oversized glyph
    (drop cap, inline math) doesn't skew the line/block font size */
const dominant = (pairs: [size: number, weight: number][]): number => {
  const byBucket = new Map<number, number>()
  for (const [size, weight] of pairs) {
    const b = Math.round(size * 2) / 2
    byBucket.set(b, (byBucket.get(b) ?? 0) + weight)
  }
  let best = pairs[0]![0]
  let bestW = -1
  for (const [b, w] of byBucket)
    if (w > bestW) {
      bestW = w
      best = b
    }
  return best
}

/** Estimated width of a line's first wrap-unit: CJK breaks after any character
    (one em, two with a kinsoku pull-down), Latin carries the whole word to the
    next line. 0.6 em per Latin char errs high on purpose so a ragged wrap
    before a wide word never reads as a paragraph end. */
function firstTokenWidth(text: string, fs: number): number {
  const t = text.trimStart()
  if (!t) return 0
  if (/^[\u2e80-\u9fff\uf900-\ufaff]/.test(t)) return fs * 2
  const token = /^\S+/.exec(t)![0]
  return token.length * fs * 0.6
}

const buildLine = (seg: Frag[]): BlockLine => {
  let text = ''
  let x1 = Infinity
  let y1 = Infinity
  let x2 = -Infinity
  let y2 = -Infinity
  let prev: Frag | null = null
  for (const f of seg) {
    if (prev && f.x - (prev.x + prev.w) > f.h * 0.15 && !/\s$/.test(text) && !/^\s/.test(f.text))
      text += ' '
    text += f.text
    prev = f
    x1 = Math.min(x1, f.x)
    // Item y is the baseline; drop the box bottom below it so descenders stay
    // inside the rect (hover hit-testing works on the last line's lower strip)
    y1 = Math.min(y1, f.y - f.h * 0.2)
    x2 = Math.max(x2, f.x + f.w)
    y2 = Math.max(y2, f.y + f.h)
  }
  return {
    rect: [x1, y1, x2, y2],
    text,
    y: seg[0]!.y,
    fontSize: dominant(seg.map((f) => [f.h, f.w])),
  }
}

interface Para {
  lines: BlockLine[]
  /** Baseline deltas between consecutive lines */
  leads: number[]
}

export function groupPageBlocks(entry: PageEntry): TextBlock[] {
  const frags: Frag[] = []
  for (const it of entry.items) {
    if (it.rot || !(it.w > 0) || !(it.h > 0)) continue
    const text = entry.text.slice(it.start, it.end)
    if (!text.trim()) continue
    frags.push({ x: it.x, y: it.y, w: it.w, h: it.h, text })
  }
  if (frags.length === 0) return []

  // Baseline rows: top of page first (y-up), tolerance 0.3 × font size
  frags.sort((a, b) => b.y - a.y || a.x - b.x)
  const rows: Frag[][] = []
  for (const f of frags) {
    const row = rows[rows.length - 1]
    const anchor = row?.[0]
    if (row && anchor && Math.abs(f.y - anchor.y) <= Math.max(f.h, anchor.h) * 0.3) row.push(f)
    else rows.push([f])
  }

  // Split each row at horizontal gaps wider than ~a line-height (column gutters);
  // word gaps are far smaller — same threshold the DOM-side line grouping uses
  const lines: BlockLine[] = []
  for (const row of rows) {
    row.sort((a, b) => a.x - b.x)
    let seg: Frag[] = [row[0]!]
    for (let i = 1; i < row.length; i++) {
      const f = row[i]!
      const last = seg[seg.length - 1]!
      if (f.x - (last.x + last.w) > Math.max(f.h, last.h) * 1.2) {
        lines.push(buildLine(seg))
        seg = [f]
      } else seg.push(f)
    }
    lines.push(buildLine(seg))
  }

  // Paragraphs: attach each line (top→bottom) to the best-overlapping open
  // paragraph. Multiple paragraphs stay open at once so interleaved columns
  // (L, R, L, R in reading order) don't break each other's runs.
  lines.sort((a, b) => b.y - a.y || a.rect[0] - b.rect[0])
  const open: Para[] = []
  for (const ln of lines) {
    let best: Para | null = null
    let bestOv = 0
    for (const p of open) {
      const last = p.lines[p.lines.length - 1]!
      const gap = last.y - ln.y
      if (gap <= 0) continue
      const fs = Math.max(last.fontSize, ln.fontSize)
      if (Math.abs(last.fontSize - ln.fontSize) > fs * 0.1) continue
      // Bootstrap allows up to 2.5× the font size: designed CJK pages set body
      // leading around 2×, which the old 1.8× cutoff split into per-line blocks.
      // Over-merges this causes are undone by the outlier re-split below.
      const maxGap = p.leads.length ? median(p.leads) * 1.5 : fs * 2.5
      if (gap > maxGap) continue
      let x1 = Infinity
      let x2 = -Infinity
      for (const l of p.lines) {
        x1 = Math.min(x1, l.rect[0])
        x2 = Math.max(x2, l.rect[2])
      }
      const ov = Math.min(x2, ln.rect[2]) - Math.max(x1, ln.rect[0])
      const narrow = Math.min(x2 - x1, ln.rect[2] - ln.rect[0])
      if (!(narrow > 0) || ov < narrow * 0.6) continue
      if (ov > bestOv) {
        bestOv = ov
        best = p
      }
    }
    if (best) {
      best.leads.push(best.lines[best.lines.length - 1]!.y - ln.y)
      best.lines.push(ln)
    } else open.push({ lines: [ln], leads: [] })
  }

  // Split after paragraph-final short lines: adjacent paragraphs that share
  // leading, font size and x-extent (labeled footer paragraphs, stacked body
  // paragraphs, list items) are indistinguishable by leading alone. A real wrap
  // leaves at most the next line's first wrap-unit unfilled, so an internal
  // line falling shorter than that is a paragraph end, not a wrapped line.
  // Left-set text only: centered/right-set lines are ragged on the right by design.
  for (let i = open.length - 1; i >= 0; i--) {
    const p = open[i]!
    if (p.lines.length < 2) continue
    let x1 = Infinity
    let x2 = -Infinity
    for (const l of p.lines) {
      x1 = Math.min(x1, l.rect[0])
      x2 = Math.max(x2, l.rect[2])
    }
    const fs = dominant(p.lines.map((l) => [l.fontSize, l.rect[2] - l.rect[0]]))
    if (inferAlign(p.lines, x1, x2, fs) !== 'left') continue
    const parts: Para[] = []
    let cur: Para = { lines: [p.lines[0]!], leads: [] }
    for (let j = 1; j < p.lines.length; j++) {
      const prev = p.lines[j - 1]!
      const ln = p.lines[j]!
      const shortfall = x2 - prev.rect[2]
      if (shortfall > firstTokenWidth(ln.text, ln.fontSize) + ln.fontSize * 1.5) {
        parts.push(cur)
        cur = { lines: [ln], leads: [] }
      } else {
        cur.leads.push(p.leads[j - 1]!)
        cur.lines.push(ln)
      }
    }
    parts.push(cur)
    if (parts.length > 1) open.splice(i, 1, ...parts)
  }

  // Re-split where an internal gap is an outlier against the paragraph's own
  // (lower-)median leading: the generous bootstrap can glue a stray same-size
  // line (a heading, a list row) to the paragraph below, but real paragraphs
  // keep a consistent leading, so a boundary shows up as a jump. Two-line
  // groups have no consensus leading and fall back to a font-size cap.
  const split: Para[] = []
  for (const p of open) {
    if (p.leads.length === 0) {
      split.push(p)
      continue
    }
    if (p.leads.length === 1) {
      // A 2-line group has no internal consensus to compare against; fall back to
      // the font size: designed leading tops out around 2× (CJK body text), so a
      // lone gap beyond that is the generous bootstrap gluing a stray line
      const fs = Math.max(p.lines[0]!.fontSize, p.lines[1]!.fontSize)
      if (p.leads[0]! > fs * 2.1) {
        split.push({ lines: [p.lines[0]!], leads: [] }, { lines: [p.lines[1]!], leads: [] })
      } else split.push(p)
      continue
    }
    const med = lowMedian(p.leads)
    let cur: Para = { lines: [p.lines[0]!], leads: [] }
    for (let i = 1; i < p.lines.length; i++) {
      const lead = p.leads[i - 1]!
      if (lead > med * 1.45) {
        split.push(cur)
        cur = { lines: [p.lines[i]!], leads: [] }
      } else {
        cur.leads.push(lead)
        cur.lines.push(p.lines[i]!)
      }
    }
    split.push(cur)
  }

  return split.map((p) => {
    let x1 = Infinity
    let y1 = Infinity
    let x2 = -Infinity
    let y2 = -Infinity
    for (const l of p.lines) {
      x1 = Math.min(x1, l.rect[0])
      y1 = Math.min(y1, l.rect[1])
      x2 = Math.max(x2, l.rect[2])
      y2 = Math.max(y2, l.rect[3])
    }
    const fontSize = dominant(p.lines.map((l) => [l.fontSize, l.rect[2] - l.rect[0]]))
    return {
      rect: [x1, y1, x2, y2],
      fontSize,
      lineHeight: p.leads.length ? median(p.leads) : fontSize * 1.2,
      align: inferAlign(p.lines, x1, x2, fontSize),
      lines: p.lines,
    }
  })
}

/** Which edge the lines agree on BEST. Left check skips the first line (indent);
    right check skips the last (a paragraph's short closing line); center uses
    every line. Similar-width centered/right lines share the left edge within the
    tolerance too, so passing the left check alone is not enough — the smallest
    deviation wins, with left taking ties (justified text agrees on every edge,
    and 'left' is the lossless default: reflow keeps x1 without lineXOffsets). */
function inferAlign(
  lines: BlockLine[],
  x1: number,
  x2: number,
  fontSize: number,
): 'left' | 'center' | 'right' {
  if (lines.length < 2) return 'left'
  const tol = fontSize * 0.6
  const mid = (x1 + x2) / 2
  // Skipping leaves a two-line block with a single left/right witness whose edge
  // defines the block extent (deviation 0 by construction), so left would always
  // win — two-line blocks check every line and get an explicit indent carve-out
  const two = lines.length === 2
  const dev = (sel: (l: BlockLine) => number, skip: (i: number) => boolean) =>
    Math.max(...lines.map((l, i) => (!two && skip(i) ? 0 : Math.abs(sel(l)))))
  const leftDev = dev(
    (l) => l.rect[0] - x1,
    (i) => i === 0,
  )
  const rightDev = dev(
    (l) => x2 - l.rect[2],
    (i) => i === lines.length - 1,
  )
  const centerDev = dev(
    (l) => (l.rect[0] + l.rect[2]) / 2 - mid,
    () => false,
  )
  if (leftDev <= tol && leftDev <= centerDev && leftDev <= rightDev) return 'left'
  if (two && rightDev <= tol) {
    // A small left-edge shift with flush right edges is a first-line indent or a
    // hanging indent — a left-set paragraph either way, not right alignment
    const shift = Math.abs(lines[0]!.rect[0] - lines[1]!.rect[0])
    if (shift > tol && shift <= fontSize * 2.5) return 'left'
  }
  if (centerDev <= tol && centerDev <= rightDev) return 'center'
  if (!two) {
    // Hanging indent: first line flush at the block edge, every body line flush at
    // one small common inset. The skip-first left check reads the inset as leftDev
    // and flush right edges would then win as 'right'. Runs after the center check:
    // a geometry that centering fully explains stays centered (a longest-first
    // centered block with similar-width followers is indistinguishable otherwise).
    const rest = lines.slice(1)
    const restX1 = Math.min(...rest.map((l) => l.rect[0]))
    const restDev = Math.max(...rest.map((l) => l.rect[0] - restX1))
    const inset = restX1 - x1
    if (lines[0]!.rect[0] - x1 <= tol && restDev <= tol && inset > tol && inset <= fontSize * 2.5) {
      return 'left'
    }
  }
  if (rightDev <= tol) return 'right'
  return 'left'
}

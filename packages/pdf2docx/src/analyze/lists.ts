/**
 * List detection (pdf2docx rule 5): paragraph blocks whose lines start with a
 * marker (bullet glyph / increasing numbers) and show list structure become
 * list-item paragraphs; the rebuild layer turns them into real docx numbering.
 *
 * The bias is MISS RATHER THAN MISFIRE:
 * - round bullets need ≥2 sibling items at the same indent OR a hanging
 *   continuation line;
 * - dash bullets (dialogue!) additionally REQUIRE hanging-indent evidence;
 * - numbers need a strictly increasing run of ≥2 (a lone "1. Introduction"
 *   is a heading, not a list);
 * - nesting deeper than 3 levels degrades to plain indented paragraphs.
 */
import { median, rectUnionAll } from '../geometry'
import type { Line, ListInfo, Span, TextBlock } from '../ir'
import { firstStrongDir } from './rtl'
import { isEastAsianScript, scriptOf } from '../script'

/** marker x positions within this many ems belong to the same nesting level */
const LEVEL_X_TOL_EMS = 0.6
/** a continuation line indented at least this many ems past the marker is "hanging" */
const HANGING_MIN_EMS = 0.8
/** …and further than this is unrelated indentation, not a hanging body */
const HANGING_MAX_EMS = 6
/** deepest real list level (0-based); deeper items stay plain paragraphs */
const MAX_LIST_LEVELS = 3
/** dash bullets without hanging evidence: the marker column must sit at least
 * this many ems past the surrounding plain text's left edge (slide sub-bullets
 * are indented; dialogue dashes start at the prose margin) */
const WEAK_INDENT_MIN_EMS = 1.2

/** unambiguous bullet glyphs */
const STRONG_BULLETS = /^([•◦▪▫‣∙·●○■□◆◇])(\s*)/u
/** dash-family bullets — also how dialogue lines start, hence "weak" */
const WEAK_BULLETS = /^([–—−-])(\s+)/u
/** "1." / "12)" */
const ORDERED_DOT_PAREN = /^(\d{1,3})([.)])(\s*)/u
/** "(3)" */
const ORDERED_PARENS = /^\((\d{1,3})\)(\s*)/u
/**
 * "3.1." / "3.1.15." — multi-level outline numbers (2–3 levels). The trailing
 * dot is REQUIRED: "3.14 is pi" is a decimal number, not an outline marker.
 */
const ORDERED_MULTI = /^(\d{1,3}(?:\.\d{1,3}){1,2})\.(\s*)/u

interface ParsedMarker {
  kind: 'bullet' | 'ordered'
  /** dash-family bullet: needs hanging-indent corroboration */
  weak: boolean
  value?: number
  /** multi style: all outline ordinals ("3.1.15." → [3, 1, 15]) */
  values?: number[]
  style?: ListInfo['style']
  /** marker text incl. trailing whitespace, in text units (to strip) */
  text: string
}

const lineText = (line: Line): string => line.spans.map((s) => s.text).join('')

const lineFontSize = (line: Line): number => median(line.spans.map((s) => s.fontSize)) || 12

/**
 * Parse a leading list marker. The marker must be separated from the body by
 * whitespace, except before East-Asian text where the PDF often sets the
 * glyphs flush (bullet immediately followed by a CJK item, no space).
 */
export function parseListMarker(line: Line): ParsedMarker | null {
  const text = lineText(line)
  const bodyOk = (matched: string, space: string): boolean => {
    if (space.length > 0) return true
    const rest = text.slice(matched.length)
    if (rest.length === 0) return false
    return isEastAsianScript(scriptOf(rest.codePointAt(0) ?? 0))
  }

  let m = STRONG_BULLETS.exec(text)
  if (m) {
    if (!bodyOk(m[0], m[2]!)) return null
    return { kind: 'bullet', weak: false, text: m[0] }
  }
  m = WEAK_BULLETS.exec(text)
  if (m && text.length > m[0].length) {
    return { kind: 'bullet', weak: true, text: m[0] }
  }
  m = ORDERED_MULTI.exec(text)
  if (m) {
    if (!bodyOk(m[0], m[2]!)) return null
    const values = m[1]!.split('.').map(Number)
    return {
      kind: 'ordered',
      weak: false,
      value: values[values.length - 1],
      values,
      style: 'multi',
      text: m[0],
    }
  }
  m = ORDERED_DOT_PAREN.exec(text)
  if (m) {
    if (!bodyOk(m[0], m[3]!)) return null
    return {
      kind: 'ordered',
      weak: false,
      value: Number(m[1]),
      style: m[2] === '.' ? 'dot' : 'paren',
      text: m[0],
    }
  }
  m = ORDERED_PARENS.exec(text)
  if (m) {
    if (!bodyOk(m[0], m[2]!)) return null
    return { kind: 'ordered', weak: false, value: Number(m[1]), style: 'parens', text: m[0] }
  }
  return null
}

/** Strip `count` leading text units (plus following whitespace) off a line's spans. */
function stripMarker(line: Line, count: number): Line {
  const spans: Span[] = []
  let remaining = count
  let atBoundary = true
  for (const span of line.spans) {
    let text = span.text
    if (remaining > 0) {
      const take = Math.min(remaining, text.length)
      text = text.slice(take)
      remaining -= take
    }
    if (remaining === 0 && atBoundary && text.length > 0) {
      text = text.replace(/^\s+/u, '')
      if (text.length > 0) atBoundary = false
    }
    if (text.length > 0) spans.push({ ...span, text })
  }
  return { ...line, spans }
}

interface ItemCand {
  /** index of the source block in the input array */
  srcIndex: number
  lines: Line[]
  marker: ParsedMarker
  /** marker line's left edge (level clustering key) */
  x0: number
  fontSize: number
  /** has a continuation line indented past the marker (hanging body) */
  hanging: boolean
}

/** exploded view of one input block: optional plain prefix + item candidates */
interface Exploded {
  srcIndex: number
  block: TextBlock
  prefixLines: Line[]
  items: ItemCand[]
}

function explodeBlock(block: TextBlock, srcIndex: number): Exploded {
  const out: Exploded = { srcIndex, block, prefixLines: [], items: [] }
  if (block.dir === 'rtl') {
    // RTL list markers sit on the right edge; detection is LTR-only for now
    out.prefixLines = block.lines
    return out
  }
  let current: ItemCand | null = null
  for (const line of block.lines) {
    const marker = parseListMarker(line)
    if (marker) {
      current = {
        srcIndex,
        lines: [line],
        marker,
        x0: line.box.x0,
        fontSize: lineFontSize(line),
        hanging: false,
      }
      out.items.push(current)
    } else if (current) {
      const indent = (line.box.x0 - current.x0) / current.fontSize
      if (indent >= HANGING_MIN_EMS && indent <= HANGING_MAX_EMS) current.hanging = true
      current.lines.push(line)
    } else {
      out.prefixLines.push(line)
    }
  }
  return out
}

/** cluster marker x positions into nesting levels (ascending x = deeper) */
function levelOf(x0: number, levels: number[], tolPt: number): number {
  for (const [i, lx] of levels.entries()) if (Math.abs(x0 - lx) <= tolPt) return i
  return levels.length
}

/**
 * Validate one region (a maximal run of consecutive item candidates) and
 * return the accepted items with their list annotations. `seq` provides
 * page-unique ids for ordered runs.
 */
function validateRegion(
  items: ItemCand[],
  seq: { next: number },
  bodyLeftX0?: number,
): Map<ItemCand, ListInfo> {
  const accepted = new Map<ItemCand, ListInfo>()
  const em = median(items.map((i) => i.fontSize)) || 12
  const tolPt = LEVEL_X_TOL_EMS * em

  // level clustering over marker x positions
  const levelXs: number[] = []
  for (const item of items) {
    const lvl = levelOf(item.x0, levelXs, tolPt)
    if (lvl === levelXs.length) levelXs.push(item.x0)
  }
  const sortedXs = [...levelXs].sort((a, b) => a - b)
  const levelFor = (item: ItemCand): number => levelOf(item.x0, sortedXs, tolPt)

  // multi-level outline numbers ("3.1.15."): the level comes from the marker's
  // own depth, not from x clustering (outline items all sit at the margin).
  // A run = same depth, same prefix ordinals, last ordinal counting up by one.
  {
    let run: ItemCand[] = []
    const flushMulti = (): void => {
      if (run.length >= 2) {
        const seqId = seq.next++
        const first = run[0]!.marker.values!
        for (const item of run) {
          accepted.set(item, {
            kind: 'ordered',
            level: first.length - 1,
            seqId,
            start: first[first.length - 1]!,
            style: 'multi',
            startValues: first,
            marker: item.marker.text,
          })
        }
      }
      run = []
    }
    for (const item of items.filter((i) => i.marker.style === 'multi')) {
      const v = item.marker.values!
      const pv = run[run.length - 1]?.marker.values
      const follows =
        pv !== undefined &&
        v.length === pv.length &&
        v.slice(0, -1).every((x, k) => x === pv[k]) &&
        v[v.length - 1] === pv[pv.length - 1]! + 1
      if (follows) {
        run.push(item)
      } else {
        flushMulti()
        run = [item]
      }
    }
    flushMulti()
  }

  for (const [level] of sortedXs.entries()) {
    if (level >= MAX_LIST_LEVELS) continue // deeper nesting degrades to plain text
    const atLevel = items.filter((i) => levelFor(i) === level && i.marker.style !== 'multi')

    // bullets
    const bullets = atLevel.filter((i) => i.marker.kind === 'bullet')
    const strongCount = bullets.filter((i) => !i.marker.weak).length
    const anyHanging = bullets.some((i) => i.hanging)
    // dash bullets without hanging evidence still count as a list when ≥2
    // siblings sit clearly indented past the surrounding plain text (P20:
    // slide sub-bullets are single short lines — dialogue never indents)
    const indented = (i: ItemCand): boolean =>
      bodyLeftX0 !== undefined && i.x0 - bodyLeftX0 >= WEAK_INDENT_MIN_EMS * Math.max(i.fontSize, 1)
    const indentedWeakCount = bullets.filter((i) => i.marker.weak && indented(i)).length
    for (const item of bullets) {
      const ok = item.marker.weak
        ? anyHanging || (indentedWeakCount >= 2 && indented(item))
        : strongCount >= 2 || item.hanging
      if (ok) {
        accepted.set(item, { kind: 'bullet', level, marker: item.marker.text })
      }
    }

    // ordered: strictly increasing runs of ≥2 with a uniform marker style.
    // A multi-line item whose continuation returns FLUSH to the margin is a
    // Chinese official-document style paragraph, not a hanging list item —
    // the numbering level's hanging indent would push every continuation
    // line right of where the source drew it (P18 C, zh-nchu)
    const ordered = atLevel.filter(
      (i) => i.marker.kind === 'ordered' && (i.lines.length < 2 || i.hanging),
    )
    let run: ItemCand[] = []
    const flushRun = (): void => {
      if (run.length >= 2) {
        const seqId = seq.next++
        const start = run[0]!.marker.value!
        for (const item of run) {
          accepted.set(item, {
            kind: 'ordered',
            level,
            seqId,
            start,
            style: item.marker.style,
            marker: item.marker.text,
          })
        }
      }
      run = []
    }
    for (const item of ordered) {
      const prev = run[run.length - 1]
      if (
        prev &&
        item.marker.value === prev.marker.value! + 1 &&
        item.marker.style === prev.marker.style
      ) {
        run.push(item)
      } else {
        flushRun()
        run = [item]
      }
    }
    flushRun()
  }
  return accepted
}

function itemToBlock(cand: ItemCand, info: ListInfo): TextBlock {
  const lines = [stripMarker(cand.lines[0]!, cand.marker.text.length), ...cand.lines.slice(1)]
  const kept = lines.filter((l) => l.spans.length > 0)
  const finalLines = kept.length > 0 ? kept : cand.lines
  return {
    kind: 'text',
    lines: finalLines,
    box: rectUnionAll(cand.lines.map((l) => l.box)),
    align: 'left',
    firstLineIndentPt: 0,
    dir: firstStrongDir(finalLines),
    list: info,
  }
}

/** plain block rebuilt from a subset of an original block's lines */
function plainBlock(lines: Line[], src: TextBlock, isPrefix: boolean): TextBlock {
  return {
    kind: 'text',
    lines,
    box: rectUnionAll(lines.map((l) => l.box)),
    align: src.align,
    firstLineIndentPt: isPrefix ? src.firstLineIndentPt : 0,
    dir: firstStrongDir(lines),
  }
}

/**
 * Detect list items across a column's paragraph blocks. Accepted items become
 * their own TextBlocks carrying `list`; everything else passes through (a
 * rejected candidate re-joins its original block unchanged).
 */
export function detectListBlocks(
  blocks: TextBlock[],
  seq: { next: number },
  fallbackBodyLeftX0?: number,
): TextBlock[] {
  if (blocks.length === 0) return blocks
  const exploded = blocks.map((b, i) => explodeBlock(b, i))
  if (exploded.every((e) => e.items.length === 0)) return blocks

  // the surrounding plain text's left edge — the weak-bullet indent evidence
  // is judged against it (median: a stray page number cannot drag it around).
  // Slide layouts pin a sub-bullet group into its own column with no plain
  // neighbours; the page-level fallback (median section-column x0) serves it.
  const plainX0s = exploded
    .filter((e) => e.items.length === 0 && e.block.lines.length > 0)
    .map((e) => e.block.box.x0)
  const bodyLeftX0 = plainX0s.length > 0 ? median(plainX0s) : fallbackBodyLeftX0

  // regions = maximal runs of consecutive items; a plain block (or a plain
  // prefix inside a block) closes the running region
  const regions: ItemCand[][] = []
  let region: ItemCand[] = []
  for (const e of exploded) {
    if ((e.items.length === 0 || e.prefixLines.length > 0) && region.length > 0) {
      regions.push(region)
      region = []
    }
    region.push(...e.items)
  }
  if (region.length > 0) regions.push(region)

  const accepted = new Map<ItemCand, ListInfo>()
  for (const r of regions) {
    for (const [cand, info] of validateRegion(r, seq, bodyLeftX0)) accepted.set(cand, info)
  }
  if (accepted.size === 0) return blocks

  // reassemble: blocks without accepted items stay untouched; others split
  const out: TextBlock[] = []
  for (const e of exploded) {
    if (e.items.length === 0 || !e.items.some((i) => accepted.has(i))) {
      out.push(e.block)
      continue
    }
    if (e.prefixLines.length > 0) out.push(plainBlock(e.prefixLines, e.block, true))
    let plainRun: Line[] = []
    const flushPlain = (): void => {
      if (plainRun.length > 0) out.push(plainBlock(plainRun, e.block, false))
      plainRun = []
    }
    for (const item of e.items) {
      const info = accepted.get(item)
      if (info) {
        flushPlain()
        out.push(itemToBlock(item, info))
      } else {
        plainRun.push(...item.lines)
      }
    }
    flushPlain()
  }
  return out
}

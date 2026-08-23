/**
 * lines → paragraph blocks + paragraph-format inference (alignment,
 * first-line indent). Median-based gap clustering follows pdftext; the
 * indent / short-line refinements follow pdf2docx's paragraph rules.
 */
import { approxEq, median, rectCenterX, rectUnionAll } from '../geometry'
import type { Line, TextBlock } from '../ir'
import { isNoSpaceScript } from '../script'
import { firstStrongDir } from './rtl'

/** page body context: the text extent the alignment inference compares against */
export interface BodyContext {
  bodyLeft: number
  bodyRight: number
}

/** paragraph split: gap beyond max(RATIO × median gap, MIN × line height) */
const GAP_MEDIAN_RATIO = 1.5
const GAP_LINE_HEIGHT_MIN = 0.45
/**
 * run-on item split (P8 C): a line ending under this share of the column
 * width did not wrap greedily — with the next line's left edge aligned to
 * its own, it reads as one item of a marker-less list (vector-drawn bullets
 * never reach the text layer). Body prose fills ≥ ~0.9 of its column on
 * every non-final line, so 0.85 keeps natural wraps whole.
 */
const ITEM_MAX_WIDTH_RATIO = 0.85
/** …and the shared left edge must sit within this many ems of the column
 * edge (list indents are shallow; centered/right stacks sit much deeper) */
const ITEM_MAX_INDENT_EMS = 4
/**
 * …but never joined across more than this many line heights. The median-gap
 * rule alone breaks down when consumed table/form regions leave sparse
 * UNIFORMLY-spaced leftover lines (median gap ≈ every gap): one "paragraph"
 * then spans the whole page and overlaps the tables between its lines.
 */
const GAP_ABS_MAX_LINE_HEIGHTS = 2
/** first-line indents between these bounds (in ems) start a new paragraph */
const INDENT_MIN_EMS = 0.8
const INDENT_MAX_EMS = 4
/** justified prose needs at least this many lines (three non-final rights) */
const JUSTIFY_MIN_LINES = 4
/** …whose non-final right edges align within this many ems */
const JUSTIFY_RIGHT_TOL_EMS = 0.3
/** a line ending this many ems before the body right edge "ends" its paragraph */
const SHORT_LINE_EMS = 3
/**
 * consecutive lines whose font sizes differ by at least this ratio never share
 * a paragraph (P14 A): a 40pt display title over its 14pt latin subtitle used
 * to ride the gap rule into one block, and once the hard break between them
 * is missed they reflow onto one line. Prose never steps size like that.
 */
const SIZE_BREAK_RATIO = 1.5

function lineFontSize(line: Line): number {
  return median(line.spans.map((s) => s.fontSize)) || 12
}

export function bodyContextOf(lines: readonly Line[]): BodyContext {
  if (lines.length === 0) return { bodyLeft: 0, bodyRight: 0 }
  return {
    bodyLeft: Math.min(...lines.map((l) => l.box.x0)),
    bodyRight: Math.max(...lines.map((l) => l.box.x1)),
  }
}

/** true → `cur` starts a new paragraph */
function isParagraphBreak(
  prev: Line,
  cur: Line,
  medianGap: number,
  medianLineH: number,
  body: BodyContext,
): boolean {
  const fontSize = Math.max(lineFontSize(prev), lineFontSize(cur))
  const tol = 0.4 * fontSize

  // 0. display-type step: title and subtitle sizes never mix in a paragraph
  const smaller = Math.min(lineFontSize(prev), lineFontSize(cur))
  if (smaller > 0 && fontSize / smaller >= SIZE_BREAK_RATIO) return true

  // 1. vertical gap clustering (pdftext: tolerance_factor × median gap),
  //    hard-capped in absolute line heights
  const gap = prev.box.y0 - cur.box.y1
  const allowed = Math.min(
    Math.max(GAP_MEDIAN_RATIO * medianGap, GAP_LINE_HEIGHT_MIN * medianLineH),
    GAP_ABS_MAX_LINE_HEIGHTS * medianLineH,
  )
  if (gap > allowed) return true

  // verse lines (P21 C) stay in one block — the pinned break IS the layout;
  // the indent/short-line rules below would shred a stanza into fragments
  // whose inter-block spacing no longer matches the line pitch
  if (cur.hardBreakBefore) return false

  // 2. first-line indent: prev sits on the body's left edge, cur is indented
  //    AND runs (nearly) full width — centered stacks fail that guard
  const indent = cur.box.x0 - prev.box.x0
  if (
    approxEq(prev.box.x0, body.bodyLeft, tol) &&
    indent > INDENT_MIN_EMS * fontSize &&
    indent < INDENT_MAX_EMS * fontSize &&
    cur.box.x1 >= body.bodyRight - 2 * fontSize
  ) {
    return true
  }

  // 3. short line: a left-anchored line ending well before the right edge
  //    finished its paragraph (both left-anchored → not centered text). The
  //    continuation must extend clearly past the short line's end — that
  //    guards against ragged-right jitter and against bodyRight being pushed
  //    out by an unrelated wide line (e.g. a large-type title)
  if (
    approxEq(prev.box.x0, body.bodyLeft, tol) &&
    approxEq(cur.box.x0, body.bodyLeft, tol) &&
    body.bodyRight - prev.box.x1 > SHORT_LINE_EMS * fontSize &&
    cur.box.x1 - prev.box.x1 > fontSize
  ) {
    return true
  }

  return false
}

/**
 * Run-on list-item split (P8 C): a group whose lines all share one shallow
 * left indent is either a wrapped paragraph or a marker-less list (vector-
 * drawn bullets never reach the text layer, so nothing else distinguishes
 * them). The tell is a line that stops short of the GROUP's own right extent
 * even though the next line's first word still had room there — a greedy
 * wrap never leaves that gap, on any of its lines. One such pair marks the
 * whole group as a list and every line becomes its own paragraph.
 *
 * Deliberately conservative: full-width prose shows no such pair (its next
 * word never fits the leftover), centered/right stacks do not share a left
 * edge, first-line-indented paragraphs misalign at the first pair, and the
 * 0.85-column-width cap keeps ragged near-full lines out.
 */
function splitRunOnItems(group: Line[], wrap: { left: number; right: number }): Line[][] {
  if (group.length < 2) return [group]
  const fontSize = median(group.map(lineFontSize)) || 12
  const tol = 0.4 * fontSize
  const left = group[0]!.box.x0
  if (left - wrap.left >= ITEM_MAX_INDENT_EMS * fontSize) return [group]
  if (!group.every((l) => approxEq(l.box.x0, left, tol))) return [group]
  const groupRight = Math.max(...group.map((l) => l.box.x1))
  const isItemEnd = (l: Line, next: Line): boolean =>
    l.box.x1 - l.box.x0 < ITEM_MAX_WIDTH_RATIO * (wrap.right - wrap.left) &&
    groupRight - l.box.x1 >= firstWordWidthPt(next) + HARD_BREAK_WORD_SLACK_EMS * fontSize
  const evidence = group.some((l, i) => i < group.length - 1 && isItemEnd(l, group[i + 1]!))
  if (!evidence) return [group]
  // each line becomes its own paragraph — a leading break flag (e.g. from the
  // verse pass, which overlaps marker-less lists by design) is meaningless
  return group.map((l) => {
    if (!l.hardBreakBefore) return [l]
    const { hardBreakBefore: _, ...rest } = l
    return [rest]
  })
}

/** alignment + first-line indent for a finished group of lines */
function inferFormat(
  lines: Line[],
  body: BodyContext,
): { align: TextBlock['align']; firstLineIndentPt: number } {
  const fontSize = median(lines.map(lineFontSize)) || 12
  const tol = Math.max(fontSize, (body.bodyRight - body.bodyLeft) * 0.02)

  if (lines.length === 1) {
    const line = lines[0]!
    const leftGap = line.box.x0 - body.bodyLeft
    const rightGap = body.bodyRight - line.box.x1
    if (leftGap <= tol) return { align: 'left', firstLineIndentPt: 0 }
    if (rightGap <= tol && leftGap > 2 * tol) return { align: 'right', firstLineIndentPt: 0 }
    if (approxEq(leftGap, rightGap, 2 * tol) && leftGap > 2 * tol) {
      return { align: 'center', firstLineIndentPt: 0 }
    }
    // left-ish with an offset: treat as an indented paragraph start
    const indent = leftGap > 0.5 * fontSize && leftGap < INDENT_MAX_EMS * fontSize ? leftGap : 0
    return { align: 'left', firstLineIndentPt: indent }
  }

  const lefts = lines.map((l) => l.box.x0)
  const rights = lines.map((l) => l.box.x1)
  const centers = lines.map((l) => rectCenterX(l.box))
  const spread = (values: number[]): number => Math.max(...values) - Math.min(...values)

  // the first line may be indented; judge left alignment on the rest
  const bodyLefts = lefts.slice(1)
  const leftAligned = spread(bodyLefts) <= tol
  const rightAligned = spread(rights) <= tol
  const centerAligned = spread(centers) <= tol

  let align: TextBlock['align'] = 'left'
  if (!leftAligned && rightAligned) align = 'right'
  else if (!leftAligned && !rightAligned && centerAligned) align = 'center'
  // justified prose (P16 G): every NON-FINAL line ends on one flush right
  // edge — a greedy ragged wrap never lines its rights up this tightly. The
  // last line is free (it may also be full). Two non-last lines minimum.
  else if (leftAligned && rightAligned === false && lines.length >= JUSTIFY_MIN_LINES) {
    const nonLast = rights.slice(0, -1)
    if (
      spread(nonLast) <= JUSTIFY_RIGHT_TOL_EMS * fontSize &&
      rights[rights.length - 1]! <= Math.max(...nonLast) + tol
    ) {
      align = 'justify'
    }
  }

  let firstLineIndentPt = 0
  if ((align === 'left' || align === 'justify') && leftAligned) {
    const paraLeft = Math.min(...bodyLefts)
    const indent = lines[0]!.box.x0 - paraLeft
    if (indent > 0.5 * fontSize && indent < INDENT_MAX_EMS * fontSize) firstLineIndentPt = indent
  }
  return { align, firstLineIndentPt }
}

// ── intra-paragraph hard breaks (P7) ──
/** the short line must leave at least this many ems free beyond the tolerance */
const HARD_BREAK_MIN_LEFTOVER_EMS = 1
/** …and the next line's first word must fit with this much word-gap slack (ems) */
const HARD_BREAK_WORD_SLACK_EMS = 0.5
/**
 * …and the line must be CLEARLY short: at least this share of the available
 * width left free. Next-word-fit alone misfires on non-greedy typesetting
 * (ragged TeX, hand-placed lines); a spaced-out title / address line stops
 * far before the edge, a prose wrap does not.
 */
const HARD_BREAK_MIN_LEFTOVER_RATIO = 0.4
/**
 * Zone-local pass (P11 b): slide layouts drop a text stack into a zone far
 * narrower than the column frame, so the ratio guard above (judged against
 * the whole frame) never fires and a four-item list rebuilds as one flowing
 * line. Inside the group's own extent the greedy-wrap argument still holds —
 * a wrap never leaves the next line's first word room — but it only applies
 * to display stacks, marked by open leading (baseline pitch well past the
 * font size); dense prose keeps the strict frame-ratio path.
 */
const HARD_BREAK_ZONE_PITCH_RATIO = 1.35
/** slide stacks at most this many lines pin all their open-leaded breaks (P14 A) */
const PIN_STACK_MAX_LINES = 3

// ── display-heading pass (P12 B) ──
/** every line of the group must be at least this big to count as display type */
const HEADING_BREAK_MIN_SIZE_PT = 24
/** …and the stack short enough to be a title, not big-print prose */
const HEADING_BREAK_MAX_LINES = 3

/**
 * Width of the next line's first wrap unit: the text up to the first space
 * (character-proportional share of the first span's box), or a single
 * character for no-space scripts where any character may wrap.
 */
function firstWordWidthPt(line: Line): number {
  const span = line.spans[0]
  if (!span) return 0
  const spanW = span.box.x1 - span.box.x0
  const text = span.text.trim()
  const chars = Math.max(1, [...text].length)
  if (isNoSpaceScript(span.script)) return spanW / chars
  const spaceIdx = text.indexOf(' ')
  if (spaceIdx <= 0) return spanW
  return spanW * (spaceIdx / chars)
}

/**
 * Mark intentional line breaks inside one paragraph: a line that stops well
 * short of the wrap edge even though the NEXT line's first word would still
 * have fit did not auto-wrap — the author broke it (spaced-out titles,
 * addresses, poetry). Wrap edges come from the surrounding column's text
 * extent; the judged edge follows the paragraph's alignment.
 */
function markHardBreaks(
  group: Line[],
  align: TextBlock['align'],
  wrapLeft: number,
  wrapRight: number,
  pinOpenLeaded: boolean,
): Line[] {
  if (group.length < 2) return group
  const groupLeft = Math.min(...group.map((l) => l.box.x0))
  const groupRight = Math.max(...group.map((l) => l.box.x1))
  // display-heading pass (P12 B): a short stack of ≥24pt lines is a title
  // whose breaks the author placed — substitute fonts run wider, so a natural
  // re-wrap lands in an ugly spot even when the source line reaches the wrap
  // edge (leftover ≈ 0 defeats the leftover gates below). Pin every boundary.
  const displayHeading =
    group.length <= HEADING_BREAK_MAX_LINES &&
    group.every((l) => lineFontSize(l) >= HEADING_BREAK_MIN_SIZE_PT)
  return group.map((line, i) => {
    if (i === 0) return line
    const prev = group[i - 1]!
    if (prev.endsWithHyphen) return line // hyphenation is a soft wrap by definition
    if (displayHeading) return { ...line, hardBreakBefore: true }
    const fontSize = lineFontSize(prev)
    let leftover: number
    if (align === 'right') leftover = prev.box.x0 - wrapLeft
    else if (align === 'center') {
      leftover = wrapRight - wrapLeft - (prev.box.x1 - prev.box.x0)
    } else leftover = wrapRight - prev.box.x1
    const need = firstWordWidthPt(line) + HARD_BREAK_WORD_SLACK_EMS * fontSize
    if (
      leftover > HARD_BREAK_MIN_LEFTOVER_EMS * fontSize &&
      leftover >= HARD_BREAK_MIN_LEFTOVER_RATIO * (wrapRight - wrapLeft) &&
      leftover >= need
    ) {
      return { ...line, hardBreakBefore: true }
    }
    // zone-local pass (P11 b): display stacks judged in the group's own extent
    const openLeaded =
      Math.abs(prev.baseline - line.baseline) >= HARD_BREAK_ZONE_PITCH_RATIO * fontSize
    // slide pages pin every open-leaded break of a SHORT stack outright
    // (P14 A): their sections hug the content, so the frame gates above are
    // blind to hand-placed breaks (the longest line defines its own frame
    // edge) and a title/quote pair would fuse into one line at the section's
    // rebuild width. Longer groups are greedy-wrapped prose even on slides —
    // pinning those wraps overflows the page once the substitute font runs
    // wider — so they keep the leftover-evidence path below.
    if (pinOpenLeaded && openLeaded && group.length <= PIN_STACK_MAX_LINES) {
      return { ...line, hardBreakBefore: true }
    }
    let zoneLeftover: number
    if (align === 'right') zoneLeftover = prev.box.x0 - groupLeft
    else if (align === 'center') {
      zoneLeftover = groupRight - groupLeft - (prev.box.x1 - prev.box.x0)
    } else zoneLeftover = groupRight - prev.box.x1
    if (
      openLeaded &&
      zoneLeftover > HARD_BREAK_MIN_LEFTOVER_EMS * fontSize &&
      zoneLeftover >= need
    ) {
      return { ...line, hardBreakBefore: true }
    }
    return line
  })
}

export interface BlockOptions {
  /**
   * pin EVERY open-leaded line break as a hard break (P14 A). Slide pages
   * pass true: their layouts are hand-placed, their sections hug the content
   * (so the wrap-frame gates cannot fire), and rebuild widths far beyond the
   * content extent would otherwise fuse the stack into one line.
   */
  pinOpenLeadedBreaks?: boolean
}

// ── text-level hyphenation detection (P21 B) ──
// PDFium's end-of-line hyphen flag misses plenty of real hyphenations (the
// rebuilt text then reads "ex pected"). Conservative text fallback: a line
// ending in a hyphen after a Latin letter whose NEXT line starts with a
// lowercase Latin letter was wrapped mid-word — join it seamlessly. A soft
// hyphen (U+00AD) exists ONLY to mark a wrap point and always joins. An
// uppercase/digit continuation stays as-is (likely a compound or a list).
const lineText = (line: Line): string => line.spans.map((s) => s.text).join('')
const LATIN_THEN_HYPHEN = /\p{Script=Latin}[-‐]$/u
const SOFT_HYPHEN_END = /­$/
const LOWER_LATIN = /^\p{Ll}$/u

function isTextHyphenation(prev: Line, next: Line): boolean {
  const prevText = lineText(prev).trimEnd()
  if (SOFT_HYPHEN_END.test(prevText)) return true
  if (!LATIN_THEN_HYPHEN.test(prevText)) return false
  const first = [...lineText(next).trimStart()][0]
  return first !== undefined && LOWER_LATIN.test(first) && /\p{Script=Latin}/u.test(first)
}

function markTextHyphenation(lines: readonly Line[]): Line[] {
  return lines.map((line, i) => {
    const next = lines[i + 1]
    if (line.endsWithHyphen || !next || !isTextHyphenation(line, next)) return line
    return { ...line, endsWithHyphen: true }
  })
}

// ── short-line (verse) runs (P21 C) ──
// Poems, lyrics and addresses are stacks of author-broken lines far narrower
// than their column; joined into a flowing paragraph they re-wrap at the
// rebuild width and every verse smears into the next. The P7/P11 leftover
// paths miss them: tight leading fails the zone gate, and the frame-ratio
// gate compares against the group's own extent (the widest verse line).
// Detection is run-level: enough consecutive left-aligned tight lines whose
// right edges are RAGGED (prose is flush — greedy wraps and justification
// both fill the measure) and where the next line's first word would have fit
// in the leftover (greedy wrap never leaves that gap). Qualified runs pin a
// hard break at every internal boundary.
/** minimum lines that make a verse stack */
const VERSE_MIN_LINES = 3
/** share of lines allowed to reach (within 1 em of) the run's right extent —
 * "saturated" stacks are wrapped prose, not verse */
const VERSE_MAX_NEAR_FULL_SHARE = 1 / 3
/** share of boundaries that must show greedy-wrap violation evidence */
const VERSE_MIN_EVIDENCE_SHARE = 0.5
/** verse pitch stays tight — larger baseline steps are separate stacks */
const VERSE_MAX_PITCH_EMS = 1.9
/** left edges must agree within this (ems) for lines to share a run */
const VERSE_LEFT_TOL_EMS = 0.5

/** consecutive lines a→b can extend one verse run */
function verseChain(a: Line, b: Line): boolean {
  if (a.endsWithHyphen) return false // a hyphen wrap is prose by definition
  const sizeA = lineFontSize(a)
  const sizeB = lineFontSize(b)
  if (Math.max(sizeA, sizeB) / Math.max(1, Math.min(sizeA, sizeB)) >= SIZE_BREAK_RATIO) {
    return false
  }
  const fontSize = Math.max(sizeA, sizeB)
  if (Math.abs(a.box.x0 - b.box.x0) > VERSE_LEFT_TOL_EMS * fontSize) return false
  const pitch = a.baseline - b.baseline
  if (pitch <= 0 || pitch > VERSE_MAX_PITCH_EMS * fontSize) return false
  const mixed = (l: Line): boolean =>
    l.spans.some((s) => isNoSpaceScript(s.script) || s.dir === 'rtl')
  return !mixed(a) && !mixed(b)
}

// second signal: verse lines end on punctuation nearly every line (each verse
// is a clause); prose wraps mid-sentence, so its NON-FINAL line ends rarely
// carry any. Longer minimum run + a looser saturation cap keep 3-line prose
// paragraphs that happen to end two lines on commas out.
const VERSE_PUNCT_MIN_LINES = 4
const VERSE_PUNCT_MIN_SHARE = 0.8
const VERSE_PUNCT_MAX_NEAR_FULL_SHARE = 1 / 2
const VERSE_EOL_PUNCT = /[.,;:!?…—–"'"'»«)\]]\s*$/u

/** the run [start..end] (inclusive) reads as verse — mark its boundaries */
function isVerseRun(run: Line[]): boolean {
  if (run.length < VERSE_MIN_LINES) return false
  const fontSize = median(run.map(lineFontSize)) || 12
  const maxRight = Math.max(...run.map((l) => l.box.x1))
  const nearFullShare = run.filter((l) => l.box.x1 >= maxRight - fontSize).length / run.length
  let evidence = 0
  for (let i = 0; i + 1 < run.length; i++) {
    const leftover = maxRight - run[i]!.box.x1
    const need = Math.max(
      HARD_BREAK_MIN_LEFTOVER_EMS * fontSize,
      firstWordWidthPt(run[i + 1]!) + HARD_BREAK_WORD_SLACK_EMS * fontSize,
    )
    if (leftover >= need) evidence++
  }
  if (
    nearFullShare <= VERSE_MAX_NEAR_FULL_SHARE &&
    evidence / (run.length - 1) >= VERSE_MIN_EVIDENCE_SHARE
  ) {
    return true
  }
  if (run.length < VERSE_PUNCT_MIN_LINES || nearFullShare > VERSE_PUNCT_MAX_NEAR_FULL_SHARE) {
    return false
  }
  const nonFinal = run.slice(0, -1)
  const punct = nonFinal.filter((l) => VERSE_EOL_PUNCT.test(lineText(l).trimEnd())).length
  return punct / nonFinal.length >= VERSE_PUNCT_MIN_SHARE
}

/**
 * A whole stack of lines reads as one verse run (P22 E) — exported for the
 * stream-table detector: side-by-side stanzas row-cluster exactly like a
 * 2-column table (flush-left, aligned baselines, wide gutter) and must be
 * vetoed BEFORE the table consumes them, because verse marking runs after.
 */
export function isVerseStack(lines: readonly Line[]): boolean {
  if (lines.length < VERSE_MIN_LINES) return false
  for (let i = 0; i + 1 < lines.length; i++) {
    if (!verseChain(lines[i]!, lines[i + 1]!)) return false
  }
  return isVerseRun([...lines])
}

/** hard breaks pinned inside detected verse runs; other lines pass through */
function markVerseRuns(lines: readonly Line[]): Line[] {
  const out = [...lines]
  let start = 0
  while (start < out.length) {
    let end = start
    while (end + 1 < out.length && verseChain(out[end]!, out[end + 1]!)) end++
    const run = out.slice(start, end + 1)
    if (isVerseRun(run)) {
      for (let i = start + 1; i <= end; i++) {
        out[i] = { ...out[i]!, hardBreakBefore: true }
      }
    }
    start = end + 1
  }
  return out
}

/** Cluster top-down ordered lines into paragraph blocks. */
export function groupIntoBlocks(
  rawLines: readonly Line[],
  body?: BodyContext,
  options: BlockOptions = {},
): TextBlock[] {
  if (rawLines.length === 0) return []
  const lines = markVerseRuns(markTextHyphenation(rawLines))
  const pinOpenLeaded = options.pinOpenLeadedBreaks ?? false
  const ctx = body ?? bodyContextOf(lines)

  const gaps: number[] = []
  for (let i = 1; i < lines.length; i++) {
    const gap = lines[i - 1]!.box.y0 - lines[i]!.box.y1
    if (gap > 0) gaps.push(gap)
  }
  const medianGap = median(gaps)
  const medianLineH = median(lines.map((l) => l.box.y1 - l.box.y0)) || 12

  // wrap edges for the short-item and hard-break judgments: the body edge
  // tightened to the real text extent — the page-level mirrored bodyRight
  // must not read every line of a narrow layout as "short"
  const wrapRight = Math.min(ctx.bodyRight, Math.max(...lines.map((l) => l.box.x1)))
  const wrapLeft = Math.max(ctx.bodyLeft, Math.min(...lines.map((l) => l.box.x0)))
  const wrap = { left: wrapLeft, right: wrapRight }

  const grouped: Line[][] = []
  let current: Line[] = [lines[0]!]
  for (let i = 1; i < lines.length; i++) {
    const prev = lines[i - 1]!
    const cur = lines[i]!
    if (isParagraphBreak(prev, cur, medianGap, medianLineH, ctx)) {
      grouped.push(current)
      current = []
    }
    current.push(cur)
  }
  grouped.push(current)
  // marker-less list items merged by the gap rules split back apart (P8 C)
  const groups = grouped.flatMap((g) => splitRunOnItems(g, wrap))

  return groups.flatMap((group) => {
    const format = inferFormat(group, ctx)
    const dir = firstStrongDir(group)
    const marked =
      dir === 'ltr'
        ? markHardBreaks(group, format.align, wrapLeft, wrapRight, pinOpenLeaded)
        : group
    // a justified paragraph must not carry w:br hard breaks — Word/LO
    // stretch the line BEFORE an intra-paragraph break to the full column,
    // filling the author's deliberate hole. Split into real paragraphs
    // instead: a paragraph end never justifies its last line (P16 G).
    const pieces: Line[][] = []
    if (format.align === 'justify') {
      let piece: Line[] = []
      for (const line of marked) {
        if (line.hardBreakBefore && piece.length > 0) {
          pieces.push(piece)
          piece = []
        }
        piece.push(line.hardBreakBefore ? { ...line, hardBreakBefore: false } : line)
      }
      if (piece.length > 0) pieces.push(piece)
    } else {
      pieces.push(marked)
    }
    return pieces.map((piece) => ({
      kind: 'text' as const,
      lines: piece,
      box: rectUnionAll(piece.map((l) => l.box)),
      align: format.align,
      firstLineIndentPt: piece === pieces[0] ? format.firstLineIndentPt : 0,
      dir,
    }))
  })
}

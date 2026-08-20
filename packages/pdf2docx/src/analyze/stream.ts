/**
 * Stream (borderless) table detection — pdf2docx rule 3 with pdfplumber's
 * mixed-evidence idea. Candidate regions are consecutive multi-segment rows;
 * column boundaries come from an x-projection (XY-Cut variant); sparse
 * strokes (booktabs rules) and fills (row banding) raise confidence.
 *
 * The bias is MISS RATHER THAN MISFIRE (plan §7 risk 2): aligned running
 * text, page columns, TOC leaders and poetry must all be rejected, so a
 * candidate has to clear every hard gate below AND a confidence threshold.
 * The confidence lands in the Table IR for P4's low-confidence downgrade.
 */
import type { Interval, Rect } from '../geometry'
import { intersectArea, median, mergeIntervals, rectUnion, rectUnionAll } from '../geometry'
import type { Fill, PageShapes, TableBlock, TableCellBlock } from '../ir'
import { isNearWhite } from './shapes'
import { groupIntoBlocks, isVerseStack } from './blocks'
import { buildStackCell } from './zones'
import { analyzeChars } from './chars'
import type { LineUnit, UnitRow } from './units'
import { clusterUnitRows } from './units'

// ── tunable gates (spec: constants, commented) ──

/** minimum grid: anything smaller is never a table */
const STREAM_MIN_ROWS = 2
const STREAM_MIN_COLS = 2
/** every column's dominant edge must align on at least this share of rows */
const COL_ALIGN_MIN_ROW_RATIO = 0.6
/** column edges within this many ems count as aligned */
const ALIGN_TOL_EMS = 0.4
/** an x-projection valley narrower than this many ems is intra-column noise */
const VALLEY_MIN_EMS = 0.4
/** a vertical gap over this multiple of the taller row's height breaks the region */
const ROW_GAP_MAX_RATIO = 1.6
/** median cell width ÷ median gap width above this = prose, not a table */
const CONTENT_GAP_RATIO_MAX = 10
/** columns whose text fills them this consistently are running paragraphs… */
const PARA_FILL_MIN = 0.85
/** …when the entries are also sentence-length (in ems — script-neutral) … */
const PROSE_WIDTH_EMS = 10
/** …or average this many words per cell (CJK counts one word per char) */
const SENTENCE_WORDS_MIN = 5
/** region holding this share of the page's units triggers the page-columns test */
const PAGE_COVER_RATIO = 0.7
/** page-columns veto also needs this many words per cell — prose is wordy (P27) */
const PAGE_COLUMN_MIN_WORDS = 4
/** row-pitch median absolute deviation below this share of the median = uniform */
const PITCH_MAD_RATIO = 0.2
/** minimum confidence to accept a candidate */
const STREAM_CONF_MIN = 0.5
/** stroke/fill evidence must span this share of the region width to count */
const EVIDENCE_SPAN_MIN = 0.5

/** one row's content merged per column */
interface CellEntry {
  units: LineUnit[]
  box: Rect
}

interface Candidate {
  rows: UnitRow[]
  box: Rect
  /** column x intervals, left → right */
  columns: Interval[]
  /** [row][col] → entry or null */
  entries: Array<Array<CellEntry | null>>
  fontSize: number
}

/** a narrow-row (wrapped-label) absorption candidate stays under this share of the run width */
const ABSORB_MAX_WIDTH_RATIO = 0.5
/** …and inside the run's x-range with this much play (pt) */
const ABSORB_X_TOL_PT = 2

/** a sub-minimum row that is a wrapped cell continuation of the active run */
function rowFitsRun(run: UnitRow[], row: UnitRow): boolean {
  const x0 = Math.min(...run.map((r) => r.box.x0))
  const x1 = Math.max(...run.map((r) => r.box.x1))
  if (row.box.x1 - row.box.x0 > ABSORB_MAX_WIDTH_RATIO * (x1 - x0)) return false
  return row.units.every(
    (u) => u.box.x0 >= x0 - ABSORB_X_TOL_PT && u.box.x1 <= x1 + ABSORB_X_TOL_PT,
  )
}

/** unit count that marks a row as strongly tabular for the sparse-leading rules (P27) */
const SPARSE_MIN_UNITS = 5
/** gap ceiling (× row height) between two strongly tabular rows */
const SPARSE_GAP_MAX_RATIO = 3.5
/** an established run continues while the row-top pitch stays within this ratio */
const PITCH_CONT_RATIO = 1.35

/** consecutive multi-unit rows (vertical gaps within reason) form candidates */
function findCandidates(
  rows: UnitRow[],
  absorbNarrowRows = false,
  maxRowGapRatio = ROW_GAP_MAX_RATIO,
  strongRunsOnly = false,
): UnitRow[][] {
  const runs: UnitRow[][] = []
  let run: UnitRow[] = []
  const flush = (): void => {
    // absorbed sub-minimum rows may not END a run — a table never finishes
    // on a wrapped-label fragment, and trailing absorbs would glue prose
    while (run.length > 0 && run[run.length - 1]!.units.length < STREAM_MIN_COLS) run.pop()
    if (run.length >= STREAM_MIN_ROWS) runs.push(run)
  }
  for (const row of rows) {
    const prev = run[run.length - 1]
    const rowH = row.box.y1 - row.box.y0
    // a strong run is unmistakably tabular — its typical row splits into many
    // units. Only such runs earn the relaxed continuation rules; a run seeded
    // by 2-3-unit prose fragments must never glue itself into a table below
    // (health.pdf: title rows + absorbed prose poisoned the whole candidate)
    const runStrong = run.length > 0 && median(run.map((r) => r.units.length)) >= SPARSE_MIN_UNITS
    let gapOk = false
    if (prev !== undefined) {
      const gap = prev.box.y0 - row.box.y1
      const refH = Math.max(prev.box.y1 - prev.box.y0, rowH)
      gapOk = gap <= maxRowGapRatio * refH
      // sparse-leading tables (P27): journal grids print 7pt digit rows at a
      // 20pt+ pitch, so the height-relative gap never holds. Two strongly
      // tabular rows still pair up under a wider ceiling…
      if (!gapOk && row.units.length >= SPARSE_MIN_UNITS && prev.units.length >= SPARSE_MIN_UNITS) {
        gapOk = gap <= SPARSE_GAP_MAX_RATIO * refH
      }
      // …and an established STRONG run keeps going while the row-top pitch
      // stays steady (wrapped-label rows sit between value rows at the same
      // rhythm)
      if (!gapOk && run.length >= 2 && runStrong) {
        const pitches: number[] = []
        for (let i = 1; i < run.length; i++) pitches.push(run[i - 1]!.box.y1 - run[i]!.box.y1)
        const pitch = median(pitches)
        gapOk = pitch > 0 && prev.box.y1 - row.box.y1 <= PITCH_CONT_RATIO * pitch
      }
    }
    if (row.units.length >= STREAM_MIN_COLS && (run.length === 0 || gapOk)) {
      run.push(row)
      continue
    }
    // multi-line cells (a two-line row label) surface as sub-minimum rows
    // BETWEEN the grid's value rows; inside an active run they continue it
    // when they fit the run's x-range (P16 E; base pass since P27 with the
    // strong-run guard — the flush() pop keeps trailing absorbs from
    // gluing prose)
    if (
      absorbNarrowRows &&
      run.length > 0 &&
      gapOk &&
      (!strongRunsOnly || runStrong) &&
      rowFitsRun(run, row)
    ) {
      run.push(row)
      continue
    }
    flush()
    run = row.units.length >= STREAM_MIN_COLS ? [row] : []
  }
  flush()
  return runs
}

/** how deep the failed-solve run subdivision may recurse (P27) */
const RUN_SPLIT_MAX_DEPTH = 2

/**
 * Solve a run into candidates, subdividing on failure (P27): the relaxed
 * continuation rules can glue two stacked tables with different column sets
 * into one run — their union bridges every x-valley and the solve collapses.
 * Splitting at the widest row gap recovers each table separately. Absorbed
 * sub-minimum rows never lead or trail a piece.
 */
function collectCandidates(run: UnitRow[], out: Candidate[], depth: number): void {
  let lo = 0
  let hi = run.length
  while (lo < hi && run[lo]!.units.length < STREAM_MIN_COLS) lo++
  while (hi > lo && run[hi - 1]!.units.length < STREAM_MIN_COLS) hi--
  const piece = run.slice(lo, hi)
  if (piece.length < STREAM_MIN_ROWS) return
  const cand = solveColumns(piece)
  if (cand) {
    out.push(cand)
    return
  }
  if (depth >= RUN_SPLIT_MAX_DEPTH || piece.length < 2 * STREAM_MIN_ROWS) return
  let cut = -1
  let widest = -Infinity
  for (let i = 1; i < piece.length; i++) {
    const gap = piece[i - 1]!.box.y0 - piece[i]!.box.y1
    if (gap > widest) {
      widest = gap
      cut = i
    }
  }
  if (cut <= 0) return
  collectCandidates(piece.slice(0, cut), out, depth + 1)
  collectCandidates(piece.slice(cut), out, depth + 1)
}

/** solve a candidate's columns by x-projection; null when under 2 survive */
function solveColumns(rows: UnitRow[]): Candidate | null {
  const units = rows.flatMap((r) => r.units)
  const box = rectUnionAll(units.map((u) => u.box))
  const fontSize = median(units.map((u) => u.fontSize)) || 12
  const columns = mergeIntervals(
    units.map((u) => ({ lo: u.box.x0, hi: u.box.x1 })),
    VALLEY_MIN_EMS * fontSize,
  )
  if (columns.length < STREAM_MIN_COLS) return null

  const entries: Candidate['entries'] = rows.map((row) => {
    const per: Array<CellEntry | null> = columns.map(() => null)
    for (const unit of row.units) {
      const cx = (unit.box.x0 + unit.box.x1) / 2
      const col = columns.findIndex((c) => cx >= c.lo && cx <= c.hi)
      if (col < 0) continue
      const entry = per[col]
      if (entry) {
        entry.units.push(unit)
        entry.box = rectUnion(entry.box, unit.box)
      } else {
        per[col] = { units: [unit], box: unit.box }
      }
    }
    return per
  })
  return { rows, box, columns, entries, fontSize }
}

/** share of PRESENT entries whose best edge (left/right/center) aligns (P27:
 * the denominator was all rows, so one sparse column — a header straddling
 * two body columns, a stray footnote mark — vetoed entire clean grids like
 * camelot mexican_towns at align 1/44; a column with a single entry has no
 * alignment evidence either way and stays neutral) */
function columnAlignRatio(cand: Candidate, col: number): number {
  const boxes = cand.entries.map((row) => row[col]?.box ?? null)
  const present = boxes.filter((b): b is Rect => b !== null)
  if (present.length < 2) return 1
  const tol = ALIGN_TOL_EMS * cand.fontSize
  const hits = (edge: (b: Rect) => number): number => {
    const m = median(present.map(edge))
    return present.filter((b) => Math.abs(edge(b) - m) <= tol).length
  }
  const best = Math.max(
    hits((b) => b.x0),
    hits((b) => b.x1),
    hits((b) => (b.x0 + b.x1) / 2),
  )
  return best / present.length
}

/** per-entry means: column fill ratio, width in ems, word count, full-line share */
function entryStats(cand: Candidate): {
  fill: number
  widthEms: number
  words: number
  fullShare: number
} {
  let fillSum = 0
  let widthSum = 0
  let wordSum = 0
  let full = 0
  let n = 0
  for (const row of cand.entries) {
    for (const [col, entry] of row.entries()) {
      if (!entry) continue
      const colWidth = Math.max(cand.columns[col]!.hi - cand.columns[col]!.lo, 1)
      const width = entry.box.x1 - entry.box.x0
      fillSum += width / colWidth
      if (width / colWidth >= PARA_FILL_MIN) full++
      widthSum += width / Math.max(cand.fontSize, 1)
      wordSum += entry.units.reduce((s, u) => s + u.wordCount, 0)
      n++
    }
  }
  if (n === 0) return { fill: 0, widthEms: 0, words: 0, fullShare: 0 }
  return { fill: fillSum / n, widthEms: widthSum / n, words: wordSum / n, fullShare: full / n }
}

/** at least this share of full-width entries also reads as prose columns */
const PARA_FULL_SHARE_MIN = 0.5

/**
 * Paragraph-like columns: text consistently fills the column AND entries are
 * sentence-length. The fill ratio alone is meaningless for short cells (the
 * column width derives from the entries themselves), hence the width gate.
 */
function isParagraphLike(cand: Candidate): boolean {
  const s = entryStats(cand)
  return s.fill >= PARA_FILL_MIN && (s.widthEms >= PROSE_WIDTH_EMS || s.words >= SENTENCE_WORDS_MIN)
}

/**
 * The MEAN fill misses prose whose ragged paragraph-final and heading lines
 * dilute it (P23: two page-column tail bands beside a footer row formed
 * phantom tables, killed the section sweep and split the page) — a majority
 * of individually full sentence-length entries is prose all the same. Runs
 * AFTER the verse check: stanzas are majority-full too, and their exact
 * per-line rebuild is the correct outcome (P22 E).
 */
function isMajorityFullProse(cand: Candidate): boolean {
  const s = entryStats(cand)
  if (s.widthEms < PROSE_WIDTH_EMS && s.words < SENTENCE_WORDS_MIN) return false
  return s.fullShare >= PARA_FULL_SHARE_MIN
}

/**
 * Whole-page column layout (the classic pdf2docx "fake table"): the candidate
 * holds (nearly) all page text, rows tick at a uniform pitch, and the cells
 * read as running lines — that is a column layout for the XY-Cut stage, not a
 * table.
 */
function isPageColumns(cand: Candidate, totalUnits: number): boolean {
  const candUnits = cand.rows.reduce((s, r) => s + r.units.length, 0)
  if (candUnits / Math.max(totalUnits, 1) < PAGE_COVER_RATIO) return false
  const pitches: number[] = []
  for (let i = 1; i < cand.rows.length; i++) {
    pitches.push(cand.rows[i - 1]!.baseline - cand.rows[i]!.baseline)
  }
  const pitch = median(pitches)
  if (pitch <= 0) return false
  const mad = median(pitches.map((p) => Math.abs(p - pitch)))
  if (mad > PITCH_MAD_RATIO * pitch) return false
  // looser sentence gate than isParagraphLike: dense full-page coverage means
  // even moderately long lines are body text, not cells. Running text is also
  // WORDY — a full-page data table whose cells are phrase-length (2-3 words)
  // must not read as newspaper columns (P27: tabula us-026, camelot issue_288)
  const s = entryStats(cand)
  return s.widthEms >= 0.8 * PROSE_WIDTH_EMS && s.words >= PAGE_COLUMN_MIN_WORDS
}

/** verse columns must read as text, not table entries — median words per cell */
const VERSE_COL_MIN_WORDS = 3

/**
 * Side-by-side stanzas (P22 E): a poem pair row-clusters exactly like a
 * 2-column table — flush-left stacks, aligned baselines, one wide gutter —
 * and P21's verse marking runs AFTER the stream pass, so the fake table
 * would consume the stanzas first. Veto a candidate when EVERY column reads
 * as one verse stack (ragged rights + greedy-wrap violations, or clause
 * punctuation) of sentence-like lines. A data table never trips this: its
 * short-entry columns fail the word gate and its cells fail the chain.
 */
/** orphan rows this close (ems) above/below a verse pair belong to a stanza */
const VERSE_ORPHAN_PITCH_EMS = 1.9

/**
 * Pull stanza lines the row clustering dropped from the run back into a
 * verse candidate (P22 E): the longest verse line often stands alone in its
 * row (the other stanza is shorter), and a hair of pitch jitter pushes it
 * past the run-gap gate — it then falls back into the flow and renders at
 * the LEFT margin below the table (poems' lone "Mein Sohn…").
 */
function absorbVerseOrphans(cand: Candidate, allRows: readonly UnitRow[]): void {
  const inRun = new Set(cand.rows)
  const maxGap = VERSE_ORPHAN_PITCH_EMS * cand.fontSize
  const xs: number[] = [cand.box.x0]
  for (let c = 1; c < cand.columns.length; c++) {
    xs.push((cand.columns[c - 1]!.hi + cand.columns[c]!.lo) / 2)
  }
  xs.push(cand.box.x1)
  const fits = (row: UnitRow): boolean =>
    row.units.length <= cand.columns.length &&
    row.units.every((u) => {
      const cx = (u.box.x0 + u.box.x1) / 2
      return cx >= cand.box.x0 - cand.fontSize && cx <= cand.box.x1 + cand.fontSize
    })
  let changed = true
  while (changed) {
    changed = false
    const top = cand.rows[0]!.baseline
    const bottom = cand.rows[cand.rows.length - 1]!.baseline
    for (const row of allRows) {
      if (inRun.has(row)) continue
      const gapAbove = row.baseline - top
      const gapBelow = bottom - row.baseline
      const adjacent = (gapAbove > 0 && gapAbove <= maxGap) || (gapBelow > 0 && gapBelow <= maxGap)
      if (!adjacent || !fits(row)) continue
      cand.rows.push(row)
      inRun.add(row)
      cand.box = rectUnion(cand.box, row.box)
      cand.rows.sort((a, b) => b.baseline - a.baseline)
      changed = true
    }
  }
}

/** one row of exact-geometry stack cells (verse pair) — see isVerseColumns */
function buildVerseTable(cand: Candidate, confidence: number): TableBlock {
  const xs: number[] = [cand.box.x0]
  for (let c = 1; c < cand.columns.length; c++) {
    xs.push((cand.columns[c - 1]!.hi + cand.columns[c]!.lo) / 2)
  }
  xs.push(cand.box.x1)
  const height = cand.box.y1 - cand.box.y0
  // re-assign units by the grid midpoints, NOT the row-clustered entries:
  // the widest verse line can straddle a column interval and land in the
  // wrong cell (poems' lone "Mein Sohn…" filed into the left stanza)
  const all = cand.rows.flatMap((r) => r.units)
  const cells = cand.columns.map((_, c) => {
    const units = all.filter((u) => {
      const cx = (u.box.x0 + u.box.x1) / 2
      return cx >= xs[c]! && cx < xs[c + 1]! + (c === cand.columns.length - 1 ? 1 : 0)
    })
    const box: Rect = { x0: xs[c]!, x1: xs[c + 1]!, y0: cand.box.y0, y1: cand.box.y1 }
    // stanzas top-anchor (measured top gap); centering a short stanza would
    // sink it below its partner's shared baselines
    return buildStackCell(units, box, height, { allowVAlignCenter: false })
  })
  return {
    kind: 'table',
    box: cand.box,
    colWidthsPt: xs.slice(1).map((x, i) => x - xs[i]!),
    rows: [cells],
    confidence,
  }
}

function isVerseColumns(cand: Candidate): boolean {
  return cand.columns.every((_, c) => {
    const entries = cand.entries
      .map((row) => row[c])
      .filter((e): e is CellEntry => e !== null && e !== undefined)
    if (entries.length < cand.rows.length) return false // holes = table, not stanza
    const words = median(entries.map((e) => e.units.reduce((s, u) => s + u.wordCount, 0)))
    if (words < VERSE_COL_MIN_WORDS) return false
    const chars = entries.flatMap((e) => e.units.flatMap((u) => u.chars))
    const lines = analyzeChars(chars).sort((a, b) => b.baseline - a.baseline)
    return isVerseStack(lines)
  })
}

interface Evidence {
  stroke: boolean
  fill: boolean
}

/** booktabs rules / row banding inside the region raise confidence */
function findEvidence(cand: Candidate, shapes: PageShapes): Evidence {
  const rowH = median(cand.rows.map((r) => r.box.y1 - r.box.y0)) || cand.fontSize
  const y0 = cand.box.y0 - rowH
  const y1 = cand.box.y1 + rowH
  const width = cand.box.x1 - cand.box.x0
  const stroke = shapes.strokes.some((s) => {
    if (s.orientation !== 'h') return false
    const cy = (s.box.y0 + s.box.y1) / 2
    if (cy < y0 || cy > y1) return false
    const span = Math.min(s.box.x1, cand.box.x1) - Math.max(s.box.x0, cand.box.x0)
    return span >= EVIDENCE_SPAN_MIN * width
  })
  const fill = shapes.fills.some((f) => {
    if (f.box.y1 > y1 || f.box.y0 < y0 - rowH) return false
    if (f.box.y1 - f.box.y0 > 2.5 * rowH) return false // page background, not banding
    const span = Math.min(f.box.x1, cand.box.x1) - Math.max(f.box.x0, cand.box.x0)
    return span >= EVIDENCE_SPAN_MIN * width
  })
  return { stroke, fill }
}

/**
 * Confidence: base for clearing the hard gates, plus row depth, alignment
 * quality and vector evidence. 2-row candidates only pass WITH evidence
 * (0.3 + 0.15 alignment < 0.5); 3+ perfectly aligned rows pass on their own.
 */
function confidenceOf(cand: Candidate, minAlign: number, evidence: Evidence): number {
  let conf = 0.3
  const rows = cand.rows.length
  if (rows >= 3) conf += 0.15 + Math.min(0.1, 0.02 * (rows - 3))
  conf += (0.15 * (minAlign - COL_ALIGN_MIN_ROW_RATIO)) / (1 - COL_ALIGN_MIN_ROW_RATIO)
  if (evidence.stroke) conf += 0.2
  if (evidence.fill) conf += 0.15
  return Math.min(1, conf)
}

/** a pool fill must cover this share of a cell to become its shading */
const CELL_FILL_MIN_COVER = 0.6

/** map row-banding / per-cell fill rects onto cell shading (P16 E zebra) */
function cellFillOf(box: Rect, fills: readonly Fill[]): string | undefined {
  const cx = (box.x0 + box.x1) / 2
  const cy = (box.y0 + box.y1) / 2
  const area = Math.max((box.x1 - box.x0) * (box.y1 - box.y0), 1)
  for (const f of fills) {
    if (cx < f.box.x0 || cx > f.box.x1 || cy < f.box.y0 || cy > f.box.y1) continue
    if (isNearWhite(f.color)) continue
    if (intersectArea(box, f.box) / area < CELL_FILL_MIN_COVER) continue
    return f.color
  }
  return undefined
}

function buildTable(cand: Candidate, confidence: number, fills: readonly Fill[] = []): TableBlock {
  // grid boundaries: valley midpoints between columns, row-gap midpoints
  const xs: number[] = [cand.box.x0]
  for (let c = 1; c < cand.columns.length; c++) {
    xs.push((cand.columns[c - 1]!.hi + cand.columns[c]!.lo) / 2)
  }
  xs.push(cand.box.x1)
  const ys: number[] = [cand.box.y1]
  for (let r = 1; r < cand.rows.length; r++) {
    ys.push((cand.rows[r]!.box.y1 + cand.rows[r - 1]!.box.y0) / 2)
  }
  ys.push(cand.box.y0)

  const rows: TableCellBlock[][] = cand.entries.map((row, r) =>
    row.map((entry, c) => {
      const box: Rect = { x0: xs[c]!, x1: xs[c + 1]!, y0: ys[r + 1]!, y1: ys[r]! }
      const cell: TableCellBlock = { box, gridSpan: 1, blocks: [] }
      const fill = cellFillOf(box, fills)
      if (fill) cell.fill = fill
      if (entry) {
        const chars = entry.units
          .slice()
          .sort((a, b) => a.box.x0 - b.box.x0)
          .flatMap((u) => u.chars)
        cell.blocks = groupIntoBlocks(analyzeChars(chars))
      }
      return cell
    }),
  )
  const colWidthsPt = xs.slice(1).map((x, i) => x - xs[i]!)
  return { kind: 'table', box: cand.box, colWidthsPt, rows, confidence }
}

export interface DetectedStreamTables {
  tables: TableBlock[]
  /** units that stay in the regular flow */
  remainingUnits: LineUnit[]
}

/**
 * Full stream pass over one page's units. `latticeBoxes` are already-detected
 * lattice table regions — candidates overlapping them are rejected outright.
 */
export function detectStreamTables(
  units: readonly LineUnit[],
  shapes: PageShapes,
  latticeBoxes: readonly Rect[] = [],
  opts: StreamOptions = {},
): DetectedStreamTables {
  // wrapped-label absorption runs in the base pass too since P27 (journal
  // tables interleave 1-unit label wraps between their value rows); the
  // findCandidates flush() pop keeps trailing absorbs from gluing prose
  const base = detectPass(units, shapes, latticeBoxes, {
    absorbNarrowRows: true,
    strongRunsOnly: true,
  })
  if (!opts.slideRegions) return base

  // slide second pass (P16 E): absolute side-by-side placement interleaves a
  // data table's rows with neighboring prose lines, so the whole-page row
  // clustering never forms the candidate run. XY-cut style: detect over the
  // region; when nothing surfaces, split at the WIDEST x-valley and recurse
  // — with the wrapped-label absorption on and the stricter strong-grid
  // gate so slide text columns never mint fake tables.
  const claimed = [...latticeBoxes, ...base.tables.map((t) => t.box)]
  const region = regionPass(base.remainingUnits, shapes, claimed, 0)
  if (region.tables.length === 0) return base
  const tables = [...base.tables, ...region.tables].sort((a, b) => b.box.y1 - a.box.y1)
  return { tables, remainingUnits: region.remainingUnits }
}

export interface StreamOptions {
  /** landscape slide: rerun over vast-valley regions with strong-grid gates (P16 E) */
  slideRegions?: boolean
}

interface PassOptions {
  /** absorb narrow label-continuation rows into an active run (wrapped cells) */
  absorbNarrowRows?: boolean
  /** demand a ≥3×3 grid — slide regions must show strong grid evidence */
  strongGrid?: boolean
  /** row-gap tolerance override: spacious slide tables pad their rows far
   * beyond running-text leading (P16 E) */
  maxRowGapRatio?: number
  /** restrict absorption to strong (many-unit) runs — base pass only (P27) */
  strongRunsOnly?: boolean
}

/** slide-region row-gap tolerance (see PassOptions.maxRowGapRatio) */
const SLIDE_ROW_GAP_MAX_RATIO = 3.5

/** never split at a valley narrower than this — sub-region recursion floor */
const REGION_VALLEY_MIN_PT = 40
/** how many binary valley splits the region recursion may take */
const REGION_MAX_DEPTH = 3
const STRONG_GRID_MIN_ROWS = 3
const STRONG_GRID_MIN_COLS = 3

/** binary split at the WIDEST x-valley of the unit set (null when none qualifies) */
function splitAtWidestValley(units: readonly LineUnit[]): [LineUnit[], LineUnit[]] | null {
  if (units.length < 2) return null
  const spans = mergeIntervals(units.map((u) => ({ lo: u.box.x0, hi: u.box.x1 })))
  let bestGap = 0
  let cut = Number.NaN
  for (let i = 1; i < spans.length; i++) {
    const gap = spans[i]!.lo - spans[i - 1]!.hi
    if (gap > bestGap) {
      bestGap = gap
      cut = (spans[i - 1]!.hi + spans[i]!.lo) / 2
    }
  }
  if (bestGap < REGION_VALLEY_MIN_PT) return null
  const left: LineUnit[] = []
  const right: LineUnit[] = []
  for (const u of units) ((u.box.x0 + u.box.x1) / 2 < cut ? left : right).push(u)
  if (left.length === 0 || right.length === 0) return null
  return [left, right]
}

/** detect within a region; when nothing surfaces, recurse into its halves */
function regionPass(
  units: readonly LineUnit[],
  shapes: PageShapes,
  claimed: readonly Rect[],
  depth: number,
): DetectedStreamTables {
  const res = detectPass(units, shapes, claimed, {
    absorbNarrowRows: true,
    strongGrid: true,
    maxRowGapRatio: SLIDE_ROW_GAP_MAX_RATIO,
  })
  if (res.tables.length > 0 || depth >= REGION_MAX_DEPTH) return res
  const halves = splitAtWidestValley(units)
  if (!halves) return res
  const a = regionPass(halves[0], shapes, claimed, depth + 1)
  const b = regionPass(halves[1], shapes, claimed, depth + 1)
  return {
    tables: [...a.tables, ...b.tables],
    remainingUnits: [...a.remainingUnits, ...b.remainingUnits],
  }
}

function detectPass(
  units: readonly LineUnit[],
  shapes: PageShapes,
  latticeBoxes: readonly Rect[],
  o: PassOptions,
): DetectedStreamTables {
  const rows = clusterUnitRows(units)
  const tables: TableBlock[] = []
  const consumed = new Set<LineUnit>()

  const cands: Candidate[] = []
  for (const run of findCandidates(
    rows,
    o.absorbNarrowRows === true,
    o.maxRowGapRatio,
    o.strongRunsOnly === true,
  )) {
    collectCandidates(run, cands, 0)
  }
  for (const cand of cands) {
    if (latticeBoxes.some((b) => intersectArea(cand.box, b) > 0)) continue
    if (
      o.strongGrid &&
      (cand.rows.length < STRONG_GRID_MIN_ROWS || cand.columns.length < STRONG_GRID_MIN_COLS)
    ) {
      continue
    }

    const alignRatios = cand.columns.map((_, c) => columnAlignRatio(cand, c))
    const minAlign = Math.min(...alignRatios)
    if (minAlign < COL_ALIGN_MIN_ROW_RATIO) continue

    const entryWidths: number[] = []
    for (const row of cand.entries) {
      for (const e of row) if (e) entryWidths.push(e.box.x1 - e.box.x0)
    }
    const gapWidths = cand.columns.slice(1).map((c, i) => c.lo - cand.columns[i]!.hi)
    if (median(entryWidths) / Math.max(median(gapWidths), 0.1) > CONTENT_GAP_RATIO_MAX) continue

    if (isParagraphLike(cand)) continue
    if (isPageColumns(cand, units.length)) continue

    const evidence = findEvidence(cand, shapes)
    const confidence = confidenceOf(cand, minAlign, evidence)
    if (confidence < STREAM_CONF_MIN) continue

    // side-by-side stanzas (P22 E): the row-clustered cell assignment
    // misfiles lines whose baselines straddle a row boundary (poems' lone
    // "Mein Sohn…" landing in the left column). Rebuild as ONE row of
    // exact-geometry stack cells instead — the table keeps its measured
    // extent, each stanza stays whole, verse breaks are per-line paragraphs.
    if (isVerseColumns(cand)) {
      absorbVerseOrphans(cand, rows)
      tables.push(buildVerseTable(cand, confidence))
    } else {
      if (isMajorityFullProse(cand)) continue
      tables.push(buildTable(cand, confidence, shapes.fills))
    }
    for (const row of cand.rows) for (const u of row.units) consumed.add(u)
  }

  tables.sort((a, b) => b.box.y1 - a.box.y1)
  return { tables, remainingUnits: units.filter((u) => !consumed.has(u)) }
}

/**
 * Rule-separated side-by-side zones (P22 A). Court captions and similar
 * front-matter set two short text stacks beside a drawn vertical rule
 * (single or double line). Column detection cannot hold that zone together:
 * rows where only ONE side has ink dissolve the gutter (a gutter needs both
 * flanks inside each slab), so the section closes and the two stacks
 * interleave by baseline into one garbled column.
 *
 * The drawn rule is the reliable signal the slab projection lacks. A zone is
 * only claimed when a tall vertical rule stands between two multi-line unit
 * groups whose x extents do not cross it — then the whole zone becomes one
 * borderless 1×2 table (the rule survives as the inside-vertical border) and
 * flows through layout as a single element. MISS RATHER THAN MISFIRE: no
 * rule, no zone — gutters without ink stay the column detector's business.
 */
import type { Rect } from '../geometry'
import { intersectArea, rectUnion } from '../geometry'
import type { PageShapes, Stroke, TableBlock, TableCellBlock, TextBlock } from '../ir'
import { analyzeChars } from './chars'
import type { LineUnit } from './units'

/** rule segments this close in x merge into one band (double rules, joints) */
const RULE_X_MERGE_PT = 4
/** a rule band wider than this is a bar/fill artifact, not a separator */
const RULE_MAX_WIDTH_PT = 8
/** minimum rule height — a separator spans several lines, not one */
const RULE_MIN_HEIGHT_PT = 48
/** merged segments must ink at least this share of the band's y extent */
const RULE_Y_COVER_MIN = 0.7
/** the rule must stand in the middle band of the page, not at an edge */
const RULE_CENTER_MIN_RATIO = 0.2
const RULE_CENTER_MAX_RATIO = 0.8
/** units whose vertical center sits within the padded rule extent join the zone */
const ZONE_Y_PAD_PT = 3
/** each side needs a real stack: units and distinct baselines */
const ZONE_MIN_UNITS_PER_SIDE = 2
const ZONE_MIN_ROWS_PER_SIDE = 2
/** the two stacks must actually sit beside each other */
const ZONE_SIDE_OVERLAP_MIN = 0.4
/** a "zone" filling most of the page is a two-column layout, not a caption */
const ZONE_MAX_PAGE_HEIGHT_RATIO = 0.75
/** the rule must span most of the zone it separates */
const ZONE_RULE_SPAN_MIN = 0.6
/** a side whose content fills less of the zone than this centers vertically */
const ZONE_VALIGN_CENTER_MAX_FILL = 0.85
/** drawn-rule evidence is strong; well above the degrade threshold */
const ZONE_CONFIDENCE = 0.9
/** measured gaps between cell paragraphs survive from this size up */
const ZONE_CELL_GAP_MIN_PT = 2
/** single-line paragraphs keep left insets from this size up */
const ZONE_CELL_INSET_MIN_PT = 6

interface RuleBand {
  box: Rect
  strokes: Stroke[]
  /** ≥2 distinct line x-centers → double rule */
  double: boolean
}

export interface DetectedZones {
  tables: TableBlock[]
  remainingUnits: LineUnit[]
  /** rule strokes consumed as table inside borders (keep away from decor) */
  consumedStrokes: Set<Stroke>
}

/** merge vertical strokes into candidate separator bands */
function ruleBands(strokes: readonly Stroke[]): RuleBand[] {
  const vertical = strokes
    .filter((s) => s.orientation === 'v')
    .sort((a, b) => (a.box.x0 + a.box.x1) / 2 - (b.box.x0 + b.box.x1) / 2)
  const bands: RuleBand[] = []
  for (const s of vertical) {
    const cx = (s.box.x0 + s.box.x1) / 2
    const open = bands[bands.length - 1]
    if (open && cx - (open.box.x0 + open.box.x1) / 2 <= RULE_X_MERGE_PT) {
      open.box = rectUnion(open.box, s.box)
      open.strokes.push(s)
    } else {
      bands.push({ box: { ...s.box }, strokes: [s], double: false })
    }
  }
  for (const band of bands) {
    const centers = band.strokes.map((s) => (s.box.x0 + s.box.x1) / 2).sort((a, b) => a - b)
    band.double = centers[centers.length - 1]! - centers[0]! >= 1.5
  }
  return bands
}

/** share of the band's y extent actually inked by its segments */
function yCoverage(band: RuleBand): number {
  const spans = band.strokes
    .map((s) => ({ lo: s.box.y0, hi: s.box.y1 }))
    .sort((a, b) => a.lo - b.lo)
  let covered = 0
  let cursor = -Infinity
  for (const sp of spans) {
    const lo = Math.max(sp.lo, cursor)
    if (sp.hi > lo) covered += sp.hi - lo
    cursor = Math.max(cursor, sp.hi)
  }
  const extent = band.box.y1 - band.box.y0
  return extent > 0 ? covered / extent : 0
}

/**
 * Build one side of a side-by-side stack as a table cell with EXACT display
 * geometry. Shared with the stream detector's verse tables (P22 E).
 */
export function buildStackCell(
  units: LineUnit[],
  box: Rect,
  zoneHeight: number,
  opts: { allowVAlignCenter?: boolean } = {},
): TableCellBlock {
  if (units.length === 0) return { box, gridSpan: 1, blocks: [] }
  const chars = units
    .slice()
    .sort((a, b) => a.box.x0 - b.box.x0)
    .flatMap((u) => u.chars)
  // content-stream order interleaves the two stacks — sort lines top→down
  // (same visual ordering the column assembler applies)
  const lines = analyzeChars(chars).sort((a, b) => b.baseline - a.baseline || a.box.x0 - b.box.x0)
  // one paragraph per line: these stacks are display typography (party lists,
  // case titles) — re-wrapping them at cell width loses the author's breaks,
  // and measured insets/gaps reproduce the exact geometry instead
  const blocks: TextBlock[] = lines.map((line) => ({
    kind: 'text' as const,
    lines: [line],
    box: line.box,
    align: 'left' as const,
    firstLineIndentPt: 0,
    dir: 'ltr' as const,
  }))
  const cell: TableCellBlock = { box, gridSpan: 1, blocks }
  const contentH = Math.max(...units.map((u) => u.box.y1)) - Math.min(...units.map((u) => u.box.y0))
  if ((opts.allowVAlignCenter ?? true) && contentH < ZONE_VALIGN_CENTER_MAX_FILL * zoneHeight) {
    cell.vAlign = 'center'
  }
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i]!
    const gap = i === 0 ? (cell.vAlign ? 0 : box.y1 - b.box.y1) : blocks[i - 1]!.box.y0 - b.box.y1
    if (gap >= ZONE_CELL_GAP_MIN_PT) b.spacingBeforePt = gap
    const inset = b.box.x0 - box.x0
    if (inset >= ZONE_CELL_INSET_MIN_PT) b.firstLineIndentPt = inset
  }
  return cell
}

/**
 * Claim rule-separated side-by-side zones as borderless 1×2 tables.
 * `claimedBoxes` are lattice/form table regions — a rule inside one is a
 * table border, not a zone separator.
 */
export function detectRuleSeparatedZones(
  units: readonly LineUnit[],
  shapes: PageShapes,
  claimedBoxes: readonly Rect[],
  opts: { pageWidthPt: number; pageHeightPt: number },
): DetectedZones {
  const none: DetectedZones = {
    tables: [],
    remainingUnits: [...units],
    consumedStrokes: new Set(),
  }
  // landscape/slide pages place content absolutely; their dividers are design
  // chrome and the canvas/stream passes own them
  if (opts.pageWidthPt > opts.pageHeightPt) return none

  const tables: TableBlock[] = []
  const consumedStrokes = new Set<Stroke>()
  let pool = [...units]

  for (const band of ruleBands(shapes.strokes)) {
    const bandW = band.box.x1 - band.box.x0
    const bandH = band.box.y1 - band.box.y0
    if (bandW > RULE_MAX_WIDTH_PT || bandH < RULE_MIN_HEIGHT_PT) continue
    if (yCoverage(band) < RULE_Y_COVER_MIN) continue
    const cx = (band.box.x0 + band.box.x1) / 2
    if (
      cx < RULE_CENTER_MIN_RATIO * opts.pageWidthPt ||
      cx > RULE_CENTER_MAX_RATIO * opts.pageWidthPt
    )
      continue
    if (claimedBoxes.some((c) => intersectArea(band.box, c) > 0)) continue

    // zone membership: unit centers within the padded rule extent
    const yLo = band.box.y0 - ZONE_Y_PAD_PT
    const yHi = band.box.y1 + ZONE_Y_PAD_PT
    const zone: LineUnit[] = []
    const left: LineUnit[] = []
    const right: LineUnit[] = []
    let crossing = false
    for (const u of pool) {
      const cy = (u.box.y0 + u.box.y1) / 2
      if (cy < yLo || cy > yHi) continue
      zone.push(u)
      if (u.box.x1 <= band.box.x0 + 0.5) left.push(u)
      else if (u.box.x0 >= band.box.x1 - 0.5) right.push(u)
      else crossing = true
    }
    if (crossing) continue
    if (left.length < ZONE_MIN_UNITS_PER_SIDE || right.length < ZONE_MIN_UNITS_PER_SIDE) continue
    const rowsOf = (side: LineUnit[]): number =>
      new Set(side.map((u) => Math.round(u.baseline))).size
    if (rowsOf(left) < ZONE_MIN_ROWS_PER_SIDE || rowsOf(right) < ZONE_MIN_ROWS_PER_SIDE) continue

    // the stacks must overlap vertically — a rule between stacked (not
    // side-by-side) content is somebody else's decoration
    const extentOf = (side: LineUnit[]): { lo: number; hi: number } => ({
      lo: Math.min(...side.map((u) => u.box.y0)),
      hi: Math.max(...side.map((u) => u.box.y1)),
    })
    const le = extentOf(left)
    const re = extentOf(right)
    const overlap = Math.min(le.hi, re.hi) - Math.max(le.lo, re.lo)
    const minH = Math.min(le.hi - le.lo, re.hi - re.lo)
    if (minH <= 0 || overlap < ZONE_SIDE_OVERLAP_MIN * minH) continue

    const contentBox = [...left, ...right].map((u) => u.box).reduce(rectUnion)
    const zoneBox = rectUnion(contentBox, band.box)
    const zoneH = zoneBox.y1 - zoneBox.y0
    if (zoneH > ZONE_MAX_PAGE_HEIGHT_RATIO * opts.pageHeightPt) continue
    if (bandH < ZONE_RULE_SPAN_MIN * zoneH) continue

    // text-only zone (anti-misfire): a diagram's vertical connector also
    // stands between two text scraps, but diagrams carry fills / boxes /
    // other rules inside the same region — a caption zone is nothing but the
    // two stacks and their separator (testPDF_protected p40 chart pages)
    const pad = 2
    const inZone = (b: Rect): boolean =>
      b.x1 > zoneBox.x0 - pad &&
      b.x0 < zoneBox.x1 + pad &&
      b.y1 > zoneBox.y0 - pad &&
      b.y0 < zoneBox.y1 + pad
    const bandSet = new Set(band.strokes)
    if (
      shapes.fills.some((f) => inZone(f.box)) ||
      (shapes.curvedFills ?? []).some((f) => inZone(f.box)) ||
      shapes.strokes.some((s) => !bandSet.has(s) && inZone(s.box))
    )
      continue

    const split = cx
    const leftBox: Rect = { x0: zoneBox.x0, x1: split, y0: zoneBox.y0, y1: zoneBox.y1 }
    const rightBox: Rect = { x0: split, x1: zoneBox.x1, y0: zoneBox.y0, y1: zoneBox.y1 }
    tables.push({
      kind: 'table',
      box: zoneBox,
      colWidthsPt: [split - zoneBox.x0, zoneBox.x1 - split],
      rows: [[buildStackCell(left, leftBox, zoneH), buildStackCell(right, rightBox, zoneH)]],
      confidence: ZONE_CONFIDENCE,
      sepRule: band.double ? 'double' : 'single',
    })
    for (const s of band.strokes) consumedStrokes.add(s)
    const consumed = new Set(zone)
    pool = pool.filter((u) => !consumed.has(u))
  }

  return tables.length > 0 ? { tables, remainingUnits: pool, consumedStrokes } : none
}

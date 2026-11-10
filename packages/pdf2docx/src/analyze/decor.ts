/**
 * Decorative-rule mapping (P7): stray horizontal thin lines that neither the
 * table pass nor the underline/strikethrough pass consumed are page design
 * elements (title rules, separators). They become paragraph borders (w:pBdr)
 * on the nearest paragraph below (w:top) or above (w:bottom); with no host in
 * range they turn into a standalone thin bordered paragraph. A vertical bar
 * hugging a text block's left flank becomes its w:left border (quote/accent
 * bars, P14 C); other vertical strays stay ignored (doc-level warning).
 * Pure geometry.
 */
import type { Rect } from '../geometry'
import { rectUnion } from '../geometry'
import type { DecorBorder, PageSection, Stroke, TextBlock } from '../ir'

/** decorative rules are thin… */
const DECOR_MAX_THICK_PT = 3
/** …and clearly elongated */
const DECOR_MIN_LEN_PT = 30
/** a host paragraph must sit within this many points of the line */
const DECOR_ATTACH_MAX_PT = 24
/**
 * fallback host range for bars that would otherwise stand alone: farther is
 * fine because the border's w:space (≤31pt) caps the drawn offset anyway —
 * a somewhat misplaced rule beats an extra flow paragraph that shifts pages
 */
const DECOR_ATTACH_FAR_PT = 100
/** looser x-overlap for the far fallback */
const DECOR_X_OVERLAP_FAR = 0.3
/** the line and its host must share at least this much of the narrower extent */
const DECOR_X_OVERLAP_MIN = 0.5
/** shapes within this distance (pt) of a table/grid region belong to it */
const TABLE_EXCLUDE_TOL = 2
/** Word clamps w:pBdr w:space to this many points */
const BORDER_SPACE_MAX_PT = 31
// ── vertical accent bars (P14 C) ──
/** the bar must sit within this many points left of its block's text edge */
const LEFT_BAR_MAX_GAP_PT = 30
/** …vertically cover at least this share of the block… */
const LEFT_BAR_MIN_BLOCK_COVER = 0.55
/** …and the block must cover at least this share of the bar (no page-height rails) */
const LEFT_BAR_MIN_BAR_COVER = 0.45
/** accent bars may be thicker than hairline rules */
const LEFT_BAR_MAX_THICK_PT = 6
/** …but must still be clearly vertical bars, not panels */
const LEFT_BAR_MIN_LEN_PT = 14
/**
 * at most this many HOSTLESS bars materialize per page. Every standalone bar
 * costs ~1–2pt of flow height that no spacing budget can always fund (bars
 * side by side stack vertically in the flow); design-heavy pages with dozens
 * of rules would creep into page overflow.
 */
const MAX_STANDALONE_BARS = 2

export interface DecorResult {
  /** vertical decorative strays seen and skipped (doc-level warning) */
  ignoredVertical: number
  /** hostless bars dropped over the per-page cap (doc-level warning) */
  droppedBars: number
}

const centerInsideAny = (x: number, y: number, boxes: readonly Rect[]): boolean =>
  boxes.some(
    (b) =>
      x >= b.x0 - TABLE_EXCLUDE_TOL &&
      x <= b.x1 + TABLE_EXCLUDE_TOL &&
      y >= b.y0 - TABLE_EXCLUDE_TOL &&
      y <= b.y1 + TABLE_EXCLUDE_TOL,
  )

const xOverlapRatio = (a: Rect, b: Rect): number => {
  const overlap = Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0)
  const narrower = Math.min(a.x1 - a.x0, b.x1 - b.x0)
  return narrower > 0 && overlap > 0 ? overlap / narrower : 0
}

interface HostRef {
  block: TextBlock
  column: { box: Rect; blocks: PageSection['columns'][number]['blocks'] }
}

/**
 * Attach a vertical accent bar as its neighbor block's w:left border (P14 C).
 * The bar must sit just left of the block, cover most of its height, and the
 * block most of the bar's length (a column divider rail fails that and stays
 * ignored). Returns true when a host took the border.
 */
function attachLeftBar(
  stroke: Stroke,
  len: number,
  columns: ReadonlyArray<{ column: HostRef['column'] }>,
): boolean {
  let best: TextBlock | null = null
  let bestGap = Infinity
  for (const { column } of columns) {
    for (const block of column.blocks) {
      if (block.kind !== 'text' || block.tocEntry || block.border) continue
      const gap = block.box.x0 - stroke.box.x1
      if (gap < -1 || gap > LEFT_BAR_MAX_GAP_PT || gap >= bestGap) continue
      const overlap = Math.min(block.box.y1, stroke.box.y1) - Math.max(block.box.y0, stroke.box.y0)
      const blockH = block.box.y1 - block.box.y0
      if (overlap < LEFT_BAR_MIN_BLOCK_COVER * blockH) continue
      if (overlap < LEFT_BAR_MIN_BAR_COVER * len) continue
      best = block
      bestGap = gap
    }
  }
  if (!best) return false
  best.border = {
    side: 'left',
    color: stroke.color,
    widthPt: stroke.widthPt,
    spacePt: Math.min(BORDER_SPACE_MAX_PT, Math.max(0, bestGap)),
  }
  return true
}

/** border payload for a hosted rule; the right inset keeps the border's extent */
function borderOf(stroke: Stroke, side: DecorBorder['side'], host: HostRef): DecorBorder {
  const y = (stroke.box.y0 + stroke.box.y1) / 2
  const textEdge = side === 'top' ? host.block.box.y1 : host.block.box.y0
  const border: DecorBorder = {
    side,
    color: stroke.color,
    widthPt: stroke.widthPt,
    spacePt: Math.min(BORDER_SPACE_MAX_PT, Math.max(0, Math.abs(y - textEdge))),
  }
  // only a rule at least as wide as its text may pin the border's right end —
  // narrowing the paragraph under wider text would re-wrap it
  if (stroke.box.x1 >= host.block.box.x1 - 2) {
    const inset = host.column.box.x1 - stroke.box.x1
    if (inset > 1) border.indentRightPt = inset
  }
  return border
}

/**
 * Attach leftover horizontal rules to paragraphs (mutates the section blocks).
 * `excludeBoxes` are table/grid regions whose strokes are table furniture;
 * `consumedStrokes` already restyled text (underline/strikethrough).
 */
export function applyDecorBorders(
  sections: PageSection[],
  strokes: readonly Stroke[],
  excludeBoxes: readonly Rect[],
  consumedStrokes: ReadonlySet<Stroke>,
): DecorResult {
  const columns = sections.flatMap((s, si) =>
    s.columns.map((column) => ({ column, sectionIndex: si })),
  )
  let ignoredVertical = 0
  const standalone: Stroke[] = []

  for (const stroke of strokes) {
    if (consumedStrokes.has(stroke)) continue
    const len =
      stroke.orientation === 'h' ? stroke.box.x1 - stroke.box.x0 : stroke.box.y1 - stroke.box.y0
    const cx = (stroke.box.x0 + stroke.box.x1) / 2
    const cy = (stroke.box.y0 + stroke.box.y1) / 2
    if (centerInsideAny(cx, cy, excludeBoxes)) continue
    if (stroke.orientation === 'v') {
      // vertical accent bar hugging a text block's left flank → w:left border
      // (P14 C: quote bars); anything else vertical stays ignored, counted
      // only within the decorative-rule dimensions (legacy warning scope)
      if (
        stroke.widthPt <= LEFT_BAR_MAX_THICK_PT &&
        len >= LEFT_BAR_MIN_LEN_PT &&
        attachLeftBar(stroke, len, columns)
      ) {
        continue
      }
      if (stroke.widthPt <= DECOR_MAX_THICK_PT && len >= DECOR_MIN_LEN_PT) ignoredVertical++
      continue
    }
    if (stroke.widthPt > DECOR_MAX_THICK_PT || len < DECOR_MIN_LEN_PT) continue

    // nearest paragraph below the line hosts it as w:top; else the one above
    // as w:bottom (y up: below = smaller y)
    let below: HostRef | null = null
    let above: HostRef | null = null
    let farBelow: HostRef | null = null
    let belowDist = Infinity
    let aboveDist = Infinity
    let farDist = Infinity
    for (const { column } of columns) {
      for (const block of column.blocks) {
        if (block.kind !== 'text' || block.tocEntry) continue
        const overlap = xOverlapRatio(stroke.box, block.box)
        if (overlap < DECOR_X_OVERLAP_FAR) continue
        const dBelow = cy - block.box.y1
        const dAbove = block.box.y0 - cy
        if (overlap >= DECOR_X_OVERLAP_MIN) {
          if (dBelow >= -1 && dBelow <= DECOR_ATTACH_MAX_PT && dBelow < belowDist) {
            belowDist = dBelow
            below = { block, column }
          }
          if (dAbove >= -1 && dAbove <= DECOR_ATTACH_MAX_PT && dAbove < aboveDist) {
            aboveDist = dAbove
            above = { block, column }
          }
        }
        if (dBelow >= -1 && dBelow <= DECOR_ATTACH_FAR_PT && dBelow < farDist) {
          farDist = dBelow
          farBelow = { block, column }
        }
      }
    }
    // far fallback: a somewhat misplaced border (w:space caps the drawn
    // offset at 31pt) beats a standalone flow paragraph that shifts pages
    const host = below ?? above ?? farBelow
    if (host) {
      // first rule wins; a second rule on the same side falls through to a
      // standalone paragraph rather than overwriting
      const side: DecorBorder['side'] = host === above ? 'bottom' : 'top'
      if (!host.block.border) {
        host.block.border = borderOf(stroke, side, host)
        continue
      }
    }

    standalone.push(stroke)
  }

  // hostless bars: longest first up to the cap, the rest are dropped
  standalone.sort((a, b) => b.box.x1 - b.box.x0 - (a.box.x1 - a.box.x0))
  let placed = 0
  for (const stroke of standalone) {
    if (placed >= MAX_STANDALONE_BARS) break
    const cx = (stroke.box.x0 + stroke.box.x1) / 2
    const cy = (stroke.box.y0 + stroke.box.y1) / 2
    // the bar joins the horizontally-containing column whose y range is
    // NEAREST — bars often sit in the gaps between sections; dumping them
    // into an arbitrary column produced page-sized spacing gaps
    let entry: (typeof columns)[number] | null = null
    let entryDy = Infinity
    for (const e of columns) {
      const c = e.column.box
      if (cx < c.x0 - 2 || cx > c.x1 + 2) continue
      const dy = cy > c.y1 ? cy - c.y1 : cy < c.y0 ? c.y0 - cy : 0
      if (dy < entryDy) {
        entryDy = dy
        entry = e
      }
    }
    if (!entry) continue
    const { column, sectionIndex } = entry
    // the bar must own its horizontal slice: interleaving with a block's y
    // range would make the spacing chain re-measure gaps from the bar's edge
    // and inflate the flow by the whole overlap
    const interleaves = column.blocks.some(
      (b) =>
        !(b.kind === 'image' && b.float) &&
        b.box.y0 < stroke.box.y1 + 1 &&
        b.box.y1 > stroke.box.y0 - 1,
    )
    if (interleaves) continue
    // …and must not become the first flow block of a NON-FIRST section's
    // column: the spacing chain would measure its gap from the previous
    // section's bottom — across all the parallel columns — and emit a
    // page-sized spacing
    const hasAbove = column.blocks.some(
      (b) => !(b.kind === 'image' && b.float) && b.box.y1 >= stroke.box.y1,
    )
    if (!hasAbove && sectionIndex > 0) continue
    placed++
    const bar: TextBlock = {
      kind: 'text',
      lines: [],
      box: stroke.box,
      align: 'left',
      firstLineIndentPt: 0,
      dir: 'ltr',
      border: {
        side: 'top',
        color: stroke.color,
        widthPt: stroke.widthPt,
        spacePt: 0,
        indentLeftPt: Math.max(0, stroke.box.x0 - column.box.x0),
        indentRightPt: Math.max(0, column.box.x1 - stroke.box.x1),
      },
    }
    const at = column.blocks.findIndex((b) => b.box.y1 < stroke.box.y1)
    if (at === -1) column.blocks.push(bar)
    else column.blocks.splice(at, 0, bar)
    // a bar outside the column's old extent grows it — the NEXT section's
    // spacing chains from this section's bottom, and the whitespace the bar
    // just consumed must not be counted again
    column.box = rectUnion(column.box, bar.box)
    const section = sections[sectionIndex]
    if (section) section.box = rectUnion(section.box, bar.box)
  }
  return { ignoredVertical, droppedBars: standalone.length - placed }
}

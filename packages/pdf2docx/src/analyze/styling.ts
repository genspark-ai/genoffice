/**
 * Text × shape style mapping (pdf2docx rule 6): a Fill overlapping text is a
 * run highlight; a horizontal Stroke through the text's middle band is a
 * strikethrough, one hugging the baseline an underline. Shapes inside
 * lattice-table regions are table furniture (borders / cell shading) and never
 * restyle text. Pure geometry; every comparison is tolerance-based.
 */
import type { Rect } from '../geometry'
import { overlapRatio, rectHeight, rectWidth, median } from '../geometry'
import type { PageShapes, PdfChar, Stroke } from '../ir'

/** a char is highlighted when this share of its box sits inside the fill */
const HIGHLIGHT_COVER_MIN = 0.6
/** fills taller than this many ems are text-box backgrounds, not highlights */
const HIGHLIGHT_MAX_HEIGHT_EMS = 2.5
/**
 * a highlight marks BODY text — display-size glyphs (stamp monograms, poster
 * titles) sit on design panels, and painting those as w:highlight turns the
 * glyph into a solid slab (P10 C)
 */
const HIGHLIGHT_MAX_FONT_PT = 36
/** fills covering more than this share of the page are page background */
const HIGHLIGHT_MAX_PAGE_RATIO = 0.5
/** near-white fills are paper, not highlight */
const WHITE_FILL = /^F[EF]F[EF]F[EF]$/i

/** stroke bands relative to the baseline, in ems (y up: positive = above) */
const STRIKE_BAND_LO_EMS = 0.15
const STRIKE_BAND_HI_EMS = 0.62
const UNDERLINE_BAND_LO_EMS = -0.35
const UNDERLINE_BAND_HI_EMS = 0.12
/** the stroke must cover this share of a char's width to style it */
const LINE_COVER_MIN = 0.6
/** a line thicker than this many ems is a decorative bar, not under/strike */
const LINE_MAX_WIDTH_EMS = 0.25
/**
 * a stroke extending further than this many ems beyond the text it covers is
 * a separator / table rule (real under/strike lines match their text's width)
 */
const LINE_EXTEND_MAX_EMS = 1.5
/** shapes within this distance (pt) of a table region count as its furniture */
const TABLE_EXCLUDE_TOL = 2

export interface StyledChars {
  chars: PdfChar[]
  warnings: string[]
  /** strokes that became run underline/strikethrough (the P7 decor pass skips them) */
  consumedStrokes: Set<Stroke>
}

const insideAny = (x: number, y: number, boxes: readonly Rect[]): boolean =>
  boxes.some(
    (b) =>
      x >= b.x0 - TABLE_EXCLUDE_TOL &&
      x <= b.x1 + TABLE_EXCLUDE_TOL &&
      y >= b.y0 - TABLE_EXCLUDE_TOL &&
      y <= b.y1 + TABLE_EXCLUDE_TOL,
  )

const isStylableChar = (c: PdfChar): boolean =>
  !c.isGenerated && c.code > 0x20 && rectWidth(c.box) > 0

interface CharStyle {
  highlight?: string
  underline?: boolean
  strike?: boolean
  underlineColorMismatch?: boolean
}

/** horizontal overlap between the stroke and the char, as a share of char width */
function xCover(stroke: Stroke, c: PdfChar): number {
  const w = rectWidth(c.box)
  if (w <= 0) return 0
  const overlap = Math.min(stroke.box.x1, c.box.x1) - Math.max(stroke.box.x0, c.box.x0)
  return overlap > 0 ? overlap / w : 0
}

/**
 * Apply fill-highlight and stroke-underline/strikethrough styling to page
 * chars. `tableBoxes` are lattice grid regions whose shapes are excluded.
 * Returns fresh char objects (input untouched) plus per-page warnings.
 */
export function applyTextShapeStyles(
  chars: readonly PdfChar[],
  shapes: PageShapes,
  tableBoxes: readonly Rect[],
  pageArea: number,
): StyledChars {
  const styles = new Map<number, CharStyle>()
  const styleOf = (i: number): CharStyle => {
    let s = styles.get(i)
    if (!s) styles.set(i, (s = {}))
    return s
  }
  const candidates = chars.map((c, i) => ({ c, i })).filter(({ c }) => isStylableChar(c))

  // ── highlights: fills overlapping text (paint order = array order, last wins) ──
  for (const fill of shapes.fills) {
    if (WHITE_FILL.test(fill.color)) continue
    if (
      pageArea > 0 &&
      rectWidth(fill.box) * rectHeight(fill.box) > HIGHLIGHT_MAX_PAGE_RATIO * pageArea
    ) {
      continue
    }
    const cx = (fill.box.x0 + fill.box.x1) / 2
    const cy = (fill.box.y0 + fill.box.y1) / 2
    if (insideAny(cx, cy, tableBoxes)) continue // cell shading context
    for (const { c, i } of candidates) {
      if (c.fontSize > HIGHLIGHT_MAX_FONT_PT) continue
      if (rectHeight(fill.box) > HIGHLIGHT_MAX_HEIGHT_EMS * Math.max(c.fontSize, 1)) continue
      if (overlapRatio(c.box, fill.box) >= HIGHLIGHT_COVER_MIN) styleOf(i).highlight = fill.color
    }
  }

  // ── underline / strikethrough: thin horizontal strokes over text bands ──
  const warnings: string[] = []
  const consumedStrokes = new Set<Stroke>()
  let colorMismatch = false
  for (const stroke of shapes.strokes) {
    if (stroke.orientation !== 'h') continue
    const strokeY = (stroke.box.y0 + stroke.box.y1) / 2
    const strokeX = (stroke.box.x0 + stroke.box.x1) / 2
    if (insideAny(strokeX, strokeY, tableBoxes)) continue // table border

    const under: Array<{ c: PdfChar; i: number }> = []
    const strike: Array<{ c: PdfChar; i: number }> = []
    for (const entry of candidates) {
      const { c } = entry
      const fs = Math.max(c.fontSize, 1)
      if (stroke.widthPt > LINE_MAX_WIDTH_EMS * fs) continue
      if (xCover(stroke, c) < LINE_COVER_MIN) continue
      const rel = (strokeY - c.originY) / fs
      if (rel >= STRIKE_BAND_LO_EMS && rel <= STRIKE_BAND_HI_EMS) strike.push(entry)
      else if (rel >= UNDERLINE_BAND_LO_EMS && rel < UNDERLINE_BAND_HI_EMS) under.push(entry)
    }
    const matched = [...under, ...strike]
    if (matched.length === 0) continue

    // a rule running well past its text is a separator, not a text decoration
    const textX0 = Math.min(...matched.map(({ c }) => c.box.x0))
    const textX1 = Math.max(...matched.map(({ c }) => c.box.x1))
    const em = median(matched.map(({ c }) => c.fontSize)) || 12
    if (
      textX0 - stroke.box.x0 > LINE_EXTEND_MAX_EMS * em ||
      stroke.box.x1 - textX1 > LINE_EXTEND_MAX_EMS * em
    ) {
      continue
    }

    for (const { i } of under) styleOf(i).underline = true
    for (const { i } of strike) styleOf(i).strike = true
    consumedStrokes.add(stroke)
    if (matched.some(({ c }) => c.color !== stroke.color)) colorMismatch = true
  }
  if (colorMismatch) {
    warnings.push(
      'underline/strikethrough color differs from text color (not representable, ignored)',
    )
  }

  if (styles.size === 0) return { chars: [...chars], warnings, consumedStrokes }
  const out = chars.map((c, i) => {
    const s = styles.get(i)
    if (!s || (!s.highlight && !s.underline && !s.strike)) return c
    const styled: PdfChar = { ...c }
    if (s.highlight) styled.highlight = s.highlight
    if (s.underline) styled.underline = true
    if (s.strike) styled.strike = true
    return styled
  })
  return { chars: out, warnings, consumedStrokes }
}

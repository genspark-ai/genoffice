/**
 * Shape normalization (pdf2docx rule 1): raw page paths → Strokes (thin
 * horizontal/vertical lines: table borders / underline candidates) and Fills
 * (plain rectangles: cell shading / highlight candidates). Curves, diagonals
 * and non-rectangular fills are counted and ignored — vector-art handling is
 * P4. Pure geometry, fully unit-testable.
 */
import type { Rect } from '../geometry'
import { approxEq, intersectArea, rectArea } from '../geometry'
import type { Fill, PageShapes, RawPath, RawSubpath, Stroke } from '../ir'

/** points this close are the same coordinate when detecting rectangles/axis lines */
const AXIS_TOL = 0.5
/** fills below this alpha are glows/tints — never shading or highlight (P10 C) */
const FILL_MIN_ALPHA = 128
/** a rect this thin is a line, not a shading box (3pt covers accent bars, P14 C) */
const THIN_MAX_PT = 3.2
/** …but only when clearly elongated (a 2×2 dot is not a line) */
const THIN_MIN_ASPECT = 3

interface Pt {
  x: number
  y: number
}

/** stroke whose centerline runs horizontally through `box` */
function hStroke(x0: number, x1: number, y: number, widthPt: number, color: string): Stroke {
  const w = Math.max(widthPt, 0.1)
  return {
    box: { x0: Math.min(x0, x1), x1: Math.max(x0, x1), y0: y - w / 2, y1: y + w / 2 },
    orientation: 'h',
    widthPt: w,
    color,
  }
}

function vStroke(y0: number, y1: number, x: number, widthPt: number, color: string): Stroke {
  const w = Math.max(widthPt, 0.1)
  return {
    box: { x0: x - w / 2, x1: x + w / 2, y0: Math.min(y0, y1), y1: Math.max(y0, y1) },
    orientation: 'v',
    widthPt: w,
    color,
  }
}

/**
 * Detect an axis-aligned rectangle: 4 corner points (closing point may repeat
 * the first) whose consecutive edges alternate horizontal/vertical.
 */
export function rectOfSubpath(sub: RawSubpath): Rect | null {
  // collapse consecutive duplicate points (pen-drop artifacts)
  let pts: Pt[] = []
  for (const p of sub.points) {
    const prev = pts[pts.length - 1]
    if (!prev || !(approxEq(p.x, prev.x, AXIS_TOL) && approxEq(p.y, prev.y, AXIS_TOL))) pts.push(p)
  }
  if (pts.length >= 2) {
    const first = pts[0]!
    const last = pts[pts.length - 1]!
    if (approxEq(first.x, last.x, AXIS_TOL) && approxEq(first.y, last.y, AXIS_TOL)) {
      pts = pts.slice(0, -1)
    }
  }
  // drop points sitting on an axis-aligned line between their cyclic
  // neighbors — outlines often start mid-edge (P16 A), leaving 5 corners
  for (let i = 0; pts.length > 4 && i < pts.length;) {
    const a = pts[(i - 1 + pts.length) % pts.length]!
    const b = pts[i]!
    const c = pts[(i + 1) % pts.length]!
    const collinear =
      (approxEq(a.x, b.x, AXIS_TOL) && approxEq(b.x, c.x, AXIS_TOL)) ||
      (approxEq(a.y, b.y, AXIS_TOL) && approxEq(b.y, c.y, AXIS_TOL))
    if (collinear) {
      pts.splice(i, 1)
      i = 0
    } else {
      i++
    }
  }
  if (pts.length !== 4) return null
  for (let i = 0; i < 4; i++) {
    const a = pts[i]!
    const b = pts[(i + 1) % 4]!
    const horizontal = approxEq(a.y, b.y, AXIS_TOL)
    const vertical = approxEq(a.x, b.x, AXIS_TOL)
    if (horizontal === vertical) return null // diagonal edge, or a degenerate double point
    // consecutive edges must alternate orientation; checking i vs i+1 covers all pairs
    const c = pts[(i + 2) % 4]!
    const nextHorizontal = approxEq(b.y, c.y, AXIS_TOL)
    if (horizontal === nextHorizontal) return null
  }
  const xs = pts.map((p) => p.x)
  const ys = pts.map((p) => p.y)
  return {
    x0: Math.min(...xs),
    y0: Math.min(...ys),
    x1: Math.max(...xs),
    y1: Math.max(...ys),
  }
}

/** all consecutive edges (incl. the closing one) are axis-aligned */
function allEdgesAxisAligned(sub: RawSubpath): boolean {
  const pts = sub.points.length > 2 ? [...sub.points, sub.points[0]!] : sub.points
  for (let i = 1; i < pts.length; i++) {
    const a = pts[i - 1]!
    const b = pts[i]!
    if (!approxEq(a.x, b.x, AXIS_TOL) && !approxEq(a.y, b.y, AXIS_TOL)) return false
  }
  return true
}

const bboxOf = (sub: RawSubpath): Rect => {
  const xs = sub.points.map((p) => p.x)
  const ys = sub.points.map((p) => p.y)
  return { x0: Math.min(...xs), y0: Math.min(...ys), x1: Math.max(...xs), y1: Math.max(...ys) }
}

/** thin + elongated → the rect IS a line */
function thinRectStroke(rect: Rect, color: string): Stroke | null {
  const w = rect.x1 - rect.x0
  const h = rect.y1 - rect.y0
  const minDim = Math.min(w, h)
  const maxDim = Math.max(w, h)
  if (minDim > THIN_MAX_PT || maxDim < THIN_MIN_ASPECT * Math.max(minDim, 0.1)) return null
  return w >= h
    ? hStroke(rect.x0, rect.x1, (rect.y0 + rect.y1) / 2, h, color)
    : vStroke(rect.y0, rect.y1, (rect.x0 + rect.x1) / 2, w, color)
}

/**
 * A rounded card/banner's outline hugs its bbox; logo art (swooshes, mascots,
 * outline glyphs) fills a small share of it. Below this polygon/bbox area
 * ratio the bbox is NOT a faithful stand-in for the shape — treating it as a
 * backdrop plate turns a faint watermark into a solid slab over the body text
 * (P29 C). Rounded rects run ≥0.95, circles/pills ≈0.79.
 */
const CURVED_FILL_MIN_COVER = 0.55

/** this many mutually-overlapping curved fills at one z mark the group as vector art */
const CURVED_ART_GROUP_MIN = 3
/** pair overlap share (of the smaller bbox) that counts two fills as one drawing */
const CURVED_ART_OVERLAP_SHARE = 0.2

/** |shoelace| area of the subpath's (control-)polygon */
function polygonArea(sub: RawSubpath): number {
  const pts = sub.points
  if (pts.length < 3) return 0
  let sum = 0
  for (let i = 0; i < pts.length; i++) {
    const a = pts[i]!
    const b = pts[(i + 1) % pts.length]!
    sum += a.x * b.y - b.x * a.y
  }
  return Math.abs(sum) / 2
}

/** translucent fills keep their alpha so background panels can preserve it (P11 A) */
const alphaOf = (path: RawPath): { alpha?: number } =>
  path.fillAlpha !== undefined && path.fillAlpha < 255 ? { alpha: path.fillAlpha } : {}

/** fills keep the source paint order for behindDoc stacking (P16 A) */
const zOf = (path: RawPath): { z?: number } => (path.z !== undefined ? { z: path.z } : {})

/** Normalize raw paths into page shapes (strokes + fills + ignored count). */
export function normalizeShapes(paths: readonly RawPath[]): PageShapes {
  const strokes: Stroke[] = []
  const fills: Fill[] = []
  const curvedFills: Fill[] = []
  let ignoredPaths = 0

  for (const path of paths) {
    const strokesBefore = strokes.length
    if (!path.filled && !path.stroked) continue
    // a translucent fill is a glow/shadow/tint (P10 C): as cell shading or
    // text highlight it would paint at FULL opacity (a 35%-alpha black glow
    // became a solid black highlight slab) — only its stroke edges survive
    const filled = path.filled && (path.fillAlpha ?? 255) >= FILL_MIN_ALPHA
    if (!filled && !path.stroked) {
      ignoredPaths++
      continue
    }
    for (const sub of path.subpaths) {
      if (sub.hasCurves) {
        // rounded cards/banners: keep the bbox as a light-text backdrop
        // candidate (P10 B) — vector-art handling still ignores the path
        if (filled && sub.points.length > 0) {
          const box = bboxOf(sub)
          const bboxArea = (box.x1 - box.x0) * (box.y1 - box.y0)
          if (bboxArea > 0 && polygonArea(sub) >= CURVED_FILL_MIN_COVER * bboxArea) {
            curvedFills.push({
              box,
              color: path.fillColor,
              ...alphaOf(path),
              ...zOf(path),
            })
          }
        }
        ignoredPaths++
        continue
      }
      const rect = rectOfSubpath(sub)
      if (rect) {
        // a thin rect is a line whether it was filled or stroked; emitting its
        // outline edges too would seed phantom perpendicular micro-strokes
        const thin = thinRectStroke(rect, filled ? path.fillColor : path.strokeColor)
        if (thin) {
          strokes.push(thin)
          continue
        }
        if (filled) fills.push({ box: rect, color: path.fillColor, ...alphaOf(path), ...zOf(path) })
        if (path.stroked) {
          strokes.push(hStroke(rect.x0, rect.x1, rect.y1, path.strokeWidth, path.strokeColor))
          strokes.push(hStroke(rect.x0, rect.x1, rect.y0, path.strokeWidth, path.strokeColor))
          strokes.push(vStroke(rect.y0, rect.y1, rect.x0, path.strokeWidth, path.strokeColor))
          strokes.push(vStroke(rect.y0, rect.y1, rect.x1, path.strokeWidth, path.strokeColor))
        }
        continue
      }
      // non-rect polyline: stroked axis-aligned edges become strokes
      if (filled) {
        // P7: a filled bar drawn with redundant collinear points (>4 corners)
        // is not a 4-point rect but still a decorative line when its bbox is
        // thin, elongated and every edge is axis-aligned — salvage it
        const thinBar = allEdgesAxisAligned(sub)
          ? thinRectStroke(bboxOf(sub), path.fillColor)
          : null
        if (thinBar) {
          strokes.push(thinBar)
          continue
        }
        ignoredPaths++ // non-rect fill shape (vector art)
      }
      if (!path.stroked) continue
      const closing =
        sub.closed && sub.points.length > 2 ? [...sub.points, sub.points[0]!] : sub.points
      let sawDiagonal = false
      for (let i = 1; i < closing.length; i++) {
        const a = closing[i - 1]!
        const b = closing[i]!
        if (approxEq(a.y, b.y, AXIS_TOL)) {
          if (!approxEq(a.x, b.x, AXIS_TOL)) {
            strokes.push(hStroke(a.x, b.x, (a.y + b.y) / 2, path.strokeWidth, path.strokeColor))
          }
        } else if (approxEq(a.x, b.x, AXIS_TOL)) {
          strokes.push(vStroke(a.y, b.y, (a.x + b.x) / 2, path.strokeWidth, path.strokeColor))
        } else {
          sawDiagonal = true
        }
      }
      if (sawDiagonal) ignoredPaths++
    }
    // strokes minted from this path inherit its form-XObject origin (P14 C)
    if (path.fromForm) {
      for (let i = strokesBefore; i < strokes.length; i++) strokes[i]!.fromForm = true
    }
  }
  return { strokes, fills, ignoredPaths, ...(curvedFills.length > 0 ? { curvedFills } : {}) }
}

/** a full-page background fill must cover at least this share of each page dimension */
const BG_COVER_RATIO = 0.94
/** channels at/above this are white enough that no background is worth emitting */
const BG_WHITE_MIN = 0xf7

/** all RGB channels near 255 — the default page color, not a background */
export function isNearWhite(hex: string): boolean {
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return false
  return (
    parseInt(hex.slice(0, 2), 16) >= BG_WHITE_MIN &&
    parseInt(hex.slice(2, 4), 16) >= BG_WHITE_MIN &&
    parseInt(hex.slice(4, 6), 16) >= BG_WHITE_MIN
  )
}

/** full-bleed tile group (P22 B): how many tiles may paint the page */
const TILE_MAX_COUNT = 6
/** each tile must be a real slab, not an accent bar */
const TILE_MIN_AREA_RATIO = 0.08
/** together the tiles must paint essentially the whole page */
const TILE_COVER_MIN = 0.95
/** coverage sampling grid (matches canvas.ts's backgroundCoverShare idea) */
const TILE_GRID = 24
/** a page with real body text is a document, not a color plate */
const TILE_MAX_CHARS = 10

/**
 * Full-bleed tile group (P22 B): a page painted edge-to-edge by a few large
 * vector slabs (color plates, poster backgrounds). One slab covering the
 * whole page is the wash (extractPageBackground); a MOSAIC of slabs never
 * hits that path — instead its hairline edges mint a ghost lattice table
 * that gets margin-clamped and shrunk. Claim the group: the tiles leave the
 * fill pool (the caller pins them as behindDoc floats at their absolute
 * coordinates), and their edge strokes leave the stroke pool so no ghost
 * grid survives. Text-bearing pages never qualify — a big banner or a
 * shaded table region always sits under body text.
 */
export function extractFullBleedTiles(
  shapes: PageShapes,
  charCount: number,
  widthPt: number,
  heightPt: number,
): Fill[] {
  if (charCount > TILE_MAX_CHARS) return []
  const pool = shapes.fills
  const candidates = pool.filter(
    (f) =>
      (f.alpha === undefined || f.alpha >= 250) &&
      (f.box.x1 - f.box.x0) * (f.box.y1 - f.box.y0) >= TILE_MIN_AREA_RATIO * widthPt * heightPt,
  )
  if (candidates.length < 2 || candidates.length > TILE_MAX_COUNT) return []

  let covered = 0
  for (let gy = 0; gy < TILE_GRID; gy++) {
    for (let gx = 0; gx < TILE_GRID; gx++) {
      const cx = ((gx + 0.5) / TILE_GRID) * widthPt
      const cy = ((gy + 0.5) / TILE_GRID) * heightPt
      if (
        candidates.some((f) => cx >= f.box.x0 && cx <= f.box.x1 && cy >= f.box.y0 && cy <= f.box.y1)
      )
        covered++
    }
  }
  if (covered / (TILE_GRID * TILE_GRID) < TILE_COVER_MIN) return []

  // consume the tiles and their edge strokes (hairline tile borders)
  const taken = new Set(candidates)
  shapes.fills = pool.filter((f) => !taken.has(f))
  const inTiles = (b: Rect): boolean =>
    candidates.some(
      (f) =>
        b.x0 >= f.box.x0 - 2 &&
        b.x1 <= f.box.x1 + 2 &&
        b.y0 >= f.box.y0 - 2 &&
        b.y1 <= f.box.y1 + 2,
    )
  shapes.strokes = shapes.strokes.filter((s) => !inTiles(s.box))
  // clamp to the page so the behindDoc offsets stay on-page
  return candidates.map((f) => ({
    ...f,
    box: {
      x0: Math.max(0, f.box.x0),
      y0: Math.max(0, f.box.y0),
      x1: Math.min(widthPt, f.box.x1),
      y1: Math.min(heightPt, f.box.y1),
    },
  }))
}

/**
 * Pull the page background out of the fill pool: every fill covering ~the
 * whole page is removed (stacked backgrounds are invisible below the topmost
 * one), and the topmost one's color is returned. Removing them keeps a colored
 * page wash from reading as cell shading / highlight / vector art downstream.
 */
export function extractPageBackground(
  fills: Fill[],
  widthPt: number,
  heightPt: number,
): string | undefined {
  let color: string | undefined
  for (let i = fills.length - 1; i >= 0; i--) {
    const b = fills[i]!.box
    if (b.x1 - b.x0 >= widthPt * BG_COVER_RATIO && b.y1 - b.y0 >= heightPt * BG_COVER_RATIO) {
      if (color === undefined) color = fills[i]!.color
      fills.splice(i, 1)
    }
  }
  return color
}

// ── region background panels (P10 B) ──

/** a panel must cover at least this share of the page area … */
const PANEL_MIN_AREA_RATIO = 0.18
/** … and at most this much (bigger = the page wash, handled above) */
const PANEL_MAX_AREA_RATIO = 0.85
/** a panel spans ~one full page dimension (spine: height, banner: width) */
const PANEL_FULL_DIM_RATIO = 0.96
/** a panel's other side is flush with a page edge within this many points */
const PANEL_EDGE_TOL_PT = 2
/** minimum characters sitting on the panel — a panel is a text backdrop */
const PANEL_MIN_CHARS = 4
/** only the first this-many fills (bottom of the z-order) can be panels */
const PANEL_Z_PREFIX = 3

/**
 * Region background panels: full-height spines / full-width banners flush
 * with a page edge, drawn at the bottom of the z-order with body text on
 * top (deck covers, chapter dividers). They cannot ride the text flow —
 * pulled out of the fill pool here and pinned as region-sized behindDoc
 * floats by the rebuild layer, so the text on them keeps its color contrast.
 * Table shading / highlight bars never qualify: too small, not edge-flush,
 * or not at the z-order bottom.
 */
export function extractBackgroundPanels(
  fills: Fill[],
  charBoxes: readonly Rect[],
  widthPt: number,
  heightPt: number,
): Fill[] {
  const pageArea = widthPt * heightPt
  if (pageArea <= 0) return []
  const panels: Fill[] = []
  const isPanel = (b: Rect, color: string): boolean => {
    if (isNearWhite(color)) return false
    const area = (b.x1 - b.x0) * (b.y1 - b.y0)
    const share = area / pageArea
    if (share < PANEL_MIN_AREA_RATIO || share > PANEL_MAX_AREA_RATIO) return false
    const fullHeight =
      b.y1 - b.y0 >= heightPt * PANEL_FULL_DIM_RATIO &&
      (b.x0 <= PANEL_EDGE_TOL_PT || b.x1 >= widthPt - PANEL_EDGE_TOL_PT)
    const fullWidth =
      b.x1 - b.x0 >= widthPt * PANEL_FULL_DIM_RATIO &&
      (b.y0 <= PANEL_EDGE_TOL_PT || b.y1 >= heightPt - PANEL_EDGE_TOL_PT)
    if (!fullHeight && !fullWidth) return false
    let chars = 0
    for (const c of charBoxes) {
      const cx = (c.x0 + c.x1) / 2
      const cy = (c.y0 + c.y1) / 2
      if (cx >= b.x0 && cx <= b.x1 && cy >= b.y0 && cy <= b.y1 && ++chars >= PANEL_MIN_CHARS) {
        return true
      }
    }
    return false
  }
  // panels sit at the bottom of the z-order: scan a short prefix and stop at
  // the first non-panel fill so mid-stack boxes (cards, callouts) never match
  for (let i = 0; i < Math.min(fills.length, PANEL_Z_PREFIX);) {
    const fill = fills[i]!
    if (!isPanel(fill.box, fill.color)) break
    panels.push(fill)
    fills.splice(i, 1)
  }
  return panels
}

// ── light-text backdrops (P10 B) ──

/** the fill must be at least this dark (relative luminance 0–1) … */
const BACKDROP_MAX_FILL_LUMA = 0.75
/** … and its text this light — without the backdrop that text is invisible */
const BACKDROP_MIN_TEXT_LUMA = 0.85
/** minimum light chars sitting on the fill */
const BACKDROP_MIN_LIGHT_CHARS = 2
/** light chars must dominate the fill's text */
const BACKDROP_MIN_LIGHT_SHARE = 0.6
/** tiny fills are bullets/legend chips, not text cards (pt²) */
const BACKDROP_MIN_AREA_PT2 = 600
/** fills near page size are the wash / a page panel, handled elsewhere */
const BACKDROP_MAX_PAGE_RATIO = 0.85
/** dark-text card panels must cover this share of the page … */
const BACKDROP_CARD_MIN_PAGE_RATIO = 0.1
/** label plates (P16 I) may be smaller … */
const BACKDROP_LABEL_MIN_PAGE_RATIO = 0.02
/** … but must stand far taller than their text — highlight bars and zebra
 * rows hug their line (~1–2×), a label plate is a block (≥2.5×) */
const BACKDROP_LABEL_MIN_TEXT_HEIGHTS = 2.5
/** … hold at least this many chars … */
const BACKDROP_CARD_MIN_CHARS = 6
/** … and keep this much text/fill luminance contrast (highlight bars are
 * smaller; losing a page-scale panel is a big visual delta either way) */
const BACKDROP_CARD_MIN_CONTRAST = 0.3

// ── small capsule backdrops (P12 C) ──
/** a capsule holds at least this many chars — its size floor is its text */
const CAPSULE_MIN_CHARS = 2
/** …contrasting with the fill (median text luma vs fill luma) */
const CAPSULE_MIN_CONTRAST = 0.3
/** capsule shape: clearly wider than tall (chips/circles stay vector art) */
const CAPSULE_MIN_ASPECT = 2
/** …and mostly occupied by its text, not a decorative swoosh with a stray label */
const CAPSULE_MIN_TEXT_WIDTH_SHARE = 0.3

/** relative luminance of a RRGGBB hex, 0 (black) – 1 (white) */
export function hexLuminance(hex: string): number {
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return 0
  const r = parseInt(hex.slice(0, 2), 16)
  const g = parseInt(hex.slice(2, 4), 16)
  const b = parseInt(hex.slice(4, 6), 16)
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255
}

/**
 * Rectangular card fills carrying light-colored text (slide tiles, colored
 * call-out bars): the fill pool has no rendering path, so dropping them
 * leaves white text invisible on the white paper. They leave the pool here
 * and pin to the page as behindDoc floats; the text on them keeps flowing.
 * Lattice-table cell shading is excluded via `gridBoxes` (the table path
 * carries it as w:shd), near-white and page-scale fills never qualify.
 */
export function extractTextBackdrops(
  fills: Fill[],
  chars: ReadonlyArray<{ box: Rect; code: number; color: string }>,
  gridBoxes: readonly Rect[],
  widthPt: number,
  heightPt: number,
  curvedFills: readonly Fill[] = [],
  hasPageBackdrop = false,
): Fill[] {
  const pageArea = widthPt * heightPt
  if (pageArea <= 0) return []
  const centerInside = (b: Rect, r: Rect): boolean => {
    const cx = (b.x0 + b.x1) / 2
    const cy = (b.y0 + b.y1) / 2
    return cx >= r.x0 && cx <= r.x1 && cy >= r.y0 && cy <= r.y1
  }
  const qualifies = (fill: Fill, rounded = false): boolean => {
    const area = (fill.box.x1 - fill.box.x0) * (fill.box.y1 - fill.box.y0)
    // the pool paths need a real card; the capsule path (rounded only) sizes
    // itself by the text it holds, and a white capsule counts only when the
    // page has a backdrop for it to contrast with (P12 C)
    const poolEligible = !isNearWhite(fill.color) && area >= BACKDROP_MIN_AREA_PT2
    const capsuleEligible = rounded && (hasPageBackdrop || !isNearWhite(fill.color))
    if (!poolEligible && !capsuleEligible) return false
    if (area <= 0 || area > BACKDROP_MAX_PAGE_RATIO * pageArea) return false
    if (gridBoxes.some((g) => centerInside(fill.box, g))) return false
    const fillLuma = hexLuminance(fill.color)
    let total = 0
    let light = 0
    const lumas: number[] = []
    const charHeights: number[] = []
    let textX0 = Infinity
    let textX1 = -Infinity
    for (const c of chars) {
      if (c.code <= 0x20 || !centerInside(c.box, fill.box)) continue
      total++
      const luma = hexLuminance(c.color)
      lumas.push(luma)
      charHeights.push(c.box.y1 - c.box.y0)
      if (luma >= BACKDROP_MIN_TEXT_LUMA) light++
      textX0 = Math.min(textX0, c.box.x0)
      textX1 = Math.max(textX1, c.box.x1)
    }
    if (poolEligible) {
      // light text on a dark card: without the backdrop the text is invisible
      if (
        fillLuma <= BACKDROP_MAX_FILL_LUMA &&
        light >= BACKDROP_MIN_LIGHT_CHARS &&
        light / Math.max(total, 1) >= BACKDROP_MIN_LIGHT_SHARE
      ) {
        return true
      }
      // page-scale card panel (divider slides, checklist boxes): any text tone,
      // as long as the panel visibly contrasts with the text on it
      if (area >= BACKDROP_CARD_MIN_PAGE_RATIO * pageArea && total >= BACKDROP_CARD_MIN_CHARS) {
        lumas.sort((a, b) => a - b)
        const medianLuma = lumas[Math.floor(lumas.length / 2)]!
        if (Math.abs(medianLuma - fillLuma) >= BACKDROP_CARD_MIN_CONTRAST) return true
      }
      // label plate (P16 I): a mid-size flat plate far taller than the short
      // label it carries (a grey "payment method" caption block) — highlight bars and
      // zebra rows hug their text line and never reach the height gate
      if (
        area >= BACKDROP_LABEL_MIN_PAGE_RATIO * pageArea &&
        total >= CAPSULE_MIN_CHARS &&
        charHeights.length > 0
      ) {
        charHeights.sort((a, b) => a - b)
        const medianCharH = charHeights[Math.floor(charHeights.length / 2)]!
        if (fill.box.y1 - fill.box.y0 >= BACKDROP_LABEL_MIN_TEXT_HEIGHTS * medianCharH) {
          lumas.sort((a, b) => a - b)
          const medianLuma = lumas[Math.floor(lumas.length / 2)]!
          if (Math.abs(medianLuma - fillLuma) >= BACKDROP_CARD_MIN_CONTRAST) return true
        }
      }
    }
    // small capsule (P12 C): a rounded pill sized to the contrasting text it
    // carries (a "lecturer: NAME" byline on a white pill over a photo cover) — too small
    // for the card path, invisible once the fill pool is dropped
    if (!capsuleEligible || total < CAPSULE_MIN_CHARS) return false
    const w = fill.box.x1 - fill.box.x0
    const h = fill.box.y1 - fill.box.y0
    if (w < CAPSULE_MIN_ASPECT * h) return false
    if (textX1 - textX0 < CAPSULE_MIN_TEXT_WIDTH_SHARE * w) return false
    lumas.sort((a, b) => a - b)
    const medianLuma = lumas[Math.floor(lumas.length / 2)]!
    return Math.abs(medianLuma - fillLuma) >= CAPSULE_MIN_CONTRAST
  }
  const backdrops: Fill[] = []
  for (let i = fills.length - 1; i >= 0; i--) {
    if (qualifies(fills[i]!)) {
      backdrops.push(fills[i]!)
      fills.splice(i, 1)
    }
  }
  // fills were scanned top-down; restore paint order (bottom first)
  backdrops.reverse()
  // rounded cards/banners (curved subpaths) live outside the fill pool — the
  // bbox stands in for the rounded shape, which is all a backdrop needs.
  // Many MUTUALLY-OVERLAPPING curved fills sharing one z come from ONE
  // drawing (a logo / mascot watermark, P29 C) — its piece bboxes are art,
  // never card plates. Disjoint fills in one form (a card grid) stay eligible.
  const byZ = new Map<number | undefined, Fill[]>()
  for (const fill of curvedFills) {
    const group = byZ.get(fill.z)
    if (group) group.push(fill)
    else byZ.set(fill.z, [fill])
  }
  const artFills = new Set<Fill>()
  for (const group of byZ.values()) {
    if (group.length < CURVED_ART_GROUP_MIN) continue
    let overlapping = 0
    for (let i = 0; i < group.length && overlapping < CURVED_ART_GROUP_MIN; i++) {
      for (let j = i + 1; j < group.length; j++) {
        const inter = intersectArea(group[i]!.box, group[j]!.box)
        const smaller = Math.min(rectArea(group[i]!.box), rectArea(group[j]!.box))
        if (smaller > 0 && inter >= CURVED_ART_OVERLAP_SHARE * smaller) {
          overlapping++
          break
        }
      }
    }
    if (overlapping + 1 >= CURVED_ART_GROUP_MIN) for (const f of group) artFills.add(f)
  }
  for (const fill of curvedFills) {
    if (artFills.has(fill)) continue
    if (qualifies(fill, true)) backdrops.push(fill)
  }
  return backdrops
}

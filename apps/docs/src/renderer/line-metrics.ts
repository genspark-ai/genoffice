/**
 * F1 line-box metrics engine
 *
 * Pure functions — input (paragraph text + run font/size + available width +
 * line-height rule + docGrid grid), output (line count + per-line heights + total block height).
 *
 * Design aligned with packages/pptx-render/src/metrics.ts (dual track: precise
 * opentype / heuristic fallback). F1 runs on heuristics first with the interface
 * ready, so we can switch seamlessly to OpentypeMetrics once font files load.
 *
 * Line-breaking rules:
 *   - Latin text breaks greedily by word (breakable at whitespace)
 *   - CJK breaks per character (Unicode CJK ranges)
 *   - Mixed text splits by Unicode segments (same logic as text-layout.ts, simplified port)
 *
 * Line-height semantics (Word's three modes):
 *   - auto    = multiple spacing; single = font natural line height (ascent+descent+lineGap)
 *   - atLeast = max(font natural line height, specified value (twips))
 *   - exact   = fixed value (twips), does not grow with the font
 *
 * docGrid:
 *   - with type='lines' or 'linesAndChars', each line height rounds up to a linePitch multiple
 *   - space-before/space-after keep their face value (Word does not grid-snap them)
 */

import { isEastAsianFaceName, isJapaneseFaceName, isKoreanFaceName } from './text-layout/face-info'
import { openDocumentText } from './text-layout/open-document'
import type { DocGrid } from '@genoffice/docx-engine'
import {
  isCjkLetterChar,
  isDoubleByteChar,
  isEastAsianChar,
  isHangulChar,
  textHasEastAsian,
} from '@genoffice/docx-engine'
import {
  SEGOE_SYMBOL_CSS,
  declaredFaceMetrics,
  hasModelledLineBox,
  isArialUnicodeFace,
  isCalibriClassFace,
  isMsFixedPitchHead,
  isSegoeUiEmojiFace,
  isTimesClassFace,
} from './text-layout/font-rules/face-classes'
import { bundledPrimary, unloadedBundledPrimary } from './text-layout/font-rules/web-faces'
import { MONO_FONT_RE } from './font-name-patterns'
import {
  batangLineFactor,
  cjkDeclaredLineFactor,
  cssFontFamily,
  docLatinChainCss,
  forgetFontAvailability,
  inheritedMarkFace,
  isEaLayoutFace,
  krLineFactor,
  latinDescent,
  latinFace,
  latinLineFactor,
  latinSlotFace,
  lineHeightFactor,
  SIMSUN_GAP_CHAR_RE,
  simsunGapLineFactor,
  yaheiLineFactor,
} from './text-layout/font-rules'
import standInBoxesJson from './fonts/stand-in-boxes.json'

// ─── Font metrics interface (same interface as pptx-render/metrics.ts) ─────

export interface RunStyle {
  fontFamily: string
  fontSizePx: number
  bold: boolean
  italic: boolean
}

export interface FontMetrics {
  ascent: number
  descent: number
  lineHeight: number
}

export interface FontMetricsProvider {
  metrics(style: RunStyle): FontMetrics
  measure(text: string, style: RunStyle): number
}

// ─── Heuristic metrics (deterministic, unit-testable) ──────────────────────

/**
 * Estimate advance by character class (as a fraction of the font size).
 * Ported directly from pptx-render/metrics.ts HeuristicMetrics.
 */

/** Constant Latin advance for monospace families (chain-head name decides); null when proportional */
export function monospaceAdvanceEm(fontFamily: string): number | null {
  const head = fontFamily.split(',')[0]
  if (isMsFixedPitchHead(head)) return 0.5
  if (!MONO_FONT_RE.test(head)) return null
  return /consolas/i.test(head) ? 0.55 : 0.6
}

function charAdvanceEm(code: number): number {
  if (
    isHangul(code) ||
    (code >= 0x3000 && code <= 0x30ff) ||
    (code >= 0x3400 && code <= 0x9fff) ||
    (code >= 0xf900 && code <= 0xfaff) ||
    (code >= 0xff00 && code <= 0xff60) ||
    (code >= 0xffe0 && code <= 0xffe6)
  ) {
    return 1.0
  }
  if (code >= 0x1f000 || (code >= 0x2600 && code <= 0x27bf)) return 1.0
  // Arabic joins cursively: isolated-form widths overshoot, so keep the estimate low
  if (
    (code >= 0x0600 && code <= 0x06ff) ||
    (code >= 0x0750 && code <= 0x077f) ||
    (code >= 0x08a0 && code <= 0x08ff)
  ) {
    return 0.35
  }
  if ("iIlj.,:;'!|".includes(String.fromCharCode(code))) return 0.28
  if (' ftr'.includes(String.fromCharCode(code))) return 0.32
  if ('mwMW'.includes(String.fromCharCode(code))) return 0.82
  return 0.52
}

export class HeuristicMetrics implements FontMetricsProvider {
  metrics(style: RunStyle): FontMetrics {
    const s = style.fontSizePx
    // Windows metric line heights ((usWinAscent+usWinDescent) / UPM) differ a lot by font:
    //   PMingLiU / MingLiU (Traditional Chinese): ~1.0em (compact)
    //   SimSun / NSimSun / SimHei families (Simplified Chinese): ~1.3em (expanded)
    //   Calibri / Arial / Times (Western): ~1.2em (standard)
    // ascent/descent keep the generic ratios (0.8/0.2 em); lineHeight uses the per-font factor
    return {
      ascent: s * 0.8,
      descent: s * 0.2,
      lineHeight: s * lineHeightFactor(style.fontFamily),
    }
  }

  measure(text: string, style: RunStyle): number {
    const monoEm = monospaceAdvanceEm(style.fontFamily)
    let em = 0
    for (const ch of text) {
      const cp = ch.codePointAt(0) ?? 0
      const base = charAdvanceEm(cp)
      // fullwidth (CJK/emoji) chars stay 1.0em even in monospace fonts
      em += monoEm !== null && base < 1 ? monoEm : base
    }
    const boldFactor = style.bold && monoEm === null ? 1.04 : 1
    return em * style.fontSizePx * boldFactor
  }
}

/**
 * Real advances from an offscreen canvas in the renderer's family chain; the
 * per-class heuristic runs ~15% wide on Times-like faces and over-grows autofit columns.
 */
export class CanvasMetrics implements FontMetricsProvider {
  private cache = new Map<string, number>()
  constructor(
    private ctx: CanvasRenderingContext2D,
    private fallback: FontMetricsProvider = new HeuristicMetrics(),
  ) {
    // document text renders unkerned by default (Word semantics, styles.css)
    if ('fontKerning' in ctx) ctx.fontKerning = 'none'
    // advances measured before a @font-face resolved are the fallback face's
    if (typeof document !== 'undefined' && document.fonts?.addEventListener) {
      document.fonts.addEventListener('loadingdone', () => this.cache.clear())
    }
  }

  metrics(style: RunStyle): FontMetrics {
    return this.fallback.metrics(style)
  }

  measure(text: string, style: RunStyle): number {
    const font = `${style.italic ? 'italic ' : ''}${style.bold ? 'bold ' : ''}${style.fontSizePx}px ${cssFontFamily(style.fontFamily)}`
    const key = `${font}|${text}`
    let w = this.cache.get(key)
    if (w === undefined) {
      this.ctx.font = font
      w = this.ctx.measureText(text).width
      if (this.cache.size > 20000) this.cache.clear()
      this.cache.set(key, w)
    }
    return w
  }
}

function canvasContext(): CanvasRenderingContext2D | null {
  try {
    return typeof document === 'undefined'
      ? null
      : document.createElement('canvas').getContext('2d')
  } catch {
    return null
  }
}

let canvasMetricsSingleton: FontMetricsProvider | null | undefined
/** CanvasMetrics when a 2D context exists (renderer), else null (jsdom/tests) */
export function canvasMetrics(): FontMetricsProvider | null {
  if (canvasMetricsSingleton !== undefined) return canvasMetricsSingleton
  const ctx = canvasContext()
  canvasMetricsSingleton =
    ctx && typeof ctx.measureText === 'function' ? new CanvasMetrics(ctx) : null
  return canvasMetricsSingleton
}

/** Word DFonts win ascent (em) of the symbol bullet faces */
const SYMBOL_FACE_ASCENT: Record<string, number> = {
  symbol: 1.0054,
  'segoe ui symbol': 1.0791,
  wingdings: 0.8989,
  'wingdings 2': 0.8433,
  'wingdings 3': 0.9277,
  webdings: 0.7998,
}

/**
 * Single-line factor of a line mixing Latin faces at one size: Word lays it at
 * the max ascent plus the max descent of the faces on it (probe 2026-11-07: a
 * Century Gothic 10pt title with a Calibri 10pt page number paces 14.64 at 1.15
 * against 14.1 for Century Gothic alone = (1.006 + 0.269) x 10 x 1.15; Times +
 * Calibri stays Calibri's 1.22). Never below the tallest face's own factor.
 */
export function mixedLatinLineFactor(fonts: ReadonlyArray<string>): number {
  let factor = 0
  let ascent = 0
  let descent = 0
  for (const font of fonts) {
    factor = Math.max(factor, lineHeightFactor(font))
    const face = latinFace(font)
    if (!face || face.metrics.eastAsian) continue
    ascent = Math.max(ascent, face.metrics.ascent)
    descent = Math.max(descent, face.metrics.descent)
  }
  return Math.round(Math.max(factor, ascent + descent) * 1e4) / 1e4
}

/**
 * Extra height (pt) a number/bullet drawn larger than the item text adds to the
 * first line, or null when it is not larger. Word probe 2026-10-20 (TNR 12pt
 * text, level w:sz 14/18/22pt): the line grows by the marker's ascent + line gap
 * over the text's, (markerPt - textPt) x (factor - descent), and the paragraph's
 * line multiple does not scale that growth (+1.8/+5.9/+9.2pt at single, 1.5 and
 * double spacing alike).
 */
export function oversizedMarkerExtraPt(
  markerFont: string,
  markerPt: number,
  textPt: number,
): number | null {
  if (!(markerPt > textPt) || !(textPt > 0)) return null
  const extra = (markerPt - textPt) * (lineHeightFactor(markerFont) - latinDescent(markerFont))
  return Math.round(extra * 1000) / 1000
}

/**
 * Ascent (em) Word stands a marker glyph on: the symbol faces' win ascent, or
 * for an East Asian face the hhea ascent plus half the leading its x1.3 line
 * adds (probe 2026-11-02: Malgun Gothic bullet on a Times New Roman 11pt item
 * 16.56 = 1.2886 x 11 + 2.38, Batang / Gulim 13.44 = 1.0098 x 11 + 2.38)
 */
function markerFaceAscent(font: string): number | undefined {
  const key = font.trim().toLowerCase()
  const symbol = SYMBOL_FACE_ASCENT[key]
  if (symbol !== undefined) return symbol
  const metrics = declaredFaceMetrics(font)
  // Arial Unicode MS markers are unprobed; the CJK vendor faces follow the split
  if (!metrics?.eastAsian || !isCjkFontName(font) || isArialUnicodeFace(key)) return undefined
  const { ascent, descent } = metrics
  return ascent + (lineHeightFactor(font) - ascent - descent) / 2
}

/** a level face whose glyph stands taller than Latin text faces (symbol and East Asian faces) */
export function markerFaceLifts(font: string): boolean {
  return markerFaceAscent(font) !== undefined
}

/**
 * Natural height (pt) of a list item's first line when a symbol-font bullet
 * sits on it, or null when the text face alone is as tall. Word sizes a line
 * by the max ascent plus the max descent of the faces on it, so Symbol (ascent
 * 1.0054em, above every Latin text face) lifts bullet lines: probe 2026-09-17,
 * Aptos 10pt text + Symbol 10.5pt bullet at 1.15 = 15.12-15.36pt against
 * 14.0pt text-only lines; Cambria 11pt + Symbol 11pt = 15.36-15.6 vs 14.9.
 */
export function symbolBulletLinePt(
  symbolFont: string,
  markerPt: number,
  textFont: string | null,
  textPt: number,
): number | null {
  const asc = markerFaceAscent(symbolFont)
  if (asc === undefined || !(markerPt > 0) || !(textPt > 0)) return null
  const line = asc * markerPt + latinDescent(textFont ?? '') * textPt
  if (line <= lineHeightFactor(textFont ?? 'Calibri') * textPt + 0.01) return null
  return Math.round(line * 1000) / 1000
}

// ─── opentype.js precise metrics (interface ready; F1 falls back to heuristics) ─────

export interface OpentypeFontLike {
  unitsPerEm: number
  ascender: number
  descender: number
  getAdvanceWidth(text: string, fontSize: number): number
  charToGlyphIndex?(char: string): number
}

export class OpentypeMetrics implements FontMetricsProvider {
  constructor(
    private fontResolver: (style: RunStyle) => OpentypeFontLike | undefined,
    private fallback: FontMetricsProvider = new HeuristicMetrics(),
  ) {}

  metrics(style: RunStyle): FontMetrics {
    const font = this.fontResolver(style)
    if (!font) return this.fallback.metrics(style)
    const scale = style.fontSizePx / font.unitsPerEm
    const ascent = font.ascender * scale
    const descent = Math.abs(font.descender) * scale
    return { ascent, descent, lineHeight: ascent + descent }
  }

  measure(text: string, style: RunStyle): number {
    const font = this.fontResolver(style)
    if (!font) return this.fallback.measure(text, style)
    try {
      if (font.charToGlyphIndex) {
        for (const ch of text) {
          if (font.charToGlyphIndex(ch) === 0) return this.fallback.measure(text, style)
        }
      }
      return font.getAdvanceWidth(text, style.fontSizePx)
    } catch {
      return this.fallback.measure(text, style)
    }
  }
}

// ─── Line-height rule semantics ─────────────────────────────────────────────

const TWIPS_TO_PX = 96 / 1440

/**
 * ε (fraction of a cell) so a needed height marginally past a cell boundary
 * stays in the lower cell. Word derives suggested grid pitches from the body
 * font's own line height, so needed == pitch equality is common and our probed
 * factors can land a hair above it: Yu Mincho 10.5pt on a 302-twip grid takes
 * one cell in Word although 1.44em = 15.12pt exceeds the 15.1pt pitch
 * (probe 2026-08-22); known legitimate two-cell cases sit ≥3% past the pitch.
 */
const GRID_SNAP_EPS = 0.004

/**
 * Ceil a line height UP to whole grid cells. Single source for docGrid line
 * snapping: the pagination model calls it directly and cssGridLineExpr emits
 * the same formula (ε included) as a CSS round(up) — keep them in lockstep.
 */
export function snapLineToPitch(heightPx: number, pitchPx: number): number {
  return pitchPx > 0 ? Math.ceil(heightPx / pitchPx - GRID_SNAP_EPS) * pitchPx : heightPx
}

/**
 * Compute a single line's height (px) per Word's line-height rules.
 *
 * @param naturalLineH the font's natural line height (ascent+descent, px)
 * @param lineRule      'auto'|'atLeast'|'exact'
 * @param lineRawTwips  raw w:spacing w:line twips (auto = multiple of 240; atLeast/exact = absolute)
 * @param docGrid       the section docGrid (optional; when present, round to linePitch)
 */
export function computeLineHeight(
  naturalLineH: number,
  lineRule: 'auto' | 'atLeast' | 'exact' | undefined,
  lineRawTwips: number | undefined,
  docGrid: DocGrid | undefined,
): number {
  // typed line grid (Word probe 2026-08-22, 15 cases): an auto multiple scales
  // the PITCH (the product does not re-snap: 1.5 x 15.6pt grid = 23.4pt), and
  // no non-exact line is shorter than its grid-snapped single height —
  // max(rule value, snapped single). Holds for mult < 1 and atLeast; exact
  // never snaps; w:snapToGrid=0 opts a paragraph out (pitch 0 here).
  const pitchPx =
    docGrid && (docGrid.type === 'lines' || docGrid.type === 'linesAndChars') && docGrid.linePitch
      ? docGrid.linePitch * TWIPS_TO_PX
      : 0
  const snapped = snapLineToPitch(naturalLineH, pitchPx)

  if (lineRule === 'exact' && lineRawTwips !== undefined) {
    // fixed line height: use the specified value regardless of font size
    return lineRawTwips * TWIPS_TO_PX
  }
  if (lineRule === 'atLeast' && lineRawTwips !== undefined) {
    // line="0" atLeast is the one non-exact form that never snaps: natural
    // height (Word probe 2026-09-11: Meiryo UI 10pt on a 350-twip grid lays
    // 16.56pt, not the 17.5pt cell; atLeast 200/330 still floor at the cell)
    if (lineRawTwips === 0) return naturalLineH
    // at least: the floor value at face value, but never below the snapped
    // single height (snapped == natural when there is no grid)
    return Math.max(lineRawTwips * TWIPS_TO_PX, snapped)
  }
  if (lineRule === 'auto' && lineRawTwips !== undefined) {
    const mult = lineRawTwips / 240
    return pitchPx > 0 ? Math.max(mult * pitchPx, snapped) : naturalLineH * mult
  }
  // default (no lineRule): single spacing on the grid
  return snapped
}

/**
 * Ascent/descent of a family chain's primary face, as percentages of the em
 * (canvas fontBoundingBox at 1000px for 0.1% granularity). Fuels the typed-grid
 * strut alias in doc-style-css; null where canvas metrics are unavailable (jsdom).
 */
const chainMetricsCache = new Map<string, { ascentPct: number; descentPct: number } | null>()

/** embedded faces registered/revoked: their availability and every chain's metrics may have flipped */
export function noteEmbeddedFontsChanged(families: readonly string[]): void {
  forgetFontAvailability(families)
  chainMetricsCache.clear()
}

const FACE_VARIANTS = ['', 'bold ', 'italic ', 'bold italic ']

/**
 * Load the bundled stand-in cuts the document's families chain to before the
 * first layout pass. Canvas measures the system fallback while a web face is
 * still loading, and the autofit table pass sizes columns from those advances
 * once: Arial Bold runs ~15% wider than Carlito, so a Calibri table's saved
 * Word grid grew its short-word columns and wrapped the long ones. A face that
 * never loads must not hold the open.
 */
export async function loadBundledStandIns(
  families: Iterable<string>,
  timeoutMs = 3000,
): Promise<void> {
  const fonts = typeof document === 'undefined' ? undefined : document.fonts
  if (!fonts?.check || !fonts.load) return
  const loads: Promise<unknown>[] = []
  const seen = new Set<string>()
  for (const family of families) {
    const face = bundledPrimary(cssFontFamily(family))
    if (!face || seen.has(face)) continue
    seen.add(face)
    for (const variant of FACE_VARIANTS) {
      const spec = `${variant}16px "${face}"`
      try {
        if (!fonts.check(spec)) loads.push(fonts.load(spec).catch(() => undefined))
      } catch {
        /* no FontFaceSet loader */
      }
    }
  }
  if (loads.length === 0) return
  await Promise.race([
    Promise.all(loads),
    new Promise<void>((resolve) => setTimeout(resolve, timeoutMs)),
  ])
}

const metricsPendingFaces = new Set<string>()

/** True once a face measured before it loaded is in: callers re-derive their metric CSS. */
export function takeSettledMetricFaces(): boolean {
  let settled = false
  for (const family of metricsPendingFaces) {
    try {
      if (!document.fonts.check(`16px "${family}"`)) continue
    } catch {
      continue
    }
    metricsPendingFaces.delete(family)
    settled = true
  }
  if (settled) chainMetricsCache.clear()
  return settled
}

/**
 * The grid strut's box for an East Asian body face. Word splits the line as
 * hhea ascent + (factor - hhea ascent - hhea descent) / 2 (probes 2026-10-22
 * DengXian / PMingLiU, 2026-11-05 SimSun / MS Mincho / Yu Mincho body and
 * textbox lines), so the strut box is that split of the face's own hhea
 * metrics and totals the line factor. A canvas measurement of the chain
 * returned the stand-in's box instead (SimSun's Latin twin 1.16/0.288 put the
 * first baseline 1.1pt low, MS Mincho GO 0.69/0.25 put it 1.5pt high);
 * faces the resolver does not know still fall back to it.
 */
export function gridStrutMetricsPct(
  face: string,
): { ascentPct: number; descentPct: number } | null {
  const m = declaredFaceMetrics(face)
  if (m && m.ascent > 0 && m.descent >= 0) {
    const lead = (lineHeightFactor(face) - m.ascent - m.descent) / 2
    const pct = (v: number) => Math.round(v * 10000) / 100
    return { ascentPct: pct(m.ascent + lead), descentPct: pct(m.descent + lead) }
  }
  return fontChainMetricsPct(cssFontFamily(face))
}

export function fontChainMetricsPct(
  chain: string,
): { ascentPct: number; descentPct: number } | null {
  const cached = chainMetricsCache.get(chain)
  if (cached !== undefined) return cached
  const pending = unloadedBundledPrimary(chain)
  if (pending) {
    metricsPendingFaces.add(pending)
    try {
      void document.fonts.load(`16px "${pending}"`).catch(() => {})
    } catch {
      /* no FontFaceSet loader */
    }
  }
  let out: { ascentPct: number; descentPct: number } | null = null
  try {
    const ctx = document.createElement('canvas').getContext('2d')
    if (ctx) {
      ctx.font = `1000px ${chain}`
      const m = ctx.measureText('M')
      if (typeof m.fontBoundingBoxAscent === 'number' && m.fontBoundingBoxAscent > 0) {
        out = { ascentPct: m.fontBoundingBoxAscent / 10, descentPct: m.fontBoundingBoxDescent / 10 }
      }
    }
  } catch {
    out = null
  }
  if (!pending) chainMetricsCache.set(chain, out)
  return out
}

/** Text contains complex-script characters (Arabic/Hebrew/Syriac/Thaana/NKo/Indic/Thai), i.e. the w:cs font slot applies */
export function textHasComplexScript(text: string): boolean {
  return /[\u0590-\u05FF\u0600-\u077F\u0780-\u07FF\u08A0-\u08FF\u0900-\u0DFF\u0E00-\u0E7F\uFB1D-\uFDFF\uFE70-\uFEFF]/.test(
    text,
  )
}

export interface RunFace {
  fontAscii?: string | null
  font?: string | null
  sizeHalfPoints?: number | null
  bold?: boolean | null
  italic?: boolean | null
  /** style-resolved values for the slots the run leaves to inheritance */
  effective?: {
    fonts?: { ascii?: string }
    sizeHalfPoints?: number
    bold?: boolean
    italic?: boolean
  }
}

/** w:w horizontal scaling approximated as letter-spacing (em, negative = condensed).
 *  With a face the per-character share is the run's measured average advance, so
 *  the run grows by exactly (w - 100)% like Word's glyph scaling (a 10.5pt Times
 *  line justified at 105% had zero slack in Word: a flat 0.52em average runs
 *  0.25pt per word wide and wraps one word early); without one the wide/narrow
 *  glyph-mix estimate stands in */
export function charScaleEm(text: string, scalePct: number, face?: RunFace | null): number {
  let n = 0
  for (const _ of text) n++
  const eff = face?.effective
  const family = face?.fontAscii ?? face?.font ?? eff?.fonts?.ascii
  let avg: number
  if (family && n > 0) {
    const half = face?.sizeHalfPoints ?? eff?.sizeHalfPoints ?? DEFAULT_FONT_SIZE_PT * 2
    const style: RunStyle = {
      fontFamily: family,
      fontSizePx: (half / 2) * (96 / 72),
      bold: face?.bold ?? eff?.bold ?? false,
      italic: face?.italic ?? eff?.italic ?? false,
    }
    avg = (canvasMetrics() ?? defaultMetrics).measure(text, style) / n / style.fontSizePx
  } else {
    let sum = 0
    for (const ch of text) sum += isDoubleByteChar(ch.codePointAt(0)!) ? 1 : 0.52
    avg = n > 0 ? sum / n : 0.52
  }
  return Math.round((scalePct / 100 - 1) * avg * 1000) / 1000
}

/** letter-spacing CSS value for a run's w:spacing / w:w pair (null = none) */
/**
 * Word pair-kerns a run only under w:kern and from its half-point threshold up
 * (Chromium kerns by default). undefined = no w:kern in effect; an unknown run
 * size counts as reaching the threshold.
 */
export function wordKerns(
  kernHalfPoints: number | undefined,
  sizeHalfPoints: number | undefined,
): boolean | undefined {
  if (kernHalfPoints === undefined) return undefined
  return kernHalfPoints > 0 && (sizeHalfPoints === undefined || sizeHalfPoints >= kernHalfPoints)
}

export function fontKerningCss(run: {
  kernHalfPoints?: number
  sizeHalfPoints?: number
}): 'normal' | 'none' | null {
  const on = wordKerns(run.kernHalfPoints, run.sizeHalfPoints)
  return on === undefined ? null : on ? 'normal' : 'none'
}

export function runLetterSpacingCss(
  run: RunFace & {
    text: string
    charSpacingTwips?: number
    charScalePct?: number
  },
): string | null {
  const spacingPt = run.charSpacingTwips ? run.charSpacingTwips / 20 : 0
  const scaleEm = run.charScalePct ? charScaleEm(run.text, run.charScalePct, run) : 0
  if (spacingPt && scaleEm) return `calc(${spacingPt}pt + ${scaleEm}em)`
  if (spacingPt) return `${spacingPt}pt`
  if (scaleEm) return `${scaleEm}em`
  return run.charSpacingTwips === 0 ? '0' : null
}

/** Text contains CJK characters (decides the line-height factor: CJK lines measure ~1.3em per Chinese font metrics) */
/** a run declaring a Latin face and a different East Asian face lays two boxes on its line */
export function facesDiffer(ascii: unknown, ea: unknown): boolean {
  if (!ea) return false
  if (!ascii) return true
  return String(ascii).trim().toLowerCase() !== String(ea).trim().toLowerCase()
}

export function textHasCjk(text: string): boolean {
  for (const ch of text) {
    if (isCjk(ch.codePointAt(0) ?? 0)) return true
  }
  return false
}

/** Text contains hangul (Korean paragraphs take the Korean line factor, not the 1.3 Chinese one) */
export function textHasHangul(text: string): boolean {
  for (const ch of text) {
    if (isHangul(ch.codePointAt(0) ?? 0)) return true
  }
  return false
}

/** Text carries non-CJK ink: any character that is neither CJK nor whitespace */
export function textHasLatinInk(text: string): boolean {
  for (const ch of text) {
    if (!/\s/.test(ch) && !isCjk(ch.codePointAt(0) ?? 0)) return true
  }
  return false
}

/**
 * Maximal CJK stretches of a text (UTF-16 offsets); any other character,
 * spaces included, ends one. lineFactorLive lifts each stretch to its East
 * Asian factor.
 */
export function cjkScriptRanges(text: string): Array<{ from: number; to: number }> {
  const ranges: Array<{ from: number; to: number }> = []
  let start = -1
  let end = -1
  let i = 0
  for (const ch of text) {
    if (isCjk(ch.codePointAt(0) ?? 0)) {
      if (start < 0) start = i
      end = i + ch.length
    } else if (start >= 0) {
      // a space between CJK words is the ascii font's glyph (Word sizes the
      // line by it like any Latin glyph); inside the lifted range its
      // Latin-metric box at the CJK line-height would protrude below it
      ranges.push({ from: start, to: end })
      start = -1
    }
    i += ch.length
  }
  if (start >= 0) ranges.push({ from: start, to: end })
  return ranges
}

/** Per-paragraph --doc-line-factor value by script (approximates Word's max-of-inline-fonts line height) */
export function paraLineFactorCss(text: string): string {
  if (textHasHangul(text)) return `var(--doc-line-factor-kr,${batangLineFactor()})`
  if (textHasCjk(text)) return 'var(--doc-line-factor-cjk,1.7)'
  return 'var(--doc-line-factor-latin,1.2)'
}

/**
 * A CJK-named face declared in the ascii/hAnsi slots only: Word draws the CJK
 * glyphs with the inherited eastAsia face and this face only renders spaces,
 * digits and Latin punctuation, lifting just the lines that contain one
 * (probe 2026-09-30: ascii-only YaHei over a DengXian theme = 14.16pt pure-CJK
 * lines, 18.0pt once a digit is on the line). Returns that ascii face.
 */
export function asciiOnlyCjkFace(attrs: {
  font?: string | null
  fontAscii?: string | null
  eastAsiaFont?: string | null
}): string | null {
  const ascii = attrs.fontAscii
  if (!ascii || attrs.eastAsiaFont || attrs.font !== ascii || !isCjkFontName(ascii)) return null
  return ascii
}

/** Inked non-CJK stretches (UTF-16 offsets): the complement of cjkScriptRanges minus whitespace-only gaps */
export function nonCjkInkRanges(text: string): Array<{ from: number; to: number }> {
  const ranges: Array<{ from: number; to: number }> = []
  let prev = 0
  const push = (from: number, to: number) => {
    if (to > from && /\S/.test(text.slice(from, to))) ranges.push({ from, to })
  }
  for (const r of cjkScriptRanges(text)) {
    push(prev, r.from)
    prev = r.to
  }
  push(prev, text.length)
  return ranges
}

/** Faces sizing a non-CJK run's line: complex-script text takes its w:cs face
 *  (Word lays Arabic in the cs font, Arial 1.15 under the Office theme, not the
 *  ascii Calibri 1.22); the ascii face still counts when Latin letters share the run */
export function runLatinFactorFaces(
  text: string,
  fontAscii: string | null | undefined,
  csFont: string | null | undefined,
  cs?: boolean | null,
): Array<string | null | undefined> {
  if (!csFont) return [fontAscii]
  // an rtl run is complex script throughout: Word lays its Latin in the cs face too
  if (cs) return [csFont]
  if (!textHasComplexScript(text)) return [fontAscii]
  return /[A-Za-z]/.test(text) ? [csFont, fontAscii] : [csFont]
}

/**
 * Word for Mac draws emoji-presentation characters in Apple Color Emoji whatever
 * face the run declares (probe 2026-11-07: no face, Arial Unicode MS; a declared
 * Segoe UI Emoji run keeps the text line) and the line takes its hhea box 800/-250.
 * Text-presentation pictographs (U+25B6, U+2714, U+26A0, probe 2026-11-08) go
 * the same way, with or without U+FE0F.
 */
export const EMOJI_LINE_FACTOR = 1.3125
const EMOJI_PRESENTATION_RE = /^\p{Emoji_Presentation}$/u
const PICTOGRAPH_RE = /^\p{Extended_Pictographic}$/u
/** WGL4 shapes, arrows and suits the Office Latin faces carry themselves */
const WGL4_SYMBOL_RE =
  /[\u2190-\u2195\u21a8\u25a0\u25a1\u25aa\u25ab\u25ac\u25b2\u25ba\u25bc\u25c4\u25ca\u25cb\u25cf\u25d8\u25d9\u25e6\u263a-\u263c\u2640\u2642\u2660\u2663\u2665\u2666\u266a\u266b]/
/** the WGL4 shapes Calibri lacks (probe 2026-11-08: U+25A0 U+25B2 U+25BC U+25BA; U+25C4 by symmetry) */
const CALIBRI_MISSING_SHAPE_RE = /[\u25a0\u25b2\u25bc\u25ba\u25c4]/

export type SymbolFallback = 'emoji' | 'segoe' | 'arial'

/** the face Word for Mac draws a symbol a Latin run's face lacks with (probes
 *  2026-11-07/08): emoji-presentation and pictographs -> Apple Color Emoji; arrows,
 *  geometric shapes, misc symbols and dingbats outside WGL4 -> Segoe UI Symbol;
 *  the WGL4 shapes Calibri lacks -> Arial */
export function symbolFallbackKind(
  cp: number,
  nextCp: number,
  textFace: string | null | undefined,
): SymbolFallback | null {
  const ch = String.fromCodePoint(cp)
  if (EMOJI_PRESENTATION_RE.test(ch)) return 'emoji'
  const pictograph = PICTOGRAPH_RE.test(ch)
  if (pictograph && nextCp === 0xfe0f) return 'emoji'
  // ballot boxes: Segoe UI Symbol with or without a declared face (probe 2026-11-07)
  if (cp >= 0x2610 && cp <= 0x2612) return 'segoe'
  // WGL4 glyphs (suits, smileys, gender signs, arrows) are the text face's own
  if (WGL4_SYMBOL_RE.test(ch)) {
    return CALIBRI_MISSING_SHAPE_RE.test(ch) && textFace && isCalibriClassFace(textFace)
      ? 'arial'
      : null
  }
  if (pictograph) return cp >= 0x2300 && cp < 0x3000 ? 'emoji' : null
  if (
    (cp >= 0x2190 && cp <= 0x21ff) ||
    (cp >= 0x25a0 && cp <= 0x27bf) ||
    (cp >= 0x2b00 && cp <= 0x2bff)
  )
    return 'segoe'
  return null
}

export interface FallbackGlyphRange {
  from: number
  to: number
  kind: SymbolFallback
}

const EMOJI_JOINER = new Set([0xfe0f, 0x200d, 0x20e3])

/** UTF-16 ranges of the glyphs Word draws with a fallback face; `family` is the run's
 *  declared ascii face (a Segoe UI Emoji run keeps its emoji on the text line),
 *  `textFace` the face the glyphs would otherwise draw in */
export function fallbackGlyphRanges(
  text: string,
  family?: string | null,
  textFace: string | null | undefined = family,
): FallbackGlyphRange[] {
  const out: FallbackGlyphRange[] = []
  const noEmoji = !!family && isSegoeUiEmojiFace(family)
  for (let i = 0; i < text.length;) {
    const cp = text.codePointAt(i)!
    const len = cp > 0xffff ? 2 : 1
    const nextCp = i + len < text.length ? text.codePointAt(i + len)! : -1
    let kind = symbolFallbackKind(cp, nextCp, textFace)
    if (kind === 'emoji' && noEmoji) kind = null
    const last = out[out.length - 1]
    if (kind) {
      if (last && last.to === i && last.kind === kind) last.to = i + len
      else out.push({ from: i, to: i + len, kind })
    } else if (last && last.to === i && EMOJI_JOINER.has(cp)) {
      // variation selectors, ZWJ and keycap marks stay with the glyph they modify
      last.to = i + len
    }
    i += len
  }
  return out
}

export function symbolFallbackLineFactor(kind: SymbolFallback): number | null {
  if (kind === 'emoji') return EMOJI_LINE_FACTOR
  if (kind === 'segoe') return lineHeightFactor('Segoe UI Symbol')
  return null
}

export function symbolFallbackFontFamily(kind: SymbolFallback): string {
  if (kind === 'emoji') return "'Apple Color Emoji','Segoe UI Emoji','Noto Color Emoji'"
  if (kind === 'segoe') return SEGOE_SYMBOL_CSS
  return docLatinChainCss('Arial')
}

/** line factor of the fallback face Word draws a run's symbols with, null when none lifts */
export function fallbackGlyphLineFactor(text: string, family?: string | null): number | null {
  let max: number | null = null
  for (const r of fallbackGlyphRanges(text, family)) {
    const f = symbolFallbackLineFactor(r.kind)
    if (f !== null && (max === null || f > max)) max = f
  }
  return max
}

/** every glyph of the run is emoji: the declared face draws nothing and lends no line box */
export function textInkIsEmoji(text: string): boolean {
  const ranges = fallbackGlyphRanges(text)
  if (!ranges.some((r) => r.kind === 'emoji')) return false
  let i = 0
  for (const r of ranges) {
    if (r.kind !== 'emoji') return false
    if (/\S/.test(text.slice(i, r.from))) return false
    i = r.to
  }
  return !/\S/.test(text.slice(i))
}

export const isKoreanFontName = isKoreanFaceName
export const isJapaneseFontName = isJapaneseFaceName

/**
 * Hangul in a run whose empty EA theme slot resolved (by themeFontLang) to a
 * non-Korean CJK face: Word's per-glyph fallback draws it in Malgun Gothic under
 * a Japanese face, Batang otherwise, and that face sizes the line. Null when the
 * rule does not apply (the inherited --doc-line-factor-kr stands).
 */
export function emptyEaSlotHangulFactor(
  family: string | null | undefined,
  text: string,
): number | null {
  if (!family || !isCjkFontName(family) || isKoreanFontName(family)) return null
  return textHasHangul(text) ? krLineFactor(family) : null
}

/**
 * Word line spacing → editing-canvas CSS line-height value (shared by extensions/docStyleCss).
 *
 * Line rules: auto = multiple × font natural line height; atLeast = max(natural, N); exact = N.
 * Auto/multiple emits a UNITLESS number (var(--doc-line-factor) × multiple):
 * numbers inherit by value and rescale with each run's own font size, matching
 * Word's tallest-run-per-line rule (a length computed at the paragraph would
 * inherit into larger runs and collapse their lines).
 *
 * docGrid (Word probe, 2026-08-22): in sections with a typed line grid
 * (w:docGrid lines/linesAndChars) an auto multiple scales the PITCH (the
 * product does not re-snap) and no non-exact line is shorter than its
 * grid-snapped single height: max(mult x pitch, snapped single). atLeast is
 * max(value, snapped single); exact never snaps; w:snapToGrid=0 opts a
 * paragraph out. (The earlier snap-then-multiply order was an LO artifact.)
 * round() needs a length, so grid snapping cannot be unitless: docStyleCss
 * declares --doc-line-grid (cssGridLineExpr) on every element of typed-grid
 * documents, and auto/multiple line heights resolve it ahead of the unitless
 * fallback. Per-element declaration keeps per-paragraph --doc-line-factor /
 * --doc-grid-pitch overrides live (a value inherited from .doc-page would
 * freeze the vars substituted there).
 */
const GRID_PITCH = 'var(--doc-grid-pitch,0.0001px)'

/** grid-snapped single-line height; docStyleCss assigns it to --doc-line-grid in typed-grid docs.
 *  CSS mirror of snapLineToPitch (same ε); 1em is the paragraph's own font size —
 *  Word snaps by the tallest run per line, one line-height per paragraph is our simplification. */
export function cssGridLineExpr(): string {
  return `round(up, calc(var(--doc-line-factor,1.2) * 1em - ${GRID_PITCH} * ${GRID_SNAP_EPS}), ${GRID_PITCH})`
}

/** typed-grid padding of an inline picture/chart block (--doc-obj-h = object
 *  height): the line takes whole cells and the object sits centred in them
 *  (Word: a 279pt chart on an 18pt grid holds 288pt, 4.5pt free above and below) */
export function cssGridObjectPadExpr(): string {
  return `calc((round(up, var(--doc-obj-h) - ${GRID_PITCH} * ${GRID_SNAP_EPS}, ${GRID_PITCH}) - var(--doc-obj-h)) / 2)`
}

/** grid line height with the auto multiple applied (Word probe 2026-08-22):
 *  max(mult x pitch, grid-snapped single). docStyleCss assigns it to
 *  --doc-line-max in typed-grid docs; emitters keep --doc-line-mult on the
 *  same element in sync with the multiple they resolve. */
export function cssGridLineMaxExpr(): string {
  return `max(calc(${GRID_PITCH} * var(--doc-line-mult,1)), ${cssGridLineExpr()})`
}

/** the auto multiple inline lifts (SimSun gap, cs digit faces) pace against:
 *  null for exact lines (pinned), 'inherited' when the paragraph declares no
 *  spacing of its own and the style / docDefaults multiple arrives as --doc-line-mult */
export function inlineLiftLineMultiple(
  lineRule: 'auto' | 'atLeast' | 'exact' | null | undefined,
  lineRawTwips: number | null | undefined,
  lineSpacing: number | null | undefined,
): number | 'inherited' | null {
  if (lineRule === 'exact') return null
  if (lineSpacing) return lineSpacing
  if (lineRule === 'auto' && lineRawTwips) return lineRawTwips / 240
  return lineRule || lineRawTwips ? 1 : 'inherited'
}

/** SimSun-gap lifted line (・/〜 runs): max(mult x pitch, grid-snapped boosted
 *  single), mirroring computeLineHeight with the YaHei lift as natural
 *  height. --doc-grid-single-mult (1 in typed-grid docs, absent otherwise)
 *  keeps the non-grid value at boosted x mult. */
export function cssSimsunGapLineExpr(multiple: number | 'inherited'): string {
  const single = `round(up, calc(${yaheiLineFactor()} * 1em - ${GRID_PITCH} * ${GRID_SNAP_EPS}), ${GRID_PITCH})`
  if (multiple === 1) return single
  // no own spacing: the style / docDefaults multiple arrives as --doc-line-mult
  const m = multiple === 'inherited' ? 'var(--doc-line-mult,1)' : String(multiple)
  return `max(calc(${GRID_PITCH} * ${m}), calc(${single} * var(--doc-grid-single-mult,${m})))`
}

/** single-spacing line height: snapped length in typed-grid docs, unitless factor otherwise */
export function cssGridLineBase(): string {
  return 'var(--doc-line-grid,var(--doc-line-factor,1.2))'
}

/**
 * Paragraph strut font-size from the smallest run-declared size. Word sizes
 * each line by the runs on it, and every run already lays its own inline box
 * (per-element line rule), so the strut only has to carry the lines made of
 * the smallest runs: a max-size strut gave the 9.5pt body lines of an
 * `11pt title <w:br/> 9.5pt body` callout the title's 11pt pitch (prod corpus
 * 023, 16.1 vs Word 13.7pt). A uniform run size grows the strut too: left at
 * the inherited size, a larger run rides above the strut's half-leading and
 * the union line box outgrows it by ~1px/line (prod-sas 003).
 */
export function strutFontCss(halfPoints: number, mixed = false): string[] {
  const pt = `${halfPoints / 2}pt`
  // mixed: runs inheriting their size take --doc-base-fs back (styles.css .doc-strut-mixed)
  const strut = mixed ? `min(${pt},var(--doc-base-fs,${pt}))` : pt
  return [`--doc-strut:${strut}`, 'font-size:var(--doc-strut)']
}

/**
 * Paragraph space before/after (pt) as a CSS margin value. Word keeps
 * paragraph spacing at face value even under a typed grid (probe 2026-08-13:
 * before 4.5/10/18/24pt all render as written on an 18pt grid); the earlier
 * round-down-to-grid-cells behavior was an LO artifact.
 */
export function cssGridSpacingPt(pt: number): string {
  return `${pt.toFixed(1)}pt`
}

/**
 * Word's HTML "auto" paragraph spacing (w:beforeAutospacing/w:afterAutospacing):
 * a fixed 14pt regardless of font size (probe 2026-08-23: 11pt and 17pt runs both
 * 14pt; adjacent auto margins collapse; 0 between two list items).
 */
export const WORD_AUTO_SPACING_PT = 14

/**
 * settings.xml w:compat doNotUseHTMLParagraphAutoSpacing (Word probe 2026-10-22):
 * adjacent paragraphs' space-after and space-before add up instead of merging to
 * the larger one, and HTML auto spacing is 0. The space-before is laid out as
 * padding so CSS cannot collapse it with the previous margin.
 */
export function additiveParagraphSpacing(): boolean {
  return openDocumentText().additiveSpacing
}
export function spaceBeforeProp(
  additive = additiveParagraphSpacing(),
): 'padding-top' | 'margin-top' {
  return additive ? 'padding-top' : 'margin-top'
}
export function autoSpacingPt(additive = additiveParagraphSpacing()): number {
  return additive ? 0 : WORD_AUTO_SPACING_PT
}

/** space-only run text: never sizes a line, the paragraph mark does (Word probe 2026-09-11) */
export const SPACE_ONLY_RE = /^[ \u00a0\u3000]+$/

/** spaces / tabs only: such a run never sizes the line it shares with ink, whatever
 *  its face or size; a no-break space does (Word probe 2026-11-08) */
export const WHITESPACE_RUN_RE = /^[ \t]+$/
export const isWhitespaceRun = (text: string): boolean => WHITESPACE_RUN_RE.test(text)

/** an exact rule caps every run's line box at the rule (--doc-lh-cap, styles.css .doc-lh-fixed span) */
export function cssExactLineCap(
  lineRule: 'auto' | 'atLeast' | 'exact' | undefined,
  lineRawTwips: number | undefined,
): string | null {
  return lineRule === 'exact' && lineRawTwips ? `${(lineRawTwips / 20).toFixed(1)}pt` : null
}

export function cssLineHeight(
  lineRule: 'auto' | 'atLeast' | 'exact' | undefined,
  lineRawTwips: number | undefined,
  lineSpacing: number | undefined,
): string | null {
  const FACTOR = 'var(--doc-line-factor,1.2)'
  if (lineRule === 'exact' && lineRawTwips) return `${(lineRawTwips / 20).toFixed(1)}pt`
  if (lineRule === 'atLeast' && lineRawTwips != null) {
    // atLeast: face value, but never below the grid-snapped single height
    // (Word probe 2026-08-22); line="0" atLeast is natural height off the grid
    if (lineRawTwips === 0) return `calc(${FACTOR} * 1em)`
    return `max(${(lineRawTwips / 20).toFixed(1)}pt, var(--doc-line-grid, calc(${FACTOR} * 1em)))`
  }
  const m = lineSpacing ?? (lineRule === 'auto' && lineRawTwips ? lineRawTwips / 240 : undefined)
  // typed-grid docs resolve --doc-line-max (mult x pitch, floored at the
  // snapped single; --doc-line-mult on the same element carries m); elsewhere
  // the unitless factor scales per run like before
  if (m) return m === 1 ? cssGridLineBase() : `var(--doc-line-max, calc(${FACTOR} * ${m}))`
  return null
}

/**
 * --doc-lead-top override of a fixed line rule (Word probe 2026-09-23: an atLeast
 * line bottom-aligns its glyphs, exact keeps the CSS centre); styles.css derives
 * the auto-multiple shift from --doc-line-mult, `initial` releases an inherited override.
 */
export function cssLeadTop(
  lineRule: 'auto' | 'atLeast' | 'exact' | undefined,
  lineRawTwips: number | undefined,
  lineSpacing: number | undefined,
): string | null {
  // exact: baseline at 0.8 x the rule (styles.css --doc-lead-exact from --doc-lh-cap
  // and the face box); the fallback keeps the CSS centre where the box is unknown
  if (lineRule === 'exact' && lineRawTwips) return 'var(--doc-lead-exact, 0px)'
  if (lineRule === 'atLeast' && lineRawTwips != null) {
    if (lineRawTwips === 0) return 'var(--doc-lead-gap, 0px)'
    return `calc(max(0px, (${(lineRawTwips / 20).toFixed(1)}pt - var(--doc-line-grid, calc(var(--doc-line-factor,1.2) * 1em))) / 2) + var(--doc-lead-gap, 0px))`
  }
  return cssAutoLineMult(lineRule, lineRawTwips, lineSpacing) ? 'initial' : null
}

interface FaceBox {
  ascent: number
  descent: number
}
/** generated by tools/build-stand-in-boxes.py: fonts.css faces and the macOS system snapshot */
const standInBoxes = standInBoxesJson as {
  bundled: Record<string, FaceBox>
  installed: Record<string, FaceBox>
}

/** hhea box of the first chain member Chromium can draw: a bundled face or a snapshot system face */
function chromiumBox(chain: string): FaceBox | null {
  for (const member of chain.split(',')) {
    const key = member.replace(/['"]/g, '').normalize('NFKC').toLowerCase().trim()
    if (!key || key.startsWith('var(')) return null
    const hit = standInBoxes.bundled[key] ?? standInBoxes.installed[key]
    if (hit) return hit
  }
  return null
}

/**
 * Content box of a family chain's primary face for the glyph shift (styles.css
 * --doc-lead-gap / --doc-lead-exact): Word rests the descent on the single line's
 * bottom and puts an exact line's baseline at 0.8 h, Chromium centres the leading.
 * Generated boxes, not a canvas probe: at parse time a probe still sees the
 * fallback face of a lazily loaded web font. Unknown faces emit nothing (CSS centre).
 */
export function fontBoxCss(chain: string): string[] {
  // the chain head is the strut face; a fallback member (Liberation Sans after
  // Arial Unicode MS, the CJK tail) must not lend its box
  const head = /^\s*(?:'([^']*)'|"([^"]*)"|([^,]*))/.exec(chain)
  const face = (head?.[1] ?? head?.[2] ?? head?.[3] ?? '').trim().toLowerCase()
  if (!face || face.startsWith('var(')) return []
  const box = hasModelledLineBox(face) ? chromiumBox(chain) : null
  // a head outside the modelled set must not keep an inherited box (CSS centre instead)
  if (!box) return ['--doc-font-box:initial', '--doc-font-skew:initial']
  const { ascent, descent } = box
  return [
    `--doc-font-box:${(ascent + descent).toFixed(4)}`,
    `--doc-font-skew:${(ascent - descent).toFixed(4)}`,
  ]
}

/**
 * line-height of a span drawn in `asciiChain` inside a paragraph whose strut is
 * `csChain` (the ASCII digit groups of a complex-script run): the line must pace
 * the ascii factor exactly. Chromium places both boxes on the baseline, so a span
 * at the target height still unions with the strut's taller ascent (Arial 0.905
 * over Carlito's 0.75 box) and overshoots; the span box is shrunk until the
 * union closes on the target. Linear in the line multiple so an inherited
 * multiple stays a var().
 */
export function cssCsDigitLiftExpr(
  asciiChain: string,
  csChain: string,
  asciiFactor: number,
  csFactor: number,
  multiple: number | 'inherited',
): string {
  const M = multiple === 'inherited' ? 'var(--doc-line-mult,1)' : String(multiple)
  const a = chromiumBox(asciiChain)
  const r = chromiumBox(csChain)
  const plain = `calc(${asciiFactor} * ${M} * 1em)`
  if (!a || !r) return plain
  const m = multiple === 'inherited' ? 1 : multiple
  const target = asciiFactor * m
  const ca = a.ascent + a.descent
  const hlR = (csFactor * m - (r.ascent + r.descent)) / 2
  const topR = r.ascent + hlR
  const botR = r.descent + hlR
  const union = (L: number) => {
    const hl = (L - ca) / 2
    return Math.max(topR, a.ascent + hl) + Math.max(botR, a.descent + hl)
  }
  if (union(target) <= target + 1e-4) return plain
  // strut dominates the top (or bottom): L = m x (2Fa - Fr) + constant
  const k1 = 2 * asciiFactor - csFactor
  const cr = r.ascent + r.descent
  for (const k2 of [
    ca + cr - 2 * r.ascent - 2 * a.descent,
    ca + cr - 2 * r.descent - 2 * a.ascent,
  ]) {
    const L = k1 * m + k2
    if (L > 0 && Math.abs(union(L) - target) < 1e-3)
      return `calc(${round4(k1)} * ${M} * 1em + ${round4(k2)}em)`
  }
  return plain
}

const round4 = (n: number) => Math.round(n * 1e4) / 1e4

/** auto/multiple factor of a spacing declaration (grid span snapping scales by it) */
export function cssAutoLineMult(
  lineRule: 'auto' | 'atLeast' | 'exact' | undefined,
  lineRawTwips: number | undefined,
  lineSpacing: number | undefined,
): number | undefined {
  if (lineRule === 'exact' || lineRule === 'atLeast') return undefined
  return lineSpacing ?? (lineRule === 'auto' && lineRawTwips ? lineRawTwips / 240 : undefined)
}

/**
 * The auto line multiple a rendered paragraph lays out with. Per-paragraph
 * declarations live in the inline style; style-level and document-default
 * multiples cascade through the computed style. Fixed-height lines (direct
 * exact/atLeast = .doc-lh-fixed, style-level = --doc-line-fixed) count as 1
 * unless a direct auto override re-declares the inline multiple.
 */
export function autoLineMultOf(el: HTMLElement): number {
  if (el.classList.contains('doc-lh-fixed')) return 1
  const inlineMult = el.style.getPropertyValue('--doc-line-mult')
  if (inlineMult) return parseFloat(inlineMult) || 1
  const cs = getComputedStyle(el)
  const fixed =
    el.style.getPropertyValue('--doc-line-fixed') || cs.getPropertyValue('--doc-line-fixed')
  if (fixed.trim() === '1') return 1
  return parseFloat(cs.getPropertyValue('--doc-line-mult')) || 1
}

/**
 * Space-before/space-after in px. Word keeps paragraph spacing at face value
 * even under a typed grid (probe 2026-08-13), so no grid quantization.
 */
export function snapSpacingToGrid(spacingTwips: number, _docGrid: DocGrid | undefined): number {
  return spacingTwips * TWIPS_TO_PX
}

// ─── CJK detection ───────────────────────────────────────────────────────────

export const isCjk = isEastAsianChar
const isHangul = isHangulChar

// ─── CJK-Latin autospace pads ────────────────────────────────────────────────
// Word's autoSpaceDE/DN gap measures ~1/4em while Chromium's text-autospace is
// fixed at 1/8em; the renderer wraps the character after each boundary in a
// .doc-autospace-pad span whose ::before collapsible space carries the whole
// gap (the space suppresses the native gap and vanishes with the seam at a line
// break, as Word's gap does); seams touching hangul keep a start margin. Pads
// only go between a directly adjacent CJK letter and a Latin letter/digit —
// never next to spaces or punctuation — matching where Chromium applies its
// native gap. Chromium cannot switch letters and digits separately (no
// ideograph-alpha / ideograph-numeric), so when only one of w:autoSpaceDE/DN
// is off the native gap goes off and the pad carries the whole 1/4em
// (.doc-autospace-pad-full).

const isCjkAutospaceSide = isCjkLetterChar

function isLatinAlnum(cp: number): boolean {
  if ((cp >= 0x30 && cp <= 0x39) || (cp >= 0x41 && cp <= 0x5a) || (cp >= 0x61 && cp <= 0x7a)) {
    return true
  }
  // Latin-1 Supplement / Extended letters, minus multiply/divide signs
  return cp >= 0xc0 && cp <= 0x24f && cp !== 0xd7 && cp !== 0xf7
}

/** which CJK-Latin gaps a paragraph gets (docx-engine AutoSpace): true = both kinds,
 *  'latin' = letters only (w:autoSpaceDN off), 'digits' = digits only (w:autoSpaceDE off) */
export type AutospaceMode = true | 'latin' | 'digits'

function isDigit(cp: number): boolean {
  return cp >= 0x30 && cp <= 0x39
}

export function needsAutospacePad(
  prevCp: number,
  nextCp: number,
  mode: AutospaceMode = true,
): boolean {
  const latin =
    isCjkAutospaceSide(prevCp) && isLatinAlnum(nextCp)
      ? nextCp
      : isLatinAlnum(prevCp) && isCjkAutospaceSide(nextCp)
        ? prevCp
        : -1
  if (latin < 0) return false
  return mode === true || (mode === 'digits') === isDigit(latin)
}

function lastCodePoint(text: string): number {
  const tail = text.charCodeAt(text.length - 1)
  if (tail >= 0xdc00 && tail <= 0xdfff && text.length > 1) {
    return text.codePointAt(text.length - 2)!
  }
  return tail
}

/** pad between two adjacent stretches of text (last char of prev vs first of next) */
export function autospacePadBetween(
  prev: string,
  next: string,
  mode: AutospaceMode = true,
): boolean {
  if (!prev || !next) return false
  return needsAutospacePad(lastCodePoint(prev), next.codePointAt(0)!, mode)
}

/** the renderer's pad mode for a paragraph's autoSpace attribute; null = pads off */
export function autospaceMode(autoSpace: unknown): AutospaceMode | null {
  if (autoSpace === false) return null
  return autoSpace === 'latin' || autoSpace === 'digits' ? autoSpace : true
}

/** pad span class: the full gap when the native text-autospace is switched off */
export function autospacePadClass(mode: AutospaceMode): string {
  return mode === true ? 'doc-autospace-pad' : 'doc-autospace-pad doc-autospace-pad-full'
}

/** UTF-16 length of the code point starting at offset */
export function codePointLengthAt(text: string, offset: number): number {
  return text.codePointAt(offset)! > 0xffff ? 2 : 1
}

/** UTF-16 offsets inside text where a pad belongs (between offset-1 and offset) */
export function autospaceBoundaries(text: string, mode: AutospaceMode = true): number[] {
  const out: number[] = []
  let prev = -1
  for (let i = 0; i < text.length;) {
    const cp = text.codePointAt(i)!
    if (prev >= 0 && needsAutospacePad(prev, cp, mode)) out.push(i)
    prev = cp
    i += cp > 0xffff ? 2 : 1
  }
  return out
}

/** seams of autospaceBoundaries with their direction (latinFirst = Latin letter/digit before the CJK letter) */
export function autospaceSeams(
  text: string,
  mode: AutospaceMode = true,
): Array<{ at: number; latinFirst: boolean }> {
  const out: Array<{ at: number; latinFirst: boolean }> = []
  let prev = -1
  for (let i = 0; i < text.length;) {
    const cp = text.codePointAt(i)!
    if (prev >= 0 && needsAutospacePad(prev, cp, mode))
      out.push({ at: i, latinFirst: isLatinAlnum(prev) })
    prev = cp
    i += cp > 0xffff ? 2 : 1
  }
  return out
}

/** the seam between two adjacent stretches has the Latin side first */
export function autospaceLatinFirst(prev: string): boolean {
  return prev !== '' && isLatinAlnum(lastCodePoint(prev))
}

/** Times New Roman's hhea ascent+descent without the line gap, (1825+443)/2048 */
export const TIMES_SCALED_LINE_FACTOR = 1.107

/**
 * Line factor of a run with character scaling (w:w != 100): Word lays Times New
 * Roman (and Liberation Serif, its substitute) at the hhea box 1.107 instead of
 * 1.15 whatever the scale, while Arial / Calibri / Cambria keep their factor
 * (probe 2026-11-08: TNR 10.5pt line 420 w:w 90-200 -> 20.4pt, w:w 100 -> 21.12)
 */
export function scaledLatinLineFactor(font: string): number {
  const face = latinSlotFace(font) ?? font
  return isTimesClassFace(face) ? TIMES_SCALED_LINE_FACTOR : latinLineFactor(font)
}

/** Word's whole seam gap in units of the 1/8 em pad base (.doc-autospace-pad-le); 2 = the flat 1/4 em */
export function autospacePadScale(gapEm: number): number {
  return Math.round((gapEm / 0.125) * 1000) / 1000
}

/** a seam touching hangul keeps the gap as a start margin: Korean word-wrap never breaks there */
export function autospaceHangulSeam(prev: string, next: string): boolean {
  return (!!prev && isHangul(lastCodePoint(prev))) || (!!next && isHangul(next.codePointAt(0)!))
}

// Korean word-wrap breaks only at spaces (prod corpus 20261020 sample 023: 94
// of 95 Word line ends sit on a space, the other inside a hyphenated URL), but
// Chromium's keep-all still breaks between punctuation and a hangul syllable
// (U+00B7 chains, `)이`, `[3]아`). A U+2060 word-joiner widget at the seam
// closes the opportunity whatever spans the two characters end up in.
const KR_NOBREAK_MIDDLE = new Set([0xb7, 0x30fb, 0x2027, 0x2022])
const KR_NOBREAK_OPEN = new Set(
  [...'([{\u300c\u300e\u3008\u300a\u3010\u3014\uff08\u201c\u2018'].map((c) => c.codePointAt(0)!),
)
const KR_NOBREAK_CLOSE = new Set(
  [...')]}\u300d\u300f\u3009\u300b\u3011\u3015\uff09\u201d\u2019'].map((c) => c.codePointAt(0)!),
)

/** UTF-16 offsets in text where a hangul/punctuation seam must not break (between
 *  offset-1 and offset); prev = the text directly before (a seam on a run boundary is offset 0) */
export function krNoBreakOffsets(text: string, prev = ''): number[] {
  const out: number[] = []
  let prevCp = prev ? lastCodePoint(prev) : -1
  for (let i = 0; i < text.length;) {
    const cp = text.codePointAt(i)!
    if (
      prevCp >= 0 &&
      ((isHangul(prevCp) && (KR_NOBREAK_MIDDLE.has(cp) || KR_NOBREAK_OPEN.has(cp))) ||
        (isHangul(cp) && (KR_NOBREAK_MIDDLE.has(prevCp) || KR_NOBREAK_CLOSE.has(prevCp))))
    ) {
      out.push(i)
    }
    prevCp = cp
    i += cp > 0xffff ? 2 : 1
  }
  return out
}

// ─── Balanced spaces ─────────────────────────────────────────────────────────
// Under settings.xml balanceSingleByteDoubleByteWidth Word lays out a space at
// 0.5em of the run size whatever the font's own space advance when it touches
// an East Asian character (Han, kana, hangul, CJK punctuation, fullwidth
// forms) or another space; a lone space between Latin letters, digits or
// punctuation keeps the glyph advance (Word probes 2026-09-24 Korean corpus,
// 2026-11-04 SimSun contract: "AB  CD" 6.75/6.75pt at 13.5pt, "A B" 3.05,
// run boundaries, w:hint and w:lang irrelevant). The flag alone is enough when
// the document's East Asian face (Normal/docDefaults, theme-resolved) is a CJK
// font; with a Latin face there (Calibri in the eastAsia slot) it also needs
// useFELayout (probe 2026-09-30: either alone left every space at the glyph
// advance). Run-level East Asian faces change nothing. The renderer wraps each
// such space in a .doc-balanced-space span whose only glyph comes from the
// half-width space face; a space left in its text node keeps its run face, so
// nothing is wrapped when the rule is off.

export function balancedSpaceWideningFor(
  settings: { balanceDbcsSpacing?: boolean; useFELayout?: boolean },
  docEaFace?: string | null,
): boolean {
  if (settings.balanceDbcsSpacing !== true) return false
  return settings.useFELayout === true || (!!docEaFace && isCjkFontName(docEaFace))
}
/** DocumentTextContext.eastAsiaLang of the open document */
export function docEastAsiaLang(): string | null {
  return openDocumentText().eastAsiaLang
}
/** DocumentTextContext.balancedSpaces of the open document */
export function balancedSpaceWideningOn(): boolean {
  return openDocumentText().balancedSpaces
}

/** cheap paragraph-level gate: any East Asian character or a double space */
export function hasBalancedSpaceContext(text: string): boolean {
  return text.includes('  ') || textHasEastAsian(text)
}

/** UTF-16 offsets of the U+0020 in text with an East Asian or space neighbour;
 *  prev/next are the texts adjacent to its ends (empty = no direct neighbour) */
export function balancedSpaceOffsets(text: string, prev = '', next = ''): number[] {
  const out: number[] = []
  let prevCp = prev ? lastCodePoint(prev) : -1
  const widens = (cp: number) => cp === 0x20 || isCjk(cp)
  for (let i = 0; i < text.length;) {
    const cp = text.codePointAt(i)!
    if (cp === 0x20) {
      const nextCp =
        i + 1 < text.length ? text.codePointAt(i + 1)! : next ? next.codePointAt(0)! : -1
      if (widens(prevCp) || widens(nextCp)) out.push(i)
    }
    prevCp = cp
    i += cp > 0xffff ? 2 : 1
  }
  return out
}

// ─── Justification symbols ──────────────────────────────────────────────────
// Chromium's Character::IsCJKIdeographOrSymbol (character_property_data.h)
// also covers these symbols outside the CJK blocks and counts them as
// justification opportunities like ideographs, so a justified line stretches
// after them; Word stretches only spaces around them (prod 012: the gap after
// U+2116 doubled). Emoji and the CJK blocks themselves are left to Chromium.
const CHROMIUM_CJK_SYMBOLS = new Set([
  0x2c7, 0x2ca, 0x2cb, 0x2d9, 0x2020, 0x2021, 0x2030, 0x203b, 0x203c, 0x2042, 0x2047, 0x2048,
  0x2049, 0x2051, 0x20dd, 0x20de, 0x2100, 0x2103, 0x2105, 0x2109, 0x210a, 0x2113, 0x2116, 0x2121,
  0x212b, 0x213b, 0x2150, 0x2151, 0x2152, 0x217f, 0x2189, 0x2307, 0x2312, 0x23ce, 0x2423, 0x25a0,
  0x25a1, 0x25a2, 0x25aa, 0x25ab, 0x25b1, 0x25b2, 0x25b3, 0x25b6, 0x25b7, 0x25bc, 0x25bd, 0x25c0,
  0x25c1, 0x25c6, 0x25c7, 0x25c9, 0x25cb, 0x25cc, 0x25ef, 0x2605, 0x2606, 0x260e, 0x2616, 0x2617,
  0x26a0, 0x2713, 0x271a, 0x273f, 0x2740, 0x2756, 0x2763, 0x2b1a,
])
const CHROMIUM_CJK_SYMBOL_RANGES: ReadonlyArray<readonly [number, number]> = [
  [0x2156, 0x215a],
  [0x2160, 0x216b],
  [0x2170, 0x217b],
  [0x23be, 0x23cc],
  [0x2460, 0x2492],
  [0x249c, 0x24ff],
  [0x25ce, 0x25d3],
  [0x25e2, 0x25e6],
  [0x2600, 0x2603],
  [0x2660, 0x266f],
  [0x2672, 0x267d],
  [0x2776, 0x277f],
]

export function isChromiumCjkSymbol(cp: number): boolean {
  if (cp < 0x2c7 || cp >= 0x2e80) return false
  if (CHROMIUM_CJK_SYMBOLS.has(cp)) return true
  return CHROMIUM_CJK_SYMBOL_RANGES.some(([lo, hi]) => cp >= lo && cp <= hi)
}

/** UTF-16 ranges of the symbols Chromium would justify-expand after (all BMP) */
export function justifySymbolRanges(text: string): Array<{ from: number; to: number }> {
  const out: Array<{ from: number; to: number }> = []
  for (let i = 0; i < text.length; i++) {
    if (isChromiumCjkSymbol(text.charCodeAt(i))) out.push({ from: i, to: i + 1 })
  }
  return out
}

// ─── Line-break simulation ───────────────────────────────────────────────────

/**
 * CJK character line-height factor (by font family).
 * Known CJK families take their LO-probe factor (lineHeightFactor); CJK text
 * inside Latin-font runs falls through LO's font fallback (~1.3 on this
 * platform), so the legacy 1.3 stays as the default.
 */
/** Font name plausibly covering CJK glyphs (drives whether a declared face may
 *  set the CJK line factor; Latin names cascade to the document's EA default). */
/** East Asian face by name (text-layout/face-info); both names stay for the migration waves */
export const isCjkNamedFont = isEastAsianFaceName
export const isCjkFontName = isEastAsianFaceName

function cjkLineHFactor(fontFamily: string): number {
  const f = fontFamily.toLowerCase()
  if (f.includes('pmingliu') || f.includes('mingliu')) return 1.0
  const declared = cjkDeclaredLineFactor(fontFamily)
  if (declared !== null) return declared
  if (isEaLayoutFace(fontFamily)) return lineHeightFactor(fontFamily)
  return 1.3
}

/**
 * Paragraph line-break simulation — computes line count (heights only, no glyph coordinates).
 *
 * Logic aligned with text-layout.ts layoutParagraph, but outputs only the list of line boxes (heights).
 * Each item in the runs array is a stretch of text with the same style (same as Block.runs).
 *
 * cjkFactor semantics:
 *   NaN (default) = decided dynamically per font family via cjkLineHFactor(style.fontFamily)
 *   0..N = fixed factor (table cells pass 1.0 = no extra boost)
 */
const CJK_LINE_HEIGHT_FACTOR = NaN

export function simulateLines(
  runs: Array<{
    text: string
    fontFamily?: string
    sizeHalfPoints?: number
    bold?: boolean
    italic?: boolean
  }>,
  availWidthPx: number,
  metrics: FontMetricsProvider,
  defaultFontSize: number,
  defaultFontFamily: string,
  cjkFactor = CJK_LINE_HEIGHT_FACTOR,
): Array<{ naturalLineH: number; text: string }> {
  if (availWidthPx <= 0)
    return [{ naturalLineH: defaultFontSize * 1.2, text: runs.map((r) => r.text).join('') }]

  // Word breaks Korean at spaces (keep-all), including Han inside Korean
  // paragraphs; the renderer applies the matching word-break:keep-all CSS
  const keepAll = runs.some((r) => textHasHangul(r.text))

  // resulting line list (records each line's natural height and text)
  const lines: Array<{ naturalLineH: number; text: string }> = []
  let curLineW = 0
  let curLineH = 0 // max natural line height of the current line
  let curText = ''
  // Word never breaks a line at a space: spaces past the right edge hang, and a
  // run of several spaces is squeezed when that lets the next word end at the
  // margin (prod-sas 047 footer: text + 172 spaces + text on one line)
  let squeezableW = 0
  let lastWasSpace = false
  // whitespace never sizes a line beside ink; a line of nothing but spaces / tabs
  // (tab-only paragraph) still takes the tallest whitespace run
  let wsH = 0

  const pushLine = (h: number) => {
    lines.push({ naturalLineH: curLineH === 0 ? Math.max(h, wsH) : h, text: curText })
    wsH = 0
    curLineW = 0
    curLineH = 0
    curText = ''
    squeezableW = 0
    lastWasSpace = false
  }

  const getStyle = (run: (typeof runs)[0]): RunStyle => ({
    fontFamily: run.fontFamily ?? defaultFontFamily,
    fontSizePx: (run.sizeHalfPoints ? run.sizeHalfPoints / 2 : defaultFontSize) * (96 / 72),
    bold: !!run.bold,
    italic: !!run.italic,
  })

  for (const run of runs) {
    if (!run.text) continue
    const style = getStyle(run)
    const m = metrics.metrics(style)
    // table-cell mode (cjkFactor=1.0): line-height cap = fontSizePx × 1.0,
    // matching the CJK chars' cjkH so Latin chars' font factor (1.2) doesn't raise the line.
    const lineH = isNaN(cjkFactor)
      ? m.lineHeight
      : Math.min(m.lineHeight, style.fontSizePx * Math.max(cjkFactor, 1.0))

    let buf = ''
    let bufCjkH = 0 // CJK line-height floor of buffered chars (keep-all mode)
    const flushWord = () => {
      if (!buf) return
      const wordH = Math.max(lineH, bufCjkH)
      bufCjkH = 0
      const w = metrics.measure(buf, style)
      if (curLineW + w > availWidthPx && curLineW > 0) {
        if (curLineW - squeezableW + w <= availWidthPx) {
          curLineW -= squeezableW
          squeezableW = 0
        } else {
          pushLine(Math.max(curLineH, lineH))
          curLineH = lineH
        }
      }
      lastWasSpace = false
      if (curLineW === 0 && w > availWidthPx) {
        // hard-break an overlong word
        let fragment = ''
        for (const ch of buf) {
          const cw = metrics.measure(fragment + ch, style)
          if (fragment && cw > availWidthPx) {
            curText = fragment
            pushLine(wordH)
            curLineH = wordH
            fragment = ch
          } else {
            fragment += ch
          }
        }
        if (fragment) {
          curLineW += metrics.measure(fragment, style)
          curLineH = Math.max(curLineH, wordH)
          curText = fragment
        }
      } else {
        curLineW += w
        curLineH = Math.max(curLineH, wordH)
        curText += buf
      }
      buf = ''
    }

    for (const ch of run.text) {
      const cp = ch.codePointAt(0) ?? 0
      if (ch === '\n') {
        flushWord()
        pushLine(Math.max(curLineH, lineH))
        curLineH = lineH
        continue
      }
      if (ch === ' ') {
        flushWord()
        const spW = metrics.measure(ch, style)
        wsH = Math.max(wsH, lineH)
        curText += ch
        if (curLineW + spW <= availWidthPx || curLineW === 0) {
          curLineW += spW
          if (lastWasSpace) squeezableW += spW
        }
        lastWasSpace = true
        continue
      }
      if (ch === '\t') {
        flushWord()
        const spW = metrics.measure(ch, style)
        if (curLineW + spW > availWidthPx && curLineW > 0) {
          pushLine(Math.max(curLineH, lineH))
          curLineH = lineH
          // swallow the tab at line start
        } else {
          curLineW += spW
          wsH = Math.max(wsH, lineH)
          curText += ch
        }
        continue
      }
      if (isCjk(cp)) {
        // CJK char line height: NaN = picked dynamically per font family (body path), a number = fixed factor (table cells)
        let runCjkFactor = isNaN(cjkFactor) ? cjkLineHFactor(style.fontFamily) : cjkFactor
        // ・/〜 under SimSun substitution lift the line to 1.7143 (probe 2026-08-13)
        if (isNaN(cjkFactor) && SIMSUN_GAP_CHAR_RE.test(ch)) {
          runCjkFactor = Math.max(runCjkFactor, simsunGapLineFactor(style.fontFamily) ?? 0)
        }
        const cjkH = style.fontSizePx * runCjkFactor
        if (keepAll) {
          buf += ch
          bufCjkH = Math.max(bufCjkH, cjkH)
          continue
        }
        flushWord()
        const cw = metrics.measure(ch, style)
        if (curLineW + cw > availWidthPx && curLineW > 0) {
          pushLine(Math.max(curLineH, lineH))
          curLineH = lineH
        }
        curLineW += cw
        curText += ch
        curLineH = Math.max(curLineH, cjkH)
        lastWasSpace = false
        continue
      }
      buf += ch
    }
    flushWord()
  }

  // wrap up (the last line)
  if (curLineH > 0 || wsH > 0 || lines.length === 0) {
    // last-line minimum:
    //   body path (cjkFactor=NaN): use the font-level line-height factor (lineHeightFactor)
    //   table-cell path (cjkFactor=1.0): match CJK chars, using cjkFactor as the factor
    const lastLineMin = isNaN(cjkFactor)
      ? defaultFontSize * (96 / 72) * lineHeightFactor(defaultFontFamily)
      : defaultFontSize * (96 / 72) * Math.max(cjkFactor, 1.0)
    pushLine(Math.max(curLineH, lastLineMin))
  }

  return lines
}

// ─── Main entry point ────────────────────────────────────────────────────────

export interface LineMetricsInput {
  runs: Array<{
    text: string
    fontFamily?: string
    sizeHalfPoints?: number
    bold?: boolean
    italic?: boolean
  }>
  availWidthPx: number
  /** paragraph line-height rule */
  lineRule?: 'auto' | 'atLeast' | 'exact'
  /** raw w:spacing w:line twips */
  lineRawTwips?: number
  /** space before, twips */
  spaceBefore?: number
  /** space after, twips */
  spaceAfter?: number
  /** section docGrid (affects line-height rounding) */
  docGrid?: DocGrid
  /** document default font size (pt, from docDefaults.sizeHalfPoints/2 or fallback 12) */
  defaultFontSizePt?: number
  /** document default font */
  defaultFontFamily?: string
  /** metrics implementation (default HeuristicMetrics) */
  metrics?: FontMetricsProvider
  /** whether the paragraph is empty (a textless paragraph occupies one line height) */
  isEmpty?: boolean
  /**
   * Table-cell mode: disables the CJK line-height boost (1.5× → 1.2×).
   * Cell line height is controlled by the line-height rule + docGrid, without extra CJK expansion.
   */
  tableCellMode?: boolean
  /** CJK line-height factor override (e.g. tuned against the baseline layout engine); defaults to tableCellMode/font-family behavior */
  cjkFactor?: number
}

export interface LineMetricsResult {
  /** line count */
  lineCount: number
  /** per-line heights (px, with line-height rules and docGrid rounding applied) */
  lineHeights: number[]
  /** per-line text (aligned with lineHeights; used to locate the page-leading line for in-block page splits) */
  lineTexts: string[]
  /** space before (px) */
  spaceBeforePx: number
  /** space after (px) */
  spaceAfterPx: number
  /** total block height = sum(lineHeights) + spaceBeforePx + spaceAfterPx */
  totalHeight: number
}

/**
 * Line box (for line-level page splitting).
 * offsetInBlock: offset of the line's top relative to the paragraph block's top, including spaceBefore (px).
 * height: the line's height (px).
 */
export interface LineBox {
  /** top offset of the line within the block (px, relative to block top) */
  offsetInBlock: number
  /** line height (px) */
  height: number
}

/**
 * Extended computeLineMetrics: additionally returns each line's LineBox (with offsets).
 * Used for line-level split decisions — the pagination algorithm can break a paragraph at any line boundary.
 */
export interface LineMetricsResultEx extends LineMetricsResult {
  /** per-line boxes (with in-block offsets), consumed by line-level page splitting */
  lineBoxes: LineBox[]
}

export const DEFAULT_FONT_FAMILY = 'Calibri'
export const DEFAULT_FONT_SIZE_PT = 12

/** global default HeuristicMetrics instance (avoids rebuilding per call) */
const defaultMetrics = new HeuristicMetrics()

/**
 * Compute a paragraph's precise line-box metrics (block height from Word's perspective).
 *
 * Used for pagination computation (replacing getBoundingClientRect); does not affect editor rendering.
 */
export function computeLineMetrics(input: LineMetricsInput): LineMetricsResultEx {
  const {
    runs,
    availWidthPx,
    lineRule,
    lineRawTwips,
    spaceBefore = 0,
    spaceAfter = 0,
    docGrid,
    defaultFontSizePt = DEFAULT_FONT_SIZE_PT,
    defaultFontFamily = DEFAULT_FONT_FAMILY,
    metrics = defaultMetrics,
    isEmpty = false,
    tableCellMode = false,
  } = input

  const defaultFontSizePx = defaultFontSizePt * (96 / 72)
  // table cells disable the extra CJK line-height boost (cjkFactor=1.0, relying on HeuristicMetrics' font-level line height)
  // body paragraphs use the font-level cjkLineHFactor (PMingLiU→1.0, others→1.3), triggered via NaN
  const cjkFactor = input.cjkFactor ?? (tableCellMode ? 1.0 : CJK_LINE_HEIGHT_FACTOR)

  // an empty paragraph still occupies one line sized by its paragraph mark's
  // font/size (Word rule); whitespace-only runs carry that style, defaults
  // apply only when no run declares one
  const markRuns = isEmpty ? runs.filter((r) => r.sizeHalfPoints || r.fontFamily) : []
  const emptyNaturalH =
    markRuns.length > 0
      ? Math.max(
          ...markRuns.map(
            (r) =>
              (r.sizeHalfPoints ? r.sizeHalfPoints / 2 : defaultFontSizePt) *
              (96 / 72) *
              lineHeightFactor(inheritedMarkFace(r.fontFamily ?? defaultFontFamily)),
          ),
        )
      : defaultFontSizePx * lineHeightFactor(inheritedMarkFace(defaultFontFamily))

  // simulate line breaking
  const lines =
    isEmpty || runs.length === 0
      ? [{ naturalLineH: emptyNaturalH, text: '' }]
      : simulateLines(runs, availWidthPx, metrics, defaultFontSizePt, defaultFontFamily, cjkFactor)

  // apply the line-height rule per line
  const lineHeights = lines.map((ln) =>
    computeLineHeight(ln.naturalLineH, lineRule, lineRawTwips, docGrid),
  )

  // space before/after
  // table-cell mode: no docGrid alignment for paragraph spacing (only line heights snap to the grid)
  const spacingDocGrid = tableCellMode ? undefined : docGrid
  const spaceBeforePx = spaceBefore > 0 ? snapSpacingToGrid(spaceBefore, spacingDocGrid) : 0
  const spaceAfterPx = spaceAfter > 0 ? snapSpacingToGrid(spaceAfter, spacingDocGrid) : 0

  const totalHeight = lineHeights.reduce((s, h) => s + h, 0) + spaceBeforePx + spaceAfterPx

  // build each line's LineBox (with in-block offsets)
  // spaceBefore counts before the first line
  const lineBoxes: LineBox[] = []
  let offset = spaceBeforePx
  for (const h of lineHeights) {
    lineBoxes.push({ offsetInBlock: offset, height: h })
    offset += h
  }

  return {
    lineCount: lineHeights.length,
    lineHeights,
    lineTexts: lines.map((ln) => ln.text),
    lineBoxes,
    spaceBeforePx,
    spaceAfterPx,
    totalHeight,
  }
}

/** default separator row (px): Word's own separator pPr (line 240, after 0) under Calibri 11 */
export const FOOTNOTE_SEPARATOR_H = 16

/**
 * Header/footer part height estimate (px): per-paragraph line-box model. Capacity
 * input for body push-down (body top = max(marginTop, headerDist + header height)).
 */
interface HfLineSpacingLike {
  lineRule?: 'auto' | 'atLeast' | 'exact'
  lineRawTwips?: number
  lineSpacing?: number
  spaceBefore?: number
  spaceAfter?: number
}

interface HfRunLike {
  text: string
  font?: string
  sizeHalfPoints?: number
  bold?: boolean
  italic?: boolean
  /** inline image on a layout-table cell run; its height joins the line box */
  image?: { heightPx?: number }
}

/** header push-down geometry of a section (twips → px) for estimateHfHeight's geom param */
export function hfHeaderGeom(
  set: { marginTop: number; headerDist?: number },
  legacyBands?: boolean,
): {
  marginTopPx: number
  headerDistPx: number
  legacyBands?: boolean
} {
  const toPx = (twips: number) => (twips / 1440) * 96
  return {
    marginTopPx: toPx(set.marginTop),
    headerDistPx: toPx(set.headerDist ?? 720),
    ...(legacyBands ? { legacyBands: true } : {}),
  }
}

/** narrowest side gap that still takes a word beside a header drawing (the
 *  body wrap engine's sliver approximation of L462 "the next word fits") */
const HF_SIDE_WRAP_SLIVER_PX = 36

/** wider of the two gaps beside a floating header drawing in its text column (px); null = unknown */
function hfSideRoomPx(
  im: { widthPx?: number; posXPx?: number; posH?: string; posHRel?: string },
  columnPx: number,
  marginLeftPx: number | undefined,
): number | null {
  const w = im.widthPx
  if (w == null) return null
  let x: number
  if (!im.posHRel || im.posHRel === 'margin') {
    x =
      im.posXPx ??
      (im.posH === 'center' ? (columnPx - w) / 2 : im.posH === 'right' ? columnPx - w : 0)
  } else if (im.posHRel === 'page' && marginLeftPx != null) {
    x = (im.posXPx ?? 0) - marginLeftPx
  } else return null
  return Math.max(x, columnPx - x - w)
}

type HfPartLike = NonNullable<Parameters<typeof estimateHfHeight>[0]>

/** flow heights (px) of a strip's paragraphs in order: text lines, table rows, border pads */
function hfFlowParaHeightsPx(part: HfPartLike, contentWidthPx: number): number[] {
  type HfPara = NonNullable<HfPartLike['paras']>[number]
  const paras: HfPara[] = part.paras?.length
    ? part.paras
    : part.text.trim()
      ? part.text.split('\n').map((t) => ({ runs: [{ text: t }] }))
      : []
  const lineH = (runs: HfRunLike[], rich?: HfLineSpacingLike, availWidthPx = contentWidthPx) =>
    computeLineMetrics({
      runs: runs.map((r) => ({
        text: r.text,
        ...(r.font ? { fontFamily: r.font } : {}),
        ...(r.sizeHalfPoints ? { sizeHalfPoints: r.sizeHalfPoints } : {}),
        ...(r.bold ? { bold: true } : {}),
        ...(r.italic ? { italic: true } : {}),
      })),
      availWidthPx,
      ...(rich?.lineRule ? { lineRule: rich.lineRule } : {}),
      ...(rich?.lineRawTwips
        ? { lineRawTwips: rich.lineRawTwips }
        : rich?.lineSpacing
          ? { lineRule: 'auto' as const, lineRawTwips: rich.lineSpacing * 240 }
          : {}),
      ...(rich?.spaceBefore ? { spaceBefore: rich.spaceBefore } : {}),
      ...(rich?.spaceAfter ? { spaceAfter: rich.spaceAfter } : {}),
      defaultFontSizePt: 10.5,
      isEmpty: runs.every((r) => !r.text.trim()),
    }).totalHeight
  const heights: number[] = []
  const flow = paras.filter((p) => !p.boxAnchored)
  const padKey = (p: HfPara | undefined) =>
    p && !p.cells?.length && p.borderPad ? JSON.stringify(p.borderPad) : ''
  for (const [i, p] of flow.entries()) {
    if (p.cells?.length) {
      // table row: the tallest cell's paragraph stack sets the row height;
      // a cell-run image (logo) grows its line box like the display layer
      const declared = (p.row?.heightTwips ?? 0) / 15
      if (declared > 0 && p.row?.heightRule === 'exact') {
        heights.push(declared)
        continue
      }
      heights.push(
        Math.max(
          declared,
          ...p.cells.map((c) =>
            (c.paras.length > 0 ? c.paras : [[]]).reduce((s, runs, k) => {
              const imgH = Math.max(0, ...runs.map((r) => r.image?.heightPx ?? 0))
              return s + Math.max(lineH(runs, c.paraProps?.[k]), imgH)
            }, 0),
          ),
        ),
      )
      continue
    }
    const indentPx = ((p.indentLeft ?? 0) + (p.indentRight ?? 0)) / 15
    const runs =
      p.runs.length === 0 && p.emptyRunSizeHalfPoints
        ? [{ text: '', sizeHalfPoints: p.emptyRunSizeHalfPoints, font: p.emptyRunFontFamily }]
        : p.runs
    const imgH = Math.max(0, ...runs.map((r) => r.image?.heightPx ?? 0))
    let h = Math.max(lineH(runs, p, Math.max(1, contentWidthPx - indentPx)), imgH)
    // same-border neighbours form one Word border group: padded at its edges only
    const key = padKey(p)
    if (key && key !== padKey(flow[i - 1])) h += ((p.borderPad?.t ?? 0) * 96) / 72
    if (key && key !== padKey(flow[i + 1])) h += ((p.borderPad?.b ?? 0) * 96) / 72
    heights.push(h)
  }
  return heights
}

export function estimateHfHeight(
  part:
    | {
        text: string
        paras?: Array<{
          runs: HfRunLike[]
          /** layout-table row: per-cell paragraph stacks (row height = tallest cell) */
          cells?: Array<{ paras: HfRunLike[][]; paraProps?: Array<HfLineSpacingLike | undefined> }>
          /** declared w:trHeight (twips): floors the row, or fixes it under hRule exact */
          row?: { heightTwips?: number; heightRule?: 'atLeast' | 'exact' }
          /** floating-textbox content: drawn at the anchor, no strip flow height */
          boxAnchored?: boolean
          lineRule?: 'auto' | 'atLeast' | 'exact'
          lineRawTwips?: number
          lineSpacing?: number
          spaceBefore?: number
          spaceAfter?: number
          /** w:ind (twips): narrows the wrap width of a stacked paragraph */
          indentLeft?: number
          indentRight?: number
          /** blank paragraph: the mark's (or style's) size and face size its line */
          emptyRunSizeHalfPoints?: number
          emptyRunFontFamily?: string
          /** w:pBdr none-side w:space (pt): undrawn padding above/below the line */
          borderPad?: Partial<Record<'t' | 'b' | 'l' | 'r', number>>
        }>
      }
    | null
    | undefined,
  contentWidthPx: number,
  /** images in the part (logos): non-floating ones count as one line of height (same as the display layer) */
  images?: Array<{
    heightPx?: number
    widthPx?: number
    floating?: boolean
    behind?: boolean
    wrap?: 'none' | 'square' | 'tight' | 'through' | 'topBottom'
    posXPx?: number
    posYPx?: number
    posH?: 'left' | 'center' | 'right'
    posHRel?: 'page' | 'margin' | 'leftMargin' | 'rightMargin'
    posVRel?: 'page' | 'margin' | 'paragraph' | 'topMargin' | 'bottomMargin'
  }> | null,
  /** header geometry (px): pass for headers so wrapped anchored images push the body below their bottom edge (Word); watermarks (wrapNone/behindDoc) still reserve nothing */
  geom?: HfFloorGeom,
): number {
  const inlineImages = (images ?? []).filter((im) => !im.floating && im.heightPx)
  const imagesHeight =
    inlineImages.length > 0 ? Math.max(...inlineImages.map((im) => im.heightPx!)) : 0
  if (!part) return Math.max(imagesHeight, hfAnchorFloorPx([], images, geom, contentWidthPx))
  const heights = hfFlowParaHeightsPx(part, contentWidthPx)
  const height = heights.reduce((a, b) => a + b, 0)
  return Math.max(height + imagesHeight, hfAnchorFloorPx(heights, images, geom, contentWidthPx))
}

/** the anchored-drawing floor of a header whose paragraph flow the DOM probe measured */
export function estimateHfAnchorFloorPx(
  part: HfPartLike | null | undefined,
  contentWidthPx: number,
  images: Parameters<typeof estimateHfHeight>[2],
  geom: Parameters<typeof estimateHfHeight>[3],
): number {
  return hfAnchorFloorPx(
    part ? hfFlowParaHeightsPx(part, contentWidthPx) : [],
    images,
    geom,
    contentWidthPx,
  )
}

type HfFloorImage = NonNullable<Parameters<typeof estimateHfHeight>[2]>[number] & {
  anchorPara?: number
}

interface HfFloorGeom {
  marginTopPx: number
  headerDistPx: number
  marginLeftPx?: number
  /** compatibilityMode < 15: the band clamps to the page top and pushes no strip line */
  legacyBands?: boolean
}

/**
 * Strip height a header's anchored drawings force (px above the header
 * distance). A wrapTopAndBottom band pushes the body below its bottom edge;
 * under compatibilityMode 15 a band that cuts into its anchor paragraph's
 * line moves that line and every later paragraph below it (probe 2026-11-03:
 * a box wholly above the paragraph does not push, one starting below the line
 * leaves the line where it is), and a negative offset is not clamped. Legacy
 * layout clamps the box to the page top and keeps the lines at the strip top.
 * Side wraps shape the lines beside them and push nothing unless no word fits
 * on either side (L260 CASE07). `paraHeights` = the strip's flow paragraphs.
 */
function hfAnchorFloorPx(
  paraHeights: readonly number[],
  images: readonly HfFloorImage[] | null | undefined,
  geom: HfFloorGeom | undefined,
  contentWidthPx = 0,
): number {
  if (!geom) return 0
  let floor = 0
  for (const im of images ?? []) {
    if (!im.floating || im.behind || !im.heightPx) continue
    if (!im.wrap || im.wrap === 'none') continue
    if (im.wrap !== 'topBottom') {
      const room = hfSideRoomPx(im, contentWidthPx, geom.marginLeftPx)
      if (room !== null && room >= HF_SIDE_WRAP_SLIVER_PX) continue
    }
    if (im.posYPx == null || !im.posVRel || im.posVRel === 'bottomMargin') continue
    // page coordinates; the anchor paragraph sits at the strip top (L851)
    const top =
      im.posVRel === 'page' || im.posVRel === 'topMargin'
        ? im.posYPx
        : im.posVRel === 'margin'
          ? geom.marginTopPx + im.posYPx
          : geom.headerDistPx + im.posYPx
    const bandTop = (geom.legacyBands ? Math.max(0, top) : top) - geom.headerDistPx
    const bandBottom = bandTop + im.heightPx
    let pushed = 0
    if (im.wrap === 'topBottom' && !geom.legacyBands && paraHeights.length) {
      const at = Math.min(im.anchorPara ?? 0, paraHeights.length - 1)
      const paraTop = paraHeights.slice(0, at).reduce((a, b) => a + b, 0)
      const cutsLine = bandTop < paraTop + paraHeights[at] && bandBottom > paraTop
      let from = bandBottom <= paraTop ? paraHeights.length : cutsLine ? at : at + 1
      // later paragraphs keep flowing in the gap above the band; the first
      // one that would cross it, and everything after, lands below the band
      let y = paraTop + (cutsLine ? 0 : paraHeights[at])
      while (from < paraHeights.length && y + paraHeights[from] <= bandTop) y += paraHeights[from++]
      pushed = paraHeights.slice(from).reduce((a, b) => a + b, 0)
    }
    floor = Math.max(floor, bandBottom + pushed)
  }
  return floor
}

/**
 * Footnote entry height estimate (10pt small text, content width narrowed 40px to approximate Word's footnote layout).
 * The referencing page reserves bottom space accordingly (same model as the parity runner).
 * The 40px narrowing is wider than the renderer's sup-number prefix, so wrap
 * mismatches err toward over-reserving (never clipping).
 */
export interface NoteStyleOpts {
  sizeHalfPoints?: number
  fontFamily?: string
  lineRule?: 'auto' | 'atLeast' | 'exact'
  lineRawTwips?: number
  spaceBeforeTwips?: number
  spaceAfterTwips?: number
}

interface NoteStyleSource {
  styles?: Map<
    string,
    {
      name?: string
      display?: {
        sizeHalfPoints?: number
        font?: string
        lineRule?: 'auto' | 'atLeast' | 'exact'
        lineRawTwips?: number
        spaceBeforeTwips?: number
        spaceAfterTwips?: number
      }
    }
  >
  docDefaults?: {
    sizeHalfPoints?: number
    lineRule?: 'auto' | 'atLeast' | 'exact'
    lineRawTwips?: number
    spaceBeforeTwips?: number
    spaceAfterTwips?: number
  }
}

/**
 * Effective paragraph metrics for a footnote/endnote entry: the note's own
 * w:pStyle chain when it references one, else Normal/docDefaults. Word only
 * applies a footnote style to note paragraphs that reference it — prod docs
 * exist where the FootnoteText style (10pt) is defined but 216/217 note
 * paragraphs carry no pStyle and render with Normal metrics (12pt, line 276).
 */
export function resolveNoteStyle(
  parsed: NoteStyleSource,
  styleId?: string,
  direct?: {
    beforeTwips?: number
    afterTwips?: number
    lineRule?: 'auto' | 'atLeast' | 'exact'
    lineRawTwips?: number
  },
): NoteStyleOpts {
  const style = (styleId ? parsed.styles?.get(styleId) : undefined) ?? parsed.styles?.get('Normal')
  const d = style?.display
  const dd = parsed.docDefaults
  // the line rule travels as a pair: a direct w:spacing w:line wins whole
  const lineRule = direct?.lineRawTwips ? direct.lineRule : (d?.lineRule ?? dd?.lineRule)
  const lineRawTwips = direct?.lineRawTwips ?? d?.lineRawTwips ?? dd?.lineRawTwips
  return {
    sizeHalfPoints: d?.sizeHalfPoints ?? dd?.sizeHalfPoints ?? 20,
    ...(d?.font ? { fontFamily: d.font } : {}),
    ...(lineRule ? { lineRule } : {}),
    ...(lineRawTwips ? { lineRawTwips } : {}),
    spaceBeforeTwips: direct?.beforeTwips ?? d?.spaceBeforeTwips ?? dd?.spaceBeforeTwips ?? 0,
    spaceAfterTwips: direct?.afterTwips ?? d?.spaceAfterTwips ?? dd?.spaceAfterTwips ?? 0,
  }
}

/** display run of a note entry (the NoteRun subset the height model reads) */
export interface NoteRunMetrics {
  text: string
  sizeHalfPoints?: number
  bold?: boolean
  italic?: boolean
  /** direct Latin font of the run */
  fontAscii?: string
}

/**
 * Note-level font/size owed to the entry's own runs: Word sizes the line boxes
 * from the runs, so a 9pt Times note under a 10pt Calibri style chain must
 * reserve Times 9pt lines. First paragraph only; the leading text run's font,
 * the paragraph's largest run size.
 */
export function noteRunStyle(
  richParas: NoteRunMetrics[][] | undefined,
): Pick<NoteStyleOpts, 'fontFamily' | 'sizeHalfPoints'> {
  const runs = richParas?.[0]?.filter((r) => r.text.trim()) ?? []
  if (runs.length === 0) return {}
  const font = runs.find((r) => r.fontAscii)?.fontAscii
  const size = Math.max(0, ...runs.map((r) => r.sizeHalfPoints ?? 0))
  return {
    ...(font ? { fontFamily: font } : {}),
    ...(size > 0 ? { sizeHalfPoints: size } : {}),
  }
}

export function estimateFootnoteHeight(
  footnoteText: string,
  contentWidthPx: number,
  docGrid: DocGrid | undefined,
  metrics?: FontMetricsProvider,
  style?: NoteStyleOpts,
  richParas?: NoteRunMetrics[][],
): number {
  const paras: NoteRunMetrics[][] =
    richParas && richParas.length > 0
      ? richParas
      : footnoteText.split('\n').map((t) => [{ text: t }])
  let total = 0
  for (const runs of paras) {
    const nonEmpty = runs
      .filter((r) => r.text)
      .map(({ fontAscii, ...r }) => (fontAscii ? { ...r, fontFamily: fontAscii } : r))
    total += computeLineMetrics({
      runs: nonEmpty,
      availWidthPx: contentWidthPx - 40,
      docGrid,
      defaultFontSizePt: style?.sizeHalfPoints ? style.sizeHalfPoints / 2 : 10,
      ...(style?.fontFamily ? { defaultFontFamily: style.fontFamily } : {}),
      ...(style?.lineRule ? { lineRule: style.lineRule } : {}),
      ...(style?.lineRawTwips ? { lineRawTwips: style.lineRawTwips } : {}),
      ...(style?.spaceBeforeTwips ? { spaceBefore: style.spaceBeforeTwips } : {}),
      ...(style?.spaceAfterTwips ? { spaceAfter: style.spaceAfterTwips } : {}),
      ...(metrics ? { metrics } : {}),
      isEmpty: nonEmpty.every((r) => !r.text.trim()),
    }).totalHeight
  }
  return total
}

/**
 * Line height (px) of one note entry under the estimate model (note style
 * resolved like estimateFootnoteHeight, docGrid-snapped). The note renderers
 * set this as the entry's CSS line-height so estimate, reservation, and
 * rendering share one line-box truth.
 */
export function noteLineHeightPx(docGrid: DocGrid | undefined, style?: NoteStyleOpts): number {
  const sizePt = style?.sizeHalfPoints ? style.sizeHalfPoints / 2 : 10
  const naturalH = sizePt * (96 / 72) * lineHeightFactor(style?.fontFamily ?? DEFAULT_FONT_FAMILY)
  return computeLineHeight(naturalH, style?.lineRule, style?.lineRawTwips, docGrid)
}

/** noteLineHeightPx of the default (styleless, 10pt) note model */
export function footnoteLineHeightPx(docGrid: DocGrid | undefined): number {
  return noteLineHeightPx(docGrid)
}

export interface NoteSeparatorRow {
  heightPx: number
  rulePx: number
}

const DEFAULT_SEPARATOR_ROW: NoteSeparatorRow = { heightPx: FOOTNOTE_SEPARATOR_H, rulePx: 11 }

/** Word lays the separator out as a paragraph of the notes part: its pPr chain's line box
 *  plus space before/after (a bare separator under Normal Cambria 11 line 288 after 140 rows 22.5pt) */
export function noteSeparatorRow(
  docGrid: DocGrid | undefined,
  style?: NoteStyleOpts,
): NoteSeparatorRow {
  if (!style) return DEFAULT_SEPARATOR_ROW
  const px = (twips?: number) => ((twips ?? 0) / 20) * (96 / 72)
  const lineH = noteLineHeightPx(docGrid, style)
  const before = px(style.spaceBeforeTwips)
  return {
    heightPx: lineH + before + px(style.spaceAfterTwips),
    rulePx: Math.round(before + lineH * 0.56),
  }
}

let separatorRows: Record<NoteKindKey, NoteSeparatorRow> = {
  footnote: DEFAULT_SEPARATOR_ROW,
  endnote: DEFAULT_SEPARATOR_ROW,
}
type NoteKindKey = 'footnote' | 'endnote'

/** the open document's separator rows (set on load; the engine, canvas gaps and preview read them) */
export function setNoteSeparatorRows(rows: Record<NoteKindKey, NoteSeparatorRow>): void {
  separatorRows = rows
}

export function noteSeparatorPx(kind: NoteKindKey): number {
  return separatorRows[kind].heightPx
}

export function noteSeparatorRulePx(kind: NoteKindKey): number {
  return separatorRows[kind].rulePx
}

export { unloadedBundledPrimary }

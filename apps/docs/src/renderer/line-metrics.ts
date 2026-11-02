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

import type { DocGrid } from '@genoffice/docx-engine'
import {
  arabicCodePage,
  latinSlotSubstitute,
  type ResolvedFace,
} from '@genoffice/font-metrics/resolve'

import {
  recordResolveDiff,
  resolveDeclaredFace,
  setResolverEmbeddedFaces,
  setResolverFontTable,
} from './font-resolve'
import { NOTO_HAN_REGIONAL_RE, SEGOE_UI_TEXT_RE, TAMIL_FONT_RE } from './font-name-patterns'
import {
  LINE_FACTOR_PINS,
  LINE_FACTOR_RULES,
  LINE_FACTORS,
  type LineFactorStage,
} from './line-factors.generated'
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
const MONO_FONT_RE =
  /consolas|courier|menlo|monaco|cascadia|sf mono|jetbrains mono|source code|fira code|(liberation|roboto|ubuntu|dejavu sans|andale) mono|lucida console/i

/** Constant Latin advance for monospace families (chain-head name decides); null when proportional */
export function monospaceAdvanceEm(fontFamily: string): number | null {
  const head = fontFamily.split(',')[0]
  if (/^'?(ms (gothic|mincho)|ＭＳ (ゴシック|\u660e\u671d))'?$/i.test(head.trim())) return 0.5
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

/** Japanese font names (kana-lettered vendor names, Windows/macOS/Noto faces) */
const JA_FONT_RE =
  /[぀-ヿ]|mincho|meiryo|hiragino|osaka|yugoth|yu (gothic|mincho)|ms (ui )?p?(gothic|mincho)|明朝|biz ud|kozuka|小塚|^hg[ps]?(\u6559\u79d1\u66f8|\u884c\u66f8|\u6b63\u6977\u66f8|\u5275\u82f1)/i

/** Korean font names (Windows/Noto/Source Han/Nanum faces + bundled subsets) */
const KO_FONT_RE =
  /malgun|맑은|batang|바탕|myeongjo|myungjo|명조|gungsuh|궁서|gulim|굴림|dotum|돋움|nanum|나눔|genoffice (sans|serif) kr|(noto|source han) (sans|serif)[^,]*\bk(r|orean)?\b/i

/**
 * Per-font single-spacing line-height factor. Re-baselined against Word for
 * Mac (batch probe 2026-08-13: docx per family/size converted by installed
 * Word, baseline pitch measured from the PDF). Latin faces matched the old
 * LO-era values exactly; the CJK entries moved a lot because Word renders
 * with its own bundled faces (SimSun/Batang/PMingLiU at 1.3029, Malgun
 * 1.7305, Meiryo 1.9429, YaHei 1.7143, Yu 1.44) instead of the macOS
 * substitutes LO measured. Unprobed names keep their LO values until probed.
 */
/** Tamil family names other than Latha (TAU-* are the Tamil Virtual Academy faces) */
/** line boxes of the open document's embedded faces, keyed like docFontTable */
let embeddedLineBoxes = new Map<string, { factor: number; descent: number }>()

/** first entry per family wins: callers list the regular cut first */
export function setEmbeddedLineMetrics(
  faces: ReadonlyArray<{ family: string; ascent: number; descent: number; lineGap: number }>,
): void {
  embeddedLineBoxes = new Map()
  for (const f of faces) {
    const key = f.family.normalize('NFKC').toLowerCase().trim()
    const factor = f.ascent + f.descent + f.lineGap
    if (embeddedLineBoxes.has(key) || !(factor > 0.5 && factor < 4)) continue
    embeddedLineBoxes.set(key, { factor: Math.round(factor * 1e4) / 1e4, descent: f.descent })
  }
  setResolverEmbeddedFaces(faces.map((f) => ({ ...f, eastAsian: isCjkFontName(f.family) })))
}

/** Names Word for Mac renders with SimSun itself (lower-cased input) */
const SIMSUN_NAME_RE = /simsun|\u5b8b\u4f53|xiaobiaosong|\u5c0f\u6807\u5b8b/

/** the hand table alone (shadow audits compare the resolver against it) */
export function tableLineHeightFactor(fontFamily: string): number {
  return knownLineFactor(fontFamily) ?? unknownLatinFactor(fontFamily)
}

export function lineHeightFactor(fontFamily: string): number {
  const known = knownLineFactor(fontFamily)
  const legacy = known ?? unknownLatinFactor(fontFamily)
  const face = resolveDeclaredFace(fontFamily)
  if (face && resolverAnswers(fontFamily, face, known)) return face.lineFactor
  if (face?.lineFactor !== legacy) recordResolveDiff(fontFamily, legacy, face)
  return legacy
}

/**
 * An East Asian face (probe 2026-10-22: every EA face declared by a CJK name
 * paces at hhea (ascent + descent) x 1.3, Malgun 1.7292, Yu 1.4330, Meiryo
 * 1.9505, YaHei 1.7151, DengXian 1.3545, JhengHei 1.7292, Songti SC 1.8197,
 * STSong 1.3003, SimSun/MS Gothic 1.2970) and a behavioural substitution
 * (missing Arabic names -> Times New Roman 1.15) take the resolver
 */
function resolverAnswers(name: string, face: ResolvedFace, known: number | undefined): boolean {
  const key = name.normalize('NFKC').toLowerCase().trim()
  if (known !== undefined && LINE_FACTOR_PINS.some((pin) => pin.test(key))) return false
  if (face.metrics.eastAsian) {
    // a face Word ships paces at its own hhea x 1.3 whatever its name; an
    // installed or embedded one only under a CJK name (the EA code-page flag
    // on Latin macOS faces is a data quirk); charset-81 entries are unprobed
    if (!face.substituted) {
      const shipped = face.source !== 'installed' && face.source !== 'system'
      return (shipped || isCjkFontName(name)) && !krCharsetRoute(name)
    }
    // a substitute reached through altName or a behavioural rule is drawn
    // wholesale, Latin and CJK alike (probe 2026-10-22: altName MS Mincho,
    // Gulim, Yu Gothic, SimHei, YaHei and SimSun under Latin and CJK names)
    return true
  }
  if (face.source === 'embedded') return false
  if (face.substituted) return face.exception && known === undefined
  return latinCutover(name, face, known)
}

/**
 * Latin faces Word lays out for real (its own DFonts and cloud copies,
 * installed faces) follow the resolver when the table never probed the name or
 * agrees with it exactly; a probed value that differs stays until a corpus run
 * adopts the file's digits. Substitutions, East Asian faces and embedded boxes
 * stay on the table paths.
 */
function latinCutover(name: string, face: ResolvedFace, known: number | undefined): boolean {
  if (face.substituted || face.metrics.eastAsian) return false
  if (isCjkFontName(name)) return false
  return known === undefined || known === face.lineFactor
}

/** Word's Calibri substitute for an unprobed Latin name with a fontTable entry, else Cambria;
 *  an altName pointing elsewhere is drawn wholesale and paces as that face */
function unknownLatinFactor(fontFamily: string): number {
  const alt = missingLatinAltName(fontFamily)
  if (alt) return lineHeightFactor(alt)
  return missingLatinSubstitute(fontFamily) === 'calibri'
    ? LINE_FACTORS.calibri!.factor
    : LINE_FACTORS.cambria!.factor
}

/** first rule of `stage` matching the lower-cased name */
function ruleFactor(f: string, stage: LineFactorStage): number | undefined {
  for (const rule of LINE_FACTOR_RULES) {
    if (rule.stage !== stage || !rule.match.test(f) || rule.exclude?.test(f)) continue
    return rule.factor
  }
  return undefined
}

/**
 * Factor of a probed/known face (packages/font-metrics/measurements/line-factors.json,
 * one source per value); undefined for a Latin name no rule covers
 */
function knownLineFactor(fontFamily: string): number | undefined {
  const f = fontFamily.toLowerCase()
  // an embedded face renders real: Word lays it with its own hhea total
  // (probe 2026-09-23, typo/win groups and USE_TYPO_METRICS ignored)
  const embedded = embeddedLineBoxes.get(f.normalize('NFKC').trim())
  // Word meters an embedded East Asian face by its EA rule, not its own box
  // (embedded Malgun Gothic: 19.05 pt at 11 pt = 1.7305, hhea total is 1.3301)
  if (embedded && !isCjkFontName(fontFamily)) return embedded.factor
  const first = ruleFactor(f, 'first')
  if (first !== undefined) return first
  // MS Gothic / SimSun class fallbacks for the declares resolverAnswers leaves
  // to the table (altName-routed, Latin slot of a split substitution)
  if (
    /mincho|\u660e\u671d|\u30b4\u30b7\u30c3\u30af|ms (ui )?p?gothic|hiragino|osaka|kozuka|\u5c0f\u585a/.test(
      f.normalize('NFKC'),
    )
  )
    return simsunLineFactor()
  if (/^(noto|source han) (sans|serif)( cjk)? ?(sc|cn|tc|tw|hk)\b/.test(f))
    return simsunLineFactor()
  const arabicCloud = ruleFactor(f, 'arabic')
  if (arabicCloud !== undefined) return arabicCloud
  const arabicAlt = arabicAltNameFace(fontFamily) ?? arabicCodePageFace(fontFamily)
  if (arabicAlt) return lineHeightFactor(arabicAlt)
  const latin = ruleFactor(f, 'latin')
  if (latin !== undefined) return latin
  // a Hangul-charset fontTable entry substitutes Malgun Gothic (Batang when
  // w:family is roman) for the Latin runs too (probe 2026-10-20)
  const krRoute = krCharsetRoute(f)
  if (krRoute) return krRoute === 'malgun' ? malgunLineFactor() : batangLineFactor()
  // missing hangul-named vendor faces substitute into the Batang/SimSun class
  // for KR and Latin runs alike (probe 2026-08-23: \ud734\uba3c\uace0\ub515 -> Batang/SimSun 1.30)
  if (/[\uac00-\ud7a3]/.test(f)) return batangLineFactor()
  // missing ideograph-named faces substitute SimSun (probe 2026-08-22/23)
  if (/[\u4e00-\u9fff]/.test(f)) return simsunLineFactor()
  // unknown Western names with a fontTable entry substitute Calibri unless
  // their PANOSE says serif (probe 2026-10-22); without an entry Word for Mac
  // substitutes Cambria (probe 2026-08-23: fabricated names and HCI Poppy
  // both render as Cambria) -- unknownLatinFactor
  return ruleFactor(f, 'last')
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
 * Declared names the line-factor table lays out as a sibling face (Times and
 * Courier carry the New faces' 1.15 / 1.1333; Helvetica's 1.2 and Arial
 * Unicode MS's Arial pacing are corpus values, not hhea), so descent and
 * xAvgCharWidth follow the same sibling until the macOS copies are probed
 */
const LEGACY_FACE_ALIAS: Array<[RegExp, string]> = [
  [/^times\b(?! new)/, 'Times New Roman'],
  [/^courier\b(?! new)/, 'Courier New'],
  [/^helvetica\b(?! neue)/, 'Arial'],
  [/^arial unicode/, 'Arial'],
]

function legacyFaceAlias(font: string): string | undefined {
  const key = font.normalize('NFKC').toLowerCase().trim()
  return LEGACY_FACE_ALIAS.find(([re]) => re.test(key))?.[1]
}

/** Word's layout face for a declared Latin name, through the table's sibling aliases */
function latinFace(font: string): ResolvedFace | null {
  if (!font.trim()) return null
  return resolveDeclaredFace(legacyFaceAlias(font) ?? font)
}

/** Descent (em) Word lays a Latin text face with: its layout face's hhea
 *  descent; East Asian faces keep 1/4 em (their box is the x1.3 rule) */
function latinDescent(font: string): number {
  const f = font.toLowerCase()
  const embedded = embeddedLineBoxes.get(f.normalize('NFKC').trim())?.descent
  if (embedded !== undefined) return embedded
  const face = latinFace(font)
  return face && !face.metrics.eastAsian ? face.metrics.descent : 0.25
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
  const asc = SYMBOL_FACE_ASCENT[symbolFont.trim().toLowerCase()]
  if (asc === undefined || !(markerPt > 0) || !(textPt > 0)) return null
  const line = asc * markerPt + latinDescent(textFont ?? '') * textPt
  if (line <= lineHeightFactor(textFont ?? 'Calibri') * textPt + 0.01) return null
  return Math.round(line * 1000) / 1000
}

/**
 * Factor for run-declared Noto/Source Han regional variants missing on this
 * platform, for CJK text only. JP/SC/TC are Word-probed (2026-08-13): Word
 * for Mac substitutes SimSun at exactly 1.3029 single spacing (10/10.5pt,
 * grid and no-grid), explicit multiples scale linearly. KR reuses the
 * KO_FONT_RE value. Latin consumers (latinParaFactor, --doc-line-factor-latin)
 * must NOT use this: Cyrillic/Latin text under these names passes through
 * unicode-range aliases to proportional Latin glyphs where 1.2 matches Word
 * (sample 40).
 */
export { NOTO_HAN_REGIONAL_RE }

export function cjkDeclaredLineFactor(fontFamily: string): number | null {
  const krName = krNameLineFactor(fontFamily)
  if (krName !== null) return krName
  const m = NOTO_HAN_REGIONAL_RE.exec(fontFamily.toLowerCase())
  if (!m) return null
  if (m[1] === 'kr') return lineHeightFactor(fontFamily)
  return simsunLineFactor()
}

/** Word's SimSun / Microsoft YaHei line factors through the resolver (the
 *  missing JP/SC/TC substitute and the face it lifts gap characters to) */
export const simsunLineFactor = (): number => resolveDeclaredFace('SimSun')?.lineFactor ?? 1.3
export const yaheiLineFactor = (): number =>
  resolveDeclaredFace('Microsoft YaHei')?.lineFactor ?? 1.7157
export const batangLineFactor = (): number => resolveDeclaredFace('Batang')?.lineFactor ?? 1.3
export const malgunLineFactor = (): number =>
  resolveDeclaredFace('Malgun Gothic')?.lineFactor ?? 1.7291

/**
 * Word probe 2026-08-13: SimSun (the 1.3029 substitute for missing JP/SC/TC
 * faces) lacks ・ U+30FB and 〜 U+301C; Word falls back to Microsoft YaHei for
 * them and lifts the whole line to 1.7143 × size (13.68pt → 18.0pt @10.5pt).
 * KR (Batang class, also 1.3029) is unprobed and excluded.
 */
export const SIMSUN_GAP_CHAR_RE = /[・〜]/

export function simsunGapLineFactor(fontFamily: string): number | null {
  const f = fontFamily.toLowerCase()
  // Korean faces (named or hangul-lettered) substitute Batang-ward, not SimSun
  if (/[가-힣ᄀ-ᇿ㄰-㆏]/.test(f.normalize('NFKC')) || KO_FONT_RE.test(f.normalize('NFKC'))) {
    return null
  }
  if (cjkDeclaredLineFactor(fontFamily) === simsunLineFactor()) return yaheiLineFactor()
  if (SIMSUN_NAME_RE.test(f)) return yaheiLineFactor()
  return null
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
 * Font family → canvas CSS font-family fallback chain.
 * When a common Word font is missing locally, fall back to metric-compatible
 * open-source fonts (registered in fonts/fonts.css); identical widths keep line
 * breaks aligned with Word and the offline pagination model. CJK families fall
 * back to macOS equivalents (CJK width is always 1em, so this is mostly glyph appearance).
 */
/**
 * Whether a font family actually resolves on this machine. document.fonts.check
 * is useless in Chromium (true for any unknown system-ish family), so this
 * measures a mixed-script sample against both generic fallbacks: a family that
 * changes neither width doesn't exist.
 */
/**
 * Bundled web fonts (fonts.css). The canvas probe resolves them, but they are
 * last-resort subset faces, so substitution decisions treat them as missing.
 */
export const BUNDLED_FONTS = new Set([
  'Noto Sans CJK SC',
  'Noto Serif CJK SC',
  'GenOffice Sans KR',
  'GenOffice Serif KR',
  'GenOffice Gothic KR',
  'GenOffice Poppins',
  'GenOffice DM Sans',
  'GenOffice Tamil',
  'GenOffice Bengali',
  'GenOffice Fullwidth TC',
  'GenOffice Songti SC',
  'GenOffice Songti TC',
  'Carlito GO',
  'Aptos GO',
  'Aptos Display GO',
  'Aptos Narrow GO',
  'Century Gothic GO',
  'Cambria Math GO',
  'Caladea',
  'Liberation Serif',
  'Liberation Sans',
  'Liberation Mono',
])
const BUNDLED_LC = new Set([...BUNDLED_FONTS].map((f) => f.toLowerCase()))
export function isBundledFont(name: string): boolean {
  return BUNDLED_LC.has(name.toLowerCase())
}

const fontAvailableCache = new Map<string, boolean>()
export function isFontAvailable(font: string): boolean {
  if (typeof document === 'undefined') return false
  const cached = fontAvailableCache.get(font)
  if (cached !== undefined) return cached
  let available = false
  try {
    const canvas = document.createElement('canvas')
    const ctx = canvas.getContext('2d')
    if (ctx) {
      const sample = '한글あア中文abcWXYmm123'
      const quoted = `"${font.replace(/"/g, '')}"`
      const widthWith = (family: string) => {
        ctx.font = `32px ${family}`
        return ctx.measureText(sample).width
      }
      available =
        widthWith(`${quoted}, monospace`) !== widthWith('monospace') ||
        widthWith(`${quoted}, serif`) !== widthWith('serif')
    }
  } catch {
    /* headless/test environments treat every font as missing */
  }
  fontAvailableCache.set(font, available)
  return available
}

/**
 * Ascent/descent of a family chain's primary face, as percentages of the em
 * (canvas fontBoundingBox at 1000px for 0.1% granularity). Fuels the typed-grid
 * strut alias in doc-style-css; null where canvas metrics are unavailable (jsdom).
 */
const chainMetricsCache = new Map<string, { ascentPct: number; descentPct: number } | null>()

/** embedded faces registered/revoked: their availability and every chain's metrics may have flipped */
export function noteEmbeddedFontsChanged(families: readonly string[]): void {
  for (const f of families) fontAvailableCache.delete(f)
  chainMetricsCache.clear()
}
const GENERIC_FAMILIES = new Set([
  'serif',
  'sans-serif',
  'monospace',
  'cursive',
  'fantasy',
  'system-ui',
])

/**
 * Bundled web face that leads a chain but has not loaded yet: canvas measures
 * the next installed family instead (an unloaded font-display:block face is
 * skipped), so the strut would carry e.g. Apple SD Gothic Neo's 0.90/0.30
 * under GenOffice Sans KR spans (1.16/0.29) and every hangul line unions
 * ~1.5px past its line-height (+1pt per Korean table row vs Word).
 */
export function unloadedBundledPrimary(chain: string): string | null {
  const fonts = typeof document === 'undefined' ? undefined : document.fonts
  if (!fonts?.check) return null
  // the chain's primary is its first face covering U+0020; canvas skips any
  // web face still loading, so every loaded one is walked past
  for (const raw of chain.split(',')) {
    const family = raw.trim().replace(/^["']|["']$/g, '')
    if (!family || GENERIC_FAMILIES.has(family.toLowerCase())) return null
    if (isWebFace(family)) {
      try {
        if (!fonts.check(`16px "${family}"`)) return family
      } catch {
        return null
      }
    } else if (isFontAvailable(family)) return null
  }
  return null
}

/**
 * Every fonts.css face: bundled stand-ins plus the metric/range aliases. An
 * unloaded 'MS Mincho GO' leaves canvas measuring the Noto tail (1.15/0.29em)
 * while layout later resolves the loaded alias (0.69/0.25em) as primary.
 */
function isWebFace(family: string): boolean {
  return isBundledFont(family) || / GO$/.test(family) || /^GenOffice /.test(family)
}

/** bundled web face a chain resolves to before any installed family or generic */
function bundledPrimary(chain: string): string | null {
  for (const raw of chain.split(',')) {
    const family = raw.trim().replace(/^["']|["']$/g, '')
    if (!family || GENERIC_FAMILIES.has(family.toLowerCase())) return null
    if (isBundledFont(family)) return family
    if (isFontAvailable(family)) return null
  }
  return null
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

// ─── Document fontTable (word/fontTable.xml substitution hints) ─────────────

/** current document's fontTable, keyed by NFKC-lowercased name */
type FontTableHints = {
  altName?: string
  panose?: string
  family?: string
  charset?: string
  csb0?: string
}
let docFontTable = new Map<string, FontTableHints>()

export function setDocFontTable(
  entries: ReadonlyArray<{ name: string } & FontTableHints> | null | undefined,
): void {
  docFontTable = new Map()
  setResolverFontTable(entries)
  for (const e of entries ?? []) {
    docFontTable.set(e.name.normalize('NFKC').toLowerCase(), {
      ...(e.altName ? { altName: e.altName } : {}),
      ...(e.panose ? { panose: e.panose } : {}),
      ...(e.family ? { family: e.family } : {}),
      ...(e.charset ? { charset: e.charset } : {}),
      ...(e.csb0 ? { csb0: e.csb0 } : {}),
    })
  }
}

/**
 * Missing face whose fontTable charset is Hangul (81): Word for Mac substitutes
 * Malgun Gothic wholesale, hangul and Latin alike, unless w:family says roman
 * (then Batang). PANOSE and the name's script do not matter (probe 2026-10-20:
 * HCR Batang serif PANOSE + modern, a Latin-named entry without PANOSE, and
 * sans PANOSE + roman went Malgun / Malgun / Batang).
 */
export function krCharsetRoute(font: string): 'malgun' | 'batang' | null {
  const head = font.split(',')[0].replace(/['"]/g, '').normalize('NFKC').trim()
  const hints = docFontTable.get(head.toLowerCase())
  if (hints?.charset !== '81' || !(isBundledFont(head) || !isFontAvailable(head))) return null
  return hints.family?.toLowerCase() === 'roman' ? 'batang' : 'malgun'
}

export function krCharsetMalgunSubstituted(font: string): boolean {
  return krCharsetRoute(font) === 'malgun'
}

/**
 * Hangul-lettered family name whose fontTable altName is Batang: Word for Mac
 * lays it out in Batang even when the named face is installed (corpus
 * 2026-09-24, a Nanum Myeongjo declared by its Korean name with altName Batang
 * pitched 1.3029 with Batang advances), so the installed head must not lead.
 */
export function krBatangAltNamed(font: string): boolean {
  const head = font.split(',')[0].replace(/['"]/g, '').normalize('NFKC').trim()
  if (!/[\uac00-\ud7a3]/.test(head)) return false
  const alt = docFontTable.get(head.toLowerCase())?.altName?.normalize('NFKC').toLowerCase()
  return !!alt && /batang|\ubc14\ud0d5/.test(alt)
}

/** PANOSE serif-style byte (2nd): 0B-0F sans, 02-0A serif, 00/01/none unknown */
function panoseSerifHint(font: string): 'serif' | 'sans' | undefined {
  const panose = docFontTable.get(font.normalize('NFKC').toLowerCase())?.panose
  const hex = panose?.replace(/[^0-9a-f]/gi, '')
  if (!hex || hex.length < 4 || /^0+$/.test(hex)) return undefined
  const serifStyle = parseInt(hex.slice(2, 4), 16)
  if (serifStyle >= 0x0b && serifStyle <= 0x0f) return 'sans'
  if (serifStyle >= 0x02 && serifStyle <= 0x0a) return 'serif'
  return undefined
}

const CJK_NAMED_RE = /[\u2E80-\u9FFF\uAC00-\uD7A3\uF900-\uFAFF\uFF00-\uFFEF]/

/**
 * Missing Latin-named face that has a fontTable entry (no altName): Word for
 * Mac substitutes Calibri wholesale, Cyrillic and Latin alike, unless the
 * PANOSE serif-style byte says serif, which sends it to Cambria. A name
 * without any fontTable entry is Cambria too (probe 2026-10-22: charset
 * CC/00/none, w:family and the docDefaults/theme fonts do not matter).
 */
export function missingLatinSubstitute(font: string): 'calibri' | 'cambria' | null {
  const head = font.split(',')[0].replace(/['"]/g, '').normalize('NFKC').trim()
  if (CJK_NAMED_RE.test(head)) return null
  const hints = docFontTable.get(head.toLowerCase())
  if (!hints || hints.charset === '81') return null
  if (hints.altName && hints.altName.normalize('NFKC').toLowerCase() !== head.toLowerCase())
    return null
  if (!(isBundledFont(head) || !isFontAvailable(head))) return null
  return panoseSerifHint(head) === 'serif' ? 'cambria' : 'calibri'
}

/** fontTable altName of a missing Latin-named face (probe 2026-10-22: Word draws the alias wholesale) */
function missingLatinAltName(font: string): string | null {
  const head = font.split(',')[0].replace(/['"]/g, '').normalize('NFKC').trim()
  if (CJK_NAMED_RE.test(head)) return null
  const hints = docFontTable.get(head.toLowerCase())
  const alt = hints?.altName?.trim()
  if (!alt || hints?.charset === '81' || alt.normalize('NFKC').toLowerCase() === head.toLowerCase())
    return null
  if (!(isBundledFont(head) || !isFontAvailable(head))) return null
  return alt
}

/**
 * Line factor for hangul-lettered vendor names outside the known Korean set:
 * a missing face lays its hangul in Batang whatever its PANOSE says (probe
 * 2026-10-22: sans, serif and absent PANOSE); only a charset-81 entry reaches
 * Malgun. Null for non-hangul or known Korean names.
 */
export function krNameLineFactor(font: string): number | null {
  const nfkc = font.normalize('NFKC')
  if (!/[가-힣ᄀ-ᇿ㄰-㆏]/.test(nfkc)) return null
  if (KO_FONT_RE.test(nfkc)) return null
  const route = krCharsetRoute(font)
  if (route) return route === 'malgun' ? malgunLineFactor() : batangLineFactor()
  return batangLineFactor()
}

/** Segoe UI text family (not the Emoji/Symbol/Historic/Print/Script cuts) */
/** Montserrat family names of Word's cloud-font cache (weight cuts are families of their own) */
const MONTSERRAT_FAMILY_RE =
  /^montserrat(?: (?:thin|extralight|light|medium|semibold|extrabold|black))?$/

/**
 * Metric alias (fonts.css) of an Office-private CJK face Chromium cannot see:
 * the stand-in's glyphs under the real face's Word line geometry. Leads the
 * face's chain, and doc-style-css gives paragraphs a strut of the same
 * geometry so CJK stretches and the strut share one line box.
 */
export function wordMetricAliasFace(font: string): string | null {
  const f = font.toLowerCase().normalize('NFKC')
  if (f.includes('dengxian') || f.includes('\u7b49\u7ebf')) return 'DengXian GO'
  if (/p?mingliu|\u7d30\u660e\u9ad4/.test(f)) return 'PMingLiU GO'
  return null
}

const CJK_SERIF = 'Noto Serif CJK SC'
const CJK_SANS = 'Noto Sans CJK SC'
const BOX = 'GenOffice Box Drawing'
const JA_SANS = ['Yu Gothic', 'GenOffice Hiragino Sans', 'Meiryo', 'Noto Sans JP']
const JA_SERIF = ['Yu Mincho', 'GenOffice Hiragino Mincho', 'GenOffice MS Mincho', 'Noto Serif JP']
const KO_SANS = ['Malgun Gothic', 'GenOffice Sans KR', 'Apple SD Gothic Neo', 'Noto Sans KR']
const KO_SERIF = ['GenOffice Batang', 'GenOffice Serif KR', 'GenOffice Myungjo', 'Noto Serif KR']

type Generic = 'serif' | 'sans-serif' | 'monospace'

/**
 * CSS chains keyed by the face Word lays out with (the resolver's target), for
 * targets whose chain does not depend on the declared spelling. Localized
 * spellings of a Word face are separate store entries, hence separate keys.
 */
const STAND_IN_CHAINS: Record<string, [string[], Generic]> = {
  calibri: [['Carlito GO', 'Calibri Arabic GO', CJK_SANS], 'sans-serif'],
  'calibri light': [
    ['Calibri Light GO', 'Carlito GO', 'Calibri Arabic GO', CJK_SANS],
    'sans-serif',
  ],
  cambria: [['Caladea', BOX, CJK_SERIF], 'serif'],
  'times new roman': [['Liberation Serif', BOX, CJK_SERIF], 'serif'],
  // Office-only serif (DFonts): per-class size-adjusted Caladea alias (manifest bookman-old-style-go)
  'bookman old style': [['Bookman Old Style GO', BOX, CJK_SERIF], 'serif'],
  // M365 cloud face Word paces at its real (condensed) advances; Chromium has no narrow stand-in
  'barlow condensed': [['Barlow Condensed GO', CJK_SANS], 'sans-serif'],
  consolas: [['Consolas GO', CJK_SANS], 'monospace'],
  'segoe ui': [['Segoe UI GO', 'Noto Sans CJK GO', 'GenOffice PUA Blank'], 'sans-serif'],
  'ms gothic': [['MS Gothic GO', ...JA_SANS], 'sans-serif'],
  'ms \u30b4\u30b7\u30c3\u30af': [['MS Gothic GO', ...JA_SANS], 'sans-serif'],
  'ms pgothic': [['MS PGothic GO', 'MS PGothic JA GO', ...JA_SANS], 'sans-serif'],
  'ms p\u30b4\u30b7\u30c3\u30af': [['MS PGothic GO', 'MS PGothic JA GO', ...JA_SANS], 'sans-serif'],
  'ms ui gothic': [['MS UI Gothic GO', 'MS UI Gothic JA GO', ...JA_SANS], 'sans-serif'],
  'ms mincho': [['MS Mincho GO', ...JA_SERIF], 'serif'],
  'ms \u660e\u671d': [['MS Mincho GO', ...JA_SERIF], 'serif'],
  'ms pmincho': [JA_SERIF, 'serif'],
  'ms p\u660e\u671d': [JA_SERIF, 'serif'],
  meiryo: [['Meiryo GO', ...JA_SANS], 'sans-serif'],
  '\u30e1\u30a4\u30ea\u30aa': [['Meiryo GO', ...JA_SANS], 'sans-serif'],
  'meiryo ui': [['Meiryo UI GO', ...JA_SANS], 'sans-serif'],
  'yu gothic': [['Yu Gothic GO', 'Yu Gothic Sans GO', ...JA_SANS], 'sans-serif'],
  '\u6e38\u30b4\u30b7\u30c3\u30af': [
    ['Yu Gothic GO', 'Yu Gothic Sans GO', ...JA_SANS],
    'sans-serif',
  ],
  'yu gothic light': [['Yu Gothic Light GO', 'Yu Gothic Sans GO', ...JA_SANS], 'sans-serif'],
  '\u6e38\u30b4\u30b7\u30c3\u30af light': [
    ['Yu Gothic Light GO', 'Yu Gothic Sans GO', ...JA_SANS],
    'sans-serif',
  ],
  'yu mincho': [['Yu Mincho GO', 'Yu Mincho Serif GO', ...JA_SERIF], 'serif'],
  '\u6e38\u660e\u671d': [['Yu Mincho GO', 'Yu Mincho Serif GO', ...JA_SERIF], 'serif'],
  'malgun gothic': [KO_SANS, 'sans-serif'],
  '\ub9d1\uc740 \uace0\ub515': [KO_SANS, 'sans-serif'],
  batang: [KO_SERIF, 'serif'],
  gulim: [KO_SANS, 'sans-serif'],
  '\uad74\ub9bc': [KO_SANS, 'sans-serif'],
  dotum: [KO_SANS, 'sans-serif'],
  '\ub3cb\uc6c0': [KO_SANS, 'sans-serif'],
  // the OS asset Chromium cannot see; the bundled real-metric subset leads
  nanumgothic: [['GenOffice Gothic KR', ...KO_SANS], 'sans-serif'],
  'nanum gothic': [['GenOffice Gothic KR', ...KO_SANS], 'sans-serif'],
  '\ub098\ub214\uace0\ub515': [['GenOffice Gothic KR', ...KO_SANS], 'sans-serif'],
  'nanumgothic extrabold': [['GenOffice Gothic KR', ...KO_SANS], 'sans-serif'],
  '\ub098\ub214\uace0\ub515 extrabold': [['GenOffice Gothic KR', ...KO_SANS], 'sans-serif'],
  nanummyeongjo: [KO_SERIF, 'serif'],
  'nanum myeongjo': [KO_SERIF, 'serif'],
  '\ub098\ub214\uba85\uc870': [KO_SERIF, 'serif'],
  nanummyeongjoextrabold: [KO_SERIF, 'serif'],
  '\ub098\ub214\uba85\uc870 extrabold': [KO_SERIF, 'serif'],
  simsun: [
    ['GenOffice SimSun Latin', 'GenOffice Songti SC', 'STSong', 'SimSun', CJK_SERIF],
    'serif',
  ],
  '\u5b8b\u4f53': [
    ['GenOffice SimSun Latin', 'GenOffice Songti SC', 'STSong', 'SimSun', CJK_SERIF],
    'serif',
  ],
  simhei: [['Heiti SC', 'STHeiti', 'SimHei', 'PingFang SC', CJK_SANS], 'sans-serif'],
  '\u9ed1\u4f53': [['Heiti SC', 'STHeiti', 'SimHei', 'PingFang SC', CJK_SANS], 'sans-serif'],
  'microsoft yahei': [['Microsoft YaHei', 'PingFang SC', CJK_SANS], 'sans-serif'],
  '\u5fae\u8f6f\u96c5\u9ed1': [['Microsoft YaHei', 'PingFang SC', CJK_SANS], 'sans-serif'],
  dengxian: [['DengXian GO', 'DengXian', 'PingFang SC', 'Microsoft YaHei', CJK_SANS], 'sans-serif'],
  '\u7b49\u7ebf': [
    ['DengXian GO', 'DengXian', 'PingFang SC', 'Microsoft YaHei', CJK_SANS],
    'sans-serif',
  ],
  kaiti: [['STKaiti', 'Kaiti SC', 'KaiTi', CJK_SERIF], 'serif'],
  '\u6977\u4f53': [['STKaiti', 'Kaiti SC', 'KaiTi', CJK_SERIF], 'serif'],
  fangsong: [['STFangsong', 'FangSong', CJK_SERIF], 'serif'],
  '\u4eff\u5b8b': [['STFangsong', 'FangSong', CJK_SERIF], 'serif'],
}

/** Chain by Word's target face; null for substitutions and unkeyed targets */
export function standInFontFamily(font: string): string | null {
  const face = resolveDeclaredFace(font)
  if (!face || face.substituted) return null
  const hit = STAND_IN_CHAINS[face.family.normalize('NFKC').toLowerCase()]
  if (!hit) return null
  const [families, generic] = hit
  const head = font.split(',')[0].replace(/['"]/g, '').trim()
  return `${[...new Set([head, ...families])].map((x) => `'${x}'`).join(',')},${generic}`
}

const cssChain = (...families: string[]) =>
  [...new Set(families)].map((x) => `'${x.replace(/'/g, '')}'`).join(',')

/** Arabic-script names: the naskh chains carry the bundled subsets Word's Times New Roman substitute needs */
const ARABIC_NAME_RE =
  /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]|naskh|kufi|arabic|urdu|geeza|amiri|scheherazade|lateef|harmattan|aldhabi|andalus|nastaliq|al bayan|baghdad|damascus|diwan|farisi|mishafi|nadeem|beirut|sakkal majalla|dubai|\b(?:b|xb|ir)[ -]?(?:mitra|nazanin|titr|lotus|zar|yekan|koodak|roya|badr|homa|traffic|compset)\b|irlotus/i

/**
 * Word follows a missing Arabic name's fontTable altName before its Times New
 * Roman rule (corpus 2026-10-27: altName Arial laid out in Arial, pitch 1.15);
 * the altName face then sets the chain and the line factor
 */
function arabicAltNameFace(font: string): string | null {
  const head = font.split(',')[0].replace(/['"]/g, '').normalize('NFKC').trim()
  if (!ARABIC_NAME_RE.test(head) || isBundledFont(head) || isFontAvailable(head)) return null
  const alt = docFontTable.get(head.toLowerCase())?.altName?.normalize('NFKC').trim()
  return alt && alt.toLowerCase() !== head.toLowerCase() && !ARABIC_NAME_RE.test(alt) ? alt : null
}

/** M365 cloud Arabic faces Word downloads and lays out with real metrics */
const ARABIC_CLOUD_RE = /dubai|simplified arabic|traditional arabic|sakkal majalla|aldhabi|andalus/i

/**
 * A missing Arabic name without altName whose fontTable sig covers the Arabic
 * code page (Latin charset) lays out in Arial, not Times New Roman (corpus
 * 2026-10-27: a 36 pt heading came back ArialMT); the resolver applies the
 * same rule to Latin-named entries
 */
function arabicCodePageFace(font: string): string | null {
  const head = font.split(',')[0].replace(/['"]/g, '').normalize('NFKC').trim()
  if (!ARABIC_NAME_RE.test(head) || isBundledFont(head) || isFontAvailable(head)) return null
  // cloud faces Word renders real keep their own aliases whatever the sig says
  if (ARABIC_CLOUD_RE.test(head)) return null
  const hints = docFontTable.get(head.toLowerCase())
  return hints && arabicCodePage(hints) && hints.charset !== 'B2' ? 'Arial' : null
}

const HAN_REGIONAL_NAME_RE =
  /^(?:noto|source han) (?:sans|serif)(?: cjk)? ?(?:sc|cn|tc|tw|hk|jp|kr|k)\b/i

/**
 * Latin-range alias standing in for the face Word draws the Latin of a
 * CJK-named missing face with (latinSlotSubstitute); the CJK stand-ins follow
 */
function latinSlotHead(sub: string): string[] {
  const k = sub.normalize('NFKC').toLowerCase()
  if (k === 'calibri') return ['Carlito GO']
  if (k === 'cambria') return ['KR Theme Latin GO']
  if (k === 'simsun' || k === '\u5b8b\u4f53') return ['GenOffice SimSun Latin']
  return latinChainHead(sub)
}

/**
 * Chain of a name Word substitutes: the declared head, then the target face's
 * own chain (Word draws the substitute wholesale, probe 2026-10-22). A
 * CJK-named face substituted by script keeps its Latin on the Latin slot's
 * substitute (hangul -> Batang + Latin Calibri/Cambria, Han -> SimSun + Latin
 * Calibri, Noto Sans CJK SC -> SimSun + its altName Calibri). Arabic names and
 * DFKai keep their richer name-keyed chains.
 */
/** Arabic and DFKai keep richer chains; math faces need their symbol face; Aptos cuts are cloud faces the store may lack */
const NAME_KEYED_SUBSTITUTE_RE = /\u6a19\u6977|dfkai|math|aptos/i

/**
 * Face Word draws the Latin of a substituted CJK-named face with when its
 * chain splits the Latin slot off (hangul / Han names landing on SimSun or
 * Batang); null when the declared face draws its own Latin
 */
function latinSlotTarget(font: string): string | null {
  const face = resolveDeclaredFace(font)
  if (!face?.substituted || !face.exception) return null
  const head = font.split(',')[0].replace(/['"]/g, '').trim()
  const nfkc = head.normalize('NFKC')
  if (ARABIC_NAME_RE.test(nfkc) || NAME_KEYED_SUBSTITUTE_RE.test(nfkc)) return null
  if (!CJK_NAMED_RE.test(nfkc) && !HAN_REGIONAL_NAME_RE.test(nfkc)) return null
  if (!/^(simsun|batang|\u5b8b\u4f53|\ubc14\ud0d5)$/i.test(face.family.normalize('NFKC')))
    return null
  return latinSlotSubstitute(head, docFontTable.get(nfkc.toLowerCase()))
}

function substituteFontFamily(font: string): string | null {
  const face = resolveDeclaredFace(font)
  if (!face?.substituted) return null
  const head = font.split(',')[0].replace(/['"]/g, '').trim()
  const nfkc = head.normalize('NFKC')
  if (ARABIC_NAME_RE.test(nfkc) || NAME_KEYED_SUBSTITUTE_RE.test(nfkc)) return null
  const slot = latinSlotTarget(font)
  const latin = slot ? latinSlotHead(slot) : []
  // Chromium would draw an installed face Word skips (hangul name with altName
  // Batang, a macOS face under its localized name): the head must not lead
  const headLeads =
    !isBundledFont(head) &&
    !krBatangAltNamed(head) &&
    !(CJK_NAMED_RE.test(nfkc) && isFontAvailable(head))
  const lead = cssChain(...(headLeads ? [head] : []), ...latin)
  const target = cssFontFamily(face.family)
  // hangul under a SimSun-substituted name is Batang (probe 2026-09-06); the
  // SimSun chain carries none, so the 1em hangul faces close it
  const hangulTail =
    /^(simsun|\u5b8b\u4f53)$/i.test(face.family.normalize('NFKC')) && !/[\uac00-\ud7a3]/.test(nfkc)
      ? `,'GenOffice Batang','GenOffice Serif KR'`
      : ''
  const generic = target.lastIndexOf(',')
  const chain = `${target.slice(0, generic)}${hangulTail}${target.slice(generic)}`
  return lead ? `${lead},${chain}` : chain
}

export function cssFontFamily(font: string): string {
  // faces Word lays out directly take the chain of their target face, names it
  // substitutes the chain of the substitute; the name-keyed branches below
  // serve unkeyed targets and the script chains
  const byFace = standInFontFamily(font) ?? substituteFontFamily(font)
  if (byFace) return byFace
  const f = font.toLowerCase()
  const head = f.split(',')[0].replace(/['"]/g, '').trim()
  const chain = cssChain
  // 'Calibri Arabic GO' (fonts.css): Calibri's own Arabic; Carlito has none
  if (f.includes('calibri'))
    return `${chain(font, 'Carlito GO', 'Calibri Arabic GO', CJK_SANS)},sans-serif`
  // Aptos (M365 cloud face, never installed locally): line metrics equal
  // Calibri's (Word probe 2026-08-22) but its advances do not, so the
  // size-adjusted Carlito aliases in fonts.css stand in (probe 2026-09-03).
  // 'GenOffice PUA Blank' keeps AI-residue PUA tokens invisible like Word
  // (Carlito would otherwise supply a box .notdef for them).
  if (f.includes('aptos')) {
    // body, Display and Narrow cuts have fitted aliases; Mono/Serif are unprobed
    const alias = /mono|serif/.test(f)
      ? 'Carlito GO'
      : f.includes('display')
        ? 'Aptos Display GO'
        : f.includes('narrow')
          ? 'Aptos Narrow GO'
          : 'Aptos GO'
    return `${chain(font, alias, 'GenOffice PUA Blank', CJK_SANS)},sans-serif`
  }
  // Cambria Math text runs lay out with Cambria's advances in Word (fonts.css
  // alias); only the symbols need a math face
  if (f.includes('cambria') && f.includes('math'))
    return `${chain(font, 'Cambria Math GO', 'STIX Two Math', 'Caladea', 'Liberation Serif', BOX, CJK_SERIF)},serif`
  // math faces would fall to the unknown-name sans fallback; STIX Two Math ships with macOS,
  // and on Windows the declared name resolves natively
  if (/stix.*math|latin modern math/.test(f))
    return `${chain(font, 'STIX Two Math', 'Caladea', 'Liberation Serif', BOX, CJK_SERIF)},serif`
  if (f.includes('cambria')) return `${chain(font, 'Caladea', BOX, CJK_SERIF)},serif`
  if (f.includes('times')) return `${chain(font, 'Liberation Serif', BOX, CJK_SERIF)},serif`
  // MS Reference Sans Serif ships with Office and shares Verdana's tables; the
  // local() alias (fonts.css) draws Verdana where the face is missing
  if (f.includes('reference sans'))
    return `${chain(font, 'MS Reference Sans Serif GO', 'Verdana', CJK_SANS)},sans-serif`
  // Merriweather is an M365 cloud face Word renders real; the size-adjusted
  // Georgia alias (fonts.css) carries its advances (sample lines 1.12x Georgia,
  // sd 0.6%) and keeps the serif class the name-guess fallback would miss
  if (f.includes('merriweather') && !f.includes('sans'))
    return `${chain(font, 'Merriweather GO', 'Georgia', BOX, CJK_SERIF)},serif`
  // Palatino Linotype/Book Antiqua ship with Office and Word renders them real
  // (probe 2026-08-23); macOS Palatino matches their Latin widths within 0.2%
  if (f.includes('palatino') || f.includes('book antiqua'))
    return `${chain(font, 'Palatino Linotype', 'Palatino', 'Book Antiqua', BOX, CJK_SERIF)},serif`
  // 'Arial Arabic GO' (fonts.css): Arial's Arabic advances where the local
  // Arial (Liberation Sans) has none
  if (f === 'arial' || f.startsWith('arial '))
    return `${chain(font, 'Arial Arabic GO', 'Liberation Sans', CJK_SANS)},sans-serif`
  // Roboto is an M365 cloud face Word renders real (Roboto-Bold spans in its
  // PDF); the generic sans-serif tail would draw Helvetica, 6-8% wider in
  // bold, so the per-class size-adjusted Liberation Sans alias (fonts.css)
  // carries its advances. Mono/Slab/Condensed cuts are other designs
  if (/^roboto$/.test(head)) return `${chain(font, 'Roboto GO', CJK_SANS)},sans-serif`
  // Garamond ships with Office and Word renders it real; it is a macOS
  // on-demand asset Chromium cannot see, and the serif tail (Times) runs 3-5%
  // wider, so the size-adjusted Liberation Serif alias (fonts.css) stands in
  if (/^garamond$/.test(head)) return `${chain(font, 'Garamond GO', BOX, CJK_SERIF)},serif`
  // Berlin Sans FB Demi ships with Office and Word renders it real, a bold
  // display cut 0.953x Arial Bold wide (probe 2026-09-17); the size-adjusted
  // bold alias (fonts.css) stands in where the face is missing. The regular
  // cut is a different (unprobed) weight and keeps the generic fallback
  if (/berlin sans fb demi/.test(f))
    return `${chain(font, 'Berlin Sans FB GO', CJK_SANS)},sans-serif`
  // Segoe UI is an M365 cloud font Word renders real; where it is missing the
  // size-adjusted Helvetica alias (fonts.css) carries its narrower advances
  if (SEGOE_UI_TEXT_RE.test(f))
    return `${chain(font, 'Segoe UI GO', 'Noto Sans CJK GO', 'GenOffice PUA Blank')},sans-serif`
  // Consolas is Office-only: the size-adjusted Menlo alias (fonts.css) carries its advances
  if (f.includes('consolas')) return `${chain(font, 'Consolas GO', CJK_SANS)},monospace`
  // Century Gothic is Office-only too: the per-case size-adjusted Liberation
  // Sans alias (fonts.css) carries its wide lowercase / narrow capitals
  if (f.includes('century gothic'))
    return `${chain(font, 'Century Gothic GO', CJK_SANS)},sans-serif`
  // Lucida Handwriting ships with Office and Windows; elsewhere the per-case
  // size-adjusted Liberation Sans Italic alias (fonts.css) carries its wide script advances
  if (f.includes('lucida handwriting'))
    return `${chain(font, 'Lucida Handwriting GO', CJK_SANS)},sans-serif`
  if (MONO_FONT_RE.test(f))
    return `${chain(font, 'Menlo', 'Courier New', 'Liberation Mono', CJK_SANS)},monospace`
  // Nunito Sans is an Office cloud font Word renders real; the size-adjusted
  // Helvetica alias (fonts.css) matches its advances (probe 2026-08-23)
  if (f.includes('nunito')) return `${chain(font, 'Nunito Sans GO', CJK_SANS)},sans-serif`
  // Poppins is an M365 cloud font Word renders real; the bundled Latin subset
  // (fonts.css) carries its true advances (probe 2026-09-01)
  if (f.includes('poppins')) return `${chain(font, 'GenOffice Poppins', CJK_SANS)},sans-serif`
  // DM Sans renders real in Word when installed; the bundled Latin subset
  // (fonts.css) carries its true advances (corpus 2026-09-30)
  if (f.includes('dm sans')) return `${chain(font, 'GenOffice DM Sans', CJK_SANS)},sans-serif`
  // Montserrat is an M365 cloud font Word renders real, ignoring its fontTable
  // altName; the per-class size-adjusted Liberation Sans alias (fonts.css)
  // carries its wide advances. Only the cloud family names qualify: a
  // PostScript-style "Montserrat-Bold" is unknown to Word and keeps the altName
  if (MONTSERRAT_FAMILY_RE.test(f)) return `${chain(font, 'Montserrat GO', CJK_SANS)},sans-serif`
  // Microsoft New Tai Lue ships with Office; its Latin is Segoe-flavored with
  // Arial-class widths (probe 2026-08-23: +0.5% vs Helvetica)
  if (f.includes('new tai lue'))
    return `${chain(font, 'Segoe UI', 'Helvetica', 'Liberation Sans', CJK_SANS)},sans-serif`
  // ZhongSong (STZhongsong, gov-document title font) before the generic SimSun branch
  if (f.includes('中宋') || f.includes('zhongsong'))
    return `${chain(font, 'STZhongsong', 'Songti SC', 'STSong', 'SimSun', CJK_SERIF)},serif`
  // 'GenOffice Songti SC' (fonts.css local() alias of Songti SC): macOS Chromium
  // refuses synthetic bold for 'Songti SC' by name at weight 600/700; the alias,
  // registered weight-normal only, lets Blink synthesize. Unresolvable elsewhere.
  // SimSun's ASCII is half-width fixed pitch (0.5em); the Songti stand-in is proportional
  if (/simsun|\u5b8b\u4f53|nsimsun/.test(f)) {
    return `${chain(font, 'GenOffice SimSun Latin', 'GenOffice Songti SC', 'STSong', 'SimSun', CJK_SERIF)},serif`
  }
  if (f.includes('simhei') || f.includes('黑体') || f.includes('细黑') || f.includes('xihei'))
    return `${chain(font, 'Heiti SC', 'STHeiti', 'SimHei', 'PingFang SC', CJK_SANS)},sans-serif`
  if (f.includes('yahei') || f.includes('雅黑'))
    return `${chain(font, 'Microsoft YaHei', 'PingFang SC', CJK_SANS)},sans-serif`
  if (f.includes('等线') || f.includes('dengxian'))
    return `${chain(font, 'DengXian GO', 'DengXian', 'PingFang SC', 'Microsoft YaHei', CJK_SANS)},sans-serif`
  if (f.includes('fangsong') || f.includes('仿宋'))
    return `${chain(font, 'STFangsong', 'FangSong', CJK_SERIF)},serif`
  if (f.includes('kaiti') || f.includes('楷体'))
    return `${chain(font, 'STKaiti', 'Kaiti SC', 'KaiTi', CJK_SERIF)},serif`
  if (f.includes('隶书') || f.includes('lisu'))
    return `${chain(font, 'Baoli SC', 'LiSu', CJK_SERIF)},serif`
  // Japanese/Korean/Traditional Chinese: fall back within the same script (win/mac family names as mutual backups) so Han glyphs don't render with Simplified forms.
  // 'GenOffice *' entries are CJK-only local() aliases (fonts.css): the underlying
  // system faces draw Cyrillic/Greek fullwidth, so those scripts must pass through
  const TC_SANS = ['Microsoft JhengHei', 'PingFang TC', 'GenOffice Heiti TC', 'Noto Sans TC']
  // 'GenOffice Fullwidth TC' (fonts.css): fullwidth U+FF0D/FF0F/FF3C/FF3F/FF5E whose Songti TC glyphs look half-width
  // 'GenOffice Songti TC' (fonts.css): weight-normal alias so bold synthesizes like Word's faux-bold PMingLiU
  const TC_SERIF = [
    'GenOffice MingLiU',
    'GenOffice Fullwidth TC',
    'GenOffice Songti TC',
    'Songti TC',
    'Noto Serif TC',
  ]
  const SC_SANS = ['PingFang SC', 'Microsoft YaHei', CJK_SANS]
  const SC_SERIF = ['GenOffice Songti SC', 'STSong', 'SimSun', CJK_SERIF]
  const nfkc = font.normalize('NFKC')
  // Arabic: bundled Noto subsets stand in for missing fonts; Chromium's silent
  // fallback is a Geeza Pro-style UI face, larger and heavier than the naskh
  // serif Word substitutes. Unknown Arabic names default to the naskh chain.
  // Iranian B/XB/IR faces (B Mitra, B Nazanin, XB Zar...) included: Word for
  // Mac substitutes all of them with Times New Roman (probe 2026-08-13)
  if (ARABIC_NAME_RE.test(nfkc)) {
    const alt = arabicAltNameFace(nfkc) ?? arabicCodePageFace(nfkc)
    if (alt) return cssFontFamily(alt)
    // Dubai is a cloud face Word renders real; the size-adjusted alias
    // (fonts.css) carries its advances and hhea box where it is missing
    if (/\bdubai\b/i.test(nfkc) && !isBundledFont(font) && !isFontAvailable(font))
      return `${chain(font, 'Dubai GO', 'Arial', 'Liberation Sans', CJK_SANS)},sans-serif`
    // an installed kufi/sans face renders as such; a missing one substitutes to
    // Times New Roman in Word like every other missing Arabic name (probe
    // 2026-09-11: Noto Kufi Arabic list items came back as TimesNewRomanPSMT).
    // The bundled Noto Sans Arabic subset answers the canvas probe but is not
    // an installed face.
    const sans =
      /kufi|sans|dubai|segoe/i.test(nfkc) && !isBundledFont(font) && isFontAvailable(font)
    // Traditional/Simplified Arabic are compact naskh faces Word renders real
    // (M365 cloud fonts); the per-face size-adjusted aliases (fonts.css) carry
    // their advances, other Arabic names keep the unscaled subset
    const compact = /\b(traditional|simplified) arabic\b/i.exec(nfkc)?.[1].toLowerCase()
    // Sakkal Majalla (cloud face, compact naskh): per-class aliases carry its
    // Arabic, digit, punctuation and Latin advances
    const sakkal = /\bsakkal majalla\b/i.test(nfkc)
    // declared Noto Arabic names would resolve to the bundled unscaled subsets
    // (~13% wider than Word); the literal head must go so the 'W' alias wins
    const scaled = /^noto (naskh|sans) arabic$/i.test(nfkc.trim())
    // Word substitutes missing naskh names wholesale with Times New Roman, so
    // ASCII punctuation/digits must take Times shapes, not the Naskh subset's
    // (sunken parens/hyphen). mac/Windows Times carries full Arabic, so on
    // those platforms Arabic letters also resolve here — same face Word uses;
    // the scaled subset below serves platforms without a system Times.
    // Calibrated aliases (compact/scaled) stay as-is.
    const missingSerif = !sans && !compact && !scaled && !sakkal && !isFontAvailable(font)
    const latinHead = missingSerif ? ['Times New Roman', 'Liberation Serif'] : []
    const chainFor = sans
      ? [scaled ? 'Noto Sans Arabic W' : 'Noto Sans Arabic', 'Geeza Pro']
      : [
          ...latinHead,
          // missing serif names: 90% TNR alias (subset shapes 1.125×/1.074×
          // wider than Word's Times, M3 probe); Arabic-Indic digits unscaled
          ...(missingSerif
            ? ['Naskh Digits GO', 'Noto Naskh Arabic TNR']
            : sakkal
              ? ['Sakkal Majalla GO', 'Sakkal Majalla Latin GO']
              : compact
                ? [
                    compact === 'traditional' ? 'Noto Naskh Arabic TA' : 'Noto Naskh Arabic SA',
                    // the subsets carry no ASCII brackets and neither does
                    // Geeza Pro, so they fell to Al Bayan (30% wider, 1.5 em
                    // box); both cloud faces draw them at Times widths
                    'Times Punct GO',
                  ]
                : [scaled ? 'Noto Naskh Arabic W' : 'Noto Naskh Arabic']),
          'Geeza Pro',
          'Al Bayan',
        ]
    return `${chain(...(scaled ? [] : [font]), ...chainFor)},${sans ? 'sans-serif' : 'serif'}`
  }
  // Noto CJK / Source Han regional variants Word lays out for real (installed;
  // the missing ones are substituted above) route by suffix
  const cjkVariant = /^(?:noto|source han) (sans|serif)(?: cjk)? ?(jp|kr|k|tc|tw|hk|sc|cn)\b/i.exec(
    nfkc,
  )
  if (cjkVariant) {
    const region = cjkVariant[2].toLowerCase()
    const serif = /serif/i.test(cjkVariant[1])
    const chainFor =
      region === 'jp'
        ? serif
          ? JA_SERIF
          : JA_SANS
        : region === 'kr' || region === 'k'
          ? serif
            ? KO_SERIF
            : KO_SANS
          : region === 'sc' || region === 'cn'
            ? serif
              ? SC_SERIF
              : SC_SANS
            : serif
              ? TC_SERIF
              : TC_SANS
    return `${chain(font, ...chainFor)},${serif ? 'serif' : 'sans-serif'}`
  }
  if (JA_FONT_RE.test(nfkc)) {
    const serif = /mincho|\u660e\u671d/i.test(nfkc)
    // Meiryo (UI): Word renders the real faces; the range-limited aliases
    // (fonts.css) rescale the Hiragino/Verdana fallbacks to Meiryo advances
    if (/meiryo|メイリオ/i.test(nfkc)) {
      const alias = /(meiryo|メイリオ) ?ui/i.test(nfkc) ? 'Meiryo UI GO' : 'Meiryo GO'
      return `${chain(font, alias, ...JA_SANS)},sans-serif`
    }
    // MS Gothic family renders real in Word (half-width mono Latin / proportional
    // kana); the range-limited aliases (fonts.css) reproduce those advances.
    // MS Mincho (Word probe 2026-09-17: real MS-Mincho, digits 0.5em, fullwidth
    // colon 1em) shares the fixed-pitch Latin alias
    const msGothic = /^ms (ui )?(p)?(gothic|ゴシック)$/i.exec(nfkc.trim())
    const msMincho = /^ms (mincho|\u660e\u671d)$/i.test(nfkc.trim())
    const msAlias = msMincho
      ? ['MS Mincho GO']
      : !msGothic
        ? []
        : msGothic[1]
          ? ['MS UI Gothic GO', 'MS UI Gothic JA GO']
          : msGothic[2]
            ? ['MS PGothic GO', 'MS PGothic JA GO']
            : ['MS Gothic GO']
    // Yu Mincho/Gothic are Office-private on macOS (Word lays out with its own
    // yumin/YuGoth faces); a missing declare takes the Latin advance aliases
    // (fonts.css) so EN/ID lines wrap where Word's do. Yu Gothic UI stays out.
    const yu =
      /^(?:yu (?:mincho|gothic)|\u6e38(?:\u660e\u671d|\u30b4\u30b7\u30c3\u30af))(?: (light|regular|medium|bold|demibold))?$/i.exec(
        nfkc.trim(),
      )
    const yuLight = yu?.[1]?.toLowerCase() === 'light'
    const yuAlias =
      yu && !isFontAvailable(font)
        ? serif
          ? [yuLight ? 'Yu Mincho Light GO' : 'Yu Mincho GO', 'Yu Mincho Serif GO']
          : [yuLight ? 'Yu Gothic Light GO' : 'Yu Gothic GO', 'Yu Gothic Sans GO']
        : []
    return `${chain(font, ...msAlias, ...yuAlias, ...(serif ? JA_SERIF : JA_SANS))},${serif ? 'serif' : 'sans-serif'}`
  }
  if (
    /[가-힣ᄀ-ᇿ㄰-㆏]|malgun|batang|gulim|dotum|gungsuh|myeongjo|myungjo|nanum|apple (sd )?gothic/i.test(
      nfkc,
    )
  ) {
    // -Che fixed-pitch faces render real in Word with half-width Latin
    // (0.5em fixed, probe 2026-08-24); ASCII rides the bundled Che face,
    // hangul falls through to the KR chains (GungsuhChe keeps the
    // calligraphic GungSeo hangul the Gungsuh branch below uses)
    if (/(?:batang|gulim|dotum|gungsuh) ?che\b|(?:바탕|굴림|돋움|궁서)체/i.test(nfkc)) {
      const cheGungsuh = /gungsuh|궁서/i.test(nfkc)
      const cheSerif = cheGungsuh || /batang|바탕/i.test(nfkc)
      return `${chain(font, 'GenOffice Che Latin KR', ...(cheGungsuh ? ['GungSeo'] : []), ...(cheSerif ? KO_SERIF : KO_SANS))},${cheSerif ? 'serif' : 'sans-serif'}`
    }
    // Gungsuh ships with Office (batang.ttc) and Word renders it real; its
    // Latin is typewriter-slab at ~0.58em advances — Courier New is the
    // closest installed face (probe 2026-08-23, +2.8% width). Hangul falls
    // through to the calligraphic GungSeo, then the serif KR chain.
    if (/gungsuh|궁서/i.test(nfkc)) {
      return `${chain(font, 'Courier New', 'GungSeo', ...KO_SERIF)},serif`
    }
    const serif = /batang|바탕|myeongjo|myungjo|명조|gungsuh|궁서/i.test(nfkc)
    return `${chain(font, ...(serif ? KO_SERIF : KO_SANS))},${serif ? 'serif' : 'sans-serif'}`
  }
  // DFKai-SB is missing on Word for Mac, which substitutes Microsoft YaHei
  // wholesale (probe 2026-09-23); the Latin alias keeps YaHei's wide digits
  // where YaHei itself is missing. Not gated on availability: macOS lists the
  // name for Apple's on-demand BiauKai asset while drawing nothing with it
  if (/\u6a19\u6977|dfkai/.test(f))
    return `${chain(font, 'Microsoft YaHei', 'GenOffice YaHei Latin', 'PingFang SC', CJK_SANS)},sans-serif`
  if (
    /jhenghei|p?mingliu|biaukai|dfkai|kaiu|正黑|細明|標楷|蘋方|-繁|繁體|pingfang (tc|hk)|(heiti|songti|kaiti) tc/i.test(
      nfkc,
    )
  ) {
    const serif = /mingliu|細明|標楷|biaukai|dfkai|kaiu|songti|kaiti|宋/i.test(nfkc)
    const alias = serif ? wordMetricAliasFace(font) : null
    return `${chain(font, ...(alias ? [alias] : []), ...(serif ? TC_SERIF : TC_SANS))},${serif ? 'serif' : 'sans-serif'}`
  }
  // Ethiopic faces (Nyala/Ebrima/Abyssinica) are missing on macOS; the size-adjusted
  // Kefa alias (fonts.css) re-centers Chromium's ~1.35x-wide Kefa fallback. On
  // Windows the declared name resolves natively ahead of the alias.
  if (/nyala|ebrima|abyssinica|ethiopic/i.test(nfkc)) {
    return `${chain(font, 'GenOffice Ethiopic')},sans-serif`
  }
  // Bengali: Nirmala UI is an M365 cloud face Word renders real; macOS has no
  // twin and Chromium's Kohinoor Bangla shapes ~8% narrower, so the bundled
  // Nirmala-metric face (fonts.css) leads. On Windows the declared name
  // resolves natively ahead of the chain.
  if (/nirmala/i.test(nfkc)) {
    return `${chain(font, 'GenOffice Bengali', 'Kohinoor Bangla', 'Bangla Sangam MN')},sans-serif`
  }
  // Tamil: Latha renders through the bundled Latha-metric face (fonts.css);
  // Vijaya and missing Tamil names lay out as Vijaya on Word for Mac (probe
  // 2026-09-30), whose advances InaiMathi matches within 1 % (Latha is 44 %
  // wider). On Windows the declared name resolves natively ahead of the chain.
  if (/latha/i.test(nfkc)) {
    return `${chain(font, 'GenOffice Tamil', 'InaiMathi', 'Tamil MN', 'Tamil Sangam MN')},sans-serif`
  }
  if (TAMIL_FONT_RE.test(nfkc)) {
    return `${chain(font, 'Vijaya', 'InaiMathi', 'Tamil Sangam MN', 'Tamil MN')},sans-serif`
  }
  // unknown font family: fontTable PANOSE decides serif-ness when the font is
  // missing, else guess by name (Song/Ming/Serif → serif fallback)
  const hint = isBundledFont(font) || !isFontAvailable(font) ? panoseSerifHint(font) : undefined
  const serifLike = hint ? hint === 'serif' : /宋|明|serif|song|ming/i.test(font)
  // Latin-named unknown faces take the range-limited CJK tail: the bundled
  // subset covers U+2018/2019/201C/201D at 1em and would otherwise swallow
  // Latin punctuation. CJK-named faces keep the full subset (fullwidth quotes
  // in CJK body text are correct).
  const latinNamed = !CJK_NAMED_RE.test(nfkc)
  const tail = serifLike
    ? latinNamed
      ? 'Noto Serif CJK GO'
      : CJK_SERIF
    : latinNamed
      ? 'Noto Sans CJK GO'
      : CJK_SANS
  // the GO aliases exclude PUA, so the blank face keeps those codepoints
  // invisible (the full subsets' blank .notdef used to catch them)
  const pua = latinNamed ? ['GenOffice PUA Blank'] : []
  return `${chain(font, tail, ...pua)},${serifLike ? 'serif' : 'sans-serif'}`
}

/**
 * Latin-capable head of a font's chain: its families minus the CJK/generic
 * tail, closed by a Latin-range backstop alias of the same class (fonts.css).
 * Word substitutes a missing ascii face with a Latin font — Latin glyphs never
 * render from the eastAsia slot (probe 2026-08-23) — so the backstop keeps
 * Latin text off the EA families when nothing earlier resolves.
 */
function latinChainHead(ascii: string): string[] {
  const fams = cssFontFamily(ascii).split(',')
  const generic = fams[fams.length - 1]
  const latin = fams.filter(
    (fam) =>
      !/noto (sans|serif) cjk/i.test(fam) &&
      !/ JA GO'$/.test(fam) &&
      !/^(serif|sans-serif|monospace)$/.test(fam),
  )
  // mono chains already carry the bundled Liberation Mono
  if (generic !== 'monospace') {
    latin.push(generic === 'serif' ? "'Latin Serif GO'" : "'Latin Sans GO'")
  }
  return latin
}

/** --doc-latin-chain value for a document/style base ascii face (EA-only runs
 *  resolve their Latin glyphs through this inherited var, see cssEaOnlyFontFamily) */
export function docLatinChainCss(ascii: string): string {
  return latinChainHead(ascii).join(',')
}

/** class-matched Latin-range backstop alias for a face's chain */
function latinBackstopFor(font: string): string {
  return cssFontFamily(font).endsWith('sans-serif') ? "'Latin Sans GO'" : "'Latin Serif GO'"
}

/**
 * Fallback chain for a dual-slot run (w:ascii ≠ w:eastAsia): Latin families first,
 * then the full East Asian chain. The Latin part drops its own CJK/generic fallbacks
 * so CJK glyphs fall through to the eastAsia font instead of a lookalike.
 */
export function cssDualFontFamily(ascii: string, eastAsia: string): string {
  if (ascii === eastAsia) return cssFontFamily(ascii)
  // Korean ascii face (e.g. theme latin = Malgun): its fallback chain covers hangul
  // and would swallow the eastAsia font, so keep only the literal family — plus
  // the Latin-range backstop so a missing KO face still keeps Latin off the EA slot
  if (KO_FONT_RE.test(ascii.normalize('NFKC'))) {
    return `'${ascii.replace(/'/g, '')}',${latinBackstopFor(ascii)},${cssFontFamily(eastAsia)}`
  }
  return `${latinChainHead(ascii).join(',')},${cssFontFamily(eastAsia)}`
}

/**
 * Chain for a run declaring only w:eastAsia: Word resolves its Latin glyphs
 * through the style-inherited ascii face, never the EA font (probe 2026-08-23,
 * docDefaults and pStyle level both). --doc-latin-chain carries the inherited
 * ascii chain (doc-style-css emits it with every base font-family); the
 * class-matched backstop covers contexts without it.
 */
export function cssEaOnlyFontFamily(ea: string): string {
  const eaChain = cssFontFamily(ea)
  const backstop = eaChain.endsWith('sans-serif') ? "'Latin Sans GO'" : "'Latin Serif GO'"
  return `var(--doc-latin-chain,${backstop}),${eaChain}`
}

/** font-family for a run's declared slots (dual / Latin-only / EA-only) */
/**
 * Alias whose bold/italic cuts carry a size-adjust far from the regular cut
 * (fonts.css): under the regular-cut strut such a span lays a box of another
 * size and the baseline-aligned union inflates the line (Word keeps the pitch).
 */
export function cutAdjustedAlias(cssFamily: string): boolean {
  return cssFamily.includes("'Merriweather GO'")
}

export function cssRunFontFamily(
  ascii: string | null | undefined,
  ea: string | null | undefined,
): string {
  if (ascii && ea) return cssDualFontFamily(ascii, ea)
  if (ea) return cssEaOnlyFontFamily(ea)
  return cssFontFamily(ascii!)
}

/** Text contains complex-script characters (Arabic/Hebrew/Syriac/Thaana/NKo/Indic/Thai), i.e. the w:cs font slot applies */
export function textHasComplexScript(text: string): boolean {
  return /[\u0590-\u05FF\u0600-\u077F\u0780-\u07FF\u08A0-\u08FF\u0900-\u0DFF\u0E00-\u0E7F\uFB1D-\uFDFF\uFE70-\uFEFF]/.test(
    text,
  )
}

/** w:w horizontal scaling approximated as letter-spacing (em, negative = condensed),
 *  weighted by the run text's wide/narrow glyph mix */
export function charScaleEm(text: string, scalePct: number): number {
  let sum = 0
  let n = 0
  for (const ch of text) {
    const cp = ch.codePointAt(0)!
    const wide =
      (cp >= 0x2e80 && cp <= 0x9fff) ||
      (cp >= 0xac00 && cp <= 0xd7af) ||
      (cp >= 0xf900 && cp <= 0xfaff) ||
      (cp >= 0xff00 && cp <= 0xff60) ||
      cp >= 0x20000
    sum += wide ? 1 : 0.52
    n++
  }
  const avg = n > 0 ? sum / n : 0.52
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

export function runLetterSpacingCss(run: {
  text: string
  charSpacingTwips?: number
  charScalePct?: number
}): string | null {
  const spacingPt = run.charSpacingTwips ? run.charSpacingTwips / 20 : 0
  const scaleEm = run.charScalePct ? charScaleEm(run.text, run.charScalePct) : 0
  if (spacingPt && scaleEm) return `calc(${spacingPt}pt + ${scaleEm}em)`
  if (spacingPt) return `${spacingPt}pt`
  if (scaleEm) return `${scaleEm}em`
  return run.charSpacingTwips === 0 ? '0' : null
}

/**
 * Fallback chain for a run whose text hits the w:cs slot: the cs chain leads
 * (minus its generic tail) so complex-script glyphs use it, then the run's
 * Latin/East Asian chain for everything else.
 */
export function cssCsFontFamily(cs: string, ascii?: string, eastAsia?: string): string {
  // eastAsia-only base: a plain class backstop instead of the --doc-latin-chain
  // var (the merge below splits the base on commas, which a var() cannot survive);
  // Latin glyphs still never render from the EA slot
  const base =
    ascii && eastAsia && ascii !== eastAsia
      ? cssDualFontFamily(ascii, eastAsia)
      : ascii
        ? cssFontFamily(ascii)
        : eastAsia
          ? `${latinBackstopFor(eastAsia)},${cssFontFamily(eastAsia)}`
          : ''
  if (!base) return cssFontFamily(cs)
  const head = cssFontFamily(cs)
    .split(',')
    .filter((f) => !/^(serif|sans-serif|monospace)$/.test(f))
  const baseFams = base.split(',')
  const generic = /^(serif|sans-serif|monospace)$/
  // Bundled Noto Arabic subsets have no Latin letters; splice the run's
  // Latin chain in right after them, before Geeza Pro whose Latin punctuation
  // sits on the Arabic baseline. Arabic glyphs resolve in the subset first, so
  // shaping is unaffected. Full Times entries must sit after the ascii chain
  // (Latin letters belong to the ascii slot per Word; Times only backstops a
  // missing ascii font), but ASCII punctuation/digits belong to the
  // substituted cs font — the unicode-range 'Times Punct GO' alias claims
  // just those ahead of the subset's sunken forms.
  const notoIdx = head.findIndex((f) =>
    /noto (naskh|sans) arabic|sakkal majalla go\b|dubai go\b|arial arabic go\b/i.test(f),
  )
  const latinSub = /times new roman|liberation serif/i
  const pre = head.slice(0, notoIdx + 1)
  const preKept = pre.filter((f) => !latinSub.test(f))
  if (notoIdx >= 0 && pre.some((f) => latinSub.test(f))) preKept.splice(-1, 0, "'Times Punct GO'")
  const merged =
    notoIdx >= 0
      ? [
          ...preKept,
          ...baseFams.filter((f) => !generic.test(f)),
          ...pre.filter((f) => latinSub.test(f)),
          ...head.slice(notoIdx + 1),
          ...baseFams.filter((f) => generic.test(f)),
        ]
      : [...head, ...baseFams]
  return [...new Set(merged)].join(',')
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

/** --doc-line-factor-kr source: the document's East Asian face when Korean, else the Batang-class default */
export function krLineFactor(fontFamily: string | undefined): number {
  if (!fontFamily) return batangLineFactor()
  const krName = krNameLineFactor(fontFamily)
  if (krName !== null) return krName
  if (isKoreanFontName(fontFamily)) return lineHeightFactor(fontFamily)
  if (krCharsetMalgunSubstituted(fontFamily)) return malgunLineFactor()
  // hangul a Japanese face cannot draw goes to Malgun Gothic (Word probe 2026-09-30)
  return isJapaneseFontName(fontFamily) ? malgunLineFactor() : batangLineFactor()
}

export function isKoreanFontName(fontFamily: string): boolean {
  return KO_FONT_RE.test(fontFamily.normalize('NFKC'))
}

export function isJapaneseFontName(fontFamily: string): boolean {
  return JA_FONT_RE.test(fontFamily.normalize('NFKC'))
}

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
let additiveSpacing = false
export function setAdditiveParagraphSpacing(on: boolean): void {
  additiveSpacing = on
}
export function additiveParagraphSpacing(): boolean {
  return additiveSpacing
}
export function spaceBeforeProp(additive = additiveSpacing): 'padding-top' | 'margin-top' {
  return additive ? 'padding-top' : 'margin-top'
}
export function autoSpacingPt(additive = additiveSpacing): number {
  return additive ? 0 : WORD_AUTO_SPACING_PT
}

/** space-only run text: never sizes a line, the paragraph mark does (Word probe 2026-09-11) */
export const SPACE_ONLY_RE = /^[ \u00a0\u3000]+$/

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

/** Latin text faces whose lead shift is modelled; other heads keep the CSS centre until a corpus run adopts them */
const LINE_BOX_FACE_RE =
  /aptos|calibri|carlito|cambria|caladea|times|liberation serif|\barial\b(?! unicode)|liberation sans|georgia|verdana|reference sans|tahoma|courier|trebuchet/

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
  const box = LINE_BOX_FACE_RE.test(face) ? chromiumBox(chain) : null
  // a head outside the modelled set must not keep an inherited box (CSS centre instead)
  if (!box) return ['--doc-font-box:initial', '--doc-font-skew:initial']
  const { ascent, descent } = box
  return [
    `--doc-font-box:${(ascent + descent).toFixed(4)}`,
    `--doc-font-skew:${(ascent - descent).toFixed(4)}`,
  ]
}

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

export function isCjk(cp: number): boolean {
  return (
    isHangul(cp) ||
    (cp >= 0x3000 && cp <= 0x30ff) ||
    (cp >= 0x3400 && cp <= 0x9fff) ||
    (cp >= 0xf900 && cp <= 0xfaff) ||
    (cp >= 0xff00 && cp <= 0xffef)
  )
}

function isHangul(cp: number): boolean {
  return (
    (cp >= 0xac00 && cp <= 0xd7a3) ||
    (cp >= 0x1100 && cp <= 0x11ff) ||
    (cp >= 0x3130 && cp <= 0x318f) ||
    (cp >= 0xa960 && cp <= 0xa97f) ||
    (cp >= 0xd7b0 && cp <= 0xd7ff)
  )
}

// ─── CJK-Latin autospace pads ────────────────────────────────────────────────
// Word's autoSpaceDE/DN gap measures ~1/4em while Chromium's text-autospace is
// fixed at 1/8em; the renderer wraps the character after each boundary in a
// .doc-autospace-pad span whose start margin supplies the other 1/8em. Pads
// only go between a directly adjacent CJK letter and a Latin letter/digit —
// never next to spaces or punctuation — matching where Chromium applies its
// native gap. Chromium cannot switch letters and digits separately (no
// ideograph-alpha / ideograph-numeric), so when only one of w:autoSpaceDE/DN
// is off the native gap goes off and the pad carries the whole 1/4em
// (.doc-autospace-pad-full).

/** Han/kana/hangul letters; CJK punctuation and full/halfwidth forms get no gap */
function isCjkAutospaceSide(cp: number): boolean {
  if (cp >= 0x3000 && cp <= 0x303f) return false
  if (cp === 0x30fb) return false
  if (cp >= 0xff00 && cp <= 0xffef) return false
  return isCjk(cp)
}

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

/** OS/2 xAvgCharWidth (em) of a declared Latin name's Word layout face; null for East Asian or unknown faces */
function latinAvgCharWidth(font: string): number | null {
  const face = latinFace(font)
  return face && !face.metrics.eastAsian ? face.metrics.xAvg : null
}
/** a Latin name Word lacks lays out in Cambria (fallback: the 1/4 em gap, doubled) */
const cambriaAvgCharWidth = () => latinAvgCharWidth('Cambria') ?? 0.5

const latinGapCache = new Map<string, number>()
/**
 * Word probe 2026-10-20: the autospace gap between a Latin letter/digit and
 * the CJK letter after it is half the Latin face's OS/2 xAvgCharWidth
 * (Cambria 0.308, Calibri 0.260, Aptos 0.280, Arial 0.221, Times New Roman
 * 0.200, Garamond 0.194 em) whatever the EA face or size; the CJK -> Latin
 * gap and CJK-capable ascii faces (Batang, Malgun, Yu Gothic) stay 1/4 em.
 * A Latin name Word lacks lays out in Cambria, as does the Latin of a missing
 * KR superfamily name (the 'KR Theme Latin GO' alias).
 */
export function latinAutospaceGapEm(font: string | null | undefined): number {
  if (!font) return 0.25
  const cached = latinGapCache.get(font)
  if (cached !== undefined) return cached
  // a hangul-named face resolves to its EA face; its Latin slot is the KR theme rule below
  const avg = isCjkNamedFont(font) && !legacyFaceAlias(font) ? null : latinAvgCharWidth(font)
  let gap: number
  if (avg !== null) gap = avg / 2
  else if (isCjkNamedFont(font)) {
    const latin = latinSlotFace(font)
    gap = latin ? (latinAvgCharWidth(latin) ?? 0.5) / 2 : 0.25
  } else if (!isBundledFont(font) && !isFontAvailable(font)) gap = cambriaAvgCharWidth() / 2
  else gap = 0.25
  latinGapCache.set(font, gap)
  return gap
}

/**
 * Latin face a CJK-named face's Latin lays out in: the split Latin slot's
 * target, or the Latin face Word substitutes wholesale (altName Arial); the
 * same face the CSS chain leads with. Null when the declared face draws its
 * own Latin or the slot lands on an East Asian face (SimSun), whose box the
 * EA rule owns. Any CJK-lettered name qualifies, not only the known families:
 * the resolver decides whether it was substituted and onto what
 */
function latinSlotFace(font: string): string | null {
  if (!isCjkNamedFont(font)) return null
  const target = latinSlotTarget(font)
  if (target) {
    const face = latinFace(target)
    return face && !face.metrics.eastAsian ? target : null
  }
  const face = resolveDeclaredFace(font)
  return face?.substituted && !face.metrics.eastAsian ? face.family : null
}

/** Latin-only lines of a missing KR superfamily name lay out in Word's Cambria substitute (1.172,
 *  prod corpus 20261020 sample 023 URL lines), not in the hangul face's 1.3029 */
export function latinLineFactor(font: string): number {
  const latin = latinSlotFace(font)
  return latin ? lineHeightFactor(latin) : lineHeightFactor(font)
}

/** multiplier of the 1/8 em pad base for a Latin -> CJK seam (.doc-autospace-pad-le); 1 = the plain pad */
export function autospacePadScale(gapEm: number, mode: AutospaceMode): number {
  const pad = gapEm - (mode === true ? 0.125 : 0)
  return Math.round((pad / 0.125) * 1000) / 1000
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

// ─── Hangul-context spaces ───────────────────────────────────────────────────
// Under settings.xml balanceSingleByteDoubleByteWidth AND useFELayout Word
// lays out a space touching a hangul character at 0.5em of the run size
// whatever the font's own space advance (corpus probe 2026-09-24: Batang
// 0.49-0.50, Malgun Gothic and a Times New Roman space inside a Korean run
// 0.51-0.52), while spaces between Latin letters, digits or punctuation keep
// the glyph's advance (0.25-0.36em). One hangul neighbour is enough (hangul
// then Latin or Latin then hangul are wide; hangul, comma, space, Latin is
// not). Either flag alone leaves every space at the glyph advance of the
// ascii-slot face (Word probe 2026-09-30: Calibri 0.224, Arial 0.276, Batang
// 0.330, Malgun 0.349, Gulim 0.330; w:hint, lang and kerning change nothing). The renderer wraps each such space in a .doc-hangul-space span
// whose only glyph comes from the half-width space face; a space left in its
// text node stays in the neighbouring hangul face, so nothing is wrapped when
// the flag is off (an isolated span would fall to the Latin face and its taller
// box would lift the line).

let hangulSpaceWidening = false
export function hangulSpaceWideningFor(settings: {
  balanceDbcsSpacing?: boolean
  useFELayout?: boolean
}): boolean {
  return settings.balanceDbcsSpacing === true && settings.useFELayout === true
}
/** hangulSpaceWideningFor(settings) of the open document */
export function setHangulSpaceWidening(on: boolean): void {
  hangulSpaceWidening = on
}
export function hangulSpaceWideningOn(): boolean {
  return hangulSpaceWidening
}

/** UTF-16 offsets of the U+0020 in text with a hangul neighbour; prev/next
 *  are the texts adjacent to its ends (empty = no direct neighbour). Word
 *  widens spaces next to kana/Han the same way (probe 2026-10-13), but the
 *  half-width space span lifts Japanese line boxes, so that stays unported. */
export function hangulSpaceOffsets(text: string, prev = '', next = ''): number[] {
  const out: number[] = []
  let prevCp = prev ? lastCodePoint(prev) : -1
  for (let i = 0; i < text.length;) {
    const cp = text.codePointAt(i)!
    if (cp === 0x20) {
      const nextCp =
        i + 1 < text.length ? text.codePointAt(i + 1)! : next ? next.codePointAt(0)! : -1
      if (isHangul(prevCp) || isHangul(nextCp)) out.push(i)
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
/** isCjkFontName widened to any name spelled with Han, hangul or kana code points */
export function isCjkNamedFont(fontFamily: string): boolean {
  return CJK_NAMED_RE.test(fontFamily.normalize('NFKC')) || isCjkFontName(fontFamily)
}

export function isCjkFontName(fontFamily: string): boolean {
  const f = fontFamily.toLowerCase()
  const nfkc = f.normalize('NFKC')
  // hangul-lettered vendor names are Korean faces even when otherwise unknown
  return (
    CJK_FONT_NAME_RE.test(f) ||
    CJK_FONT_NAME_RE.test(nfkc) ||
    KO_FONT_RE.test(nfkc) ||
    /[가-힣ᄀ-ᇿ㄰-㆏]/.test(nfkc)
  )
}

const CJK_FONT_NAME_RE =
  /\u5b8b|\u9ed1|\u6977|\u4eff|\u660e|\u96c5\u9ed1|\u7b49\u7ebf|simsun|simhei|simkai|simfang|kaiti|fangsong|songti|stsong|yahei|dengxian|mingliu|jhenghei|biaukai|dfkai|kaiu|mincho|ms (ui )?p?gothic|yu gothic|\u30b4\u30b7\u30c3\u30af|meiryo|\u30e1\u30a4\u30ea\u30aa|hiragino|\u30d2\u30e9\u30ae\u30ce|osaka|kozuka|biz ud|\u6e38|\u6559\u79d1\u66f8\u4f53|\u884c\u66f8\u4f53|\u5275\u82f1|arial unicode|(noto|source han) (sans|serif)( cjk)? ?(sc|cn|jp|tc|tw|hk|kr)\b/i

function cjkLineHFactor(fontFamily: string): number {
  const f = fontFamily.toLowerCase()
  if (f.includes('pmingliu') || f.includes('mingliu')) return 1.0
  const declared = cjkDeclaredLineFactor(fontFamily)
  if (declared !== null) return declared
  if (isCjkFontName(fontFamily)) return lineHeightFactor(fontFamily)
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

  const pushLine = (h: number) => {
    lines.push({ naturalLineH: h, text: curText })
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
        curLineH = Math.max(curLineH, lineH)
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
          curLineH = Math.max(curLineH, lineH)
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
  if (curLineH > 0 || lines.length === 0) {
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

/**
 * Widest unbreakable token width (px) across runs: whitespace splits words;
 * CJK breaks per character unless Hangul is present (keep-all, same rule as
 * simulateLines). Min-content input for table autofit; scanning is capped so
 * huge cells stay cheap.
 */
export function maxWordWidthPx(
  runs: Array<{
    text: string
    fontFamily?: string
    sizeHalfPoints?: number
    bold?: boolean
    italic?: boolean
  }>,
  metrics: FontMetricsProvider = defaultMetrics,
  maxWords = 200,
): number {
  const keepAll = runs.some((r) => textHasHangul(r.text))
  let max = 0
  let cur = 0 // width of the current token's completed segments (may span runs)
  let words = 0
  let scanned = 0
  let breakAfter: '' | '-' | '/' = '' // pending break opportunity behind a '-' or '/'
  const endWord = () => {
    if (cur > max) max = cur
    if (cur > 0) words++
    cur = 0
  }
  for (const run of runs) {
    if (!run.text) continue
    const style: RunStyle = {
      fontFamily: run.fontFamily ?? DEFAULT_FONT_FAMILY,
      fontSizePx: (run.sizeHalfPoints ? run.sizeHalfPoints / 2 : DEFAULT_FONT_SIZE_PT) * (96 / 72),
      bold: !!run.bold,
      italic: !!run.italic,
    }
    let buf = ''
    const flushSeg = () => {
      if (!buf) return
      cur += metrics.measure(buf, style)
      buf = ''
    }
    for (const ch of run.text) {
      if (words >= maxWords || scanned >= maxWords * 60) {
        flushSeg()
        endWord()
        return max
      }
      scanned++
      // Word breaks after '-' even before a digit ("10-|20", "DDR5-|6400");
      // '/' keeps UAX14's no-break before a digit ("1/2")
      if (breakAfter) {
        const digit = ch >= '0' && ch <= '9'
        if (breakAfter === '-' || !digit) {
          flushSeg()
          endWord()
        }
        breakAfter = ''
      }
      if (ch === ' ' || ch === '\t' || ch === '\n') {
        flushSeg()
        endWord()
        continue
      }
      if (ch === '-' || ch === '/') breakAfter = ch
      if (!keepAll && isCjk(ch.codePointAt(0) ?? 0)) {
        flushSeg()
        endWord()
        cur = metrics.measure(ch, style)
        endWord()
        continue
      }
      buf += ch
    }
    flushSeg()
  }
  endWord()
  return max
}

/**
 * Unwrapped width (px) of the runs laid on one line; hard breaks ('\n') start
 * a new line and the widest wins. Max-content input for table autofit.
 */
export function maxLineWidthPx(
  runs: Array<{
    text: string
    fontFamily?: string
    sizeHalfPoints?: number
    bold?: boolean
    italic?: boolean
  }>,
  metrics: FontMetricsProvider = defaultMetrics,
): number {
  let max = 0
  let cur = 0
  for (const run of runs) {
    if (!run.text) continue
    const style: RunStyle = {
      fontFamily: run.fontFamily ?? DEFAULT_FONT_FAMILY,
      fontSizePx: (run.sizeHalfPoints ? run.sizeHalfPoints / 2 : DEFAULT_FONT_SIZE_PT) * (96 / 72),
      bold: !!run.bold,
      italic: !!run.italic,
    }
    const lines = run.text.split('\n')
    lines.forEach((line, i) => {
      if (i > 0) {
        max = Math.max(max, cur)
        cur = 0
      }
      if (line) cur += metrics.measure(line, style)
    })
  }
  return Math.max(max, cur)
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

const DEFAULT_FONT_FAMILY = 'Calibri'
const DEFAULT_FONT_SIZE_PT = 12

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
              lineHeightFactor(r.fontFamily ?? defaultFontFamily),
          ),
        )
      : defaultFontSizePx * lineHeightFactor(defaultFontFamily)

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

/** reserved height for the footnote separator line (px) */
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
export function hfHeaderGeom(set: { marginTop: number; headerDist?: number }): {
  marginTopPx: number
  headerDistPx: number
} {
  const toPx = (twips: number) => (twips / 1440) * 96
  return { marginTopPx: toPx(set.marginTop), headerDistPx: toPx(set.headerDist ?? 720) }
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
  geom?: { marginTopPx: number; headerDistPx: number; marginLeftPx?: number },
): number {
  const inlineImages = (images ?? []).filter((im) => !im.floating && im.heightPx)
  const imagesHeight =
    inlineImages.length > 0 ? Math.max(...inlineImages.map((im) => im.heightPx!)) : 0
  let anchoredPx = 0
  if (geom) {
    for (const im of images ?? []) {
      if (!im.floating || im.behind || !im.heightPx) continue
      if (!im.wrap || im.wrap === 'none') continue
      // a side-wrapped drawing shapes the lines beside it and pushes nothing
      // (L260 CASE07); only one that leaves no room for a word on either side
      // blocks like wrapTopAndBottom. The DOM probe measures the real flow;
      // without the drawing's width the old full-band floor stays
      if (im.wrap !== 'topBottom') {
        const room = hfSideRoomPx(im, contentWidthPx, geom.marginLeftPx)
        if (room !== null && room >= HF_SIDE_WRAP_SLIVER_PX) continue
      }
      if (im.posYPx == null || !im.posVRel || im.posVRel === 'bottomMargin') continue
      // top edge in page coordinates (Word keeps the drawing on the page: a
      // negative paragraph offset clamps at the page top); the anchor paragraph
      // sits at the header strip top
      const top =
        im.posVRel === 'page' || im.posVRel === 'topMargin'
          ? im.posYPx
          : im.posVRel === 'margin'
            ? geom.marginTopPx + im.posYPx
            : geom.headerDistPx + im.posYPx
      // effectiveTopPx adds headerDist back: dist + this = the image bottom
      anchoredPx = Math.max(anchoredPx, Math.max(0, top) + im.heightPx - geom.headerDistPx)
    }
  }
  if (!part) return Math.max(imagesHeight, anchoredPx)
  type HfPara = NonNullable<typeof part.paras>[number]
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
  let height = 0
  const flow = paras.filter((p) => !p.boxAnchored)
  const padKey = (p: HfPara | undefined) =>
    p && !p.cells?.length && p.borderPad ? JSON.stringify(p.borderPad) : ''
  for (const [i, p] of flow.entries()) {
    if (p.cells?.length) {
      // table row: the tallest cell's paragraph stack sets the row height;
      // a cell-run image (logo) grows its line box like the display layer
      const declared = (p.row?.heightTwips ?? 0) / 15
      if (declared > 0 && p.row?.heightRule === 'exact') {
        height += declared
        continue
      }
      height += Math.max(
        declared,
        ...p.cells.map((c) =>
          (c.paras.length > 0 ? c.paras : [[]]).reduce((s, runs, k) => {
            const imgH = Math.max(0, ...runs.map((r) => r.image?.heightPx ?? 0))
            return s + Math.max(lineH(runs, c.paraProps?.[k]), imgH)
          }, 0),
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
    height += Math.max(lineH(runs, p, Math.max(1, contentWidthPx - indentPx)), imgH)
    // same-border neighbours form one Word border group: padded at its edges only
    const key = padKey(p)
    if (key && key !== padKey(flow[i - 1])) height += ((p.borderPad?.t ?? 0) * 96) / 72
    if (key && key !== padKey(flow[i + 1])) height += ((p.borderPad?.b ?? 0) * 96) / 72
  }
  return Math.max(height + imagesHeight, anchoredPx)
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

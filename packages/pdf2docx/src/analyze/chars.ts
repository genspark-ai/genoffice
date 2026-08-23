/**
 * chars → IR lines: the composed per-line pipeline (cluster marks → visual
 * lines → RTL logical reorder where needed → words → spans). Shared by the
 * page-level flow and by table-cell content assembly.
 */
import { median, overlapRatio, rectCenterX, rectCenterY } from '../geometry'
import type { Line } from '../ir'
import type { ExtractedPage } from '../extract'
import { clusterCombiningMarks, groupIntoLines } from './lines'
import { lineHasRtl, normalizeArabicForms, reorderVisualToLogical } from './rtl'
import { buildSpans } from './spans'
import { groupIntoWords } from './words'

// ── effective font-size calibration (P8) ──
/** metric/declared ratios inside this band mark a font whose declared Tf size
 * is bogus; outside it the declared size is kept (below = degenerate loose
 * boxes, PDFium metrics unusable; above = normal font, metrics ≈ 1.0–1.3×) */
const CALIBRATE_MIN_RATIO = 0.3
const CALIBRATE_MAX_RATIO = 0.9

/**
 * Replace declared font sizes with metric-derived effective sizes where the
 * two wildly disagree. AI-generated PDFs routinely declare a Tf size the
 * embedded font never fills — glyphs cover ~0.65 em (loose/advance box ≈
 * 0.65 × declared) instead of the normal ≥ 1.0. Honoring the declared size
 * rebuilds the text ~1.5× larger than the source rendering: lines clip
 * against their measured exact pitch, wrap early, and any font-based line
 * floor inflates page counts. The loose (advance) box height IS the font's
 * real ascent+descent extent at rendered scale, so a per-font median of it
 * is the size the source actually shows.
 *
 * Calibration is grouped by (family, declared size) — per-char values would
 * split same-styled runs — and mutates the chars in place.
 */
export function calibrateFontSizes(chars: ExtractedPage['chars']): void {
  const groups = new Map<string, number[]>()
  const keyOf = (c: ExtractedPage['chars'][number]): string =>
    `${c.fontFamily}|${c.fontSize.toFixed(1)}`
  for (const c of chars) {
    if (c.fontSize <= 0 || c.text.trim() === '') continue
    const looseH = c.looseBox.y1 - c.looseBox.y0
    if (looseH <= 0) continue
    let group = groups.get(keyOf(c))
    if (!group) groups.set(keyOf(c), (group = []))
    group.push(looseH)
  }
  const corrected = new Map<string, number>()
  for (const [key, heights] of groups) {
    const declared = Number(key.slice(key.lastIndexOf('|') + 1))
    const effective = median(heights)
    const ratio = effective / declared
    if (ratio >= CALIBRATE_MIN_RATIO && ratio <= CALIBRATE_MAX_RATIO) {
      corrected.set(key, effective)
    }
  }
  if (corrected.size === 0) return
  for (const c of chars) {
    const size = corrected.get(keyOf(c))
    if (size !== undefined) c.fontSize = size
  }
}

// ── double-drawn char dedup (P11 B) ──
/** twin glyphs whose centers sit within this share of the font size apart */
const DOUBLE_DRAW_MAX_DIST_RATIO = 0.3
/** …and whose sizes agree within this band are one double-drawn char */
const DOUBLE_DRAW_MIN_SIZE_RATIO = 0.8
const DOUBLE_DRAW_MAX_SIZE_RATIO = 1.25
/** …and whose ink boxes truly overprint — narrow letters set tight ('ll',
 * 'tt' at ~0.22 em advance) pass the center gate but their ink never overlaps */
const DOUBLE_DRAW_MIN_BOX_OVERLAP = 0.5
/** box-corner agreement (pt) treated as "the same glyph box": PDFium copies
 * one glyph's box verbatim onto every char a ligature expands to, so the
 * copies are bit-identical; genuine re-strokes sit ≥ ~0.5pt apart */
const LIGATURE_BOX_EPSILON = 0.05

/** adjacent chars carrying the exact same glyph box are one ligature glyph
 * ('tt', 'ff') expanded through ToUnicode — not a double draw */
function isLigatureExpansionPair(
  ia: number,
  ib: number,
  a: ExtractedPage['chars'][number],
  b: ExtractedPage['chars'][number],
): boolean {
  if (Math.abs(ia - ib) !== 1) return false
  return (
    Math.abs(a.box.x0 - b.box.x0) <= LIGATURE_BOX_EPSILON &&
    Math.abs(a.box.x1 - b.box.x1) <= LIGATURE_BOX_EPSILON &&
    Math.abs(a.box.y0 - b.box.y0) <= LIGATURE_BOX_EPSILON &&
    Math.abs(a.box.y1 - b.box.y1) <= LIGATURE_BOX_EPSILON
  )
}

/**
 * Drop the ghost twin of double-drawn chars (P11 B): decks re-stroke display
 * text for fake bold / soft shadows, leaving two near-coincident glyphs of
 * the same character. Rebuilt as flow text the copies can't overprint
 * exactly, so the twin shows as a visible ghost and doubles the token count.
 * The later-drawn glyph wins (it's the one visible in the source) unless the
 * earlier one is clearly bigger. Normal tight kerning is safe: two same-size
 * glyphs advance ≥ ~0.5 em apart, well over the 0.3 em gate. Ligature glyphs
 * expanded to repeated chars ('tt', 'ff') share one verbatim box on adjacent
 * indices and are kept. Mutates in place.
 */
export function dedupeDoubleDrawnChars(chars: ExtractedPage['chars']): void {
  const byCode = new Map<number, number[]>()
  for (let i = 0; i < chars.length; i++) {
    const c = chars[i]!
    if (c.isGenerated || c.text.trim() === '') continue
    let list = byCode.get(c.code)
    if (!list) byCode.set(c.code, (list = []))
    list.push(i)
  }
  const dropped = new Set<number>()
  for (const list of byCode.values()) {
    for (let a = 0; a < list.length; a++) {
      const ia = list[a]!
      if (dropped.has(ia)) continue
      const ca = chars[ia]!
      const sizeA = Math.max(ca.fontSize, 1)
      for (let b = a + 1; b < list.length; b++) {
        const ib = list[b]!
        if (dropped.has(ib)) continue
        const cb = chars[ib]!
        const sizeB = Math.max(cb.fontSize, 1)
        const ratio = sizeB / sizeA
        if (ratio < DOUBLE_DRAW_MIN_SIZE_RATIO || ratio > DOUBLE_DRAW_MAX_SIZE_RATIO) continue
        const dx = rectCenterX(ca.box) - rectCenterX(cb.box)
        const dy = rectCenterY(ca.box) - rectCenterY(cb.box)
        if (Math.hypot(dx, dy) >= DOUBLE_DRAW_MAX_DIST_RATIO * Math.max(sizeA, sizeB)) continue
        if (overlapRatio(ca.box, cb.box) < DOUBLE_DRAW_MIN_BOX_OVERLAP) continue
        if (isLigatureExpansionPair(ia, ib, ca, cb)) continue
        dropped.add(sizeA > sizeB * 1.01 ? ib : ia)
        if (dropped.has(ia)) break
      }
    }
  }
  if (dropped.size === 0) return
  let w = 0
  for (let i = 0; i < chars.length; i++) {
    if (!dropped.has(i)) chars[w++] = chars[i]!
  }
  chars.length = w
}

// ── regional font-name artifacts (P16 D) ──

/** high-frequency Han pairs whose simplified and traditional forms differ */
const SC_INDICATORS = '为对说这读们还进关问间门业务经条现发见车书长风华万与从众体单声义写办产乐让论'
const TC_INDICATORS = '為對說這讀們還進關問間門業務經條現發見車書長風華萬與從眾體單聲義寫辦產樂讓論'
/** rewrite only with this much simplified-form evidence on the page */
const REGIONAL_MIN_SC_CHARS = 10

/** cross-platform families by shape class for TC-artifact rewrites */
const TC_ARTIFACT_FAMILIES: Array<[RegExp, string]> = [
  [/hei|pingfang/i, 'SimHei'],
  [/kai/i, 'KaiTi'],
  [/song|sung|ming/i, 'SimSun'],
]

/**
 * A Traditional-Chinese regional family ("Yuanti TC", "Songti TC") carrying
 * strictly Simplified text is an export-chain substitution artifact: the
 * authoring face (mainland document fonts) was missing at print time and an
 * Apple TC display font got stamped into the metadata. Word/LibreOffice then
 * resolve the accidental name and the whole document changes face. Rewrite
 * such families to the cross-platform mainland family of their shape class;
 * shapeless ones (Yuanti — the common gov-paper chain) become FangSong.
 */
export function normalizeRegionalFontArtifacts(chars: ExtractedPage['chars']): void {
  let sc = 0
  let tc = 0
  for (const c of chars) {
    const ch = String.fromCodePoint(c.code)
    if (SC_INDICATORS.includes(ch)) sc++
    else if (TC_INDICATORS.includes(ch)) tc++
  }
  if (sc < REGIONAL_MIN_SC_CHARS || tc > 0) return
  for (const c of chars) {
    if (!/\sTC$/.test(c.fontFamily)) continue
    const mapped = TC_ARTIFACT_FAMILIES.find(([re]) => re.test(c.fontFamily))
    c.fontFamily = mapped ? mapped[1] : 'FangSong'
  }
}

/** chars → IR lines (cluster marks → visual lines → words → spans) */
export function analyzeChars(chars: ExtractedPage['chars']): Line[] {
  const clustered = clusterCombiningMarks(normalizeArabicForms(chars))
  const lines: Line[] = []
  for (const raw of groupIntoLines(clustered)) {
    const rtl = lineHasRtl(raw.chars)
    const ordered = rtl ? reorderVisualToLogical(raw.chars) : raw.chars
    const spans = buildSpans(groupIntoWords(ordered, { inferSpaces: !rtl }))
    if (spans.length === 0) continue
    lines.push({
      spans,
      box: raw.box,
      baseline: raw.baseline,
      endsWithHyphen: raw.endsWithHyphen,
    })
  }
  return lines
}

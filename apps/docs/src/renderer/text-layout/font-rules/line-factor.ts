/**
 * Single-spacing line factor of a declared name: the resolver's face metrics
 * arbitrated against the probed table (line-factors.json), embedded boxes, the
 * East Asian and Korean routes (INVENTORY-fonts L1-L22).
 */
import type { ResolvedFace } from '@genoffice/font-metrics/resolve'
import {
  CJK_LETTERED_RE as CJK_NAMED_RE,
  isEastAsianFaceName as isCjkFontName,
  isJapaneseFaceName as isJapaneseFontName,
  KO_FACE_RE as KO_FONT_RE,
} from '../face-info'
import { NOTO_HAN_REGIONAL_RE } from '../../font-name-patterns'
import {
  recordResolveDiff,
  resolveDeclaredFace,
  setResolverEmbeddedFaces,
} from '../../font-resolve'
import {
  LINE_FACTOR_PINS,
  LINE_FACTOR_RULES,
  type LineFactorStage,
} from '../../line-factors.generated'
import { arabicAltNameFace, arabicCodePageFace } from './css-chains'
import {
  fontTableHint,
  installedEaCharsetFace,
  isKoreanLayoutFace,
  krCharsetMalgunSubstituted,
  krCharsetRoute,
  krNameLineFactor,
  missingLatinAltName,
  missingLatinSubstitute,
} from './font-table'

/** line boxes of the open document's embedded faces, keyed like docFontTable */
let embeddedLineBoxes = new Map<string, { factor: number; descent: number }>()

/** the embed Word lays out with: none when Word ships the family itself (resolve.ts direct) */
export function embeddedBox(fontFamily: string): { factor: number; descent: number } | undefined {
  const box = embeddedLineBoxes.get(fontFamily.normalize('NFKC').toLowerCase().trim())
  if (!box) return undefined
  return resolveDeclaredFace(fontFamily)?.source === 'embedded' ? box : undefined
}

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

/** the pin table alone (shadow audits compare the resolver against it) */
export function tableLineHeightFactor(fontFamily: string): number {
  return knownLineFactor(fontFamily)?.factor ?? unknownLatinFactor(fontFamily).factor
}

/**
 * Where a line factor came from: the resolver's face metrics, a probed pin the
 * file does not give, the document's embedded box, a name-class route of the
 * table, or the Calibri / Cambria fallback
 */
export type FactorSource = 'resolver' | 'pin' | 'embedded' | 'table' | 'fallback'
export interface LineFactorDecision {
  factor: number
  source: FactorSource
}

/**
 * Pins U resolver (line F wave 3): a face Word lays out for real or reaches by
 * a behavioural rule paces at the resolver's metrics; the table only holds
 * values Word confirmed that the file does not give (LINE_FACTOR_PINS), the
 * embedded boxes and the East Asian / Korean routes keep their own paths
 */
export function lineFactorDecision(fontFamily: string): LineFactorDecision {
  const known = knownLineFactor(fontFamily)
  const face = resolveDeclaredFace(fontFamily)
  // a style word or MT tag dropped from the name is the base family's own
  // answer, table pins included (probe 2026-10-22: Arial Bold -> ArialMT 1.15)
  if (face?.rules[0] === 'styleSuffix' && face.path.length === 1)
    return lineFactorDecision(face.family)
  if (face && resolverAnswers(fontFamily, face, known?.factor))
    return { factor: face.lineFactor, source: 'resolver' }
  const legacy = known ?? unknownLatinFactor(fontFamily)
  if (face?.lineFactor !== legacy.factor) recordResolveDiff(fontFamily, legacy.factor, face)
  return legacy
}

export function lineHeightFactor(fontFamily: string): number {
  return lineFactorDecision(fontFamily).factor
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
  // a pin fixes a missing face's substitute; an installed copy lays out for real
  const real = face.source === 'installed' && !face.substituted && face.metrics.eastAsian
  if (known !== undefined && !real && LINE_FACTOR_PINS.some((pin) => pin.test(key))) return false
  if (face.metrics.eastAsian) {
    // a face Word ships paces at its own hhea x 1.3 whatever its name; an
    // installed or embedded one only under a CJK name (the EA code-page flag
    // on Latin macOS faces is a data quirk); charset-81 entries are unprobed
    if (!face.substituted) {
      const shipped = face.source !== 'installed' && face.source !== 'system'
      return (
        (shipped || isCjkFontName(name) || installedEaCharsetFace(name)) && !krCharsetRoute(name)
      )
    }
    // a substitute reached through altName or a behavioural rule is drawn
    // wholesale, Latin and CJK alike (probe 2026-10-22: altName MS Mincho,
    // Gulim, Yu Gothic, SimHei, YaHei and SimSun under Latin and CJK names)
    return true
  }
  if (face.source === 'embedded') return false
  if (face.substituted) {
    // a CJK-lettered name routed by its fontTable altName paces as the alias face
    if (CJK_NAMED_RE.test(key) && altNameRouted(key, face)) return true
    return face.exception
  }
  // a Latin face Word lays out for real follows the resolver; a CJK-named
  // Latin face keeps the table routes
  return !isCjkFontName(name)
}

function altNameRouted(key: string, face: ResolvedFace): boolean {
  const alt = fontTableHint(key)?.altName
  return !!alt && alt.normalize('NFKC').toLowerCase().trim() === face.path[1]
}

/** Word's Calibri substitute for an unprobed Latin name with a fontTable entry, else Cambria;
 *  an altName pointing elsewhere is drawn wholesale and paces as that face */
function unknownLatinFactor(fontFamily: string): LineFactorDecision {
  const alt = missingLatinAltName(fontFamily)
  if (alt) return lineFactorDecision(alt)
  const fallback = missingLatinSubstitute(fontFamily) === 'calibri' ? 'Calibri' : 'Cambria'
  return { factor: resolveDeclaredFace(fallback)?.lineFactor ?? 1, source: 'fallback' }
}

/** first rule of `stage` matching the lower-cased name */
function ruleFactor(f: string, stage: LineFactorStage): LineFactorDecision | undefined {
  for (const rule of LINE_FACTOR_RULES) {
    if (rule.stage !== stage || !rule.match.test(f) || rule.exclude?.test(f)) continue
    return { factor: rule.factor, source: rule.pin ? 'pin' : 'table' }
  }
  return undefined
}

const route = (factor: number): LineFactorDecision => ({ factor, source: 'table' })

/**
 * Factor of a probed/known face (packages/font-metrics/measurements/line-factors.json,
 * one source per value); undefined for a Latin name no rule covers
 */
function knownLineFactor(fontFamily: string): LineFactorDecision | undefined {
  const f = fontFamily.toLowerCase()
  // an embedded face renders real: Word lays it with its own hhea total
  // (probe 2026-09-23, typo/win groups and USE_TYPO_METRICS ignored)
  const embedded = embeddedBox(fontFamily)
  // Word meters an embedded East Asian face by its EA rule, not its own box
  // (embedded Malgun Gothic: 19.05 pt at 11 pt = 1.7305, hhea total is 1.3301)
  if (embedded && !isCjkFontName(fontFamily)) return { factor: embedded.factor, source: 'embedded' }
  const first = ruleFactor(f, 'first')
  if (first !== undefined) return first
  // MS Gothic / SimSun class fallbacks for the declares resolverAnswers leaves
  // to the table (altName-routed, Latin slot of a split substitution)
  if (
    /mincho|\u660e\u671d|\u30b4\u30b7\u30c3\u30af|ms (ui )?p?gothic|hiragino|osaka|kozuka|\u5c0f\u585a/.test(
      f.normalize('NFKC'),
    )
  )
    return route(simsunLineFactor())
  if (/^(noto|source han) (sans|serif)( cjk)? ?(sc|cn|tc|tw|hk)\b/.test(f))
    return route(simsunLineFactor())
  const arabicCloud = ruleFactor(f, 'arabic')
  if (arabicCloud !== undefined) return arabicCloud
  const arabicAlt = arabicAltNameFace(fontFamily) ?? arabicCodePageFace(fontFamily)
  if (arabicAlt) return lineFactorDecision(arabicAlt)
  const latin = ruleFactor(f, 'latin')
  if (latin !== undefined) return latin
  // a Hangul-charset fontTable entry substitutes Malgun Gothic (Batang when
  // w:family is roman) for the Latin runs too (probe 2026-10-20)
  const krRoute = krCharsetRoute(f)
  if (krRoute) return route(krRoute === 'malgun' ? malgunLineFactor() : batangLineFactor())
  // missing hangul-named vendor faces substitute into the Batang/SimSun class
  // for KR and Latin runs alike (probe 2026-08-23: \ud734\uba3c\uace0\ub515 -> Batang/SimSun 1.30)
  if (/[\uac00-\ud7a3]/.test(f)) return route(batangLineFactor())
  // missing ideograph-named faces substitute SimSun (probe 2026-08-22/23)
  if (/[\u4e00-\u9fff]/.test(f)) return route(simsunLineFactor())
  // unknown Western names with a fontTable entry substitute Calibri unless
  // their PANOSE says serif (probe 2026-10-22); without an entry Word for Mac
  // substitutes Cambria (probe 2026-08-23: fabricated names and HCI Poppy
  // both render as Cambria) -- unknownLatinFactor
  return ruleFactor(f, 'last')
}

/**
 * Declared names Word lays out as a sibling face (probe 2026-11-09: Times is
 * drawn in Times New Roman; Courier carries Courier New's 1.1333 unprobed), so
 * descent and xAvgCharWidth follow the sibling too
 */
const LEGACY_FACE_ALIAS: Array<[RegExp, string]> = [
  [/^times\b(?! new)/, 'Times New Roman'],
  [/^courier\b(?! new)/, 'Courier New'],
]

export function legacyFaceAlias(font: string): string | undefined {
  const key = font.normalize('NFKC').toLowerCase().trim()
  return LEGACY_FACE_ALIAS.find(([re]) => re.test(key))?.[1]
}

/** Word's layout face for a declared Latin name, through the table's sibling aliases */
export function latinFace(font: string): ResolvedFace | null {
  if (!font.trim()) return null
  return resolveDeclaredFace(legacyFaceAlias(font) ?? font)
}

/** Descent (em) Word lays a Latin text face with: its layout face's hhea
 *  descent; East Asian faces keep 1/4 em (their box is the x1.3 rule) */
export function latinDescent(font: string): number {
  const embedded = embeddedBox(font)?.descent
  if (embedded !== undefined) return embedded
  const face = latinFace(font)
  return face && !face.metrics.eastAsian ? face.metrics.descent : 0.25
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
  // an installed copy lays out for real at its own EA rule (corpus 2026-11-01:
  // Noto Sans JP 12pt on an 18pt grid takes two pitches in Word, 1.448 x 1.3)
  const face = resolveDeclaredFace(fontFamily)
  if (face && face.source === 'installed' && !face.substituted && face.metrics.eastAsian)
    return lineHeightFactor(fontFamily)
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

/** --doc-line-factor-kr source: the document's East Asian face when Korean, else the Batang-class default */
export function krLineFactor(fontFamily: string | undefined): number {
  if (!fontFamily) return batangLineFactor()
  const krName = krNameLineFactor(fontFamily)
  if (krName !== null) return krName
  if (isKoreanLayoutFace(fontFamily)) return lineHeightFactor(fontFamily)
  if (krCharsetMalgunSubstituted(fontFamily)) return malgunLineFactor()
  // hangul a Japanese face cannot draw goes to Malgun Gothic (Word probe 2026-09-30),
  // except under the Yu faces, which hand it to Batang (Word probe 2026-11-03)
  if (!isJapaneseFontName(fontFamily)) return batangLineFactor()
  return /yu (gothic|mincho)|yugoth|\u6e38/i.test(fontFamily.normalize('NFKC'))
    ? batangLineFactor()
    : malgunLineFactor()
}

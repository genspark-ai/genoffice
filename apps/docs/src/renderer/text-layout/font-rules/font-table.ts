/**
 * The open document's word/fontTable.xml hints and the routes Word takes from
 * them for a missing face: charset 81 / 80, altName, w:family roman, the Arabic
 * code page (INVENTORY-fonts R7-R17, renderer side; the resolver in
 * packages/font-metrics runs the same rules on its own path).
 */
import {
  CJK_LETTERED_RE as CJK_NAMED_RE,
  isEastAsianFaceName as isCjkFontName,
  isKoreanFaceName as isKoreanFontName,
  KO_FACE_RE as KO_FONT_RE,
} from '../face-info'
import {
  batangAltNamed,
  faceKey,
  latinFallback,
  type ResolvedFace,
  type SubstituteRule,
} from '@genoffice/font-metrics/resolve'
import { resolveDeclaredFace, setResolverFontTable } from '../../font-resolve'
import { declaredFaceMissing, isBundledFont } from './availability'
import { batangLineFactor, malgunLineFactor } from './line-factor'

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

/** hints of the open document's fontTable entry, by NFKC-lowercased name */
export function fontTableHint(key: string): FontTableHints | undefined {
  return docFontTable.get(key)
}

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

const headOf = (font: string) => font.split(',')[0].replace(/['"]/g, '').normalize('NFKC').trim()

/** the resolver's first substitution step of a face Chromium does not draw, else null */
function firstStep(
  head: string,
): { rule: SubstituteRule; target: string; face: ResolvedFace } | null {
  if (!declaredFaceMissing(head)) return null
  const face = resolveDeclaredFace(head)
  const rule = face?.rules[0]
  return face && rule && face.path[1] ? { rule, target: face.path[1], face } : null
}

/**
 * Missing face whose fontTable charset is Hangul (81): Word for Mac substitutes
 * Malgun Gothic wholesale, hangul and Latin alike, unless w:family says roman
 * (then Batang). PANOSE and the name's script do not matter (probe 2026-10-20:
 * HCR Batang serif PANOSE + modern, a Latin-named entry without PANOSE, and
 * sans PANOSE + roman went Malgun / Malgun / Batang). A face Word ships or the
 * machine installs lays out for real, not routed (probes 2026-11-02 Gulim,
 * 2026-11-03 Pretendard): the resolver finds it before the charset step.
 */
export function krCharsetRoute(font: string): 'malgun' | 'batang' | null {
  const step = firstStep(headOf(font))
  if (step?.rule !== 'charset81') return null
  return step.target === 'batang' ? 'batang' : 'malgun'
}

/**
 * Latin-named face filed under the Hangul charset (81) that is installed here:
 * Word lays it out for real at hhea (ascent + descent) x 1.3 like any East
 * Asian face (corpus 2026-11-03: Pretendard 1.1934 x 1.3 = 1.551, 11 pt
 * line=276 at 20.0 pt, 10 pt table rows 16 pt) instead of the Malgun route.
 */
export function installedEaCharsetFace(font: string): boolean {
  const head = font.split(',')[0].replace(/['"]/g, '').normalize('NFKC').trim()
  if (docFontTable.get(head.toLowerCase())?.charset !== '81' || isBundledFont(head)) return false
  const face = resolveDeclaredFace(head)
  return !!face && !face.substituted && face.source === 'installed' && face.metrics.eastAsian
}

/** a declared name that sets the East Asian line factor: CJK-named, or an installed charset-81 face */
export function isEaLayoutFace(font: string): boolean {
  return isCjkFontName(font) || installedEaCharsetFace(font)
}

/** Korean face by name or by an installed charset-81 fontTable entry */
export function isKoreanLayoutFace(font: string): boolean {
  return isKoreanFontName(font) || installedEaCharsetFace(font)
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
  const key = headOf(font).toLowerCase()
  return batangAltNamed(key, docFontTable.get(key))
}

/** PANOSE serif-style byte (2nd): 0B-0F sans, 02-0A serif, 00/01/none unknown */
export function panoseSerifHint(font: string): 'serif' | 'sans' | undefined {
  const panose = docFontTable.get(font.normalize('NFKC').toLowerCase())?.panose
  const hex = panose?.replace(/[^0-9a-f]/gi, '')
  if (!hex || hex.length < 4 || /^0+$/.test(hex)) return undefined
  const serifStyle = parseInt(hex.slice(2, 4), 16)
  if (serifStyle >= 0x0b && serifStyle <= 0x0f) return 'sans'
  if (serifStyle >= 0x02 && serifStyle <= 0x0a) return 'serif'
  return undefined
}

/**
 * Missing Latin-named face that has a fontTable entry (no altName): Word for
 * Mac substitutes Calibri wholesale, Cyrillic and Latin alike, unless w:family
 * says roman, which sends it to Cambria. A name without any fontTable entry
 * is Cambria too (probes 2026-10-22 / 2026-11-08: PANOSE, charset, sig code
 * pages, run language and the docDefaults/theme fonts do not matter).
 */
export function missingLatinSubstitute(font: string): 'calibri' | 'cambria' | null {
  const head = headOf(font)
  if (CJK_NAMED_RE.test(head)) return null
  const hints = docFontTable.get(head.toLowerCase())
  if (!hints || hints.charset === '81' || missingLatinAltName(font)) return null
  if (!declaredFaceMissing(head)) return null
  return faceKey(latinFallback(hints)) === 'cambria' ? 'cambria' : 'calibri'
}

/** fontTable altName of a missing Latin-named face (probe 2026-10-22: Word draws the alias wholesale) */
export function missingLatinAltName(font: string): string | null {
  const head = headOf(font)
  if (CJK_NAMED_RE.test(head)) return null
  const hints = docFontTable.get(head.toLowerCase())
  const alt = hints?.altName?.trim()
  if (!alt || hints?.charset === '81' || faceKey(alt) === head.toLowerCase()) return null
  return declaredFaceMissing(head) ? alt : null
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

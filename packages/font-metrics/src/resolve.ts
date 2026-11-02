/**
 * Word's face resolution, as pure logic: which face Word for Mac lays a
 * declared font name out with, and the single-spacing line factor of that
 * face. Metrics come from an injected context (embedded faces, the installed
 * font index, the generated word-faces.json) so the same algorithm runs in the
 * renderer, in node scripts and in jsdom tests.
 *
 * Word's rules (word-truths): an embedded, installed or Word-private face is
 * metered by its hhea total; an East Asian face by (ascent + descent) x 1.3;
 * a missing name follows the fontTable altName, then charset/script/PANOSE
 * substitution, ending in Calibri (entry present) or Cambria (no entry).
 */

export interface FaceMetrics {
  /** (hhea ascender - descender + lineGap) / unitsPerEm */
  hheaTotal: number
  ascent: number
  descent: number
  /** OS/2 xAvgCharWidth / unitsPerEm */
  xAvg: number | null
  /** OS/2 ulCodePageRange1 bits 17-21 (JIS / GB / KR / Big5 / Johab) */
  eastAsian: boolean
}

export interface FaceHints {
  altName?: string
  panose?: string
  family?: string
  charset?: string
  csb0?: string
}

/** fontTable w:sig csb0 bit 6: the face covers the Arabic code page (1256) */
export function arabicCodePage(hints: FaceHints | undefined): boolean {
  const bits = parseInt(hints?.csb0 ?? '', 16)
  return Number.isFinite(bits) && (bits & 0x40) !== 0
}

/** system = the generated macOS snapshot standing in for the installed index; asset = macOS on-demand font Word lays out with */
export type FaceSource = 'embedded' | 'installed' | 'system' | 'dfonts' | 'cloud' | 'asset'

export interface StoredFace extends FaceMetrics {
  family: string
  source: FaceSource
}

export interface ResolvedFace {
  /** display name of the face Word lays out with */
  family: string
  source: FaceSource
  /** the declared name itself was not laid out */
  substituted: boolean
  /** the substitute came from a behavioural rule, not the Calibri/Cambria fallback */
  exception: boolean
  metrics: FaceMetrics
  lineFactor: number
  /** declared name, substitutes, final face */
  path: string[]
}

export interface ResolveContext {
  embedded(key: string): FaceMetrics | null
  installed(key: string): StoredFace | null
  wordFace(key: string): StoredFace | null
  hints(key: string): FaceHints | undefined
}

export const faceKey = (name: string): string =>
  name.split(',')[0].replace(/['"]/g, '').normalize('NFKC').toLowerCase().trim()

const round4 = (v: number) => Math.round(v * 1e4) / 1e4

/** East Asian faces: Word drops hhea lineGap and applies its 1.3 multiplier */
export function lineFactorOf(m: FaceMetrics): number {
  return round4(m.eastAsian ? (m.ascent + m.descent) * 1.3 : m.hheaTotal)
}

/** Word's own symbol-encoded faces win over the macOS copies of the same name */
const WORD_FIRST = new Set(['symbol', 'wingdings', 'wingdings 2', 'wingdings 3', 'webdings'])

const HAN_RE = /[\u2e80-\u9fff\uf900-\ufaff]/
const HANGUL_RE = /[\uac00-\ud7a3\u1100-\u11ff\u3130-\u318f]/
const KANA_RE = /[\u3040-\u30ff\uff66-\uff9f]/
/** Word for Mac never matches a macOS face by its localized family name */
const localizedName = (key: string) => HAN_RE.test(key) || HANGUL_RE.test(key) || KANA_RE.test(key)
const HAN_REGIONAL_RE = /^(noto|source han) (sans|serif)( cjk)? ?(sc|cn|tc|tw|hk|jp)\b/
const HAN_KR_RE = /^(noto|source han) (sans|serif)( cjk)? ?(kr|k)\b/
const HG_RE = new RegExp(
  '^hg[ps]?(\u6559\u79d1\u66f8\u4f53|\u660e\u671db|\u884c\u66f8\u4f53|\u6b63\u6977\u66f8\u4f53|\u5275\u82f1\u89d2\u30dd\u30c3\u30d7\u4f53|\u5275\u82f1\u30d7\u30ec\u30bc\u30f3\u30b9|\u30b4\u30b7\u30c3\u30afm)|^\u30d2\u30e9\u30ae\u30ce\u89d2\u30b4|biz ud',
)
const ARABIC_RE =
  /naskh|kufi|arabic|amiri|scheherazade|\b(b|xb|ir)[ -]?(mitra|nazanin|titr|lotus|zar|yekan|koodak|roya|badr|homa|traffic|compset)\b|irlotus/
const TAMIL_RE = /tamil|vijaya|inaimathi|^tau[- ]/
/** Word drops a style word or the MT tag and lays out the base family (probe 2026-10-22: Arial Bold, Courier New Regular, Arial MT) */
const STYLE_SUFFIX_RE = / (?:regular|bold|italic|bold italic|mt)$/

/** Behavioural substitutions Word applies to a missing name before the generic rules */
const SUBSTITUTIONS: Array<[RegExp, string]> = [
  [/nanum ?barun ?gothic|\ub098\ub214\ubc14\ub978\uace0\ub515/, 'Batang'],
  [/xiaobiaosong|\u5c0f\u6807\u5b8b/, 'SimSun'],
  [/(kaiti|\u6977\u4f53).*(gb2312|gbk)|dfkai|\u6a19\u6977/, 'Microsoft YaHei'],
  [HAN_REGIONAL_RE, 'SimSun'],
  [HG_RE, 'Yu Gothic'],
  [/dejavu/, 'Verdana'],
  // the LibreOffice metric clones (probe 2026-10-22): Word lays them out with the
  // face they clone, Liberation Mono alone falls through to Calibri
  [/^liberation sans\b/, 'Arial'],
  [/^liberation serif\b/, 'Times New Roman'],
  [/^carlito\b/, 'Calibri'],
  [/^caladea\b/, 'Cambria'],
  [/iranyekan/, 'Arial'],
  [ARABIC_RE, 'Times New Roman'],
  [TAMIL_RE, 'Vijaya'],
]

/** PANOSE serif-style byte: 0B-0F sans, 02-0A serif */
export function panoseSerif(panose: string | undefined): 'serif' | 'sans' | undefined {
  const hex = panose?.replace(/[^0-9a-f]/gi, '')
  if (!hex || hex.length < 4 || /^0+$/.test(hex)) return undefined
  const b = parseInt(hex.slice(2, 4), 16)
  return b >= 0x0b && b <= 0x0f ? 'sans' : b >= 0x02 && b <= 0x0a ? 'serif' : undefined
}

function substituteFor(key: string, hints: FaceHints | undefined): [string, boolean] {
  if (hints?.charset === '81')
    return [hints.family?.toLowerCase() === 'roman' ? 'Batang' : 'Malgun Gothic', true]
  if (HAN_KR_RE.test(key))
    return [panoseSerif(hints?.panose) === 'sans' ? 'Malgun Gothic' : 'Batang', true]
  const alt = hints?.altName && faceKey(hints.altName) !== key ? hints.altName : undefined
  // a missing Arabic name follows its altName before the Times New Roman rule
  // (corpus 2026-10-27: altName Arial laid out in Arial)
  if (alt && ARABIC_RE.test(key) && !ARABIC_RE.test(faceKey(alt))) return [alt, false]
  for (const [re, target] of SUBSTITUTIONS) if (re.test(key)) return [target, true]
  if (alt) return [alt, false]
  // a missing name whose fontTable sig covers the Arabic code page lays out in
  // Arial, Latin and Arabic alike (corpus 2026-10-27: two Latin-charset Arabic
  // sans families came back ArialMT in two documents); Arabic charset (B2)
  // entries keep the Times New Roman rule
  if (arabicCodePage(hints) && hints?.charset !== 'B2') return ['Arial', true]
  if (HAN_RE.test(key)) return ['SimSun', true]
  // hangul goes to Batang whatever the PANOSE says (probe 2026-10-22: sans,
  // serif and absent PANOSE all Batang); only charset 81 reaches Malgun
  if (HANGUL_RE.test(key)) return ['Batang', true]
  return [latinFallback(hints), false]
}

/** Calibri with a fontTable entry, Cambria for a serif PANOSE, a roman family without PANOSE, or no entry (truths 2026-10-22) */
function latinFallback(hints: FaceHints | undefined): string {
  if (!hints) return 'Cambria'
  const serif = panoseSerif(hints.panose)
  if (serif === 'serif' || (!serif && hints.family?.toLowerCase() === 'roman')) return 'Cambria'
  return 'Calibri'
}

/**
 * Face Word draws the Latin of a CJK-named missing face with: the altName
 * when there is one, else the Latin fallback (probe 2026-10-22: hangul names
 * with sans / absent PANOSE -> Calibri, serif PANOSE -> Cambria, Noto Sans
 * CJK SC with altName Calibri -> Calibri); a hangul name without any
 * fontTable entry draws its Latin in SimSun (truth 2026-08-23)
 */
export function latinSlotSubstitute(name: string, hints: FaceHints | undefined): string {
  const key = faceKey(name)
  const alt = hints?.altName
  if (alt && faceKey(alt) !== key) return alt
  if (!hints && HANGUL_RE.test(key)) return 'SimSun'
  return latinFallback(hints)
}

/** hangul-named face with altName Batang lays out in Batang even when installed */
function batangAltNamed(key: string, hints: FaceHints | undefined): boolean {
  const alt = hints?.altName && faceKey(hints.altName)
  return HANGUL_RE.test(key) && !!alt && /batang|\ubc14\ud0d5/.test(alt)
}

/**
 * Word's own copy wins for the symbol faces and for East Asian faces (Yu
 * Gothic lays out at DFonts 1.1021 x 1.3 = 1.4327, not the macOS copy's 1.3);
 * a Latin name takes the installed copy first (Times New Roman 1.15).
 */
function direct(key: string, display: string, ctx: ResolveContext): StoredFace | null {
  const embedded = ctx.embedded(key)
  if (embedded) return { ...embedded, family: display, source: 'embedded' }
  const word = ctx.wordFace(key)
  if (word && (WORD_FIRST.has(key) || word.eastAsian)) return word
  const installed = localizedName(key) ? null : ctx.installed(key)
  if (installed) return installed.family === key ? { ...installed, family: display } : installed
  return word
}

/** the base family of a name Word matches after dropping a style suffix */
function styleStripped(key: string, display: string, ctx: ResolveContext): StoredFace | null {
  const base = key.replace(STYLE_SUFFIX_RE, '')
  if (base === key || !base) return null
  return direct(base, display.replace(/ [^ ]+( italic)?$/i, ''), ctx)
}

/** null when the chain ends in a face the context has no metrics for */
export function resolveFace(name: string, ctx: ResolveContext): ResolvedFace | null {
  const path: string[] = []
  let key = faceKey(name)
  let display = name.split(',')[0].replace(/['"]/g, '').trim()
  let exception = false
  for (let depth = 0; depth < 4; depth++) {
    path.push(key)
    const hints = ctx.hints(key)
    const found = batangAltNamed(key, hints) ? null : direct(key, display, ctx)
    const stripped = found ? null : styleStripped(key, display, ctx)
    const face = found ?? stripped
    if (face) {
      return {
        family: face.family,
        source: face.source,
        substituted: depth > 0 || !!stripped,
        exception,
        metrics: face,
        lineFactor: lineFactorOf(face),
        path,
      }
    }
    const [target, behavioural] = substituteFor(key, hints)
    const next = faceKey(target)
    if (next === key) return null
    exception ||= behavioural
    key = next
    display = target
  }
  return null
}

export interface WordFacesData {
  faces: Record<string, StoredFace>
  system: Record<string, StoredFace>
}

/** Context over the generated word-faces.json, with optional live installed metrics */
export function storeContext(
  data: WordFacesData,
  live: {
    installed?: (key: string) => StoredFace | null | undefined
    hints?: (key: string) => FaceHints | undefined
    embedded?: (key: string) => FaceMetrics | null
  } = {},
): ResolveContext {
  return {
    embedded: (key) => live.embedded?.(key) ?? null,
    // the macOS snapshot stands in while the installed index has no answer
    installed: (key) => {
      const hit = live.installed?.(key)
      return hit === undefined ? (data.system[key] ?? null) : hit
    },
    wordFace: (key) => data.faces[key] ?? null,
    hints: (key) => live.hints?.(key),
  }
}

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

import substitutions from '../measurements/substitutions.json'
import substitutionSteps from '../measurements/substitution-steps.json'

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
  /** English family of a localized spelling of the same face (store twins) */
  canonical?: string
}

/** the rule that took one step of the substitution path */
export type SubstituteRule =
  | 'charset81'
  | 'charset80'
  | 'hanKr'
  | 'altName'
  | 'arabicCodePage'
  | 'table'
  | 'han'
  | 'hangul'
  | 'fallback'
  | 'styleSuffix'

export interface ResolvedFace {
  /** display name of the face Word lays out with */
  family: string
  /** the family under its English name: one key for every spelling of a face */
  canonical: string
  source: FaceSource
  /** the declared name itself was not laid out */
  substituted: boolean
  /** the substitute came from a behavioural rule, not the Calibri/Cambria fallback */
  exception: boolean
  metrics: FaceMetrics
  lineFactor: number
  /** declared name, substitutes, final face */
  path: string[]
  /** rules[i] took path[i] to path[i + 1]; empty when the declared name laid out itself */
  rules: SubstituteRule[]
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
const HAN_KR_RE = /^(noto|source han) (sans|serif)( cjk)? ?(kr|k)\b/
const HG_RE = new RegExp(
  '^hg[ps]?(\u6559\u79d1\u66f8\u4f53|\u660e\u671db|\u884c\u66f8\u4f53|\u6b63\u6977\u66f8\u4f53|\u5275\u82f1\u89d2\u30dd\u30c3\u30d7\u4f53|\u5275\u82f1\u30d7\u30ec\u30bc\u30f3\u30b9|\u30b4\u30b7\u30c3\u30afm)|^\u30d2\u30e9\u30ae\u30ce\u89d2\u30b4|biz ud',
)
const ARABIC_RE =
  /naskh|kufi|arabic|amiri|scheherazade|\b(b|xb|ir)[ -]?(mitra|nazanin|titr|lotus|zar|yekan|koodak|roya|badr|homa|traffic|compset)\b|irlotus/
const TAMIL_RE = /tamil|vijaya|inaimathi|^tau[- ]/
/** Thai legacy families (Windows UPC/New faces, the SIPA TH * national fonts); Tahoma and Leelawadee are Latin-metric and stay out */
const THAI_RE =
  /^th |sarabun|angsana|cordia|browallia|eucrosia|freesia|iris ?upc|jasmine|kodchiang|lily ?upc|dillenia|\bupc$|thai/
/** M365 cloud Arabic faces Word downloads and lays out with real metrics */
const ARABIC_CLOUD_RE = /dubai|simplified arabic|traditional arabic|sakkal majalla|aldhabi|andalus/
/** Word drops a style word or the MT tag and lays out the base family (probe 2026-10-22: Arial Bold, Courier New Regular, Arial MT) */
const STYLE_SUFFIX_RE = / (?:regular|bold|italic|bold italic|mt)$/

export interface SubstitutionRule {
  id: string
  match: Pattern
  target: string
  source: { kind: string; date: string; probe: string }
}

/** named patterns a substitutions.json / substitution-steps.json row may reference instead of spelling a regex */
const PATTERN_REFS = {
  HG_RE,
  ARABIC_RE,
  TAMIL_RE,
  THAI_RE,
  HAN_RE,
  HANGUL_RE,
  HAN_KR_RE,
  ARABIC_CLOUD_RE,
}
type Pattern = string | { ref: keyof typeof PATTERN_REFS }
const patternOf = (m: Pattern): RegExp =>
  typeof m === 'string' ? new RegExp(m) : PATTERN_REFS[m.ref]

export const SUBSTITUTION_RULES = (substitutions as { rules: SubstitutionRule[] }).rules

/** Behavioural substitutions Word applies to a missing name before the generic rules (measurements/substitutions.json, in order) */
const SUBSTITUTIONS: Array<[RegExp, string]> = SUBSTITUTION_RULES.map((r) => [
  patternOf(r.match),
  r.target,
])

/** one step of measurements/substitution-steps.json (field semantics in its _about) */
export interface SubstitutionStep {
  id: string
  rule: SubstituteRule
  when?: {
    charset?: string
    altName?: boolean
    name?: Pattern
    nameNot?: Pattern
    altNameNot?: Pattern
    arabicSig?: boolean
  }
  target:
    | string
    | { altName: true }
    | { byFamily: { roman: string; else: string; none?: string } }
    | { byPanose: { sans: string; else: string } }
    | { table: string }
  behavioural: boolean
  source: { kind: string; date: string; probe: string }
}

export const SUBSTITUTION_STEPS = (substitutionSteps as { steps: SubstitutionStep[] }).steps

/** PANOSE serif-style byte: 0B-0F sans, 02-0A serif */
export function panoseSerif(panose: string | undefined): 'serif' | 'sans' | undefined {
  const hex = panose?.replace(/[^0-9a-f]/gi, '')
  if (!hex || hex.length < 4 || /^0+$/.test(hex)) return undefined
  const b = parseInt(hex.slice(2, 4), 16)
  return b >= 0x0b && b <= 0x0f ? 'sans' : b >= 0x02 && b <= 0x0a ? 'serif' : undefined
}

type Step = [target: string, behavioural: boolean, rule: SubstituteRule]

/** a missing name whose fontTable sig covers the Arabic code page, Arabic charset (B2) entries excepted */
const latinArabicSig = (hints: FaceHints | undefined) =>
  arabicCodePage(hints) && hints?.charset !== 'B2'

/** the fontTable altName when it names another face */
const altNameOf = (key: string, hints: FaceHints | undefined): string | undefined =>
  hints?.altName && faceKey(hints.altName) !== key ? hints.altName : undefined

function stepApplies(
  when: NonNullable<SubstitutionStep['when']>,
  key: string,
  hints: FaceHints | undefined,
  alt: string | undefined,
): boolean {
  // font-rules: allow -- the interpreter's predicate, not a rule
  if (when.charset !== undefined && hints?.charset !== when.charset) return false
  if (when.altName !== undefined && !!alt !== when.altName) return false
  if (when.name && !patternOf(when.name).test(key)) return false
  if (when.nameNot && patternOf(when.nameNot).test(key)) return false
  if (when.altNameNot && (!alt || patternOf(when.altNameNot).test(faceKey(alt)))) return false
  if (when.arabicSig !== undefined && latinArabicSig(hints) !== when.arabicSig) return false
  return true
}

const familyChoice = (
  choice: { roman: string; else: string; none?: string },
  hints: FaceHints | undefined,
): string => {
  if (!hints && choice.none) return choice.none
  return hints?.family?.toLowerCase() === 'roman' ? choice.roman : choice.else
}

/** the step's face; undefined when a table step has no row for the name */
function stepTarget(
  target: SubstitutionStep['target'],
  key: string,
  hints: FaceHints | undefined,
  alt: string | undefined,
): string | undefined {
  if (typeof target === 'string') return target
  if ('altName' in target) return alt
  if ('byFamily' in target) return familyChoice(target.byFamily, hints)
  if ('byPanose' in target)
    return panoseSerif(hints?.panose) === 'sans' ? target.byPanose.sans : target.byPanose.else
  for (const [re, face] of SUBSTITUTIONS) if (re.test(key)) return face
  return undefined
}

const FALLBACK_STEP = SUBSTITUTION_STEPS.find((s) => s.id === 'fallback')!

/** the first step of substitution-steps.json that applies and yields a face */
function substituteFor(key: string, hints: FaceHints | undefined): Step {
  const alt = altNameOf(key, hints)
  for (const step of SUBSTITUTION_STEPS) {
    if (step.when && !stepApplies(step.when, key, hints, alt)) continue
    const target = stepTarget(step.target, key, hints, alt)
    if (target !== undefined) return [target, step.behavioural, step.rule]
  }
  return [latinFallback(hints), false, 'fallback']
}

/**
 * Cambria for no fontTable entry or a w:family roman entry, Calibri for every
 * other entry: PANOSE (serif or sans), charset and the sig code pages do not
 * take part (probes 2026-10-22, 2026-11-08: a serif PANOSE alone came back
 * Calibri, a sans PANOSE with family roman Cambria)
 */
export function latinFallback(hints: FaceHints | undefined): string {
  const target = FALLBACK_STEP.target
  return typeof target === 'object' && 'byFamily' in target
    ? familyChoice(target.byFamily, hints)
    : 'Cambria'
}

/**
 * Face Word draws the Latin of a CJK-named missing face with: the altName
 * when there is one, else the Latin fallback (probe 2026-10-22: hangul names
 * -> Calibri, roman family -> Cambria, Noto Sans CJK SC with altName Calibri
 * -> Calibri); a hangul name without any fontTable entry draws its Latin in
 * SimSun (truth 2026-08-23)
 */
export function latinSlotSubstitute(name: string, hints: FaceHints | undefined): string {
  const key = faceKey(name)
  const alt = hints?.altName
  if (alt && faceKey(alt) !== key) return alt
  if (!hints && HANGUL_RE.test(key)) return 'SimSun'
  return latinFallback(hints)
}

/** hangul-named face with altName Batang lays out in Batang even when installed */
export function batangAltNamed(key: string, hints: FaceHints | undefined): boolean {
  const alt = hints?.altName && faceKey(hints.altName)
  return HANGUL_RE.test(key) && !!alt && /batang|\ubc14\ud0d5/.test(alt)
}

/**
 * Word's own copy wins for the symbol faces and for East Asian faces (Yu
 * Gothic lays out at DFonts 1.1021 x 1.3 = 1.4327, not the macOS copy's 1.3);
 * a Latin name takes the installed copy first (Times New Roman 1.15).
 */
function direct(key: string, display: string, ctx: ResolveContext): StoredFace | null {
  const word = ctx.wordFace(key)
  const embedded = ctx.embedded(key)
  // Word's shipped copy beats the document's embed (corpus 2026-11-01: a docx
  // embedding Century Gothic 2.37g, hhea 1.4282, printed Word's own 2.35 at 1.2261)
  if (embedded && word?.source !== 'dfonts')
    return { ...embedded, family: display, source: 'embedded' }
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
  const rules: SubstituteRule[] = []
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
      if (stripped) rules.push('styleSuffix')
      return {
        family: face.family,
        canonical: face.canonical ?? face.family,
        source: face.source,
        substituted: depth > 0 || !!stripped,
        exception,
        metrics: face,
        lineFactor: lineFactorOf(face),
        path,
        rules,
      }
    }
    const [target, behavioural, rule] = substituteFor(key, hints)
    const next = faceKey(target)
    if (next === key) return null
    exception ||= behavioural
    rules.push(rule)
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

/**
 * CSS font-family chain of a declared name: the target-keyed stand-ins, the
 * substituted path, the name-keyed rules of face-chains.ts, then the
 * unknown-name tail; plus the dual / East Asian-only / complex-script run
 * chains (INVENTORY-fonts C1-C11).
 */
import { faceKey, latinSlotSubstitute } from '@genoffice/font-metrics/resolve'
import { ARABIC_NAME_RE, THAI_FONT_RE } from '../../font-name-patterns'
import { resolveDeclaredFace } from '../../font-resolve'
import {
  BOX,
  CJK_SANS,
  CJK_SERIF,
  cssChain,
  JA_SERIF,
  KO_SANS,
  KO_SERIF,
  namedFaceChain,
} from '../face-chains'
import { CJK_LETTERED_RE as CJK_NAMED_RE, KO_FACE_RE as KO_FONT_RE } from '../face-info'
import {
  declaredFaceMissing,
  GENERIC_FAMILIES,
  isBundledFont,
  isFontAvailable,
} from './availability'
import { fontTableHint, krBatangAltNamed, panoseSerifHint, wordMetricAliasFace } from './font-table'
import { STAND_IN_ALIAS_CHAINS } from '../../stand-in-chains.generated'

type Generic = 'serif' | 'sans-serif' | 'monospace'

/**
 * CSS chains keyed by the face Word lays out with (the resolver's canonical
 * target). Rows a metric alias can derive (alias, companions, class tail) come
 * from the manifests' standInChain; the hand rows below need a companion or
 * tail the manifests do not express.
 */
const HAND_STAND_IN_CHAINS: Record<string, [string[], Generic]> = {
  calibri: [['Carlito GO', 'Calibri Arabic GO', CJK_SANS], 'sans-serif'],
  'calibri light': [
    ['Calibri Light GO', 'Carlito GO', 'Calibri Arabic GO', CJK_SANS],
    'sans-serif',
  ],
  cambria: [['Caladea', 'Cambria Cyrillic GO', BOX, CJK_SERIF], 'serif'],
  'times new roman': [['Liberation Serif', BOX, CJK_SERIF], 'serif'],
  'segoe ui': [['Segoe UI GO', 'Noto Sans CJK GO', 'GenOffice PUA Blank'], 'sans-serif'],
  'ms pmincho': [JA_SERIF, 'serif'],
  'malgun gothic': [KO_SANS, 'sans-serif'],
  batang: [KO_SERIF, 'serif'],
  gulim: [KO_SANS, 'sans-serif'],
  dotum: [KO_SANS, 'sans-serif'],
  nanumgothic: [['GenOffice Gothic KR', ...KO_SANS], 'sans-serif'],
  'nanum gothic': [['GenOffice Gothic KR', ...KO_SANS], 'sans-serif'],
  'nanumgothic extrabold': [['GenOffice Gothic KR', ...KO_SANS], 'sans-serif'],
  nanummyeongjo: [KO_SERIF, 'serif'],
  'nanum myeongjo': [KO_SERIF, 'serif'],
  nanummyeongjoextrabold: [KO_SERIF, 'serif'],
  simsun: [
    ['GenOffice SimSun Latin', 'GenOffice Songti SC', 'STSong', 'SimSun', CJK_SERIF],
    'serif',
  ],
  simhei: [['Heiti SC', 'STHeiti', 'SimHei', 'PingFang SC', CJK_SANS], 'sans-serif'],
  'microsoft yahei': [['Microsoft YaHei', 'PingFang SC', CJK_SANS], 'sans-serif'],
  dengxian: [['DengXian GO', 'DengXian', 'PingFang SC', 'Microsoft YaHei', CJK_SANS], 'sans-serif'],
  kaiti: [['STKaiti', 'Kaiti SC', 'KaiTi', CJK_SERIF], 'serif'],
  fangsong: [['STFangsong', 'FangSong', CJK_SERIF], 'serif'],
}

export const STAND_IN_CHAINS: Record<string, [string[], Generic]> = {
  ...STAND_IN_ALIAS_CHAINS,
  ...HAND_STAND_IN_CHAINS,
}

/** a hand row keyed like a generated one would silently shadow the manifest */
export const shadowedStandInRows = (): string[] =>
  Object.keys(HAND_STAND_IN_CHAINS).filter((k) => k in STAND_IN_ALIAS_CHAINS)

/** Chain by Word's target face; null for substitutions and unkeyed targets */
export function standInFontFamily(font: string): string | null {
  const face = resolveDeclaredFace(font)
  if (!face || face.substituted) return null
  const hit = STAND_IN_CHAINS[faceKey(face.canonical)]
  if (!hit) return null
  const [families, generic] = hit
  const head = font.split(',')[0].replace(/['"]/g, '').trim()
  return `${[...new Set([head, ...families])].map((x) => `'${x}'`).join(',')},${generic}`
}

/**
 * Word follows a missing Arabic name's fontTable altName before its Times New
 * Roman rule (corpus 2026-10-27: altName Arial laid out in Arial, pitch 1.15);
 * the altName face then sets the chain and the line factor
 */
export function arabicAltNameFace(font: string): string | null {
  const head = font.split(',')[0].replace(/['"]/g, '').normalize('NFKC').trim()
  if (!ARABIC_NAME_RE.test(head) || !declaredFaceMissing(head)) return null
  if (resolveDeclaredFace(head)?.rules[0] !== 'altName') return null
  const alt = fontTableHint(head.toLowerCase())?.altName?.normalize('NFKC').trim()
  return alt && !ARABIC_NAME_RE.test(alt) ? alt : null
}

/** M365 cloud Arabic faces Word downloads and lays out with real metrics */
const ARABIC_CLOUD_RE = /dubai|simplified arabic|traditional arabic|sakkal majalla|aldhabi|andalus/i

/**
 * A missing Arabic name without altName whose fontTable sig covers the Arabic
 * code page (Latin charset) lays out in Arial, not Times New Roman (corpus
 * 2026-10-27: a 36 pt heading came back ArialMT): the resolver's
 * arabicCodePage step, read for Arabic names. Cloud faces the store lacks
 * (Andalus, Aldhabi) keep their own chains: routed through Arial they
 * paginated into a non-converging loop (same corpus)
 */
export function arabicCodePageFace(font: string): string | null {
  const head = font.split(',')[0].replace(/['"]/g, '').normalize('NFKC').trim()
  if (!ARABIC_NAME_RE.test(head) || ARABIC_CLOUD_RE.test(head) || !declaredFaceMissing(head))
    return null
  return resolveDeclaredFace(head)?.rules[0] === 'arabicCodePage' ? 'Arial' : null
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
export function latinSlotTarget(font: string): string | null {
  const face = resolveDeclaredFace(font)
  if (!face?.substituted || !face.exception) return null
  const head = font.split(',')[0].replace(/['"]/g, '').trim()
  const nfkc = head.normalize('NFKC')
  if (ARABIC_NAME_RE.test(nfkc) || NAME_KEYED_SUBSTITUTE_RE.test(nfkc)) return null
  if (!CJK_NAMED_RE.test(nfkc) && !HAN_REGIONAL_NAME_RE.test(nfkc)) return null
  if (!/^(simsun|batang|\u5b8b\u4f53|\ubc14\ud0d5)$/i.test(face.family.normalize('NFKC')))
    return null
  return latinSlotSubstitute(head, fontTableHint(nfkc.toLowerCase()))
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
  // Batang, a macOS face under its localized name) or, for a CSS generic
  // keyword used as a font name, the platform UI face: the head must not lead
  const headLeads =
    !isBundledFont(head) &&
    !GENERIC_FAMILIES.has(nfkc.toLowerCase()) &&
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

/** declared name -> the installed face's other spellings, canonical first;
    null until the main process answers (genoffice#2449) */
let fontAliasLookup: ((name: string) => readonly string[] | null) | null = null

export function setFontAliasLookup(fn: (name: string) => readonly string[] | null): void {
  fontAliasLookup = fn
}

export function cssFontFamily(font: string): string {
  // a localized name the browser cannot match takes the face's first spelling
  // it can; a name that resolves itself keeps its own weight ("PingFang SC
  // Semibold" must not collapse to the family)
  const declared = font.split(',')[0].replace(/['"]/g, '').trim()
  if (!isFontAvailable(declared)) {
    const candidate = fontAliasLookup?.(declared)?.find((name) => isFontAvailable(name))
    if (candidate) font = candidate
  }
  // faces Word lays out directly take the chain of their target face, names it
  // substitutes the chain of the substitute; the name-keyed branches below
  // serve unkeyed targets and the script chains
  // Thai: Word ships TH SarabunPSK and Cordia New and substitutes Cordia New
  // for every other Thai family (probe 2026-11-03); macOS has neither and
  // Chromium's Thonburi advances ~1.75x wider, so the bundled metric faces
  // (tools/build-thai-font.py) lead. On Windows the declared name resolves first.
  if (THAI_FONT_RE.test(font.normalize('NFKC').toLowerCase())) {
    const face = /^th sarabunpsk$/i.test(font.normalize('NFKC'))
      ? 'GenOffice Sarabun'
      : 'GenOffice Cordia'
    return `${cssChain(font, face, 'Thonburi')},sans-serif`
  }
  const byFace = standInFontFamily(font) ?? substituteFontFamily(font)
  if (byFace) return byFace
  const named = namedFaceChain({
    font,
    available: isFontAvailable(font),
    bundled: isBundledFont(font),
    aliasFace: wordMetricAliasFace,
    arabicAlt: (nfkc) => arabicAltNameFace(nfkc) ?? arabicCodePageFace(nfkc),
    recurse: cssFontFamily,
  })
  if (named) return named.css
  const nfkc = font.normalize('NFKC')
  const chain = cssChain
  // unknown font family: fontTable PANOSE decides serif-ness when the font is
  // missing, else guess by name (Song/Ming/Serif → serif fallback)
  const hint = declaredFaceMissing(font) ? panoseSerifHint(font) : undefined
  const serifLike = hint ? hint === 'serif' : /\u5b8b|\u660e|serif|song|ming/i.test(font)
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
  // the JA chains' hangul tail follows the Noto JP face: a dual-slot run's hangul
  // belongs to its eastAsia face (Korean chains keep their own KR stand-ins)
  const jp = fams.findIndex((fam) => /^'Noto (?:Sans|Serif) JP'$/.test(fam))
  const latin = fams.filter(
    (fam, i) =>
      !(jp >= 0 && i > jp && /^'GenOffice (?:Sans KR|Serif KR|Batang)'$/.test(fam)) &&
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
  const timesSub = pre.some((f) => latinSub.test(f))
  // the system Times draws the substituted Arabic; the scaled alias serves hosts without it
  const preKept = pre.filter((f) => !latinSub.test(f) && !(timesSub && f === head[notoIdx]))
  if (timesSub) preKept.push("'Times Punct GO'")
  // the ascii slot's Arabic companions (Calibri/Arial Arabic GO) must not claim the cs glyphs
  const companion = /arabic go\b/i
  const merged =
    notoIdx >= 0
      ? [
          ...preKept,
          ...baseFams.filter((f) => !generic.test(f) && !companion.test(f)),
          ...pre.filter((f) => latinSub.test(f)),
          ...(timesSub ? [head[notoIdx]] : []),
          ...head.slice(notoIdx + 1),
          ...baseFams.filter((f) => companion.test(f)),
          ...baseFams.filter((f) => generic.test(f)),
        ]
      : [...head, ...baseFams]
  return [...new Set(merged)].join(',')
}

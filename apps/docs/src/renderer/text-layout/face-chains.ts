/**
 * Name-keyed font-family chains: the table cssFontFamily walks after the
 * resolver has had its say (Thai metric faces, Word stand-ins and
 * substitutions come first; the unknown-name fallback comes last). One rule
 * per face family, in priority order, each with the Word probe that fixed its
 * chain. Rules read only the name context; availability on this machine and
 * the alias / Arabic helpers come in through `NameContext` so this module
 * never imports line-metrics.
 */
import { JA_FACE_RE } from './face-info'
import {
  ARABIC_NAME_RE,
  MONO_FONT_RE,
  MONTSERRAT_FAMILY_RE,
  SEGOE_UI_TEXT_RE,
  TAMIL_FONT_RE,
} from '../font-name-patterns'

export const CJK_SERIF = 'Noto Serif CJK SC'
export const CJK_SANS = 'Noto Sans CJK SC'
export const BOX = 'GenOffice Box Drawing'
export const JA_SANS = ['Yu Gothic', 'GenOffice Hiragino Sans', 'Meiryo', 'Noto Sans JP']
export const JA_SERIF = [
  'Yu Mincho',
  'GenOffice Hiragino Mincho',
  'GenOffice MS Mincho',
  'Noto Serif JP',
]
export const KO_SANS = ['Malgun Gothic', 'GenOffice Sans KR', 'Apple SD Gothic Neo', 'Noto Sans KR']
export const KO_SERIF = [
  'GenOffice Batang',
  'GenOffice Serif KR',
  'GenOffice Myungjo',
  'Noto Serif KR',
]
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

export const cssChain = (...families: string[]): string =>
  [...new Set(families)].map((x) => `'${x.replace(/'/g, '')}'`).join(',')

export interface NameContext {
  /** the declared family as written */
  font: string
  /** the face is installed on this machine (canvas probe) */
  available: boolean
  /** the face is one of the bundled fonts.css families */
  bundled: boolean
  /** metric alias (fonts.css) of an Office-private face, else null */
  aliasFace: (font: string) => string | null
  /** Word's substitute for a missing Arabic name (fontTable altName, code page), else null */
  arabicAlt: (nfkc: string) => string | null
  /** cssFontFamily itself, for rules that route to another name's chain */
  recurse: (font: string) => string
}

type Generic = 'serif' | 'sans-serif' | 'monospace'

interface ChainRule {
  id: string
  /** lower-cased name `f`, lower-cased chain head `head`, NFKC name `nfkc` */
  when: (f: string, head: string, nfkc: string, n: NameContext) => boolean
  chain: (f: string, head: string, nfkc: string, n: NameContext) => [string[], Generic] | string
}

const KO_NAME_RE =
  /[가-힣ᄀ-ᇿ\u3130-\u318f]|malgun|batang|gulim|dotum|gungsuh|myeongjo|myungjo|nanum|apple (sd )?gothic/i
const TC_NAME_RE =
  /jhenghei|p?mingliu|biaukai|dfkai|kaiu|\u6b63\u9ed1|\u7d30\u660e|\u6a19\u6977|\u860b\u65b9|-\u7e41|\u7e41\u9ad4|pingfang (tc|hk)|(heiti|songti|kaiti) tc/i

/** rules in priority order; the first whose `when` holds decides the chain */
export const FACE_CHAIN_RULES: ChainRule[] = [
  // 'Calibri Arabic GO' (fonts.css): Calibri's own Arabic; Carlito has none
  {
    id: 'calibri',
    when: (f) => f.includes('calibri'),
    chain: () => [['Carlito GO', 'Calibri Arabic GO', CJK_SANS], 'sans-serif'],
  },
  // Aptos (M365 cloud face, never installed locally): line metrics equal Calibri's
  // (Word probe 2026-08-22) but its advances do not, so the size-adjusted Carlito
  // aliases in fonts.css stand in (probe 2026-09-03); 'GenOffice PUA Blank' keeps
  // AI-residue PUA tokens invisible like Word. Body, Display and Narrow cuts have
  // fitted aliases; Mono/Serif are unprobed
  {
    id: 'aptos',
    when: (f) => f.includes('aptos'),
    chain: (f) => {
      const alias = /mono|serif/.test(f)
        ? 'Carlito GO'
        : f.includes('display')
          ? 'Aptos Display GO'
          : f.includes('narrow')
            ? 'Aptos Narrow GO'
            : 'Aptos GO'
      return [[alias, 'GenOffice PUA Blank', CJK_SANS], 'sans-serif']
    },
  },
  // Cambria Math text runs lay out with Cambria's advances in Word (fonts.css alias); only the symbols need a math face
  {
    id: 'cambria-math',
    when: (f) => f.includes('cambria') && f.includes('math'),
    chain: () => [
      ['Cambria Math GO', 'STIX Two Math', 'Caladea', 'Liberation Serif', BOX, CJK_SERIF],
      'serif',
    ],
  },
  // math faces would fall to the unknown-name sans fallback; STIX Two Math ships with macOS
  {
    id: 'math',
    when: (f) => /stix.*math|latin modern math/.test(f),
    chain: () => [['STIX Two Math', 'Caladea', 'Liberation Serif', BOX, CJK_SERIF], 'serif'],
  },
  {
    id: 'cambria',
    when: (f) => f.includes('cambria'),
    chain: () => [['Caladea', BOX, CJK_SERIF], 'serif'],
  },
  {
    id: 'times',
    when: (f) => f.includes('times'),
    chain: () => [['Liberation Serif', BOX, CJK_SERIF], 'serif'],
  },
  // MS Reference Sans Serif ships with Office and shares Verdana's tables; the local() alias draws Verdana where the face is missing
  {
    id: 'reference-sans',
    when: (f) => f.includes('reference sans'),
    chain: () => [['MS Reference Sans Serif GO', 'Verdana', CJK_SANS], 'sans-serif'],
  },
  // Merriweather is an M365 cloud face Word renders real; the size-adjusted Georgia alias carries its advances (sample lines 1.12x Georgia, sd 0.6%)
  {
    id: 'merriweather',
    when: (f) => f.includes('merriweather') && !f.includes('sans'),
    chain: () => [['Merriweather GO', 'Georgia', BOX, CJK_SERIF], 'serif'],
  },
  // Palatino Linotype/Book Antiqua ship with Office and Word renders them real (probe 2026-08-23); macOS Palatino matches their Latin widths within 0.2%
  {
    id: 'palatino',
    when: (f) => f.includes('palatino') || f.includes('book antiqua'),
    chain: () => [['Palatino Linotype', 'Palatino', 'Book Antiqua', BOX, CJK_SERIF], 'serif'],
  },
  // 'Arial Arabic GO' (fonts.css): Arial's Arabic advances where the local Arial (Liberation Sans) has none
  {
    id: 'arial',
    when: (f) => f === 'arial' || f.startsWith('arial '),
    chain: () => [['Arial Arabic GO', 'Liberation Sans', CJK_SANS], 'sans-serif'],
  },
  // Roboto is an M365 cloud face Word renders real; the generic sans-serif tail would draw Helvetica,
  // 6-8% wider in bold, so the per-class size-adjusted Liberation Sans alias carries its advances
  {
    id: 'roboto',
    when: (_f, head) => /^roboto$/.test(head),
    chain: () => [['Roboto GO', CJK_SANS], 'sans-serif'],
  },
  // Garamond ships with Office and Word renders it real; it is a macOS on-demand asset Chromium
  // cannot see, and the serif tail (Times) runs 3-5% wider, so the size-adjusted alias stands in
  {
    id: 'garamond',
    when: (_f, head) => /^garamond$/.test(head),
    chain: () => [['Garamond GO', BOX, CJK_SERIF], 'serif'],
  },
  // Berlin Sans FB Demi ships with Office, a bold display cut 0.953x Arial Bold wide (probe 2026-09-17)
  {
    id: 'berlin-sans-demi',
    when: (f) => /berlin sans fb demi/.test(f),
    chain: () => [['Berlin Sans FB GO', CJK_SANS], 'sans-serif'],
  },
  // Segoe UI is an M365 cloud font Word renders real; the size-adjusted Helvetica alias carries its narrower advances
  {
    id: 'segoe-ui',
    when: (f) => SEGOE_UI_TEXT_RE.test(f),
    chain: () => [['Segoe UI GO', 'Noto Sans CJK GO', 'GenOffice PUA Blank'], 'sans-serif'],
  },
  // Consolas is Office-only: the size-adjusted Menlo alias carries its advances
  {
    id: 'consolas',
    when: (f) => f.includes('consolas'),
    chain: () => [['Consolas GO', CJK_SANS], 'monospace'],
  },
  // Century Gothic is Office-only too: the advance-matched Liberation Sans derivative carries its wide lowercase / narrow capitals
  {
    id: 'century-gothic',
    when: (f) => f.includes('century gothic'),
    chain: () => [['Century Gothic GO', CJK_SANS], 'sans-serif'],
  },
  // Lucida Handwriting ships with Office and Windows; elsewhere the per-case size-adjusted Liberation Sans Italic alias carries its wide script advances
  {
    id: 'lucida-handwriting',
    when: (f) => f.includes('lucida handwriting'),
    chain: () => [['Lucida Handwriting GO', CJK_SANS], 'sans-serif'],
  },
  {
    id: 'mono',
    when: (f) => MONO_FONT_RE.test(f),
    chain: () => [['Menlo', 'Courier New', 'Liberation Mono', CJK_SANS], 'monospace'],
  },
  // Nunito Sans is an Office cloud font Word renders real; the size-adjusted Helvetica alias matches its advances (probe 2026-08-23)
  {
    id: 'nunito',
    when: (f) => f.includes('nunito'),
    chain: () => [['Nunito Sans GO', CJK_SANS], 'sans-serif'],
  },
  // Poppins is an M365 cloud font Word renders real; the bundled Latin subset carries its true advances (probe 2026-09-01)
  {
    id: 'poppins',
    when: (f) => f.includes('poppins'),
    chain: () => [['GenOffice Poppins', CJK_SANS], 'sans-serif'],
  },
  // DM Sans renders real in Word when installed; the bundled Latin subset carries its true advances (corpus 2026-09-30)
  {
    id: 'dm-sans',
    when: (f) => f.includes('dm sans'),
    chain: () => [['GenOffice DM Sans', CJK_SANS], 'sans-serif'],
  },
  // Montserrat is an M365 cloud font Word renders real, ignoring its fontTable altName; only the
  // cloud family names qualify: a PostScript-style "Montserrat-Bold" is unknown to Word and keeps the altName
  {
    id: 'montserrat',
    when: (f) => MONTSERRAT_FAMILY_RE.test(f),
    chain: () => [['Montserrat GO', CJK_SANS], 'sans-serif'],
  },
  // Microsoft New Tai Lue ships with Office; its Latin is Segoe-flavored with Arial-class widths (probe 2026-08-23)
  {
    id: 'new-tai-lue',
    when: (f) => f.includes('new tai lue'),
    chain: () => [['Segoe UI', 'Helvetica', 'Liberation Sans', CJK_SANS], 'sans-serif'],
  },
  // ZhongSong (STZhongsong, gov-document title font) before the generic SimSun branch
  {
    id: 'zhongsong',
    when: (f) => f.includes('\u4e2d\u5b8b') || f.includes('zhongsong'),
    chain: () => [['STZhongsong', 'Songti SC', 'STSong', 'SimSun', CJK_SERIF], 'serif'],
  },
  // 'GenOffice Songti SC' (fonts.css local() alias of Songti SC): macOS Chromium refuses synthetic
  // bold for 'Songti SC' by name; the alias lets Blink synthesize. SimSun's ASCII is half-width
  // fixed pitch (0.5em); the Songti stand-in is proportional, so the Latin alias leads
  {
    id: 'simsun',
    when: (f) => /simsun|\u5b8b\u4f53|nsimsun/.test(f),
    chain: () => [
      ['GenOffice SimSun Latin', 'GenOffice Songti SC', 'STSong', 'SimSun', CJK_SERIF],
      'serif',
    ],
  },
  {
    id: 'simhei',
    when: (f) =>
      f.includes('simhei') ||
      f.includes('\u9ed1\u4f53') ||
      f.includes('\u7ec6\u9ed1') ||
      f.includes('xihei'),
    chain: () => [['Heiti SC', 'STHeiti', 'SimHei', 'PingFang SC', CJK_SANS], 'sans-serif'],
  },
  {
    id: 'yahei',
    when: (f) => f.includes('yahei') || f.includes('\u96c5\u9ed1'),
    chain: () => [['Microsoft YaHei', 'PingFang SC', CJK_SANS], 'sans-serif'],
  },
  {
    id: 'dengxian',
    when: (f) => f.includes('\u7b49\u7ebf') || f.includes('dengxian'),
    chain: () => [
      ['DengXian GO', 'DengXian', 'PingFang SC', 'Microsoft YaHei', CJK_SANS],
      'sans-serif',
    ],
  },
  {
    id: 'fangsong',
    when: (f) => f.includes('fangsong') || f.includes('\u4eff\u5b8b'),
    chain: () => [['STFangsong', 'FangSong', CJK_SERIF], 'serif'],
  },
  {
    id: 'kaiti',
    when: (f) => f.includes('kaiti') || f.includes('\u6977\u4f53'),
    chain: () => [['STKaiti', 'Kaiti SC', 'KaiTi', CJK_SERIF], 'serif'],
  },
  {
    id: 'lisu',
    when: (f) => f.includes('\u96b6\u4e66') || f.includes('lisu'),
    chain: () => [['Baoli SC', 'LiSu', CJK_SERIF], 'serif'],
  },
  // Arabic: bundled Noto subsets stand in for missing fonts; Chromium's silent fallback is a Geeza
  // Pro-style UI face, larger and heavier than the naskh serif Word substitutes. Unknown Arabic
  // names default to the naskh chain. Iranian B/XB/IR faces included: Word for Mac substitutes
  // all of them with Times New Roman (probe 2026-08-13)
  {
    id: 'arabic',
    when: (_f, _h, nfkc) => ARABIC_NAME_RE.test(nfkc),
    chain: (_f, _h, nfkc, n) => {
      const alt = n.arabicAlt(nfkc)
      if (alt) return n.recurse(alt)
      // Dubai is a cloud face Word renders real; the size-adjusted alias carries its advances and hhea box where it is missing
      if (/\bdubai\b/i.test(nfkc) && !n.bundled && !n.available)
        return [['Dubai GO', 'Arial', 'Liberation Sans', CJK_SANS], 'sans-serif']
      // an installed kufi/sans face renders as such; a missing one substitutes to Times New Roman
      // in Word like every other missing Arabic name (probe 2026-09-11). The bundled Noto Sans
      // Arabic subset answers the canvas probe but is not an installed face.
      const sans = /kufi|sans|dubai|segoe/i.test(nfkc) && !n.bundled && n.available
      // Traditional/Simplified Arabic are compact naskh faces Word renders real (M365 cloud fonts);
      // the per-face size-adjusted aliases carry their advances, other Arabic names keep the unscaled subset
      const compact = /\b(traditional|simplified) arabic\b/i.exec(nfkc)?.[1].toLowerCase()
      // Sakkal Majalla (cloud face, compact naskh): per-class aliases carry its Arabic, digit, punctuation and Latin advances
      const sakkal = /\bsakkal majalla\b/i.test(nfkc)
      // declared Noto Arabic names would resolve to the bundled unscaled subsets (~13% wider than Word); the literal head must go so the 'W' alias wins
      const scaled = /^noto (naskh|sans) arabic$/i.test(nfkc.trim())
      // Word substitutes missing naskh names wholesale with Times New Roman, so ASCII punctuation/digits
      // must take Times shapes, not the Naskh subset's. mac/Windows Times carries full Arabic, so on those
      // platforms Arabic letters also resolve here; the scaled subset serves platforms without a system Times.
      const missingSerif = !sans && !compact && !scaled && !sakkal && !n.available
      const latinHead = missingSerif ? ['Times New Roman', 'Liberation Serif'] : []
      const fams = sans
        ? [scaled ? 'Noto Sans Arabic W' : 'Noto Sans Arabic', 'Geeza Pro']
        : [
            ...latinHead,
            // missing serif names: 90% TNR alias (subset shapes 1.125x/1.074x wider than Word's Times, M3 probe); Arabic-Indic digits unscaled
            ...(missingSerif
              ? ['Naskh Digits GO', 'Noto Naskh Arabic TNR']
              : sakkal
                ? ['Sakkal Majalla GO', 'Sakkal Majalla Latin GO']
                : compact
                  ? [
                      compact === 'traditional' ? 'Noto Naskh Arabic TA' : 'Noto Naskh Arabic SA',
                      // the subsets carry no ASCII brackets and neither does Geeza Pro, so they fell to Al Bayan (30% wider, 1.5 em box); both cloud faces draw them at Times widths
                      'Times Punct GO',
                    ]
                  : [scaled ? 'Noto Naskh Arabic W' : 'Noto Naskh Arabic']),
            'Geeza Pro',
            'Al Bayan',
          ]
      return `${cssChain(...(scaled ? [] : [n.font]), ...fams)},${sans ? 'sans-serif' : 'serif'}`
    },
  },
  // Noto CJK / Source Han regional variants Word lays out for real (installed; the missing ones are substituted earlier) route by suffix
  {
    id: 'noto-cjk-variant',
    when: (_f, _h, nfkc) =>
      /^(?:noto|source han) (sans|serif)(?: cjk)? ?(jp|kr|k|tc|tw|hk|sc|cn)\b/i.test(nfkc),
    chain: (_f, _h, nfkc) => {
      const m = /^(?:noto|source han) (sans|serif)(?: cjk)? ?(jp|kr|k|tc|tw|hk|sc|cn)\b/i.exec(
        nfkc,
      )!
      const region = m[2].toLowerCase()
      const serif = /serif/i.test(m[1])
      const fams =
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
      return [fams, serif ? 'serif' : 'sans-serif']
    },
  },
  {
    id: 'japanese',
    when: (_f, _h, nfkc) => JA_FACE_RE.test(nfkc),
    chain: (_f, _h, nfkc, n) => {
      const serif = /mincho|\u660e\u671d/i.test(nfkc)
      // Meiryo (UI): Word renders the real faces; the range-limited aliases rescale the Hiragino/Verdana fallbacks to Meiryo advances
      if (/meiryo|\u30e1\u30a4\u30ea\u30aa/i.test(nfkc)) {
        const alias = /(meiryo|\u30e1\u30a4\u30ea\u30aa) ?ui/i.test(nfkc)
          ? 'Meiryo UI GO'
          : 'Meiryo GO'
        return [[alias, ...JA_SANS], 'sans-serif']
      }
      // MS Gothic family renders real in Word (half-width mono Latin / proportional kana); the
      // range-limited aliases reproduce those advances. MS Mincho (Word probe 2026-09-17: real
      // MS-Mincho, digits 0.5em, fullwidth colon 1em) shares the fixed-pitch Latin alias
      const msGothic = /^ms (ui )?(p)?(gothic|\u30b4\u30b7\u30c3\u30af)$/i.exec(nfkc.trim())
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
      // Yu Mincho/Gothic are Office-private on macOS (Word lays out with its own yumin/YuGoth faces);
      // a missing declare takes the Latin advance aliases so EN/ID lines wrap where Word's do. Yu Gothic UI stays out.
      const yu =
        /^(?:yu (?:mincho|gothic)|\u6e38(?:\u660e\u671d|\u30b4\u30b7\u30c3\u30af))(?: (light|regular|medium|bold|demibold))?$/i.exec(
          nfkc.trim(),
        )
      const yuLight = yu?.[1]?.toLowerCase() === 'light'
      const yuAlias =
        yu && !n.available
          ? serif
            ? [yuLight ? 'Yu Mincho Light GO' : 'Yu Mincho GO', 'Yu Mincho Serif GO']
            : [yuLight ? 'Yu Gothic Light GO' : 'Yu Gothic GO', 'Yu Gothic Sans GO']
          : []
      return [
        [...msAlias, ...yuAlias, ...(serif ? JA_SERIF : JA_SANS)],
        serif ? 'serif' : 'sans-serif',
      ]
    },
  },
  {
    id: 'korean',
    when: (_f, _h, nfkc) => KO_NAME_RE.test(nfkc),
    chain: (_f, _h, nfkc) => {
      // -Che fixed-pitch faces render real in Word with half-width Latin (0.5em fixed, probe 2026-08-24);
      // ASCII rides the bundled Che face, hangul falls through to the KR chains (GungsuhChe keeps the calligraphic GungSeo hangul)
      if (/(?:batang|gulim|dotum|gungsuh) ?che\b|(?:바탕|굴림|돋움|궁서)체/i.test(nfkc)) {
        const cheGungsuh = /gungsuh|궁서/i.test(nfkc)
        const cheSerif = cheGungsuh || /batang|바탕/i.test(nfkc)
        return [
          [
            'GenOffice Che Latin KR',
            ...(cheGungsuh ? ['GungSeo'] : []),
            ...(cheSerif ? KO_SERIF : KO_SANS),
          ],
          cheSerif ? 'serif' : 'sans-serif',
        ]
      }
      // Gungsuh ships with Office (batang.ttc) and Word renders it real; its Latin is typewriter-slab at
      // ~0.58em advances, Courier New the closest installed face (probe 2026-08-23, +2.8% width); hangul
      // falls through to the calligraphic GungSeo, then the serif KR chain
      if (/gungsuh|궁서/i.test(nfkc)) return [['Courier New', 'GungSeo', ...KO_SERIF], 'serif']
      const serif = /batang|바탕|myeongjo|myungjo|명조|gungsuh|궁서/i.test(nfkc)
      return [serif ? KO_SERIF : KO_SANS, serif ? 'serif' : 'sans-serif']
    },
  },
  // DFKai-SB is missing on Word for Mac, which substitutes Microsoft YaHei wholesale (probe 2026-09-23); the
  // Latin alias keeps YaHei's wide digits where YaHei itself is missing. Not gated on availability: macOS
  // lists the name for Apple's on-demand BiauKai asset while drawing nothing with it
  {
    id: 'dfkai',
    when: (f) => /\u6a19\u6977|dfkai/.test(f),
    chain: () => [
      ['Microsoft YaHei', 'GenOffice YaHei Latin', 'PingFang SC', CJK_SANS],
      'sans-serif',
    ],
  },
  {
    id: 'traditional-chinese',
    when: (_f, _h, nfkc) => TC_NAME_RE.test(nfkc),
    chain: (_f, _h, nfkc, n) => {
      const serif =
        /mingliu|\u7d30\u660e|\u6a19\u6977|biaukai|dfkai|kaiu|songti|kaiti|\u5b8b/i.test(nfkc)
      const alias = serif ? n.aliasFace(n.font) : null
      return [
        [...(alias ? [alias] : []), ...(serif ? TC_SERIF : TC_SANS)],
        serif ? 'serif' : 'sans-serif',
      ]
    },
  },
  // Ethiopic faces (Nyala/Ebrima/Abyssinica) are missing on macOS; the size-adjusted Kefa alias re-centers Chromium's ~1.35x-wide Kefa fallback
  {
    id: 'ethiopic',
    when: (_f, _h, nfkc) => /nyala|ebrima|abyssinica|ethiopic/i.test(nfkc),
    chain: () => [['GenOffice Ethiopic'], 'sans-serif'],
  },
  // Bengali: Nirmala UI is an M365 cloud face Word renders real; macOS has no twin and Chromium's Kohinoor Bangla shapes ~8% narrower, so the bundled Nirmala-metric face leads
  {
    id: 'nirmala',
    when: (_f, _h, nfkc) => /nirmala/i.test(nfkc),
    chain: () => [['GenOffice Bengali', 'Kohinoor Bangla', 'Bangla Sangam MN'], 'sans-serif'],
  },
  // Tamil: Latha renders through the bundled Latha-metric face; Vijaya and missing Tamil names lay out as Vijaya on Word for Mac (probe 2026-09-30), whose advances InaiMathi matches within 1 %
  {
    id: 'latha',
    when: (_f, _h, nfkc) => /latha/i.test(nfkc),
    chain: () => [['GenOffice Tamil', 'InaiMathi', 'Tamil MN', 'Tamil Sangam MN'], 'sans-serif'],
  },
  {
    id: 'tamil',
    when: (_f, _h, nfkc) => TAMIL_FONT_RE.test(nfkc),
    chain: () => [['Vijaya', 'InaiMathi', 'Tamil Sangam MN', 'Tamil MN'], 'sans-serif'],
  },
]

/** the chain of the first matching rule (declared name first, families, generic), or null */
export function namedFaceChain(n: NameContext): { id: string; css: string } | null {
  const f = n.font.toLowerCase()
  const head = f.split(',')[0].replace(/['"]/g, '').trim()
  const nfkc = n.font.normalize('NFKC')
  for (const rule of FACE_CHAIN_RULES) {
    if (!rule.when(f, head, nfkc, n)) continue
    const out = rule.chain(f, head, nfkc, n)
    if (typeof out === 'string') return { id: rule.id, css: out }
    const [fams, generic] = out
    return { id: rule.id, css: `${cssChain(n.font, ...fams)},${generic}` }
  }
  return null
}

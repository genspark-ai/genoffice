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

import type { DocGrid, Run } from '@genoffice/docx-engine'

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

/** Korean font names (Windows/Noto/Source Han/Nanum faces + bundled subsets) */
const KO_FONT_RE =
  /malgun|맑은|batang|바탕|myeongjo|myungjo|명조|gungsuh|궁서|gulim|굴림|dotum|돋움|nanum|나눔|genoffice (sans|serif) kr|(noto|source han) (sans|serif)[^,]*\bk(r|orean)?\b/i

/**
 * Per-font single-spacing line-height factor. Re-baselined against Word for
 * Mac (batch probe 2026-08-13: docx per family/size converted by installed
 * Word, baseline pitch measured from the PDF). Latin faces matched the old
 * LO-era values exactly; the CJK entries moved a lot because Word renders
 * with its own bundled faces (SimSun/Batang/PMingLiU at 1.3029, Malgun
 * 1.7371, Meiryo 1.9429, YaHei 1.7143, Yu 1.44) instead of the macOS
 * substitutes LO measured. Unprobed names keep their LO values until probed.
 */
/** Word's own symbol faces (Office DFonts symbol.ttf 5.00i, Wingdings 5.03,
 *  Webdings 5.00), hhea = win totals over 2048: Symbol 2059+450, Wingdings
 *  1841+432, Wingdings 2 1727+432, Wingdings 3 1900+432, Webdings 1638+410.
 *  Word row pitch 16.33px @10pt on a regression sample = the
 *  Symbol run (the macOS Symbol.ttf, 1436+612, is not what Word lays out with) */
const SYMBOL_FACE_FACTORS: Record<string, number> = {
  symbol: 1.2251,
  wingdings: 1.1099,
  'wingdings 2': 1.0542,
  'wingdings 3': 1.1387,
  webdings: 1.0,
}

/** Segoe UI text family (not the Emoji/Symbol/Historic/Print/Script cuts) */
const SEGOE_UI_TEXT_RE = /^segoe ui( (light|semilight|semibold|black|variable))?$/

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
}

export function lineHeightFactor(fontFamily: string): number {
  const f = fontFamily.toLowerCase()
  // an embedded face renders real: Word lays it with its own hhea total
  // (probe 2026-09-23, typo/win groups and USE_TYPO_METRICS ignored)
  const embedded = embeddedLineBoxes.get(f.normalize('NFKC').trim())
  if (embedded) return embedded.factor
  const symbolFace = SYMBOL_FACE_FACTORS[f.trim()]
  if (symbolFace !== undefined) return symbolFace
  // Aptos (M365 cloud face): Word probe 2026-08-22 measured 1.22, same as Calibri
  if (f.includes('aptos')) return 1.22
  // PMingLiU/MingLiU (Word probe: Office ships the real face at 1.3029)
  if (f.includes('pmingliu') || f.includes('mingliu') || f.includes('細明體')) {
    return 1.3029
  }
  // NanumGothic: mac Word lays out with the OS downloadable asset — factor =
  // its hhea total (1.15em) × Word's 1.3 EA multiplier (the same rule
  // reproduces Batang 1.3029, Malgun 1.7371, Meiryo 1.9429, Yu 1.44);
  // ink-band measured 20.24pt at 10pt/1.35x on regression sample run7/19
  if (krBatangAltNamed(fontFamily)) return 1.3029
  if (/nanum ?gothic|나눔 ?고딕/.test(f)) return 1.495
  // NanumMyeongjo: the OS downloadable asset renders real (Word probe
  // 2026-08-24: 18pt @12pt and 53.9pt @36pt = 1.5, its hhea 1.154 x 1.3);
  // NanumBarunGothic instead substitutes to Batang (probe 1.3042) and keeps
  // the KO_FONT_RE value below
  if (/nanum ?myeongjo|나눔 ?명조/.test(f)) return 1.5
  // Korean faces (Word probe 2026-08-13: Malgun 1.7371, Batang class 1.3029)
  if (KO_FONT_RE.test(f)) {
    return /malgun|맑은/.test(f) || krSansVariantSubstituted(fontFamily) ? 1.7371 : 1.3029
  }
  // Japanese faces: MS (P)Mincho/Gothic substitute into the Hiragino class
  // (1.7); Meiryo 1.775. Yu Mincho/Gothic: Word probe 2026-08-13 measured
  // exactly 1.44 at 10.5/12pt with Office's own yumin.ttf (the old 2.2667 was
  // an LO-baseline value and doubled every empty line under an 18pt docGrid)
  if (/游|yu (gothic|mincho)|yugoth|yumin/.test(f)) return 1.44
  // Meiryo UI is the compact UI cut, far below Meiryo (Word probe 2026-08-22)
  if (/(meiryo|メイリオ) ?ui/.test(f)) return 1.65
  if (/meiryo|メイリオ/.test(f)) return 1.9429
  // BIZ UD: Word for Mac lacks the UDP cuts and substitutes Yu Gothic wholesale
  // (probe 2026-09-03: 1.44 at 10.5/11/12pt); an installed face renders real
  if (f.includes('biz ud')) return bizUdSubstituted(fontFamily) ? 1.44 : 1.3029
  if (hgSubstituted(fontFamily)) return 1.44
  if (/mincho|明朝|ゴシック|ms (ui )?p?gothic|hiragino|osaka|kozuka|小塚/.test(f.normalize('NFKC')))
    return 1.3029
  // missing Noto/Source Han SC: Word substitutes SimSun at 1.3029 (probe
  // 2026-08-13; bare 'Noto Sans SC' presumed same substitution)
  if (/^noto sans sc$/.test(f)) return 1.3029
  if (/^(noto|source han) (sans|serif)( cjk)? ?(sc|cn|tc|tw|hk)\b/.test(f)) return 1.3029
  // Songti class (installed Songti SC / STSong)
  if (f.includes('simsun') || f.includes('nsimsun') || f.includes('宋体')) return 1.3029
  // FangSong renders in the SimSun class, not PingFang (Word probe 2026-08-22)
  if (f.includes('仿宋') || f.includes('fangsong') || f.includes('simfang')) return 1.3029
  if (/(^|\s)(songti|stsong)\b/.test(f)) return 1.7
  // STZhongsong ships with Office and renders real (probe 2026-08-23: hhea x 1.3)
  if (/zhongsong|中宋/.test(f)) return 1.725
  // STXihei (HuaWen XiHei) is a macOS system face Word renders real (probe
  // 2026-08-23: 21.5pt @12pt, 64.5pt @36pt); other vendors' thin-hei cuts
  // are missing faces that substitute into the SimSun class instead
  if (/stxihei|华文细黑/.test(f)) return 1.79
  // FZ XiaoBiaoSong and KaiTi GB2312/GBK: Word for Mac lacks them and substitutes
  // Microsoft YaHei wholesale (probe 2026-08-23, gov-doc sample confirmed)
  if (/xiaobiaosong|小标宋/.test(f)) return 1.7143
  if ((f.includes('楷体') || f.includes('kaiti')) && /gb2312|gbk/.test(f)) return 1.7143
  // SimHei ships with Office; Word renders it at the SimSun-class pitch
  // (probe 2026-08-23: 15.6pt @12pt — the old 1.0 was a macOS-Heiti LO value)
  if (f.includes('黑体') || f.includes('simhei')) return 1.3029
  // KaiTi ships with Office for Mac and renders real at the SimSun-class pitch
  // (probe 2026-08-23: 15.6pt @12pt and 46.7pt @36pt, Latin runs the same —
  // the old 1.0 was a macOS-Kaiti LO value)
  if ((f.includes('楷体') || f.includes('kaiti')) && !/gb2312|_gbk|gbk/.test(f)) return 1.3029
  if (f.includes('microsoft yahei') || f.includes('microsoftyahei') || f.includes('雅黑'))
    return 1.7143
  // DengXian ships with Office for Mac and renders real (text probe
  // 2026-08-13 and empty-line probe 2026-08-25 both 16.32pt @12pt = 1.36;
  // under a 15.6pt grid 10.5pt fits one cell and 12pt takes two, matching
  // Word — the old 1.775 PingFang-class value doubled the 10.5pt rows)
  if (f.includes('dengxian') || f.includes('等线')) return 1.36
  // missing GB faces and other zh names substitute into the PingFang class
  if (
    f.includes('楷体') ||
    f.includes('kaiti') ||
    f.includes('宋') || // Song-family display faces and related variants
    f.includes('microsoft yahei') ||
    f.includes('microsoftyahei') ||
    f.includes('雅黑') ||
    f.includes('simkai')
  ) {
    return 1.775
  }
  // DFKai-SB is missing on Word for Mac, which lays it out as Microsoft YaHei
  // (probe 2026-09-23: 20.6pt @12pt)
  if (/\u6a19\u6977|dfkai/.test(f)) return 1.7143
  // Traditional Chinese sans/kai faces use the PingFang TC class.
  if (/jhenghei|正黑|標楷|biaukai|dfkai|kaiu/.test(f)) return 1.775
  // Traditional/Simplified Arabic are M365 cloud fonts Word downloads and
  // renders with their real metrics (Word probe 2026-08-22)
  if (f.includes('simplified arabic')) return 1.66
  if (f.includes('traditional arabic')) return 1.5
  // Arabic faces: Word for Mac substitutes missing naskh names with Times
  // New Roman (probe 2026-08-13); Iranian B/XB/IR faces (B Mitra, XB Zar…)
  // all substitute the same way (probe 2026-08-13, factor 1.14)
  if (
    /naskh|kufi|arabic|amiri|scheherazade|\b(b|xb|ir)[ -]?(mitra|nazanin|titr|lotus|zar|yekan|koodak|roya|badr|homa|traffic|compset)\b|irlotus/.test(
      f,
    )
  )
    return 1.1429
  // Western single-line factors follow each font's hhea metrics (LO probe):
  // a 4% surplus per line cascades into whole-paragraph pagination drift,
  // so the big Office faces get their real values.
  if (f.includes('times') || f.includes('liberation serif')) return 1.15
  // Georgia Pro (M365 cloud face) renders real with its typo metrics (Word
  // probe 2026-09-17: 9.72/10.8/11.76pt at 10/11/12pt)
  if (f.includes('georgia pro')) return 0.98
  if (f.includes('georgia')) return 1.1375
  // 1.172 = Cambria's Win metrics (1172/1000), probe 2026-08-23 measured
  // 1.1724 at 36pt; the unknown-name fallback below substitutes to Cambria
  // and must stay in lockstep
  if (f.includes('cambria') || f.includes('caladea')) return 1.172
  // Palatino Linotype ships with Office and Word renders it real (probe
  // 2026-08-23: 1.35 = its Win metrics); macOS Palatino is the same design
  // with compact metrics (probed 1.105)
  if (f.includes('palatino linotype')) return 1.35
  if (f.includes('palatino')) return 1.105
  // Nunito Sans is an Office cloud font Word downloads and renders real
  // (probe 2026-08-23: 1.365 = its Win metrics)
  if (f.includes('nunito')) return 1.365
  // Poppins is an M365 cloud font Word downloads and renders real (probe
  // 2026-09-01: exactly 1.500 at 10/12/16/28pt, regular and bold = hhea/typo)
  if (f.includes('poppins')) return 1.5
  // Microsoft New Tai Lue ships with Office (probe 2026-08-23: 1.31 = Win metrics)
  if (f.includes('new tai lue')) return 1.31
  // Helvetica renders real on Word for Mac (PDF embeds it) at 12.6pt @10.5pt
  // single spacing = 1.2 (corpus 2026-09-24); Helvetica Neue stays unprobed
  if (f.includes('helvetica neue')) return 1.0
  if (f.includes('helvetica')) return 1.2
  // Arial Unicode MS carries CJK glyphs with tall metrics (Word probe 2026-08-22)
  if (f.includes('arial unicode')) return 1.74
  if (f === 'arial' || f.startsWith('arial ') || f.includes('liberation sans')) return 1.15
  if (f.includes('calibri') || f.includes('carlito')) return 1.22
  if (f.includes('tahoma')) return 1.2083
  // DejaVu faces are missing and substitute to Verdana (probe 2026-08-23)
  if (f.includes('verdana') || f.includes('dejavu')) return 1.2167
  // consolas first: the mono css chain carries 'Courier New' as fallback
  if (f.includes('consolas')) return 1.1667
  if (f.includes('courier')) return 1.1333
  // Century Gothic / Century ship with Office and render real (probe 2026-08-23)
  if (f.includes('century gothic')) return 1.226
  if (f.includes('century') && !f.includes('gothic')) return 1.206
  // Book Antiqua ships with Office (probe 2026-08-23; the old 1.1 was LO-era)
  if (f.includes('book antiqua')) return 1.21
  // Segoe UI text cuts are M365 cloud fonts Word downloads and renders real:
  // hhea 2210/514/0 over 2048, Word PDF pitch 15.96pt @12pt (probe 2026-09-11);
  // the symbol/emoji and script cuts keep the old value
  if (SEGOE_UI_TEXT_RE.test(f)) return 1.3301
  if (f.includes('segoe')) return 1.15
  if (/nyala|ebrima|abyssinica|ethiopic/.test(f)) return 1.0514
  // Tamil faces: Word renders missing Noto Sans Tamil / Latha with Latha
  // metrics (probe 2026-08-13)
  if (/tamil|latha|vijaya|inaimathi/.test(f)) return 1.6686
  // Faces Word for Mac resolves for real — macOS-installed or Office cloud
  // fonts — keep their own metrics (probe 2026-08-23, prod-corpus sweep)
  if (f.includes('lato')) return 1.2
  if (f.includes('trebuchet')) return 1.1615
  // macOS Supplemental + Office face, hhea 2257/597/0 over 2048; Word pitch
  // 15.75pt @11pt on a regression sample (incl. 0.5pt row border)
  if (f.includes('comic sans')) return 1.3936
  if (f === 'inter' || f.startsWith('inter ')) return 1.21
  if (f.includes('ubuntu')) return 1.15
  if (f.includes('avenir')) return 1.369
  // Office DFonts face Word renders real: hhea 2098/727/0 over 2048 (Word
  // PDF pitch 16.55pt @12pt, 2026-09-24)
  if (f.includes('lucida handwriting')) return 1.3794
  if (f.includes('lucida console')) return 1.0
  if (f.includes('lucida sans')) return 1.179
  if (f.includes('bradley hand')) return 1.25
  if (f.includes('open sans')) return 1.365
  // Roboto and Montserrat are M365 cloud faces Word renders real at their hhea
  // totals (probe 2026-09-17: 11.76/12.96/14.16 and 12.24/13.44/14.64pt at 10/11/12pt)
  if (f.includes('roboto')) return 1.172
  if (f.includes('montserrat')) return 1.219
  // Merriweather (cloud face, hhea 984/-273/0; probe 2026-09-23: 15.12pt @12pt)
  if (f.includes('merriweather')) return 1.257
  if (f.includes('shonar bangla')) return 1.0
  // IRANYekan substitutes to Arial like the Iranian B/XB faces (probe 2026-08-23)
  if (f.includes('iranyekan')) return 1.15
  // missing hangul-named vendor faces substitute into the Batang/SimSun class
  // for KR and Latin runs alike (probe 2026-08-23: 휴먼고딕 → Batang/SimSun 1.30)
  if (/[가-힣]/.test(f)) return 1.3029
  // missing ideograph-named faces substitute SimSun (probe 2026-08-22/23)
  if (/[一-鿿]/.test(f)) return 1.3029
  // unknown Western names: Word for Mac substitutes Cambria (probe 2026-08-23:
  // fabricated names and HCI Poppy both render as Cambria) — same factor as
  // the cambria branch above
  return 1.172
}

/** Word DFonts win ascent (em) of the symbol bullet faces */
const SYMBOL_FACE_ASCENT: Record<string, number> = {
  symbol: 1.0054,
  wingdings: 0.8989,
  'wingdings 2': 0.8433,
  'wingdings 3': 0.9277,
  webdings: 0.7998,
}

/** Descent (em) Word lays the common Latin text faces with (hhea/win, whichever its factor follows) */
const LATIN_DESCENT: Array<[RegExp, number]> = [
  [/calibri|carlito/, 0.2686],
  [/aptos/, 0.2817],
  [/cambria|caladea/, 0.2222],
  [/times|liberation serif/, 0.2163],
  [/arial|helvetica|liberation sans/, 0.2119],
  [/georgia/, 0.2192],
  [/consolas/, 0.251],
  [/courier/, 0.3003],
  [/verdana/, 0.21],
  [/tahoma/, 0.2065],
  [/segoe ui/, 0.251],
  [/roboto/, 0.25],
  [/lato/, 0.213],
  [/montserrat/, 0.251],
  [/merriweather/, 0.273],
  [/open sans/, 0.293],
]

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
  const f = (textFont ?? '').toLowerCase()
  const desc =
    embeddedLineBoxes.get(f.normalize('NFKC').trim())?.descent ??
    LATIN_DESCENT.find(([re]) => re.test(f))?.[1] ??
    0.25
  const line = asc * markerPt + desc * textPt
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
export const NOTO_HAN_REGIONAL_RE =
  /^(?:noto|source han) (?:sans|serif)(?: cjk)? ?(jp|kr|sc|cn|tc|tw|hk)\b/

export function cjkDeclaredLineFactor(fontFamily: string): number | null {
  const krName = krNameLineFactor(fontFamily)
  if (krName !== null) return krName
  const m = NOTO_HAN_REGIONAL_RE.exec(fontFamily.toLowerCase())
  if (!m) return null
  if (m[1] === 'kr') return lineHeightFactor(fontFamily)
  return 1.3029
}

/**
 * Word probe 2026-08-13: SimSun (the 1.3029 substitute for missing JP/SC/TC
 * faces) lacks ・ U+30FB and 〜 U+301C; Word falls back to Microsoft YaHei for
 * them and lifts the whole line to 1.7143 × size (13.68pt → 18.0pt @10.5pt).
 * KR (Batang class, also 1.3029) is unprobed and excluded.
 */
export const SIMSUN_GAP_CHAR_RE = /[・〜]/
export const SIMSUN_GAP_LINE_FACTOR = 1.7143

export function simsunGapLineFactor(fontFamily: string): number | null {
  const f = fontFamily.toLowerCase()
  // Korean faces (named or hangul-lettered) substitute Batang-ward, not SimSun
  if (/[가-힣ᄀ-ᇿ㄰-㆏]/.test(f.normalize('NFKC')) || KO_FONT_RE.test(f.normalize('NFKC'))) {
    return null
  }
  if (cjkDeclaredLineFactor(fontFamily) === 1.3029) return SIMSUN_GAP_LINE_FACTOR
  if (f.includes('simsun') || f.includes('nsimsun') || f.includes('宋体')) {
    return SIMSUN_GAP_LINE_FACTOR
  }
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
  'GenOffice Tamil',
  'GenOffice Fullwidth TC',
  'GenOffice Songti SC',
  'GenOffice Songti TC',
  'Carlito GO',
  'Aptos GO',
  'Aptos Display GO',
  'Century Gothic GO',
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
export function fontChainMetricsPct(
  chain: string,
): { ascentPct: number; descentPct: number } | null {
  const cached = chainMetricsCache.get(chain)
  if (cached !== undefined) return cached
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
  chainMetricsCache.set(chain, out)
  return out
}

// ─── Document fontTable (word/fontTable.xml substitution hints) ─────────────

/** current document's fontTable, keyed by NFKC-lowercased name */
let docFontTable = new Map<string, { altName?: string; panose?: string }>()

export function setDocFontTable(
  entries: ReadonlyArray<{ name: string; altName?: string; panose?: string }> | null | undefined,
): void {
  docFontTable = new Map()
  for (const e of entries ?? []) {
    docFontTable.set(e.name.normalize('NFKC').toLowerCase(), {
      ...(e.altName ? { altName: e.altName } : {}),
      ...(e.panose ? { panose: e.panose } : {}),
    })
  }
}

/** fontTable altName of a locally missing font (Word substitutes it wholesale) */
function fontTableAltName(font: string): string | undefined {
  const key = font.normalize('NFKC').toLowerCase()
  const alt = docFontTable.get(key)?.altName
  if (!alt || alt.normalize('NFKC').toLowerCase() === key) return undefined
  if (!isBundledFont(font) && isFontAvailable(font)) return undefined
  return alt
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

/**
 * Line factor for hangul-lettered vendor names outside the known Korean set:
 * a missing face whose fontTable PANOSE says sans substitutes into the Malgun
 * class (Word probe 2026-08-22: 1.7371, not Batang's 1.3029); everything else
 * stays Batang-ward. Null for non-hangul or known Korean names.
 */
export function krNameLineFactor(font: string): number | null {
  const nfkc = font.normalize('NFKC')
  if (!/[가-힣ᄀ-ᇿ㄰-㆏]/.test(nfkc)) return null
  if (KO_FONT_RE.test(nfkc)) return null
  const missing = isBundledFont(font) || !isFontAvailable(font)
  return missing && panoseSerifHint(font) === 'sans' ? 1.7371 : 1.3029
}

/**
 * Missing Noto/Source Han KR variant whose fontTable PANOSE says sans: Word
 * substitutes Malgun Gothic wholesale (probe 2026-09-11: Noto Sans KR with
 * PANOSE 020B.. lays hangul and Latin in MalgunGothic at 1.7371), while the
 * same name without a fontTable hint stays Batang-ward (probe 2026-08-13).
 */
export function krSansVariantSubstituted(font: string): boolean {
  const head = font.split(',')[0].replace(/['"]/g, '').normalize('NFKC').trim()
  if (!/^(?:noto|source han) (?:sans|serif)(?: cjk)? ?(?:kr|k)\b/i.test(head)) return false
  return (isBundledFont(head) || !isFontAvailable(head)) && panoseSerifHint(head) === 'sans'
}

/** missing BIZ UD face: Word substitutes Yu Gothic (sans, 1.44) whatever the name says */
export function bizUdSubstituted(font: string): boolean {
  const head = font.split(',')[0].replace(/['"]/g, '').trim()
  if (!/biz ud/i.test(head)) return false
  return isBundledFont(head) || !isFontAvailable(head)
}

/**
 * Ricoh HG faces of the Windows JP Office bundle that Word for Mac does not
 * ship (its DFonts carry only Gothic E, Mincho E, Soei Kaku Gothic UB and Maru
 * Gothic M-PRO, rendered real at the MS-class pitch): Word substitutes Yu
 * Gothic wholesale, kana, kanji and Latin alike, whatever the name says (probe
 * 2026-09-16: 14.4pt @10pt = 1.44 for HG(P/S) Kyokashotai, HG Mincho B, HG
 * Gyoshotai, HG Sei-Kaishotai-PRO, HG Soei Kaku Pop-tai, HGS Soei Presence EB,
 * HG Gothic M; the JP-named Hiragino Kaku Gothic Pro W3 the same).
 */
const HG_SUBSTITUTED_RE = new RegExp(
  `^hg[ps]?(\u6559\u79d1\u66f8\u4f53|\u660e\u671db|\u884c\u66f8\u4f53|\u6b63\u6977\u66f8\u4f53|\u5275\u82f1\u89d2\u30dd\u30c3\u30d7\u4f53|\u5275\u82f1\u30d7\u30ec\u30bc\u30f3\u30b9|\u30b4\u30b7\u30c3\u30afm)|^\u30d2\u30e9\u30ae\u30ce\u89d2\u30b4`,
)

export function hgSubstituted(font: string): boolean {
  const head = font.split(',')[0].replace(/['"]/g, '').trim()
  if (!HG_SUBSTITUTED_RE.test(head.normalize('NFKC').toLowerCase())) return false
  return isBundledFont(head) || !isFontAvailable(head)
}

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

export function cssFontFamily(font: string, followAltName = true): string {
  const f = font.toLowerCase()
  const chain = (...families: string[]) =>
    [...new Set(families)].map((x) => `'${x.replace(/'/g, '')}'`).join(',')
  // CJK fallback at chain end (GB2312 subset bundled in fonts.css): no tofu even without system Chinese fonts
  const CJK_SERIF = 'Noto Serif CJK SC'
  const CJK_SANS = 'Noto Sans CJK SC'
  // 'GenOffice Box Drawing' (fonts.css): U+2500 rules would otherwise fall to
  // the fullwidth CJK subset and overflow the column
  const BOX = 'GenOffice Box Drawing'
  if (f.includes('calibri')) return `${chain(font, 'Carlito GO', CJK_SANS)},sans-serif`
  // Aptos (M365 cloud face, never installed locally): line metrics equal
  // Calibri's (Word probe 2026-08-22) but its advances do not, so the
  // size-adjusted Carlito aliases in fonts.css stand in (probe 2026-09-03).
  // 'GenOffice PUA Blank' keeps AI-residue PUA tokens invisible like Word
  // (Carlito would otherwise supply a box .notdef for them).
  if (f.includes('aptos')) {
    // the aliases are calibrated for the body and Display cuts only; Aptos
    // Narrow measures ~Carlito unscaled (0.92x Aptos), Mono/Serif are unprobed
    const alias = /narrow|mono|serif/.test(f)
      ? 'Carlito GO'
      : f.includes('display')
        ? 'Aptos Display GO'
        : 'Aptos GO'
    return `${chain(font, alias, 'GenOffice PUA Blank', CJK_SANS)},sans-serif`
  }
  // math faces would fall to the unknown-name sans fallback; STIX Two Math ships with macOS,
  // and on Windows the declared name resolves natively
  if ((f.includes('cambria') && f.includes('math')) || /stix.*math|latin modern math/.test(f))
    return `${chain(font, 'STIX Two Math', 'Caladea', 'Liberation Serif', BOX, CJK_SERIF)},serif`
  if (f.includes('cambria')) return `${chain(font, 'Caladea', BOX, CJK_SERIF)},serif`
  if (f.includes('times')) return `${chain(font, 'Liberation Serif', BOX, CJK_SERIF)},serif`
  // Palatino Linotype/Book Antiqua ship with Office and Word renders them real
  // (probe 2026-08-23); macOS Palatino matches their Latin widths within 0.2%
  if (f.includes('palatino') || f.includes('book antiqua'))
    return `${chain(font, 'Palatino Linotype', 'Palatino', 'Book Antiqua', BOX, CJK_SERIF)},serif`
  // a missing "Arial MT Black"-style name with a fontTable altName substitutes
  // like any unknown face (Word draws its Courier New alias)
  if ((f === 'arial' || f.startsWith('arial ')) && !(followAltName && fontTableAltName(font)))
    return `${chain(font, 'Liberation Sans', CJK_SANS)},sans-serif`
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
  // FZ XiaoBiaoSong and KaiTi GB2312/GBK: Word for Mac lacks them and substitutes
  // Microsoft YaHei wholesale (probe 2026-08-23, gov-doc sample confirmed)
  if (
    f.includes('小标宋') ||
    f.includes('xiaobiaosong') ||
    ((f.includes('楷体') || f.includes('kaiti')) && /gb2312|gbk/.test(f))
  )
    return `${chain(font, 'Microsoft YaHei', 'PingFang SC', CJK_SANS)},sans-serif`
  // 'GenOffice Songti SC' (fonts.css local() alias of Songti SC): macOS Chromium
  // refuses synthetic bold for 'Songti SC' by name at weight 600/700; the alias,
  // registered weight-normal only, lets Blink synthesize. Unresolvable elsewhere.
  // SimSun's ASCII is half-width fixed pitch (0.5em); the Songti stand-in is proportional
  if (f.includes('simsun') || f.includes('宋体') || f.includes('nsimsun')) {
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
  const JA_SANS = ['Yu Gothic', 'GenOffice Hiragino Sans', 'Meiryo', 'Noto Sans JP']
  const JA_SERIF = [
    'Yu Mincho',
    'GenOffice Hiragino Mincho',
    'GenOffice MS Mincho',
    'Noto Serif JP',
  ]
  const KO_SANS = ['Malgun Gothic', 'GenOffice Sans KR', 'Apple SD Gothic Neo', 'Noto Sans KR']
  const KO_SERIF = ['GenOffice Batang', 'GenOffice Serif KR', 'GenOffice Myungjo', 'Noto Serif KR']
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
  if (
    /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/.test(font) ||
    // Iranian B/XB/IR faces (B Mitra, B Nazanin, XB Zar…): Word for Mac
    // substitutes all of them with Times New Roman (probe 2026-08-13), so
    // they take the naskh serif chain
    /naskh|kufi|arabic|urdu|geeza|amiri|scheherazade|lateef|harmattan|aldhabi|andalus|nastaliq|al bayan|baghdad|damascus|diwan|farisi|mishafi|nadeem|beirut|\b(?:b|xb|ir)[ -]?(?:mitra|nazanin|titr|lotus|zar|yekan|koodak|roya|badr|homa|traffic|compset)\b|irlotus/i.test(
      nfkc,
    )
  ) {
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
    // declared Noto Arabic names would resolve to the bundled unscaled subsets
    // (~13% wider than Word); the literal head must go so the 'W' alias wins
    const scaled = /^noto (naskh|sans) arabic$/i.test(nfkc.trim())
    // Word substitutes missing naskh names wholesale with Times New Roman, so
    // ASCII punctuation/digits must take Times shapes, not the Naskh subset's
    // (sunken parens/hyphen). mac/Windows Times carries full Arabic, so on
    // those platforms Arabic letters also resolve here — same face Word uses;
    // the scaled subset below serves platforms without a system Times.
    // Calibrated aliases (compact/scaled) stay as-is.
    const missingSerif = !sans && !compact && !scaled && !isFontAvailable(font)
    const latinHead = missingSerif ? ['Times New Roman', 'Liberation Serif'] : []
    const chainFor = sans
      ? [scaled ? 'Noto Sans Arabic W' : 'Noto Sans Arabic', 'Geeza Pro']
      : [
          ...latinHead,
          // missing serif names: 90% TNR alias (subset shapes 1.125×/1.074×
          // wider than Word's Times, M3 probe); Arabic-Indic digits unscaled
          ...(missingSerif
            ? ['Naskh Digits GO', 'Noto Naskh Arabic TNR']
            : [
                compact
                  ? compact === 'traditional'
                    ? 'Noto Naskh Arabic TA'
                    : 'Noto Naskh Arabic SA'
                  : scaled
                    ? 'Noto Naskh Arabic W'
                    : 'Noto Naskh Arabic',
              ]),
          'Geeza Pro',
          'Al Bayan',
        ]
    return `${chain(...(scaled ? [] : [font]), ...chainFor)},${sans ? 'sans-serif' : 'serif'}`
  }
  // Word substitutes a *missing* East Asian font with the locale's default face —
  // a serif (Mincho/Batang) — regardless of the requested font's classification.
  // Classic Windows faces (Malgun/Meiryo/Yu Gothic...) have solid macOS
  // equivalents in the chains and keep their classification.
  const missingLocally = () => isBundledFont(font) || !isFontAvailable(font)
  // Noto CJK / Source Han / Nanum regional variants route by suffix; the generic
  // fallback tail below would otherwise land them on the bundled Simplified-only subset
  const cjkVariant = /^(?:noto|source han) (sans|serif)(?: cjk)? ?(jp|kr|k|tc|tw|hk|sc|cn)\b/i.exec(
    nfkc,
  )
  if (cjkVariant) {
    const region = cjkVariant[2].toLowerCase()
    const isKr = region === 'kr' || region === 'k'
    const sansSub = isKr && krSansVariantSubstituted(font)
    const serif = /serif/i.test(cjkVariant[1]) || (missingLocally() && !sansSub)
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
    // a bundled face at the chain head would win over the substitution; it stays only as the tail safety net
    const head = isBundledFont(font) ? [] : [font]
    // missing KR names: Word moves Latin/digits to the theme Latin face while
    // hangul stays Batang (M3 probe sample 13); the alias claims printable
    // ASCII only, everything else falls through per glyph. Installed names and
    // Malgun/Batang declares keep the Latin-normalized subsets untouched.
    const krLatin = isKr && missingLocally() && !sansSub ? ['KR Theme Latin GO'] : []
    // Word substitutes a missing face per script, whatever region the name
    // claims: hangul lands on Batang (1em, Word probe 2026-09-06). The
    // SC/TC/JP chains carry no hangul, so without this tail Chromium falls to
    // the system sans (Apple SD Gothic Neo, 0.865em) and lines wrap late
    const hangulTail = !isKr && missingLocally() ? ['GenOffice Batang', 'GenOffice Serif KR'] : []
    return `${chain(...head, ...krLatin, ...chainFor, ...hangulTail)},${serif ? 'serif' : 'sans-serif'}`
  }
  if (
    /[぀-ヿ]|mincho|meiryo|hiragino|osaka|yugoth|yu (gothic|mincho)|ms (ui )?p?(gothic|mincho)|明朝|biz ud|kozuka|小塚|^hg[ps]?(\u6559\u79d1\u66f8|\u884c\u66f8|\u6b63\u6977\u66f8|\u5275\u82f1)/i.test(
      nfkc,
    )
  ) {
    const serif =
      /mincho|\u660e\u671d/i.test(nfkc) && !bizUdSubstituted(font) && !hgSubstituted(font)
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
    return `${chain(font, ...msAlias, ...(serif ? JA_SERIF : JA_SANS))},${serif ? 'serif' : 'sans-serif'}`
  }
  if (
    /[가-힣ᄀ-ᇿ㄰-㆏]|malgun|batang|gulim|dotum|gungsuh|myeongjo|myungjo|nanum|apple (sd )?gothic/i.test(
      nfkc,
    )
  ) {
    // Nanum gothic-class declares: mac Word lays out with the OS downloadable
    // Nanum asset Chromium cannot see (M3 probe sample 19), so the bundled
    // real-metric subset leads the sans chain instead of the Batang-ward
    // substitution. Installed Nanum still resolves at the literal head.
    if (krBatangAltNamed(font)) return `${chain(...KO_SERIF)},serif`
    if (/nanum ?gothic|나눔 ?고딕/i.test(nfkc)) {
      return `${chain(font, 'GenOffice Gothic KR', ...KO_SANS)},sans-serif`
    }
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
    // other vendor faces missing locally follow Word's Batang-ward substitution
    // unless their fontTable PANOSE says sans (Word substitutes Malgun class);
    // Windows core faces (Malgun/Gulim/Dotum) map cleanly to the sans chain
    const knownCore = /malgun|맑은|gulim|굴림|dotum|돋움|apple (sd )?gothic/i.test(nfkc)
    const serif =
      /batang|바탕|myeongjo|myungjo|명조|gungsuh|궁서/i.test(nfkc) ||
      (!knownCore && missingLocally() && panoseSerifHint(font) !== 'sans')
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
  // Tamil: Word substitutes missing Tamil families with Latha; the bundled
  // Latha-metric face (fonts.css) keeps line breaks aligned. On Windows the
  // declared name resolves natively ahead of it; macOS system faces stay as
  // coverage tails (the subset ships no Latin letters).
  if (/tamil|latha|vijaya|inaimathi/i.test(nfkc)) {
    return `${chain(font, 'GenOffice Tamil', 'InaiMathi', 'Tamil MN', 'Tamil Sangam MN')},sans-serif`
  }
  // unknown missing font with a fontTable altName: Word substitutes the alias
  // wholesale, so the alias's whole chain follows the declared head. Hei-class
  // aliases are excluded: Word's SimHei-class Latin is half-width, while the
  // Heiti chain's macOS Latin is wide proportional — worse than the plain
  // fallback below, so those keep the name-guess path.
  const altName = followAltName ? fontTableAltName(font) : undefined
  if (altName && !/simhei|黑体|细黑|xihei/.test(altName.toLowerCase())) {
    return `${chain(font)},${cssFontFamily(altName, false)}`
  }
  // unknown font family: fontTable PANOSE decides serif-ness when the font is
  // missing, else guess by name (Song/Ming/Serif → serif fallback)
  const hint = missingLocally() ? panoseSerifHint(font) : undefined
  const serifLike = hint ? hint === 'serif' : /宋|明|serif|song|ming/i.test(font)
  // Latin-named unknown faces take the range-limited CJK tail: the bundled
  // subset covers U+2018/2019/201C/201D at 1em and would otherwise swallow
  // Latin punctuation. CJK-named faces keep the full subset (fullwidth quotes
  // in CJK body text are correct).
  const latinNamed = !/[\u2E80-\u9FFF\uAC00-\uD7A3\uF900-\uFAFF\uFF00-\uFFEF]/.test(nfkc)
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
  const notoIdx = head.findIndex((f) => /noto (naskh|sans) arabic/i.test(f))
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
 * Maximal CJK stretches of a text (UTF-16 offsets); spaces between two CJK
 * characters join them so a Korean phrase is one stretch, while any other
 * character ends it. lineFactorLive lifts each stretch to its East Asian factor.
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
    } else if (start >= 0 && ch !== ' ' && ch !== '\u00a0') {
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
  if (textHasHangul(text)) return 'var(--doc-line-factor-kr,1.3029)'
  if (textHasCjk(text)) return 'var(--doc-line-factor-cjk,1.7)'
  return 'var(--doc-line-factor-latin,1.2)'
}

/** Run[] port of extensions' latinParaFactor (declared run fonts override the doc factor). */
export function latinRunsFactor(runs: Run[], scriptVar: string): string {
  let declaredMax = 0
  let undeclared = false
  for (const run of runs) {
    // ascii slot only, like extensions' latinParaFactor: an eastAsia-only
    // declaration must not set a Latin line's factor
    const family = run.fontAscii
    if (family) declaredMax = Math.max(declaredMax, lineHeightFactor(family))
    else undeclared = true
  }
  if (declaredMax <= 0) return scriptVar
  return undeclared ? `max(${scriptVar}, ${declaredMax})` : String(declaredMax)
}

/** Per-paragraph --doc-line-factor from runs (Run[] port of extensions' paraLineFactor). */
export function runsLineFactor(runs: Run[], text: string): string {
  const scriptVar = paraLineFactorCss(text)
  if (!textHasCjk(text)) return latinRunsFactor(runs, scriptVar)
  let declaredMax = 0
  let undeclaredCjk = false
  for (const run of runs) {
    if (!textHasCjk(run.text)) continue
    const family = run.eaSlotEmpty === true ? null : (run.font ?? run.fontAscii)
    if (family && isCjkFontName(family)) {
      declaredMax = Math.max(declaredMax, cjkDeclaredLineFactor(family) ?? lineHeightFactor(family))
    } else undeclaredCjk = true
  }
  if (declaredMax <= 0) return scriptVar
  return undeclaredCjk ? `max(${scriptVar}, ${declaredMax})` : String(declaredMax)
}

/** --doc-line-factor-kr source: the document's East Asian face when Korean, else the Batang-class default */
export function krLineFactor(fontFamily: string | undefined): number {
  if (!fontFamily) return 1.3029
  const krName = krNameLineFactor(fontFamily)
  if (krName !== null) return krName
  return isKoreanFontName(fontFamily) ? lineHeightFactor(fontFamily) : 1.3029
}

export function isKoreanFontName(fontFamily: string): boolean {
  return KO_FONT_RE.test(fontFamily.normalize('NFKC'))
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
 *  single), mirroring computeLineHeight with the 1.7143em lift as natural
 *  height. --doc-grid-single-mult (1 in typed-grid docs, absent otherwise)
 *  keeps the non-grid value at boosted x mult. */
export function cssSimsunGapLineExpr(multiple: number): string {
  const single = `round(up, calc(${SIMSUN_GAP_LINE_FACTOR} * 1em - ${GRID_PITCH} * ${GRID_SNAP_EPS}), ${GRID_PITCH})`
  if (multiple === 1) return single
  return `max(calc(${GRID_PITCH} * ${multiple}), calc(${single} * var(--doc-grid-single-mult,${multiple})))`
}

/** single-spacing line height: snapped length in typed-grid docs, unitless factor otherwise */
export function cssGridLineBase(): string {
  return 'var(--doc-line-grid,var(--doc-line-factor,1.2))'
}

/**
 * Paragraph strut font-size from run-declared sizes. Word sizes each line by
 * the runs on it: mixed run sizes keep the shrink-only clamp (a grown strut
 * would inflate smaller-run lines), but a uniform run size IS every line's
 * size, so the strut grows to match — left at the inherited size, a larger
 * uniform run rides above the strut's half-leading and the union line box
 * outgrows its height by ~1px/line (under a docGrid that pushed prod-sas 003's
 * last two page-1 lines out).
 */
export function strutFontCss(strut: { halfPoints: number; uniform: boolean }): string[] {
  return [
    `--doc-strut:${strut.halfPoints / 2}pt`,
    strut.uniform ? 'font-size:var(--doc-strut)' : 'font-size:min(var(--doc-strut), 1em)',
  ]
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
  if (lineRule === 'exact' && lineRawTwips) return '0px'
  if (lineRule === 'atLeast' && lineRawTwips != null) {
    if (lineRawTwips === 0) return '0px'
    return `max(0px, (${(lineRawTwips / 20).toFixed(1)}pt - var(--doc-line-grid, calc(var(--doc-line-factor,1.2) * 1em))) / 2)`
  }
  return cssAutoLineMult(lineRule, lineRawTwips, lineSpacing) ? 'initial' : null
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
// native gap.

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

export function needsAutospacePad(prevCp: number, nextCp: number): boolean {
  return (
    (isCjkAutospaceSide(prevCp) && isLatinAlnum(nextCp)) ||
    (isLatinAlnum(prevCp) && isCjkAutospaceSide(nextCp))
  )
}

function lastCodePoint(text: string): number {
  const tail = text.charCodeAt(text.length - 1)
  if (tail >= 0xdc00 && tail <= 0xdfff && text.length > 1) {
    return text.codePointAt(text.length - 2)!
  }
  return tail
}

/** pad between two adjacent stretches of text (last char of prev vs first of next) */
export function autospacePadBetween(prev: string, next: string): boolean {
  if (!prev || !next) return false
  return needsAutospacePad(lastCodePoint(prev), next.codePointAt(0)!)
}

/** UTF-16 length of the code point starting at offset */
export function codePointLengthAt(text: string, offset: number): number {
  return text.codePointAt(offset)! > 0xffff ? 2 : 1
}

/** UTF-16 offsets inside text where a pad belongs (between offset-1 and offset) */
export function autospaceBoundaries(text: string): number[] {
  const out: number[] = []
  let prev = -1
  for (let i = 0; i < text.length;) {
    const cp = text.codePointAt(i)!
    if (prev >= 0 && needsAutospacePad(prev, cp)) out.push(i)
    prev = cp
    i += cp > 0xffff ? 2 : 1
  }
  return out
}

// ─── Hangul-context spaces ───────────────────────────────────────────────────
// Word lays out a space touching a hangul character at 0.5em of the run size
// whatever the font's own space advance (corpus probe 2026-09-24: Batang
// 0.49-0.50, Malgun Gothic and a Times New Roman space inside a Korean run
// 0.51-0.52), while spaces between Latin letters, digits or punctuation keep
// the glyph's advance (0.25-0.36em). One hangul neighbour is enough (hangul
// then Latin or Latin then hangul are wide; hangul, comma, space, Latin is not).
// The renderer wraps each such space in a .doc-hangul-space span whose only
// glyph comes from the half-width space face.

/** UTF-16 offsets of the U+0020 in text with a hangul neighbour; prev/next
 *  are the texts adjacent to its ends (empty = no direct neighbour) */
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
  /宋|黑|楷|仿|明|雅黑|等线|simsun|simhei|simkai|simfang|kaiti|fangsong|songti|stsong|yahei|dengxian|mingliu|jhenghei|biaukai|dfkai|kaiu|mincho|ms (ui )?p?gothic|yu gothic|ゴシック|meiryo|メイリオ|hiragino|\u30d2\u30e9\u30ae\u30ce|osaka|kozuka|biz ud|游|\u6559\u79d1\u66f8\u4f53|\u884c\u66f8\u4f53|\u5275\u82f1|arial unicode|(noto|source han) (sans|serif)( cjk)? ?(sc|cn|jp|tc|kr)\b/i

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
        }>
      }
    | null
    | undefined,
  contentWidthPx: number,
  /** images in the part (logos): non-floating ones count as one line of height (same as the display layer) */
  images?: Array<{
    heightPx?: number
    floating?: boolean
    behind?: boolean
    wrap?: 'none' | 'square' | 'tight' | 'through' | 'topBottom'
    posYPx?: number
    posVRel?: 'page' | 'margin' | 'paragraph' | 'topMargin' | 'bottomMargin'
  }> | null,
  /** header geometry (px): pass for headers so wrapped anchored images push the body below their bottom edge (Word); watermarks (wrapNone/behindDoc) still reserve nothing */
  geom?: { marginTopPx: number; headerDistPx: number },
): number {
  const inlineImages = (images ?? []).filter((im) => !im.floating && im.heightPx)
  const imagesHeight =
    inlineImages.length > 0 ? Math.max(...inlineImages.map((im) => im.heightPx!)) : 0
  let anchoredPx = 0
  if (geom) {
    for (const im of images ?? []) {
      if (!im.floating || im.behind || !im.heightPx) continue
      if (!im.wrap || im.wrap === 'none') continue
      if (im.posYPx == null || !im.posVRel || im.posVRel === 'bottomMargin') continue
      // bottom edge in page coordinates; the anchor paragraph sits at the header strip top
      const bottom =
        im.posVRel === 'page' || im.posVRel === 'topMargin'
          ? im.posYPx + im.heightPx
          : im.posVRel === 'margin'
            ? geom.marginTopPx + im.posYPx + im.heightPx
            : geom.headerDistPx + im.posYPx + im.heightPx
      // effectiveTopPx adds headerDist back: dist + this = the image bottom
      anchoredPx = Math.max(anchoredPx, bottom - geom.headerDistPx)
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
  for (const p of paras) {
    if (p.boxAnchored) continue
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
    height += lineH(runs, p, Math.max(1, contentWidthPx - indentPx))
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

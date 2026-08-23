/**
 * Cell-font fallback for the grid canvas (document data, theme-independent).
 * Univer draws/measures with a bare `italic bold 11pt "Family"` string, so an
 * unresolvable family drops to Chromium's default *serif* chain while Excel's
 * substitution is always sans. Two layers: a ctx.font setter patch appends
 * `sans-serif` as the last resort, and alias FontFaces map well-known Office
 * names to local faces — pinning serif-intent names (Mincho/Song/Ming/Batang/
 * Cambria…) to real serif faces so the sans last-resort never flips them.
 * Where no alias local() face exists at all (e.g. Linux), alias-known serif
 * names still get `serif` appended instead of `sans-serif`; unrecognized
 * names always fall to sans, matching Excel's substitution.
 */

const GENERIC_FAMILY =
  /(?:^|[\s,])(?:serif|sans-serif|monospace|cursive|fantasy|system-ui|math|ui-serif|ui-sans-serif|ui-monospace|ui-rounded)$/i

const SERIF_INTENT =
  /mincho|明朝|simsun|songti|宋体|宋體|mingliu|明體|細明|batang|바탕|myeongjo|명조|roman|georgia|garamond|cambria|constantia|palatino|antiqua|didot|bodoni|baskerville|caslon|goudy|bookman|(?<!sans[-\s])serif/i

/// Generic families never cover emoji code points on the canvas — without an
/// explicit color-emoji face at the end of every chain, U+274C & friends draw
/// as tofu. Listed after the generic so primary-font metrics never change.
const EMOJI_FALLBACK = '"Apple Color Emoji", "Segoe UI Emoji", "Noto Color Emoji"'

/// The family name inside one comma segment of a ctx.font string: the quoted
/// name if quoted, otherwise the last whitespace token (the first segment
/// carries the style/size prefix; Univer quotes any family with spaces).
function segmentFamily(segment: string): string {
  const quoted = /["']([^"']*)["']/.exec(segment)
  if (quoted) return quoted[1]!.trim()
  const tokens = segment.trim().split(/\s+/)
  return tokens[tokens.length - 1] ?? ''
}

/// Only alias-known names carry serif intent (first hit in the list wins).
/// Excel substitutes its sans default for any name it cannot resolve — even
/// myeongjo/mincho-keyworded ones ("휴먼명조,한컴돋움" renders sans in the
/// Excel reference for prod_059, and Univer's per-glyph fallback re-probes
/// each member alone, so the single-name path must agree). Alias-known serif
/// names keep serif so the last resort never flips them where the alias
/// local() faces are absent (e.g. Linux).
function serifIntent(trimmed: string): boolean {
  for (const segment of trimmed.split(',')) {
    const family = segmentFamily(segment)
    if (ALIAS_FAMILY_NAMES.has(family.toLowerCase())) return SERIF_INTENT.test(family)
  }
  return false
}

export function withSansSerifFallback(font: string): string {
  const trimmed = font.trimEnd()
  if (!trimmed) return font
  if (trimmed.endsWith(EMOJI_FALLBACK)) return font
  if (GENERIC_FAMILY.test(trimmed)) return `${trimmed}, ${EMOJI_FALLBACK}`
  return `${trimmed}, ${serifIntent(trimmed) ? 'serif' : 'sans-serif'}, ${EMOJI_FALLBACK}`
}

export interface CellFontAlias {
  readonly family: string
  /// local() face names, tried in order; genuine (Windows) names first so the
  /// alias is a no-op where the real font exists.
  readonly regular: readonly string[]
  /// Real bold faces only — never a regular face, which would suppress
  /// synthetic bold where no true bold exists.
  readonly bold?: readonly string[]
}

const JP_SANS = ['Hiragino Sans', 'HiraginoSans-W3', 'Hiragino Kaku Gothic ProN'] as const
const JP_SANS_BOLD = ['HiraginoSans-W6', 'Hiragino Sans W6'] as const
const JP_SERIF = ['Hiragino Mincho ProN', 'HiraMinProN-W3'] as const
const JP_SERIF_BOLD = ['HiraMinProN-W6', 'Hiragino Mincho ProN W6'] as const
const YU_GOTHIC = ['Yu Gothic Regular', 'YuGothic-Regular', 'YuGothic Medium', 'YuGo-Medium']
const YU_GOTHIC_BOLD = ['Yu Gothic Bold', 'YuGothic-Bold', 'YuGothic Bold', 'YuGo-Bold']
const YU_GOTHIC_ALL_BOLD = [...YU_GOTHIC_BOLD, ...JP_SANS_BOLD]
const YU_MINCHO = ['Yu Mincho Regular', 'YuMincho-Regular', 'YuMincho Medium', 'YuMin-Medium']
const YU_MINCHO_BOLD = ['Yu Mincho Demibold', 'YuMincho-Demibold', 'YuMin-Demibold']
const YU_MINCHO_ALL_BOLD = [...YU_MINCHO_BOLD, ...JP_SERIF_BOLD]
const SONG = ['SimSun', 'Songti SC', 'STSongti-SC-Regular']
const SONG_BOLD = ['STSongti-SC-Bold', 'Songti SC Bold']
const KAI = ['KaiTi', 'Kaiti SC', 'STKaitiSC-Regular', 'STKaiti']
const MING_TC = ['PMingLiU', 'Songti TC', 'Apple LiSung']
const KR_SANS = ['Malgun Gothic', 'Apple SD Gothic Neo', 'AppleGothic']
const KR_SANS_BOLD = ['Malgun Gothic Bold', 'AppleSDGothicNeo-Bold', 'Apple SD Gothic Neo Bold']
const KR_SERIF = ['Batang', 'AppleMyungjo', 'Nanum Myeongjo']
const TIMES_BOLD = ['Times New Roman Bold', 'TimesNewRomanPS-BoldMT']

function jpGothic(family: string, genuine: readonly string[]): CellFontAlias {
  return { family, regular: [...genuine, ...JP_SANS], bold: [...JP_SANS_BOLD] }
}

function jpMincho(family: string, genuine: readonly string[]): CellFontAlias {
  return { family, regular: [...genuine, ...JP_SERIF], bold: [...JP_SERIF_BOLD] }
}

export const CELL_FONT_ALIASES: readonly CellFontAlias[] = [
  // JP gothic (sans intent), incl. fullwidth spellings
  jpGothic('ＭＳ Ｐゴシック', ['MS PGothic']),
  jpGothic('ＭＳ ゴシック', ['MS Gothic']),
  jpGothic('MS PGothic', ['MS PGothic']),
  jpGothic('MS Gothic', ['MS Gothic']),
  jpGothic('MS UI Gothic', ['MS UI Gothic']),
  { family: 'メイリオ', regular: ['Meiryo', ...JP_SANS], bold: ['Meiryo Bold', ...JP_SANS_BOLD] },
  { family: 'Meiryo', regular: ['Meiryo', ...JP_SANS], bold: ['Meiryo Bold', ...JP_SANS_BOLD] },
  {
    family: 'Meiryo UI',
    regular: ['Meiryo UI', ...JP_SANS],
    bold: ['Meiryo UI Bold', ...JP_SANS_BOLD],
  },
  { family: '游ゴシック', regular: [...YU_GOTHIC, ...JP_SANS], bold: YU_GOTHIC_ALL_BOLD },
  { family: '游ゴシック体', regular: [...YU_GOTHIC, ...JP_SANS], bold: YU_GOTHIC_ALL_BOLD },
  { family: 'Yu Gothic', regular: [...YU_GOTHIC, ...JP_SANS], bold: YU_GOTHIC_ALL_BOLD },
  {
    family: 'Yu Gothic UI',
    regular: ['Yu Gothic UI', ...YU_GOTHIC, ...JP_SANS],
    bold: ['Yu Gothic UI Bold', ...YU_GOTHIC_ALL_BOLD],
  },
  // JP mincho (serif intent — keep serif under the sans last-resort)
  jpMincho('ＭＳ 明朝', ['MS Mincho']),
  jpMincho('ＭＳ Ｐ明朝', ['MS PMincho']),
  jpMincho('MS Mincho', ['MS Mincho']),
  jpMincho('MS PMincho', ['MS PMincho']),
  { family: '游明朝', regular: [...YU_MINCHO, ...JP_SERIF], bold: YU_MINCHO_ALL_BOLD },
  { family: 'Yu Mincho', regular: [...YU_MINCHO, ...JP_SERIF], bold: YU_MINCHO_ALL_BOLD },
  // Simplified CJK
  { family: '宋体', regular: SONG, bold: SONG_BOLD },
  { family: 'SimSun', regular: SONG, bold: SONG_BOLD },
  { family: 'NSimSun', regular: SONG, bold: SONG_BOLD },
  { family: '黑体', regular: ['SimHei', 'Heiti SC', 'PingFang SC'] },
  { family: 'SimHei', regular: ['SimHei', 'Heiti SC', 'PingFang SC'] },
  { family: '仿宋', regular: ['FangSong', 'STFangsong'] },
  { family: 'FangSong', regular: ['FangSong', 'STFangsong'] },
  { family: '楷体', regular: KAI },
  { family: 'KaiTi', regular: KAI },
  { family: '楷体_GB2312', regular: KAI },
  // Traditional CJK
  { family: '新細明體', regular: MING_TC },
  { family: 'PMingLiU', regular: MING_TC },
  { family: '細明體', regular: ['MingLiU', ...MING_TC] },
  { family: 'MingLiU', regular: ['MingLiU', ...MING_TC] },
  { family: '標楷體', regular: ['DFKai-SB', 'BiauKai', 'Kaiti TC', 'Kaiti SC'] },
  { family: 'DFKai-SB', regular: ['DFKai-SB', 'BiauKai', 'Kaiti TC', 'Kaiti SC'] },
  // Korean
  { family: 'Malgun Gothic', regular: KR_SANS, bold: KR_SANS_BOLD },
  { family: 'Gulim', regular: ['Gulim', 'Dotum', ...KR_SANS] },
  { family: '굴림', regular: ['Gulim', 'Dotum', ...KR_SANS] },
  { family: 'Dotum', regular: ['Dotum', ...KR_SANS] },
  { family: 'Batang', regular: KR_SERIF },
  { family: '바탕', regular: KR_SERIF },
  { family: 'Gungsuh', regular: ['Gungsuh', 'AppleMyungjo'] },
  { family: '궁서', regular: ['Gungsuh', 'AppleMyungjo'] },
  // Office Latin serif faces absent on macOS (serif intent)
  {
    family: 'Cambria',
    regular: ['Cambria', 'Times New Roman', 'Georgia'],
    bold: ['Cambria Bold', ...TIMES_BOLD],
  },
  { family: 'Constantia', regular: ['Constantia', 'Georgia', 'Times New Roman'] },
  {
    family: 'Garamond',
    regular: ['Garamond', 'Times New Roman'],
    bold: ['Garamond Bold', ...TIMES_BOLD],
  },
  { family: 'Palatino Linotype', regular: ['Palatino Linotype', 'Palatino', 'Book Antiqua'] },
  { family: 'Book Antiqua', regular: ['Book Antiqua', 'Palatino'] },
  {
    family: 'Times New Roman',
    regular: ['Times New Roman', 'Times', 'Georgia'],
    bold: TIMES_BOLD,
  },
  { family: 'PT Serif', regular: ['PT Serif', 'Times New Roman', 'Georgia'] },
]

const ALIAS_FAMILY_NAMES: ReadonlySet<string> = new Set(
  CELL_FONT_ALIASES.map((alias) => alias.family.toLowerCase()),
)

function patchFontSetter(proto: object): void {
  const desc = Object.getOwnPropertyDescriptor(proto, 'font')
  if (!desc?.set || !desc.configurable) return
  const nativeSet = desc.set
  Object.defineProperty(proto, 'font', {
    ...desc,
    set(value: string) {
      nativeSet.call(this, withSansSerifFallback(String(value)))
    },
  })
}

export function installCanvasFontFallback(): void {
  patchFontSetter(CanvasRenderingContext2D.prototype)
  if (typeof OffscreenCanvasRenderingContext2D !== 'undefined')
    patchFontSetter(OffscreenCanvasRenderingContext2D.prototype)
}

/// Register + load alias faces before the first skeleton: canvas fillText
/// never triggers lazy @font-face loads, and Univer measures only once.
export function registerCellFontAliases(): Promise<unknown> {
  const loads: Promise<unknown>[] = []
  for (const alias of CELL_FONT_ALIASES) {
    const variants: [readonly string[], string][] = [[alias.regular, '400']]
    if (alias.bold) variants.push([alias.bold, '700'])
    for (const [locals, weight] of variants) {
      try {
        const face = new FontFace(alias.family, locals.map((n) => `local('${n}')`).join(', '), {
          weight,
        })
        document.fonts.add(face)
        loads.push(face.load().catch(() => {}))
      } catch {
        /* invalid descriptor on this platform — keep the rest */
      }
    }
  }
  return Promise.all(loads)
}

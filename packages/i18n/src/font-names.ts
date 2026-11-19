/**
 * CJK font-family spellings per script — the locale dictionary behind the
 * suite-wide font pickers (@genoffice/ui/src/font-list.ts consumes it).
 *
 * The localized spellings (Han, Kana, Hangul) are functional data, not UI copy:
 * they are the family names the OS font enumeration reports on CJK-locale
 * systems and the names CJK documents expect in their font slots. They live in
 * this i18n package — the one place the repo keeps locale strings — while the
 * picker logic that consumes them stays in @genoffice/ui.
 */

// GB/T 9704 official-document fonts included so government documents can be
// authored from scratch (values are free-typed either way; the pickers accept any name)
export const SIMPLIFIED_CJK_FAMILIES: readonly string[] = [
  '等线',
  '等线 Light',
  '宋体',
  '黑体',
  '微软雅黑',
  '楷体',
  '仿宋',
  'PingFang SC',
  'Songti SC',
  'Kaiti SC',
  '仿宋_GB2312',
  '楷体_GB2312',
  '方正小标宋简体',
  'Noto Sans SC',
  'Noto Serif SC',
]

export const JAPANESE_FAMILIES: readonly string[] = [
  'Yu Gothic',
  'Yu Mincho',
  'Meiryo',
  'Hiragino Sans',
  'Hiragino Mincho',
  'MS Gothic',
  'MS Mincho',
  'Noto Sans JP',
  'Noto Serif JP',
]

export const KOREAN_FAMILIES: readonly string[] = [
  'Malgun Gothic',
  'Apple SD Gothic Neo',
  'Batang',
  'Gulim',
  'Dotum',
  'Noto Sans KR',
  'Noto Serif KR',
]

export const TRADITIONAL_CJK_FAMILIES: readonly string[] = [
  'Microsoft JhengHei',
  'PMingLiU',
  'PingFang TC',
  'Noto Sans TC',
  'Noto Serif TC',
]

/**
 * Localized ⇄ English spellings of the same CJK family, one entry per family.
 *
 * Chromium reports family names in the system's language: a Windows outside the
 * CJK locales enumerates `SimSun`, not the localized spelling, while a
 * Japanese-locale Windows enumerates the kana spelling of Yu Gothic, not
 * `Yu Gothic`. Matching the literal candidate spelling alone would hide a font
 * the machine really has, so presence checks every spelling in the family's
 * entry, in either direction. Extend the table rather than adding copy-pasted
 * spellings to the lists above.
 */
export const CJK_FAMILY_ALIASES: Readonly<Record<string, readonly string[]>> = {
  // Simplified CJK (zh spellings)
  宋体: ['SimSun'],
  黑体: ['SimHei'],
  微软雅黑: ['Microsoft YaHei'],
  楷体: ['KaiTi'],
  仿宋: ['FangSong'],
  仿宋_GB2312: ['FangSong_GB2312'],
  楷体_GB2312: ['KaiTi_GB2312'],
  等线: ['DengXian'],
  '等线 Light': ['DengXian Light'],
  // Traditional CJK (zh-TW spellings)
  微软正黑体: ['Microsoft JhengHei', '微軟正黑體'],
  新細明體: ['PMingLiU'],
  標楷體: ['DFKai-SB'],
  // Japanese
  メイリオ: ['Meiryo'],
  游ゴシック: ['Yu Gothic'],
  游明朝: ['Yu Mincho'],
  'ＭＳ ゴシック': ['MS Gothic'],
  'ＭＳ 明朝': ['MS Mincho'],
  'Hiragino Mincho ProN': ['Hiragino Mincho', 'ヒラギノ明朝 ProN'],
  // Korean
  '맑은 고딕': ['Malgun Gothic'],
  바탕: ['Batang'],
  돋움: ['Dotum'],
  굴림: ['Gulim'],
}

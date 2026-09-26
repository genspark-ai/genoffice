import type { Lang } from '@genoffice/i18n'

/**
 * Font dropdown candidates grouped by script, ordered per UI language so the
 * fonts a user is most likely to want appear at the top (e.g. Western fonts
 * for the US market, Japanese fonts for the Japanese market). Shared by every
 * app's font pickers so the suite offers one consistent list. Windows and
 * macOS names for the same script sit side by side — pickers hide the ones
 * the running machine proves absent (see partitionFontFamilies), so a Windows
 * user never sees the macOS names and vice versa.
 */
const LATIN = [
  'Aptos',
  'Calibri',
  'Calibri Light',
  'Arial',
  'Times New Roman',
  'Georgia',
  'Verdana',
  'Tahoma',
  'Cambria',
  'Garamond',
  'Trebuchet MS',
  'Segoe UI',
  'Courier New',
  'Impact',
]
// GB/T 9704 official-document fonts included so government documents can be
// authored from scratch (values are free-typed either way; the pickers accept any name)
const SIMPLIFIED_CHINESE = [
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
const JAPANESE = [
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
const KOREAN = [
  'Malgun Gothic',
  'Apple SD Gothic Neo',
  'Batang',
  'Gulim',
  'Dotum',
  'Noto Sans KR',
  'Noto Serif KR',
]
const TRADITIONAL_CHINESE = [
  'Microsoft JhengHei',
  'PMingLiU',
  'PingFang TC',
  'Noto Sans TC',
  'Noto Serif TC',
]

export const BUILTIN_FONT_FAMILIES: readonly string[] = [
  ...LATIN,
  ...SIMPLIFIED_CHINESE,
  ...JAPANESE,
  ...KOREAN,
  ...TRADITIONAL_CHINESE,
]

export function fontFamiliesFor(lang: Lang): readonly string[] {
  switch (lang) {
    case 'zh':
      return [...SIMPLIFIED_CHINESE, ...LATIN, ...JAPANESE, ...KOREAN, ...TRADITIONAL_CHINESE]
    case 'zh-TW':
      return [...TRADITIONAL_CHINESE, ...LATIN, ...SIMPLIFIED_CHINESE, ...JAPANESE, ...KOREAN]
    case 'ja':
      return [...JAPANESE, ...LATIN, ...SIMPLIFIED_CHINESE, ...KOREAN, ...TRADITIONAL_CHINESE]
    case 'ko':
      return [...KOREAN, ...LATIN, ...JAPANESE, ...SIMPLIFIED_CHINESE, ...TRADITIONAL_CHINESE]
    default:
      return [...LATIN, ...JAPANESE, ...SIMPLIFIED_CHINESE, ...KOREAN, ...TRADITIONAL_CHINESE]
  }
}

/**
 * Localized ⇄ English spellings of the same family.
 *
 * Chromium reports family names in the system's language: a non-Chinese Windows
 * enumerates `SimSun`, not the localized Chinese name. Matching the literal spelling
 * alone would hide a font the machine really has, so a candidate counts as present when
 * *any* known spelling of it is reported. Extend the table rather than adding copy-pasted
 * spellings to the candidate lists.
 */
const FAMILY_ALIASES: Readonly<Record<string, readonly string[]>> = {
  // Simplified Chinese
  宋体: ['SimSun'],
  黑体: ['SimHei'],
  微软雅黑: ['Microsoft YaHei'],
  楷体: ['KaiTi'],
  仿宋: ['FangSong'],
  仿宋_GB2312: ['FangSong_GB2312'],
  楷体_GB2312: ['KaiTi_GB2312'],
  等线: ['DengXian'],
  等线_Light: ['DengXian Light'],
  // Traditional Chinese
  微软正黑体: ['Microsoft JhengHei'],
  新細明體: ['PMingLiU'],
  標楷體: ['DFKai-SB'],
  // Japanese
  メイリオ: ['Meiryo'],
  游ゴシック: ['Yu Gothic'],
  游明朝: ['Yu Mincho'],
  // Korean
  '맑은 고딕': ['Malgun Gothic'],
  바탕: ['Batang'],
  돋움: ['Dotum'],
  굴림: ['Gulim'],
}

/** Every spelling we know for a family, its own name first. */
function familySpellings(family: string): readonly string[] {
  const aliases = FAMILY_ALIASES[family]
  return aliases ? [family, ...aliases] : [family]
}

function isPresent(family: string, known: ReadonlySet<string>): boolean {
  return familySpellings(family).some((spelling) => known.has(spelling))
}

/**
 * Split a queryLocalFonts family list into picker sections: `builtin` holds the
 * candidates that exist on this machine, `system` holds the rest, both keeping the
 * caller's ordering. `knownAvailable` lists families the enumeration cannot see but
 * that are known installed (app-installed catalog fonts live in a private dir); they
 * keep their builtin slot without leaking into the system section.
 *
 * Presence is alias-aware (see FAMILY_ALIASES): a localized Chinese candidate counts as
 * present when the machine reports its English family name, so it is never dropped for a
 * spelling difference. When the enumeration is unavailable, denied, or still loading
 * (`known.size === 0`) every candidate is kept — offering a dead name beats hiding a
 * real one.
 */
export function partitionFontFamilies(
  candidates: readonly string[],
  systemFamilies: readonly string[],
  knownAvailable: readonly string[] = [],
): { builtin: readonly string[]; system: readonly string[] } {
  const known = new Set(systemFamilies)
  for (const f of knownAvailable) known.add(f)
  if (known.size === 0) return { builtin: [...candidates], system: [] }
  const candidateSet = new Set(candidates)
  return {
    builtin: candidates.filter((f) => isPresent(f, known)),
    system: systemFamilies.filter((f) => !candidateSet.has(f)),
  }
}

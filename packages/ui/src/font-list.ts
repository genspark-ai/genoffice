import type { Lang } from '@genoffice/i18n'
import {
  CJK_FAMILY_ALIASES,
  JAPANESE_FAMILIES,
  KOREAN_FAMILIES,
  SIMPLIFIED_CJK_FAMILIES,
  TRADITIONAL_CJK_FAMILIES,
} from '@genoffice/i18n'

/**
 * Font dropdown candidates grouped by script, ordered per UI language so the
 * fonts a user is most likely to want appear at the top (e.g. Western fonts
 * for the US market, Japanese fonts for the Japanese market). Shared by every
 * app's font pickers so the suite offers one consistent list. Windows and
 * macOS names for the same script sit side by side — pickers hide the ones
 * the running machine proves absent (see partitionFontFamilies), so a Windows
 * user never sees the macOS names and vice versa. The localized name data
 * lives in the i18n locale dictionary (packages/i18n/src/font-names.ts).
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

export const BUILTIN_FONT_FAMILIES: readonly string[] = [
  ...LATIN,
  ...SIMPLIFIED_CJK_FAMILIES,
  ...JAPANESE_FAMILIES,
  ...KOREAN_FAMILIES,
  ...TRADITIONAL_CJK_FAMILIES,
]

export function fontFamiliesFor(lang: Lang): readonly string[] {
  switch (lang) {
    case 'zh':
      return [
        ...SIMPLIFIED_CJK_FAMILIES,
        ...LATIN,
        ...JAPANESE_FAMILIES,
        ...KOREAN_FAMILIES,
        ...TRADITIONAL_CJK_FAMILIES,
      ]
    case 'zh-TW':
      return [
        ...TRADITIONAL_CJK_FAMILIES,
        ...LATIN,
        ...SIMPLIFIED_CJK_FAMILIES,
        ...JAPANESE_FAMILIES,
        ...KOREAN_FAMILIES,
      ]
    case 'ja':
      return [
        ...JAPANESE_FAMILIES,
        ...LATIN,
        ...SIMPLIFIED_CJK_FAMILIES,
        ...KOREAN_FAMILIES,
        ...TRADITIONAL_CJK_FAMILIES,
      ]
    case 'ko':
      return [
        ...KOREAN_FAMILIES,
        ...LATIN,
        ...JAPANESE_FAMILIES,
        ...SIMPLIFIED_CJK_FAMILIES,
        ...TRADITIONAL_CJK_FAMILIES,
      ]
    default:
      return [
        ...LATIN,
        ...JAPANESE_FAMILIES,
        ...SIMPLIFIED_CJK_FAMILIES,
        ...KOREAN_FAMILIES,
        ...TRADITIONAL_CJK_FAMILIES,
      ]
  }
}

/**
 * candidate spelling → every other known spelling of the same family, built
 * once from CJK_FAMILY_ALIASES: each table entry is one family's spelling set,
 * so within an entry every spelling maps to every other (a candidate under its
 * English name must also match the localized spelling, and vice versa).
 */
const FAMILY_SPELLINGS: ReadonlyMap<string, readonly string[]> = (() => {
  const map = new Map<string, string[]>()
  for (const [key, aliases] of Object.entries(CJK_FAMILY_ALIASES)) {
    const group = [key, ...aliases]
    for (const a of group) {
      const list = map.get(a)
      if (list) {
        for (const b of group) if (b !== a && !list.includes(b)) list.push(b)
      } else {
        map.set(
          a,
          group.filter((b) => b !== a),
        )
      }
    }
  }
  return map
})()

/** Every spelling we know for a family, its own name first. */
function familySpellings(family: string): readonly string[] {
  return [family, ...(FAMILY_SPELLINGS.get(family) ?? [])]
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
 * Presence is spelling-aware (see CJK_FAMILY_ALIASES in the i18n locale dictionary):
 * a candidate counts as present when *any* known spelling of it is reported, in
 * either direction — a localized zh candidate on a machine that reports `SimSun`,
 * or the English-named ja candidates on a Japanese Windows that reports the
 * localized name — so a candidate is never dropped for a spelling difference.
 * When the enumeration is unavailable, denied, or still loading (`known.size === 0`)
 * every candidate is kept — offering a dead name beats hiding a real one.
 */
export function partitionFontFamilies(
  candidates: readonly string[],
  systemFamilies: readonly string[],
  knownAvailable: readonly string[] = [],
): { builtin: readonly string[]; system: readonly string[] } {
  const known = new Set(systemFamilies)
  for (const f of knownAvailable) known.add(f)
  if (known.size === 0) return { builtin: [...candidates], system: [] }
  return {
    builtin: candidates.filter((f) => isPresent(f, known)),
    system: systemFamiliesBesidesCandidates(candidates, systemFamilies),
  }
}

/**
 * The machine-reported families a picker should list beside the candidates:
 * everything the enumeration returned that no candidate already covers by its
 * own name. Pickers that keep the full candidate list visible (the docs ribbon
 * and Font dialog) build their system section with this instead of
 * partitionFontFamilies, so no builtin shows up twice.
 */
export function systemFamiliesBesidesCandidates(
  candidates: readonly string[],
  systemFamilies: readonly string[],
): readonly string[] {
  const candidateSet = new Set(candidates)
  return systemFamilies.filter((f) => !candidateSet.has(f))
}

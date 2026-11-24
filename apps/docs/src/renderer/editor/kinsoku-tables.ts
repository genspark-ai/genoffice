/**
 * Word's Standard kinsoku lists per East Asian language, as it applied them to
 * 620 candidate characters (probe 2026-11-18). Korean refuses nothing.
 */
export interface KinsokuLists {
  /** may not start a line */
  before: string
  /** may not end a line */
  after: string
}

export const KINSOKU_LANGS = ['ja-JP', 'zh-CN', 'zh-TW', 'ko-KR'] as const
export type KinsokuLang = (typeof KINSOKU_LANGS)[number]

export const STANDARD_KINSOKU: Readonly<Record<KinsokuLang, KinsokuLists>> = {
  'ja-JP': {
    before:
      '!%),.:;?]}\u00a2\u00b0\u2019\u201d\u2030\u2032' +
      '\u2033\u2103\u3001\u3002\u3005\u3009\u300b\u300d\u300f\u3011\u3015\u309b\u309c\u309d\u309e\u30fb' +
      '\u30fd\u30fe\uff01\uff05\uff09\uff0c\uff0e\uff1a\uff1b\uff1f\uff3d\uff5d\uff61\uff63\uff64\uff65' +
      '\uff9e\uff9f\uffe0',
    after:
      '$([\\{\u00a3\u00a5\u2018\u201c\u3008\u300a\u300c\u300e\u3010\u3014\uff04' +
      '\uff08\uff3b\uff5b\uff62\uffe1\uffe5',
  },
  'zh-CN': {
    before:
      '!%),.:;>?]}\u00a2\u00a8\u00b0\u00b7\u02c7' +
      '\u02c9\u2015\u2016\u2019\u201d\u2026\u2030\u2032\u2033\u203a\u2103\u2236\u3001\u3002\u3003\u3009' +
      '\u300b\u300d\u300f\u3011\u3015\u3017\u301e\ufe36\ufe3a\ufe3e\ufe40\ufe44\ufe5a\ufe5c\ufe5e\uff01' +
      '\uff02\uff05\uff07\uff09\uff0c\uff0e\uff1a\uff1b\uff1f\uff3d\uff40\uff5c\uff5d\uff5e\uffe0',
    after:
      '$([{\u00a3\u00a5\u00b7\u2018\u201c\u3008\u300a\u300c\u300e\u3010\u3014\u3016' +
      '\u301d\ufe59\ufe5b\ufe5d\uff04\uff08\uff0e\uff3b\uff5b\uffe1\uffe5',
  },
  'zh-TW': {
    before:
      '!),.:;?]}\u00a2\u00b7\u2019\u201d\u2022\u2025\u2026' +
      '\u2027\u2032\u3001\u3002\u3009\u300b\u300d\u300f\u3011\u3015\u301e\ufe30\ufe31\ufe33\ufe34\ufe36' +
      '\ufe38\ufe3a\ufe3c\ufe3e\ufe40\ufe42\ufe44\ufe4f\ufe50\ufe51\ufe52\ufe54\ufe55\ufe56\ufe57\ufe5a' +
      '\ufe5c\ufe5e\uff01\uff09\uff0c\uff0e\uff1a\uff1b\uff1f\uff3d\uff5c\uff5d\uff64',
    after:
      '([{\u00a3\u00a5\u2018\u201c\u2035\u3008\u300a\u300c\u300e\u3010\u3014\u301d\ufe35' +
      '\ufe37\ufe39\ufe3b\ufe3d\ufe3f\ufe41\ufe43\ufe59\ufe5b\ufe5d\uff08\uff5b',
  },
  'ko-KR': { before: '', after: '' },
}

/** what the Strict level adds to the Japanese line-start list */
export const STRICT_JA_LINE_START = new Set(
  '\u3041\u3043\u3045\u3047\u3049\u3063\u3083\u3085\u3087\u308e\u30a1\u30a3\u30a5\u30a7\u30a9\u30c3' +
    '\u30e3\u30e5\u30e7\u30ee\u30f5\u30f6\u30fc\uff67\uff68\uff69\uff6a\uff6b\uff6c\uff6d\uff6e\uff6f' +
    '\uff70',
)

/** the table a language tag selects; text with no East Asian language took the zh-CN one */
export function kinsokuLang(tag: string | null | undefined): KinsokuLang | null {
  if (!tag) return 'zh-CN'
  const lower = tag.toLowerCase()
  if (lower.startsWith('ja')) return 'ja-JP'
  if (lower.startsWith('ko')) return 'ko-KR'
  if (!lower.startsWith('zh')) return null
  return /^zh-(tw|hk|mo|hant)/.test(lower) ? 'zh-TW' : 'zh-CN'
}

const union = (pick: (lists: KinsokuLists) => string, extra = '') =>
  new Set([...KINSOKU_LANGS.flatMap((lang) => [...pick(STANDARD_KINSOKU[lang])]), ...extra])

/** every character some standard list keeps off the start / the end of a line */
export const ANY_STANDARD_LINE_START = union((l) => l.before, [...STRICT_JA_LINE_START].join(''))
export const ANY_STANDARD_LINE_END = union((l) => l.after)

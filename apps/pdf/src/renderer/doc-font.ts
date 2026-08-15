/** Approximate a PDF font for on-screen previews. The fonts pdf.js loads can't
 * render typed unicode text (their cmaps are rebuilt around the content
 * stream's own charcodes), but the PostScript name survives — map it to a local
 * stack that reads like the document: the real family first, then a
 * same-classification CJK/Latin face. Display-only; the save-side engine
 * matches fonts by itself. */

export interface DocFontStyle {
  /** CSS font-family stack */
  css: string
  /** Inferred CSS weight (only when the name says so) */
  weight?: number
  italic?: true
}

/** 'NotoSerifCJKsc' → 'Noto Serif CJKsc' (lower→upper boundaries only: 'CJK'
    stays glued so acronyms don't shatter) */
const deCamel = (s: string): string => s.replace(/([a-z0-9])([A-Z])/g, '$1 $2')

const WEIGHTS: [RegExp, number][] = [
  [/(heavy|black)/i, 900],
  [/extra ?bold|ultra ?bold/i, 800],
  [/semi ?bold|demi ?bold|demi(?![a-z])/i, 600],
  [/bold/i, 700],
  [/medium/i, 500],
  [/extra ?light|ultra ?light|thin/i, 200],
  [/light/i, 300],
]

export function mapDocFont(psName: string): DocFontStyle {
  // Subset tag ('ABCDEF+') carries no meaning
  const name = psName.replace(/^[A-Z]{6}\+/, '')
  const spaced = deCamel(name.replace(/[-_,]/g, ' ')).replace(/\s+/g, ' ').trim()

  let weight: number | undefined
  for (const [re, w] of WEIGHTS) {
    if (re.test(spaced)) {
      weight = w
      break
    }
  }
  const italic = /italic|oblique/i.test(spaced) ? (true as const) : undefined

  // Family name without style tokens, both spaced and glued: local font matching
  // is by family name, and either form may be installed
  const familySpaced = spaced
    .replace(
      /\b(regular|italic|oblique|heavy|black|extra ?bold|ultra ?bold|semi ?bold|demi ?bold|demi|bold|medium|extra ?light|ultra ?light|thin|light|mt|ps|std|pro)\b/gi,
      ' ',
    )
    .replace(/\s+/g, ' ')
    .trim()
  const candidates = [familySpaced, familySpaced.replace(/ /g, '')].filter(
    (f, i, a) => f && a.indexOf(f) === i,
  )

  // Classification for the approximation tail. CJK first: CJK names often
  // contain latin keywords too ('Noto SERIF CJK'), so the CJK check must win.
  const cjk =
    /cjk|han(s|t)?(?![a-z])|sc(?![a-z])|tc(?![a-z])|gb(?![a-z])|song|sung|ming|hei(?![a-z])|kai|fang ?song|yahei|pingfang|deng ?xian|source ?han|noto (serif|sans) cjk|sim ?sun|sim ?hei|st ?(song|hei|kai|fangsong)/i.test(
      spaced,
    )
  let tail: string
  if (cjk) {
    if (/kai/i.test(spaced)) tail = '"Kaiti SC", "STKaiti", "KaiTi", serif'
    else if (/fang ?song/i.test(spaced)) tail = '"STFangsong", "FangSong", serif'
    else if (/song|sung|ming|serif|sim ?sun/i.test(spaced))
      tail = '"Songti SC", "STSong", "SimSun", serif'
    else tail = '"PingFang SC", "Microsoft YaHei", "SimHei", sans-serif'
  } else if (/courier|mono/i.test(spaced)) {
    tail = '"Courier New", monospace'
  } else if (
    /times|georgia|garamond|palatino|baskerville|caslon|charter|bookman|book antiqua|century|serif|roman/i.test(
      spaced,
    )
  ) {
    tail = 'Georgia, "Times New Roman", serif'
  } else {
    tail = 'Helvetica, Arial, sans-serif'
  }

  const css = [...candidates.map((f) => `"${f}"`), tail].join(', ')
  return { css, ...(weight ? { weight } : {}), ...(italic ? { italic } : {}) }
}

/** <a:buAutoNum type> numbering shared by the pptx text layout and xlsx shape text. */
function toRoman(n: number): string {
  const table: Array<[number, string]> = [
    [1000, 'm'],
    [900, 'cm'],
    [500, 'd'],
    [400, 'cd'],
    [100, 'c'],
    [90, 'xc'],
    [50, 'l'],
    [40, 'xl'],
    [10, 'x'],
    [9, 'ix'],
    [5, 'v'],
    [4, 'iv'],
    [1, 'i'],
  ]
  let out = ''
  for (const [v, s] of table)
    while (n >= v) {
      out += s
      n -= v
    }
  return out
}

function toAlpha(n: number): string {
  let out = ''
  while (n > 0) {
    n--
    out = String.fromCharCode(97 + (n % 26)) + out
    n = Math.floor(n / 26)
  }
  return out
}

const CJK_DIGITS = '\u3007\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d'
function toCjkNum(n: number): string {
  if (n <= 10) return n === 10 ? '\u5341' : CJK_DIGITS[n]!
  if (n < 20) return '\u5341' + CJK_DIGITS[n % 10]!
  if (n < 100)
    return CJK_DIGITS[Math.floor(n / 10)]! + '\u5341' + (n % 10 ? CJK_DIGITS[n % 10]! : '')
  return String(n)
}

/**
 * <a:buAutoNum type> → numbered-bullet glyph (ST_TextAutonumberScheme). Circled numbers
 * only exist up to ⑳/⓴; PowerPoint falls back to plain arabic beyond that.
 */
export function formatAutoNum(n: number, numType: string | undefined): string {
  const t = numType ?? 'arabicPeriod'
  if (t.startsWith('circleNum')) {
    if (t === 'circleNumWdBlackPlain')
      return n <= 10
        ? String.fromCodePoint(0x2775 + n) // ❶–❿
        : n <= 20
          ? String.fromCodePoint(0x24eb + (n - 11)) // ⓫–⓴
          : String(n)
    // circleNumWdWhitePlain included: Wingdings white circled digits are single-ring, i.e. ①–⑳
    return n <= 20 ? String.fromCodePoint(0x245f + n) : String(n) // ①–⑳
  }
  let body: string
  if (t.startsWith('alphaLc')) body = toAlpha(n)
  else if (t.startsWith('alphaUc')) body = toAlpha(n).toUpperCase()
  else if (t.startsWith('romanLc')) body = toRoman(n)
  else if (t.startsWith('romanUc')) body = toRoman(n).toUpperCase()
  else if (t.startsWith('arabicDb'))
    body = [...String(n)].map((d) => String.fromCodePoint(0xff10 + Number(d))).join('') // fullwidth １２３
  else if (t.startsWith('ea1Chs') || t.startsWith('ea1Cht')) body = toCjkNum(n)
  else body = String(n)
  if (t.endsWith('ParenBoth')) return `(${body})`
  if (t.endsWith('ParenR')) return `${body})`
  if (t.endsWith('Period')) return `${body}.`
  if (t.endsWith('Plain')) return body
  return `${body}.`
}

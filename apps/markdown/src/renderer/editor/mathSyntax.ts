/**
 * Stricter inline math than the upstream default (`$...$` with any content):
 * the content must not start or end with whitespace, the closing `$` must not
 * be followed by a digit, and `<amount> <words> <amount>` prose (`$5 and
 * 10$`) is currency rather than a formula; any operator or symbol between the
 * amounts (`$2 + 3$`, `$3 \times 4$`) keeps it math.
 */
const INLINE_MATH_RE = /\$(?!\s)([^$\n]*[^\s$])\$(?!\d)/g
const CURRENCY_SPAN_RE = /^\d[\d.,]*(?:\s+[\p{L}\p{N}]+)*\s+\d[\d.,]*$/u

export function matchInlineMath(src: string): { raw: string; latex: string } | undefined {
  INLINE_MATH_RE.lastIndex = 0
  const match = INLINE_MATH_RE.exec(src)
  if (!match || match.index !== 0 || CURRENCY_SPAN_RE.test(match[1])) return undefined
  return { raw: match[0], latex: match[1].trim() }
}

/** true when the text, written back as-is, would tokenize a `$...$` or `$$...$$` span */
export function containsMathSyntax(text: string): boolean {
  if (!text.includes('$')) return false
  if (text.includes('$$')) return true
  for (let i = text.indexOf('$'); i >= 0; i = text.indexOf('$', i + 1)) {
    if (matchInlineMath(text.slice(i))) return true
  }
  return false
}

/**
 * Stricter inline math than the upstream default (`$...$` with any content):
 * the content must not start or end with whitespace, the closing `$` must not
 * be followed by a digit, and `<amount> <words> <amount>` prose (`$5 and
 * 10$`) is currency rather than a formula; any operator or symbol between the
 * amounts (`$2 + 3$`, `$3 \times 4$`) keeps it math. An escaped dollar never
 * delimits, so the content also may not end in a backslash (`$5\$` is not a
 * formula with a literal `$`, it is not a formula at all), and the closing `$`
 * is not one of a run (`$x$$` is currency-adjacent prose, not `$x$` followed by
 * a stray dollar).
 */
const INLINE_MATH_RE = /\$(?!\s)([^$\n]*[^\\\s$])\$(?![$\d])/g
const CURRENCY_SPAN_RE = /^\d[\d.,]*(?:\s+[\p{L}\p{N}]+)*\s+\d[\d.,]*$/u

/** An escaped dollar ("costs \$5") never opens a formula, so the tokenizer has to
 * start at the first unescaped one instead of the first `$` in the source. */
const UNESCAPED_DOLLAR_RE = /(?<!\\)(?:\\\\)*\$/

export function strictInlineMathStart(src: string): number {
  const match = UNESCAPED_DOLLAR_RE.exec(src)
  return match ? match.index + match[0].length - 1 : -1
}

export function matchInlineMath(src: string): { raw: string; latex: string } | undefined {
  INLINE_MATH_RE.lastIndex = 0
  const match = INLINE_MATH_RE.exec(src)
  if (!match || match.index !== 0 || CURRENCY_SPAN_RE.test(match[1])) return undefined
  return { raw: match[0], latex: match[1].trim() }
}

/** true when the text, written back as-is, would tokenize a `$...$` or `$$...$$` span */
export function containsMathSyntax(text: string): boolean {
  for (let i = strictInlineMathStart(text); i >= 0;) {
    const rest = text.slice(i)
    if (rest.startsWith('$$') || matchInlineMath(rest)) return true
    const next = strictInlineMathStart(text.slice(i + 1))
    i = next < 0 ? -1 : i + 1 + next
  }
  return false
}

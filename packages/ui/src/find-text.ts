export interface FindOptions {
  matchCase: boolean
  wholeWord: boolean
}

export const isWordChar = (ch: string | undefined): boolean => !!ch && /[\p{L}\p{N}_]/u.test(ch)

/** length-preserving lowercase: chars whose lowercase grows ('İ' → 'i̇') stay as-is so match offsets never shift */
export function foldCase(s: string): string {
  let out = ''
  for (const ch of s) {
    const lower = ch.toLowerCase()
    out += lower.length === ch.length ? lower : ch
  }
  return out
}

/** start offsets of every non-overlapping occurrence of `query` in `text` */
export function findInText(text: string, query: string, opts: FindOptions): number[] {
  const found: number[] = []
  if (!query) return found
  const needle = opts.matchCase ? query : foldCase(query)
  const haystack = opts.matchCase ? text : foldCase(text)
  let i = 0
  while ((i = haystack.indexOf(needle, i)) !== -1) {
    const isWhole =
      !opts.wholeWord || (!isWordChar(text[i - 1]) && !isWordChar(text[i + query.length]))
    if (isWhole) {
      found.push(i)
      i += query.length
    } else {
      i += 1
    }
  }
  return found
}

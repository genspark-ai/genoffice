/**
 * Pre-tokenizer feeding SQLite FTS5 (unicode61 has no CJK segmentation: a run of
 * Han/kana/hangul is one token, so a two-character word never matches inside a longer run).
 * CJK runs become character bigrams — the shape Obsidian's search plugins and
 * Logseq use — while Latin/digit runs stay whole words. A parallel unigram
 * stream serves one-character CJK queries, which bigrams cannot answer.
 */

export interface Tokenized {
  /** bigram stream: CJK runs as overlapping pairs, other runs as whole words */
  bi: string[]
  /** unigram stream: CJK runs as single characters, other runs as whole words */
  uni: string[]
}

const CJK_RANGES: ReadonlyArray<readonly [number, number]> = [
  [0x2e80, 0x2fdf], // CJK radicals
  [0x3005, 0x3007], // 々 〆 〇
  [0x3040, 0x30ff], // hiragana, katakana
  [0x3100, 0x312f], // bopomofo
  [0x3130, 0x318f], // hangul compatibility jamo
  [0x31a0, 0x31ff], // bopomofo ext, katakana ext
  [0x3400, 0x4dbf], // CJK ext A
  [0x4e00, 0x9fff], // CJK unified
  [0xa960, 0xa97f], // hangul jamo ext A
  [0xac00, 0xd7ff], // hangul syllables, jamo ext B
  [0xf900, 0xfaff], // CJK compat ideographs
  [0xff66, 0xff9f], // halfwidth katakana
  [0x1100, 0x11ff], // hangul jamo
  [0x20000, 0x323af], // CJK ext B–H
]

export function isCjk(cp: number): boolean {
  for (const [lo, hi] of CJK_RANGES) if (cp >= lo && cp <= hi) return true
  return false
}

const WORD_CHAR = /^[\p{L}\p{N}\p{M}]$/u

type RunKind = 'cjk' | 'word'

interface Run {
  kind: RunKind
  /** code points as strings (surrogate-safe) */
  chars: string[]
}

/** NFKC folds fullwidth forms ("ＡＢＣ", "２０２６") and compatibility ideographs onto their base */
function fold(text: string): string {
  return text.normalize('NFKC').toLowerCase()
}

function splitRuns(text: string): Run[] {
  const runs: Run[] = []
  let current: Run | null = null
  for (const ch of fold(text)) {
    const cp = ch.codePointAt(0)!
    let kind: RunKind | null = null
    if (isCjk(cp)) kind = 'cjk'
    else if (WORD_CHAR.test(ch)) kind = 'word'
    if (kind === null) {
      current = null
      continue
    }
    if (current && current.kind === kind) current.chars.push(ch)
    else {
      current = { kind, chars: [ch] }
      runs.push(current)
    }
  }
  return runs
}

export function tokenize(text: string): Tokenized {
  const bi: string[] = []
  const uni: string[] = []
  for (const run of splitRuns(text)) {
    if (run.kind === 'word') {
      const word = run.chars.join('')
      bi.push(word)
      uni.push(word)
      continue
    }
    uni.push(...run.chars)
    if (run.chars.length === 1) bi.push(run.chars[0]!)
    else for (let i = 0; i + 1 < run.chars.length; i++) bi.push(run.chars[i]! + run.chars[i + 1]!)
  }
  return { bi, uni }
}

/** the space-joined form FTS5's unicode61 tokenizer re-splits one-to-one */
export function toIndexText(tokens: string[]): string {
  return tokens.join(' ')
}

export interface QueryTerm {
  /** the user's text for this term, folded, used for snippet matching */
  text: string
  /** true when every CJK run has two or more characters, so the bigram columns answer it */
  bigramSafe: boolean
  tokens: Tokenized
  /** the term ends in a Latin/digit word the user may still be typing */
  prefix: boolean
}

/**
 * Whitespace splits terms (AND); "quoted text" keeps one phrase; a leading
 * `-` excludes the term. Column filters and other operators come later.
 */
export function parseQuery(input: string): { include: QueryTerm[]; exclude: QueryTerm[] } {
  const include: QueryTerm[] = []
  const exclude: QueryTerm[] = []
  const re = /(-)?(?:"([^"]*)"|(\S+))/g
  let m: RegExpExecArray | null
  while ((m = re.exec(input))) {
    const raw = (m[2] ?? m[3] ?? '').trim()
    if (!raw) continue
    const term = makeTerm(raw, m[2] === undefined)
    if (!term) continue
    ;(m[1] ? exclude : include).push(term)
  }
  return { include, exclude }
}

function makeTerm(raw: string, allowPrefix: boolean): QueryTerm | null {
  const runs = splitRuns(raw)
  if (runs.length === 0) return null
  const tokens = tokenize(raw)
  const bigramSafe = runs.every((r) => r.kind === 'word' || r.chars.length >= 2)
  const last = runs[runs.length - 1]!
  const prefix = allowPrefix && last.kind === 'word' && last.chars.length >= 2
  return { text: fold(raw), bigramSafe, tokens, prefix }
}

const BI_COLUMNS = '{name path body}'
const UNI_COLUMNS = '{name_u path_u body_u}'

function phrase(tokens: string[], prefix: boolean): string {
  const quoted = `"${tokens.map((t) => t.replace(/"/g, '""')).join(' ')}"`
  return prefix ? `${quoted} *` : quoted
}

export function termExpr(term: QueryTerm): string {
  return term.bigramSafe
    ? `${BI_COLUMNS}: ${phrase(term.tokens.bi, term.prefix)}`
    : `${UNI_COLUMNS}: ${phrase(term.tokens.uni, term.prefix)}`
}

/** FTS5 MATCH expression for the parsed query, or null when nothing is searchable */
export function toMatchExpression(
  parsed: ReturnType<typeof parseQuery>,
  join: 'AND' | 'OR' = 'AND',
): string | null {
  if (parsed.include.length === 0) return null
  let expr = parsed.include.map((t) => `(${termExpr(t)})`).join(` ${join} `)
  for (const t of parsed.exclude) expr = `(${expr}) NOT (${termExpr(t)})`
  return expr
}

export interface TokenExpr {
  /** folded token text, the needle the highlighter lights up */
  text: string
  /** FTS5 expression matching this token alone */
  expr: string
  /** character offsets of the term this token covers, across every place it occurs */
  at: number[]
}

/**
 * One expression per distinct token of a term, with the characters each one
 * covers: coverage is measured in characters because a CJK phrase yields
 * bigrams straddling word boundaries that no file holds, and they must not
 * outvote the words that do match.
 */
export function tokenExprs(term: QueryTerm): TokenExpr[] {
  const cols = term.bigramSafe ? BI_COLUMNS : UNI_COLUMNS
  const out = new Map<string, TokenExpr>()
  const add = (tok: string, at: number[]): void => {
    const cur = out.get(tok)
    if (cur) cur.at.push(...at)
    else out.set(tok, { text: tok, expr: '', at })
  }
  let pos = 0
  let last = ''
  for (const run of splitRuns(term.text)) {
    const n = run.chars.length
    if (run.kind === 'word') add((last = run.chars.join('')), range(pos, n))
    else if (!term.bigramSafe || n === 1) run.chars.forEach((c, i) => add(c, [pos + i]))
    else
      for (let i = 0; i + 1 < n; i++) add(run.chars[i]! + run.chars[i + 1]!, [pos + i, pos + i + 1])
    pos += n
  }
  for (const t of out.values())
    t.expr = `${cols}: ${phrase([t.text], term.prefix && t.text === last)}`
  return [...out.values()]
}

function range(start: number, n: number): number[] {
  return Array.from({ length: n }, (_, i) => start + i)
}

/** code points of the term that tokens can cover, the unit `TokenExpr.at` indexes */
export function termChars(term: QueryTerm): number {
  return splitRuns(term.text).reduce((n, r) => n + r.chars.length, 0)
}

export { fold as foldText }

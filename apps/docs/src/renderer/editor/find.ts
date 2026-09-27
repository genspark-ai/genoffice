/**
 * Find & Replace engine: query compilation (literal / Word wildcards / regex),
 * width and case folding, document scanning and replacement expansion. Pure
 * functions over ProseMirror documents so the panel stays thin and testable.
 */
import type { Node as PmNode } from '@tiptap/pm/model'

export type FindMode = 'literal' | 'wildcard' | 'regex'

export interface FindOptions {
  matchCase: boolean
  wholeWord: boolean
  mode?: FindMode
  /** Word's "match byte": half- and full-width forms are distinct only when on */
  matchWidth?: boolean
}

export interface FindMatch {
  from: number
  to: number
  /** matched text as it appears in the document */
  text: string
  /** capture groups (regex / wildcard modes) */
  groups: string[] | null
  /** flattened text of the owning block, for result previews */
  blockText: string
  /** match offset inside blockText */
  offset: number
}

export interface CompiledQuery {
  mode: FindMode
  /** null only for the plain literal fast path */
  re: RegExp | null
  needle: string
  /** true when the pattern must sit at the start / end of a paragraph (^p anchors) */
  anchorStart: boolean
  anchorEnd: boolean
}

export type CompileResult = { ok: true; query: CompiledQuery } | { ok: false; error: string }

/** how many matches the results list renders before it says "showing first N" */
export const RESULTS_LIST_CAP = 500
/** above this many hits only the active one is decorated (decorating thousands of ranges stalls typing) */
export const HIGHLIGHT_CAP = 2000

const WORD_CHAR = '[\\p{L}\\p{N}_]'
/** flattened stand-in for a non-text leaf (image, field, note mark): never part of a match */
export const LEAF_CHAR = '\u0000'
/** flattened stand-in for a hard break; `^l` finds it, `?` / `*` do not */
export const BREAK_CHAR = '\n'
/** wildcard `?`: any character that is not a leaf or a break */
const ANY_CHAR = '[^\\u0000\\n]'
const isWordChar = (ch: string | undefined) => !!ch && /[\p{L}\p{N}_]/u.test(ch)

/** length-preserving lowercase: chars whose lowercase grows ('İ' → 'i̇') stay as-is so match offsets never shift */
export function foldCase(s: string): string {
  let out = ''
  for (const ch of s) {
    const lower = ch.toLowerCase()
    out += lower.length === ch.length ? lower : ch
  }
  return out
}

// half-width katakana U+FF66–FF9D → full-width code points, in order
const HALF_KANA = String.fromCharCode(
  0x30f2,
  0x30a1,
  0x30a3,
  0x30a5,
  0x30a7,
  0x30a9,
  0x30e3,
  0x30e5,
  0x30e7,
  0x30c3,
  0x30fc,
  0x30a2,
  0x30a4,
  0x30a6,
  0x30a8,
  0x30aa,
  0x30ab,
  0x30ad,
  0x30af,
  0x30b1,
  0x30b3,
  0x30b5,
  0x30b7,
  0x30b9,
  0x30bb,
  0x30bd,
  0x30bf,
  0x30c1,
  0x30c4,
  0x30c6,
  0x30c8,
  0x30ca,
  0x30cb,
  0x30cc,
  0x30cd,
  0x30ce,
  0x30cf,
  0x30d2,
  0x30d5,
  0x30d8,
  0x30db,
  0x30de,
  0x30df,
  0x30e0,
  0x30e1,
  0x30e2,
  0x30e4,
  0x30e6,
  0x30e8,
  0x30e9,
  0x30ea,
  0x30eb,
  0x30ec,
  0x30ed,
  0x30ef,
  0x30f3,
)

/**
 * Length-preserving width folding (Word's "match byte" off): full-width ASCII
 * and the ideographic space become their half-width forms, half-width katakana
 * their full-width base letters. Voiced-sound marks stay separate characters so
 * offsets never shift (a half-width GA folds to KA + mark, not to one GA).
 */
export function foldWidth(s: string): string {
  let out = ''
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i)
    if (c >= 0xff01 && c <= 0xff5e) out += String.fromCharCode(c - 0xfee0)
    else if (c === 0x3000) out += ' '
    else if (c >= 0xff66 && c <= 0xff9d) out += HALF_KANA[c - 0xff66]
    else if (c === 0xff9e) out += String.fromCharCode(0x309b)
    else if (c === 0xff9f) out += String.fromCharCode(0x309c)
    else out += s[i]
  }
  return out
}

const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')

/**
 * Word special codes shared by all modes (`^t` tab, `^#` digit, …). `^p` is
 * returned as the marker 'P' so the caller can turn it into a paragraph anchor.
 */
function caretCode(code: string): string | 'P' | null {
  switch (code) {
    case 'p':
    case '13':
      return 'P'
    case 't':
    case '9':
      return '\\t'
    case 'l':
    case '11':
      return '\\n'
    case 's':
      return '\\u00a0'
    case '~':
      return '\\u2011'
    case '-':
      return '\\u00ad'
    case '^':
      return '\\^'
    case '#':
      return '\\d'
    case '$':
      return '\\p{L}'
    case 'w':
      return '[ \\t\\u00a0]+'
    case '?':
      return ANY_CHAR
    default: {
      if (/^\d{1,4}$/.test(code)) return `\\u{${Number(code).toString(16)}}`
      return null
    }
  }
}

/** reads a `^` code at index i (after the caret); returns [translation, consumed] */
function readCaret(src: string, i: number): [string | 'P' | null, number] {
  const digits = /^\d{1,4}/.exec(src.slice(i))
  if (digits) {
    const t = caretCode(digits[0])
    if (t) return [t, digits[0].length]
  }
  const one = src[i]
  if (one === undefined) return [null, 0]
  return [caretCode(one), 1]
}

interface Translated {
  source: string
  anchorStart: boolean
  anchorEnd: boolean
}

/**
 * Word wildcard syntax → JS regex source. `?` any character, `*` any run
 * (shortest, like Word), `[a-z]` sets, `[!x]` negation, `<` `>` word start /
 * end, `@` one or more, `{n,m}` repeats, `(…)` groups, `\x` literal.
 */
export function wildcardToRegExp(
  pattern: string,
  fold: (ch: string) => string = (ch) => ch,
): Translated | { error: string } {
  let out = ''
  let anchorStart = false
  let anchorEnd = false
  let depth = 0
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i]
    switch (ch) {
      case '\\': {
        const next = pattern[i + 1]
        if (next === undefined) return { error: 'trailing-backslash' }
        out += escapeRe(next)
        i++
        break
      }
      case '?':
        out += ANY_CHAR
        break
      case '*':
        out += `${ANY_CHAR}*?`
        break
      case '@':
        if (!out) return { error: 'nothing-to-repeat' }
        out += '+'
        break
      case '<':
        out += `(?<!${WORD_CHAR})(?=${WORD_CHAR})`
        break
      case '>':
        out += `(?<=${WORD_CHAR})(?!${WORD_CHAR})`
        break
      case '(':
        depth++
        out += '('
        break
      case ')':
        if (depth === 0) return { error: 'unbalanced-paren' }
        depth--
        out += ')'
        break
      case '[': {
        const close = pattern.indexOf(']', i + 2)
        if (close === -1) return { error: 'unbalanced-bracket' }
        let body = pattern.slice(i + 1, close)
        let negate = false
        if (body.startsWith('!')) {
          negate = true
          body = body.slice(1)
        }
        if (!body) return { error: 'empty-set' }
        const set = fold(body).replace(/[\\\]^]/g, '\\$&')
        out += negate ? `[^${set}\\u0000\\n]` : `[${set}]`
        i = close
        break
      }
      case '{': {
        const m = /^\{(\d+)(,(\d*))?\}/.exec(pattern.slice(i))
        if (!m || !out) return { error: 'bad-repeat' }
        out += m[0]
        i += m[0].length - 1
        break
      }
      case '^': {
        const [t, used] = readCaret(pattern, i + 1)
        if (t === null) return { error: 'bad-caret-code' }
        if (t === 'P') {
          if (i === 0) anchorStart = true
          else if (i + 1 + used === pattern.length) anchorEnd = true
          else return { error: 'paragraph-mark-position' }
        } else out += t
        i += used
        break
      }
      default:
        out += escapeRe(fold(ch))
    }
  }
  if (depth !== 0) return { error: 'unbalanced-paren' }
  return { source: out, anchorStart, anchorEnd }
}

/** literal text with Word's `^` codes (`^t`, `^p`, `^?`, …) → regex source */
function literalToRegExp(
  text: string,
  fold: (ch: string) => string,
): Translated | { error: string } {
  let out = ''
  let anchorStart = false
  let anchorEnd = false
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (ch !== '^') {
      out += escapeRe(fold(ch))
      continue
    }
    const [t, used] = readCaret(text, i + 1)
    if (t === null) return { error: 'bad-caret-code' }
    if (t === 'P') {
      if (i === 0) anchorStart = true
      else if (i + 1 + used === text.length) anchorEnd = true
      else return { error: 'paragraph-mark-position' }
    } else out += t
    i += used
  }
  return { source: out, anchorStart, anchorEnd }
}

export function compileFind(query: string, opts: FindOptions): CompileResult {
  const mode: FindMode = opts.mode ?? 'literal'
  if (!query) return { ok: false, error: 'empty' }
  // `d` exposes group indices so captures can be read from the unfolded text
  const flags = `dgu${opts.matchCase ? '' : 'i'}`
  const fold = opts.matchWidth ? (ch: string) => ch : foldWidth
  const finish = (t: Translated): CompileResult => {
    const source = `${t.anchorStart ? '^' : ''}${t.source}${t.anchorEnd ? '$' : ''}`
    try {
      return {
        ok: true,
        query: {
          mode,
          re: new RegExp(source, flags),
          needle: query,
          anchorStart: t.anchorStart,
          anchorEnd: t.anchorEnd,
        },
      }
    } catch {
      return { ok: false, error: 'invalid-pattern' }
    }
  }
  if (mode === 'regex') {
    return finish({ source: foldPatternWidth(query, fold), anchorStart: false, anchorEnd: false })
  }
  if (mode === 'wildcard') {
    const t = wildcardToRegExp(query, fold)
    return 'error' in t ? { ok: false, error: t.error } : finish(t)
  }
  if (!query.includes('^')) {
    return {
      ok: true,
      query: { mode, re: null, needle: query, anchorStart: false, anchorEnd: false },
    }
  }
  const t = literalToRegExp(query, fold)
  return 'error' in t ? { ok: false, error: t.error } : finish(t)
}

/**
 * Width-fold a regex source without changing its meaning: a full-width char
 * whose half-width twin is a regex metacharacter stays as typed.
 */
export function foldPatternWidth(source: string, fold: (ch: string) => string): string {
  let out = ''
  for (const ch of source) {
    const f = fold(ch)
    out += f !== ch && /[.*+?^${}()|[\]\\/]/.test(f) ? ch : f
  }
  return out
}

interface FlatBlock {
  text: string
  /** document position of each character of `text` */
  posAt: number[]
}

/** flatten a textblock's inline content so matches spanning marks are found */
function flattenBlock(node: PmNode, pos: number): FlatBlock {
  let text = ''
  const posAt: number[] = []
  node.forEach((child, offset) => {
    if (child.isText && child.text) {
      for (let k = 0; k < child.text.length; k++) posAt.push(pos + 1 + offset + k)
      text += child.text
    } else {
      posAt.push(pos + 1 + offset)
      text += child.type.name === 'hardBreak' ? BREAK_CHAR : LEAF_CHAR
    }
  })
  return { text, posAt }
}

function foldHaystack(text: string, opts: FindOptions, forRegex: boolean): string {
  let h = text
  if (!opts.matchWidth) h = foldWidth(h)
  // regexes carry the `i` flag; the literal path folds both sides itself
  if (!forRegex && !opts.matchCase) h = foldCase(h)
  return h
}

function scanBlock(flat: FlatBlock, q: CompiledQuery, opts: FindOptions, out: FindMatch[]): void {
  const { text, posAt } = flat
  /** records the hit; false when it is rejected so the caller can retry from the next char */
  const push = (start: number, len: number, groups: string[] | null): boolean => {
    if (len <= 0) return false
    // a regex `.` or `[^x]` can swallow an image / field stand-in; those are not characters
    if (text.slice(start, start + len).includes(LEAF_CHAR)) return false
    if (opts.wholeWord && (isWordChar(text[start - 1]) || isWordChar(text[start + len])))
      return false
    out.push({
      from: posAt[start],
      to: posAt[start + len - 1] + 1,
      text: text.slice(start, start + len),
      groups,
      blockText: text,
      offset: start,
    })
    return true
  }
  if (!q.re) {
    const haystack = foldHaystack(text, opts, false)
    const needle = foldHaystack(q.needle, opts, false)
    let i = 0
    while ((i = haystack.indexOf(needle, i)) !== -1) {
      i += push(i, needle.length, null) ? needle.length : 1
    }
    return
  }
  const haystack = foldHaystack(text, opts, true)
  q.re.lastIndex = 0
  let m: RegExpExecArray | null
  while ((m = q.re.exec(haystack)) !== null) {
    if (m[0].length === 0) {
      q.re.lastIndex++
      continue
    }
    // folding is one-to-one per char, so folded indices address the original text
    const spans = m.indices ?? []
    const groups =
      m.length > 1 ? spans.slice(1).map((sp) => (sp ? text.slice(sp[0], sp[1]) : '')) : null
    // a rejected span may hide a shorter valid hit inside it: rescan from its second char
    if (!push(m.index, m[0].length, groups)) q.re.lastIndex = m.index + 1
  }
}

/** protected regions are read-only: Word skips them, so do we */
const isSearchable = (node: PmNode) => node.type.name !== 'docProtected'

/** collect matches in every textblock of `doc` (synchronous) */
export function findInDoc(doc: PmNode, q: CompiledQuery, opts: FindOptions): FindMatch[] {
  const out: FindMatch[] = []
  doc.descendants((node, pos) => {
    if (!isSearchable(node)) return false
    if (!node.isTextblock) return true
    scanBlock(flattenBlock(node, pos), q, opts, out)
    return false
  })
  return out
}

/** documents above this node size scan in time-sliced chunks so typing stays responsive (tests lower it) */
export const SCAN_THRESHOLD = { nodeSize: 400_000 }
const SLICE_BUDGET_MS = 12

/**
 * Time-sliced scan over the top-level blocks; resolves with all matches, or
 * null when `signal` is aborted (a newer query superseded this one).
 */
export function findInDocAsync(
  doc: PmNode,
  q: CompiledQuery,
  opts: FindOptions,
  signal: { aborted: boolean },
): Promise<FindMatch[] | null> {
  return new Promise((resolve) => {
    const out: FindMatch[] = []
    let index = 0
    let offset = 0
    const step = () => {
      if (signal.aborted) return resolve(null)
      const t0 = performance.now()
      while (index < doc.childCount) {
        const child = doc.child(index)
        if (isSearchable(child)) {
          if (child.isTextblock) scanBlock(flattenBlock(child, offset), q, opts, out)
          else
            child.descendants((node, pos) => {
              if (!isSearchable(node)) return false
              if (!node.isTextblock) return true
              scanBlock(flattenBlock(node, offset + 1 + pos), q, opts, out)
              return false
            })
        }
        offset += child.nodeSize
        index++
        if (performance.now() - t0 > SLICE_BUDGET_MS) return void setTimeout(step, 0)
      }
      resolve(out)
    }
    step()
  })
}

/**
 * Replacement text for one match. Segments are separated where the template
 * asks for a paragraph mark (`^p`): the caller splits the block there.
 * Regex mode: `$1`…`$9`, `$&`, `$$`, `\1`…`\9`. Wildcard / literal: `\1`…`\9`,
 * `^&` (the found text), `^t`, `^p`, `\\`.
 */
export function expandReplacement(template: string, m: FindMatch, mode: FindMode): string[] {
  const segs: string[] = ['']
  const add = (s: string) => {
    segs[segs.length - 1] += s
  }
  const group = (n: number) => (n === 0 ? m.text : (m.groups?.[n - 1] ?? ''))
  for (let i = 0; i < template.length; i++) {
    const ch = template[i]
    const next = template[i + 1]
    if (ch === '\\' && next !== undefined) {
      if (/[1-9]/.test(next)) add(group(Number(next)))
      else add(next)
      i++
      continue
    }
    if (mode === 'regex' && ch === '$' && next !== undefined) {
      if (/[1-9]/.test(next)) add(group(Number(next)))
      else if (next === '&') add(m.text)
      else if (next === '$') add('$')
      else {
        add(ch)
        continue
      }
      i++
      continue
    }
    if (ch === '^' && next !== undefined) {
      if (next === '&') add(m.text)
      else if (next === 'p') segs.push('')
      else if (next === 't') add('\t')
      else if (next === 'l') add('\n')
      else if (next === 's') add(' ')
      else if (next === '~') add('‑')
      else if (next === '^') add('^')
      else {
        add(ch)
        continue
      }
      i++
      continue
    }
    add(ch)
  }
  return segs
}

/** `radius` characters of block text on either side of a match, with ellipses */
export function matchContext(
  m: FindMatch,
  radius = 30,
): { before: string; hit: string; after: string } {
  const start = Math.max(0, m.offset - radius)
  const end = Math.min(m.blockText.length, m.offset + m.text.length + radius)
  const clean = (s: string) => s.replaceAll(LEAF_CHAR, ' ').replaceAll(BREAK_CHAR, ' ')
  return {
    before: (start > 0 ? '…' : '') + clean(m.blockText.slice(start, m.offset)),
    hit: clean(m.text),
    after:
      clean(m.blockText.slice(m.offset + m.text.length, end)) +
      (end < m.blockText.length ? '…' : ''),
  }
}

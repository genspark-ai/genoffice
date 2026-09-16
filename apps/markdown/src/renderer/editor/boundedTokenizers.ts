import type { MarkdownTokenizer } from '@tiptap/core'

/**
 * The stock list / table markdown tokenizers split the *whole remaining
 * source* on every block boundary (`src.split('\n')`), which makes parsing a
 * document quadratic — a 1.8 MB file spends ~7 s in three tokenizers. These
 * wrappers keep the upstream logic but hand it a bounded prefix, and reject
 * non-candidates by looking at the first line only.
 */

type Tokenize = MarkdownTokenizer['tokenize']

/** Loose superset of the upstream item patterns — must never reject a line the original tokenizer accepts. */
const ORDERED_ITEM = /^\s*\w+[.)]\s/
const TASK_ITEM = /^\s*[-+*]\s+\[[ xX]\]\s/

function firstNonBlankLine(src: string): string {
  let pos = 0
  while (pos < src.length) {
    const nl = src.indexOf('\n', pos)
    const line = nl < 0 ? src.slice(pos) : src.slice(pos, nl)
    if (line.trim()) return line
    if (nl < 0) break
    pos = nl + 1
  }
  return ''
}

/**
 * A list only continues past a blank line with an indented line or another
 * item, so the first blank line followed by a flush non-item line is a cut the
 * upstream tokenizer never crosses. Lazy continuation (no blank line) is kept
 * intact because it never triggers the cut.
 */
function listExtent(src: string, item: RegExp): string {
  let pos = 0
  let prevBlank = false
  while (pos < src.length) {
    const nl = src.indexOf('\n', pos)
    const line = nl < 0 ? src.slice(pos) : src.slice(pos, nl)
    const blank = line.trim() === ''
    if (prevBlank && !blank && !/^\s/.test(line) && !item.test(line)) return src.slice(0, pos)
    prevBlank = blank
    if (nl < 0) break
    pos = nl + 1
  }
  return src
}

function boundList(base: MarkdownTokenizer, item: RegExp): MarkdownTokenizer {
  const tokenize: Tokenize = function (this: unknown, src, tokens, lexer) {
    if (!item.test(firstNonBlankLine(src))) return undefined
    return base.tokenize.call(this, listExtent(src, item), tokens, lexer)
  }
  return { ...base, tokenize }
}

export function boundOrderedList(base: MarkdownTokenizer): MarkdownTokenizer {
  return boundList(base, ORDERED_ITEM)
}

export function boundTaskList(base: MarkdownTokenizer): MarkdownTokenizer {
  return boundList(base, TASK_ITEM)
}

/** The upstream table tokenizer already limits itself to the text before the
 *  first blank line except for one trailing `src.split('\n')`; `start` reads
 *  the first two lines but splits everything. */
export function boundTable(base: MarkdownTokenizer): MarkdownTokenizer {
  const start = (src: string): number => {
    const a = src.indexOf('\n')
    if (a < 0) return -1
    const b = src.indexOf('\n', a + 1)
    const sep = src.slice(a + 1, b < 0 ? undefined : b)
    if (!/^[ \t|:]*-[ \t|:-]*$/.test(sep) || !sep.includes('|')) return -1
    return src.slice(0, a).includes('|') ? 0 : -1
  }
  const tokenize: Tokenize = function (this: unknown, src, tokens, lexer) {
    const blank = src.indexOf('\n\n')
    return base.tokenize.call(this, blank >= 0 ? src.slice(0, blank) : src, tokens, lexer)
  }
  return { ...base, start, tokenize }
}

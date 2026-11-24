import type { Node as PmNode } from '@tiptap/pm/model'

export interface SpellWord {
  from: number
  to: number
  word: string
}

const WORD = /\p{L}[\p{L}\p{M}'’]*/gu
/** scripts written without spaces between words: the dictionaries hold no entries for them */
const UNSPACED =
  /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Thai}]/u

export function wordsOfBlock(block: PmNode, blockPos: number, out: SpellWord[]): void {
  let text = ''
  let start = 0
  const flush = () => {
    for (const m of text.matchAll(WORD)) {
      const word = m[0].replace(/['’]+$/, '')
      if (UNSPACED.test(word)) continue
      out.push({ from: start + m.index, to: start + m.index + word.length, word })
    }
    text = ''
  }
  block.forEach((child, offset) => {
    const checked =
      child.isText &&
      !child.marks.some(
        (mark) =>
          mark.type.name === 'del' || (mark.type.name === 'docTextStyle' && mark.attrs.noProof),
      )
    if (!checked) return flush()
    if (text === '') start = blockPos + 1 + offset
    text += child.text
  })
  flush()
}

/** the words a spelling check reads, in document order; deleted and no-proofing text is left out */
export function spellWords(doc: PmNode): SpellWord[] {
  const out: SpellWord[] = []
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true
    wordsOfBlock(node, pos, out)
    return false
  })
  return out
}

/** the first word ending in (from, to] the dictionary does not know: a word `from` is inside counts */
export function nextMisspelling(
  doc: PmNode,
  from: number,
  to: number,
  misspelled: (word: string) => boolean,
  dismissed?: (word: SpellWord) => boolean,
): SpellWord | null {
  const known = new Map<string, boolean>()
  for (const word of spellWords(doc)) {
    if (word.to <= from || word.to > to || dismissed?.(word)) continue
    let wrong = known.get(word.word)
    if (wrong === undefined) known.set(word.word, (wrong = misspelled(word.word)))
    if (wrong) return word
  }
  return null
}

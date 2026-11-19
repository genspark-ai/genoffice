/**
 * Word-parity text statistics.
 *
 * Word's CJK rule: Words = Asian characters (counted one by one, punctuation
 * included) + non-Asian words (whitespace/punct-delimited). The dialog also
 * reports the two addends separately — Chinese users care about the
 * Asian-character figure.
 */

import { eqLayoutText } from './editor/eq-layout'
import type { TableModel } from '@genoffice/docx-engine'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { Selection } from '@tiptap/pm/state'

// Han (incl. radicals/compat/ext-B+), kana, hangul, bopomofo, CJK symbols
// and punctuation (U+3001 up: the ideographic space stays whitespace),
// fullwidth forms. The surrogate pair covers CJK Extensions B-H
// (U+20000-U+3134F, high surrogates D840-D8BF); it stops at D8BF so emoji
// (D83C-D83E high surrogates) and the planes 4-16 private-use area and
// variation selectors supplement (D8C0-DBFF) are never counted as Asian chars.
const ASIAN_RE =
  /[ᄀ-ᇿ⺀-⿟、-〿぀-ヿ㄀-ㄯ㄰-㆏㇀-ㇿ㐀-䶿一-鿿가-힯豈-﫿！-｠￠-￦]|[\uD840-\uD8BF][\uDC00-\uDFFF]/g

// Latin Extended Additional (Vietnamese, …) plus the space-delimited
// non-Latin scripts Word counts as words: Greek, Cyrillic, Hebrew, Arabic,
// Armenian, Devanagari, Thai, Lao, Tibetan, Myanmar, Georgian, Khmer.
// Deliberately excludes the ASIAN_RE ranges above (counted char-by-char).
/* eslint-disable no-misleading-character-class -- Thai/Lao etc ranges contain combining marks but are intentional word characters */
const NON_ASIAN_WORD_RE =
  /[A-Za-z0-9À-ɏͰ-ϿЀ-ӿ֐-׿؀-ۿḀ-ỿ\u0E00-\u0EFF\u1000-\u109F\u10A0-\u10FF\u1780-\u17FF\u0530-\u058F\u0900-\u097F\u0F00-\u0FFF]+(?:['-][A-Za-z0-9À-ɏͰ-ϿЀ-ӿ֐-׿؀-ۿḀ-ỿ\u0E00-\u0EFF\u1000-\u109F\u10A0-\u10FF\u1780-\u17FF\u0530-\u058F\u0900-\u097F\u0F00-\u0FFF]+)*/g
/* eslint-enable no-misleading-character-class */

export function asianCharCount(text: string): number {
  return (text.match(ASIAN_RE) ?? []).length
}

export function nonAsianWordCount(text: string): number {
  return (text.match(NON_ASIAN_WORD_RE) ?? []).length
}

/** Word's Words figure: asian chars + non-asian words */
export function countWords(text: string): number {
  return asianCharCount(text) + nonAsianWordCount(text)
}

export interface TextStats {
  words: number
  asianChars: number
  nonAsianWords: number
  charsNoSpace: number
  charsWithSpace: number
  paragraphs: number
}

/** the Word Count figures of plain text whose paragraphs are separated by \n */
export function textStats(text: string): TextStats {
  return {
    words: countWords(text),
    asianChars: asianCharCount(text),
    nonAsianWords: nonAsianWordCount(text),
    charsNoSpace: text.replace(/\s/g, '').length,
    charsWithSpace: text.replace(/\n/g, '').length,
    paragraphs: text.split('\n').filter((line) => line.trim() !== '').length,
  }
}

/**
 * Words inside the current selection, for the status bar's "selected/total"
 * figure. Word shows that pair while a selection exists and the plain total
 * once the caret collapses; GenOffice reported only the total, so the selected
 * figure was reachable from the AI panel alone.
 *
 * Every no-selection case falls out of the counting rule itself: a collapsed
 * caret and a NodeSelection (an image or a table is selected as a node, not
 * a text range) both yield no text, and `countWords('')` is 0 — no separate
 * guard needed.
 */
export function selectedWordCount(doc: ProseMirrorNode, selection: Selection): number {
  // the same leaf reader as the total: an inline formula's text counts in both
  return countWords(doc.textBetween(selection.from, selection.to, ' ', leafTextForWordCount))
}

export function documentTextForWordCount(doc: ProseMirrorNode): string {
  // ProseMirror textContent skips Tiptap renderText serializers on atomic leaves.
  return doc.textBetween(0, doc.content.size, ' ', leafTextForWordCount)
}

/** a nested table's paragraph texts, cell by cell, deeper nested tables included */
export function tableModelText(model: TableModel | null | undefined): string {
  if (!model?.rows) return ''
  const parts: string[] = []
  for (const row of model.rows) {
    for (const cell of row) {
      parts.push(...(cell.paras ?? []))
      for (const inner of cell.nestedTables ?? []) parts.push(tableModelText(inner))
    }
  }
  return parts.join(' ')
}

/**
 * Text an atomic leaf contributes to a word count. No doc schema defines
 * `spec.leafText`, so the nodes with text of their own are named here: inline
 * math, a ruby base, and a table nested in a cell (an atom drawn from its
 * model, whose words Word counts like any other); every other atom
 * contributes nothing.
 */
export function leafTextForWordCount(leaf: ProseMirrorNode): string {
  if (leaf.type.name === 'docInlineMath') return String(leaf.attrs.text ?? '')
  if (leaf.type.name === 'docRuby') return String(leaf.attrs.base ?? '')
  if (leaf.type.name === 'docEqLayout') return eqLayoutText(leaf.attrs)
  if (leaf.type.name === 'docNestedTable') {
    return tableModelText(leaf.attrs.model as TableModel | null)
  }
  return leaf.type.spec.leafText?.(leaf) ?? ''
}

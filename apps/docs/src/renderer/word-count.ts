/**
 * Word-parity text statistics.
 *
 * Word's CJK rule: Words = Asian characters (counted one by one, punctuation
 * included) + non-Asian words (whitespace/punct-delimited). The dialog also
 * reports the two addends separately — Chinese users care about the
 * Asian-character figure.
 */

import {
  computeListMarkerInfos,
  type ListItemRef,
  type NumberingDef,
  type TableModel,
} from '@genoffice/docx-engine'
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

/**
 * Word's non-Asian words (probed against Word for Mac 16, 2026-11-19): any run
 * of characters between whitespace is one word, punctuation and symbols
 * included, so "0.75", "p<0.05", "a/b", "[1]", "&" and a URL each count once.
 * - whitespace: space, tab, line breaks, NBSP and the ideographic space. The
 *   thin, en and other typographic spaces join their neighbours ("a\u2009b" is
 *   one word).
 * - en dash, em dash, the bullet and the private-use symbol-font glyphs (a
 *   Symbol or Wingdings bullet, U+F0B7) split words and are not words
 *   themselves: "3248–3265" is two words, a lone "—" none.
 * - the Unicode hyphen (U+2010) and the arrows split words and are words of
 *   their own: "a→b" is three. The ASCII hyphen joins ("state-of-the-art").
 * Asian characters are counted one by one (asianCharCount) and split the runs
 * around them ("2026年8月2日": three non-Asian words).
 */
// explicit, not \s: JS counts the en/thin/hair spaces as whitespace, Word does not
const WORD_SPACE_RE = /[ \t\n\r\f\v\u00a0\u3000\u2013\u2014\u2022\ue000-\uf8ff]+/
const OWN_WORD_RE = /([\u2010\u2190-\u21ff])/

export function asianCharCount(text: string): number {
  return (text.match(ASIAN_RE) ?? []).length
}

export function nonAsianWordCount(text: string): number {
  let n = 0
  for (const run of text.replace(ASIAN_RE, ' ').split(WORD_SPACE_RE)) {
    if (!run) continue
    for (const part of run.split(OWN_WORD_RE)) if (part) n++
  }
  return n
}

/** Word's Words figure: asian chars + non-asian words */
export function countWords(text: string): number {
  return asianCharCount(text) + nonAsianWordCount(text)
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

/**
 * The list labels Word counts as words: an item's number ("1.", "a)") and a
 * bullet written as text (Word's Courier New "o"), but not a symbol-font bullet
 * (Word's U+F0B7: a private-use glyph, no word) nor a picture bullet. Word for
 * Mac 16 probe 2026-11-19: three numbered items "alpha", "beta", "gamma" are
 * six words, the same items under a Symbol bullet three, under "o" six. The
 * labels come from the numbering engine the editor draws them with.
 */
export function listLabelText(
  doc: ProseMirrorNode,
  defs: Map<string, NumberingDef> | null | undefined,
): string {
  if (!defs?.size) return ''
  const refs: ListItemRef[] = []
  doc.descendants((node) => {
    if (node.type.name !== 'docListItem') return true
    const deleted =
      !!node.attrs.paraMarkDel ||
      (node.attrs.blockRevision as { kind?: string } | null)?.kind === 'del'
    refs.push({
      numId: (node.attrs.numId as string | null) ?? null,
      ilvl: Number(node.attrs.ilvl) || 0,
      ...(deleted ? { deleted: true } : {}),
    })
    return false
  })
  if (refs.length === 0) return ''
  return computeListMarkerInfos(refs, defs)
    .map((m, i) => {
      if (!m || m.picBulletSrc || refs[i]!.deleted) return ''
      // the glyph as authored: a symbol-font bullet stays in the private-use area
      return m.symbolFont && m.symbolChar ? m.symbolChar : m.text
    })
    .join(' ')
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
  if (leaf.type.name === 'docNestedTable') {
    return tableModelText(leaf.attrs.model as TableModel | null)
  }
  return leaf.type.spec.leafText?.(leaf) ?? ''
}

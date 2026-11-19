/**
 * Word-parity text statistics.
 *
 * Word's CJK rule: Words = Asian characters (counted one by one, punctuation
 * included) + non-Asian words (whitespace/punct-delimited). The dialog also
 * reports the two addends separately — Chinese users care about the
 * Asian-character figure.
 */

import { eqLayoutText } from './editor/eq-layout'
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
//
// Word also counts four symbols as Asian characters wherever they stand: the
// two-dot leader, the reference mark, prime and double prime (probe: "a" + one
// of them + "b" is three words).
const ASIAN_RE =
  /^(?:[\u1100-\u11ff\u2e80-\u2fdf\u3001-\u303f\u3040-\u30ff\u3100-\u312f\u3130-\u318f\u31c0-\u31ff\u3400-\u4dbf\u4e00-\u9fff\uac00-\ud7af\uf900-\ufaff\uff01-\uff60\uffe0-\uffe6\u2025\u203b\u2032\u2033]|[\uD840-\uD8BF][\uDC00-\uDFFF])$/

/**
 * Symbols Word counts as Asian characters when they follow one (or follow such
 * a symbol that does): curly quotes, the ellipsis, the middle dot and other
 * East Asian "ambiguous width" marks. After a Latin letter they join its word
 * ("a…b" is one word); after Han each is a character of its own (Han, two
 * ellipses, Han is four). Probed one by one in Word for Mac 16, 2026-11-19;
 * the guillemets, ASCII punctuation and the dashes stay as they are.
 */
const AFTER_ASIAN = new Set([
  ...'\u00b7\u00b0\u00a7\u2122\u00d7\u00b1\u20ac\u2018\u2019\u201c\u201d\u2026\u00b6\u2020\u2030',
  ...'\u00a8\u00b4\u02c6\u02dc\u00a1\u00bf\u2039\u203a',
])

/** zero-width characters: no word of their own, no split ("a" ZWSP "b" is one word) */
const ZERO_WIDTH = new Set(['\u200b', '\ufeff'])

/**
 * Word's non-Asian words (probed against Word for Mac 16, 2026-11-19): any run
 * of characters between whitespace is one word, punctuation and symbols
 * included, so "0.75", "p<0.05", "a/b", "[1]", "&" and a URL each count once.
 * - whitespace: space, tab, line breaks, NBSP and the ideographic space. The
 *   thin, en and other typographic spaces join their neighbours ("a\u2009b" is
 *   one word).
 * - en dash, em dash and the bullet, also as Symbol's private-use U+F0B7,
 *   split words and are not words themselves: "3248–3265" is two words, a
 *   lone "—" none. Other private-use glyphs (Wingdings bullets) are words.
 * - the Unicode hyphen (U+2010) and the arrows split words and are words of
 *   their own: "a→b" is three. The ASCII hyphen joins ("state-of-the-art").
 * Asian characters are counted one by one (asianCharCount) and split the runs
 * around them (a date written 2026, Han "year", 8, Han "month", 2, Han "day"
 * is three non-Asian words and three Asian characters).
 */
// explicit, not \s: JS counts the en/thin/hair spaces as whitespace, Word does not
const WORD_SPACE_RE = /[ \t\n\r\f\v\u00a0\u3000\u2013\u2014\u2022\uf0b7]+/
const OWN_WORD_RE = /([\u2010\u2190-\u21ff])/

/** one pass over the text: the Asian characters, and the rest with them blanked out */
function splitAsian(text: string): { asian: number; rest: string } {
  let asian = 0
  let rest = ''
  let afterAsian = false
  for (const ch of text) {
    if (ZERO_WIDTH.has(ch)) continue
    if (ASIAN_RE.test(ch) || (afterAsian && AFTER_ASIAN.has(ch))) {
      asian++
      rest += ' '
      afterAsian = true
    } else {
      rest += ch
      afterAsian = false
    }
  }
  return { asian, rest }
}

function wordsOf(rest: string): number {
  let n = 0
  for (const run of rest.split(WORD_SPACE_RE)) {
    if (!run) continue
    for (const part of run.split(OWN_WORD_RE)) if (part) n++
  }
  return n
}

export function asianCharCount(text: string): number {
  return splitAsian(text).asian
}

export function nonAsianWordCount(text: string): number {
  return wordsOf(splitAsian(text).rest)
}

/** Word's Words figure: asian chars + non-asian words */
export function countWords(text: string): number {
  const { asian, rest } = splitAsian(text)
  return asian + wordsOf(rest)
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
export function selectedWordCount(
  doc: ProseMirrorNode,
  selection: Selection,
  defs?: Map<string, NumberingDef> | null,
): number {
  // the same leaf reader as the total: an inline formula's text counts in both
  const text = doc.textBetween(selection.from, selection.to, ' ', leafTextForWordCount)
  if (selection.empty) return countWords(text)
  // a list item's label counts when the selection holds the start of its text,
  // so Select All shows the document total, as in Word
  const labels = listLabels(doc, defs).filter(
    (l) => l.pos + 1 >= selection.from && l.pos + 1 < selection.to,
  )
  return countWords(text) + countWords(labels.map((l) => l.text).join(' '))
}

/**
 * The list labels Word counts as words, by position: an item's number ("1.",
 * "a)") and a bullet that is a word under the word rule ("o", a Wingdings
 * square), but not Symbol's U+F0B7 bullet, a "•" or a dash bullet, nor a
 * picture bullet. Word for Mac 16 probe 2026-11-19: three numbered items
 * "alpha", "beta", "gamma" are six words, three under a Symbol bullet, six
 * under "o" or a Wingdings square. The labels come from the numbering engine
 * the editor draws them with, over the same items (the numbered anchor
 * paragraph of a textbox block included).
 */
export function listLabels(
  doc: ProseMirrorNode,
  defs: Map<string, NumberingDef> | null | undefined,
): Array<{ pos: number; text: string }> {
  if (!defs?.size) return []
  // docs and defs are replaced, never mutated, on change: a selection change
  // re-reads the last walk instead of walking the document again
  const hit = labelCache.get(doc)
  if (hit && hit.defs === defs) return hit.labels
  const labels = walkListLabels(doc, defs)
  labelCache.set(doc, { defs, labels })
  return labels
}

const labelCache = new WeakMap<
  ProseMirrorNode,
  { defs: Map<string, NumberingDef>; labels: Array<{ pos: number; text: string }> }
>()

function walkListLabels(
  doc: ProseMirrorNode,
  defs: Map<string, NumberingDef>,
): Array<{ pos: number; text: string }> {
  const refs: ListItemRef[] = []
  const at: number[] = []
  doc.descendants((node, pos) => {
    if (node.type.name === 'docListItem') {
      const deleted =
        !!node.attrs.paraMarkDel ||
        (node.attrs.blockRevision as { kind?: string } | null)?.kind === 'del'
      refs.push({
        numId: (node.attrs.numId as string | null) ?? null,
        ilvl: Number(node.attrs.ilvl) || 0,
        ...(deleted ? { deleted: true } : {}),
      })
      at.push(pos)
      return false
    }
    const stray =
      node.type.name === 'docProtected' ? (node.attrs.strayList as ListItemRef | null) : null
    if (stray) {
      refs.push({ numId: stray.numId, ilvl: stray.ilvl })
      at.push(pos)
      return false
    }
    // a nested table is one node: its list paragraphs live in the model, in
    // reading order, and take their numbers in the document's sequence
    if (node.type.name === 'docNestedTable') {
      for (const list of nestedTableLists(node.attrs.model as TableModel | null)) {
        refs.push(list)
        at.push(pos)
      }
      return false
    }
    return true
  })
  if (refs.length === 0) return []
  const out: Array<{ pos: number; text: string }> = []
  computeListMarkerInfos(refs, defs).forEach((m, i) => {
    if (!m || m.picBulletSrc || refs[i]!.deleted) return
    // the glyph as authored: a symbol-font bullet stays in the private-use area
    const text = m.symbolFont && m.symbolChar ? m.symbolChar : m.text
    if (text) out.push({ pos: at[i]!, text })
  })
  return out
}

/** a nested table's list paragraphs in reading order, deeper tables placed at their anchors */
function nestedTableLists(model: TableModel | null | undefined): ListItemRef[] {
  const out: ListItemRef[] = []
  for (const row of model?.rows ?? []) {
    for (const cell of row) {
      if (cell.vMerge === 'continue') continue
      const paras = cell.richParas ?? []
      const tables = cell.nestedTables ?? []
      const anchors = cell.nestedTableAnchors ?? []
      const tablesAt = (i: number) => tables.filter((_, t) => (anchors[t] ?? paras.length) === i)
      for (let i = 0; i <= paras.length; i++) {
        for (const inner of tablesAt(i)) out.push(...nestedTableLists(inner))
        const list = paras[i]?.list
        if (list) out.push({ numId: list.numId, ilvl: list.ilvl })
      }
    }
  }
  return out
}

/** every list label's text, for the document's word count */
export function listLabelText(
  doc: ProseMirrorNode,
  defs: Map<string, NumberingDef> | null | undefined,
): string {
  return listLabels(doc, defs)
    .map((l) => l.text)
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
  if (leaf.type.name === 'docEqLayout') return eqLayoutText(leaf.attrs)
  if (leaf.type.name === 'docNestedTable') {
    return tableModelText(leaf.attrs.model as TableModel | null)
  }
  return leaf.type.spec.leafText?.(leaf) ?? ''
}

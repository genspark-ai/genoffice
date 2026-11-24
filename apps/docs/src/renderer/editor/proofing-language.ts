import type { Editor } from '@tiptap/core'
import { wordRangeAtCaret } from './comments'

/** the languages Tools > Language offers, as the tags Word writes */
export const PROOFING_LANGUAGES = [
  'en-US',
  'en-GB',
  'zh-CN',
  'zh-TW',
  'ja-JP',
  'ko-KR',
  'fr-FR',
  'de-DE',
  'es-ES',
  'pt-BR',
  'pt-PT',
  'it-IT',
  'nl-NL',
  'pl-PL',
  'cs-CZ',
  'hu-HU',
  'ru-RU',
  'uk-UA',
  'tr-TR',
  'el-GR',
  'sv-SE',
  'da-DK',
  'nb-NO',
  'fi-FI',
  'ar-SA',
  'he-IL',
  'hi-IN',
  'th-TH',
  'vi-VN',
  'id-ID',
  'ms-MY',
] as const

export type LangSlots = { val?: string; eastAsia?: string; bidi?: string }
type Slot = keyof LangSlots

/** the w:lang slot Word files a language under: East Asian, right-to-left, or the rest */
export function languageSlot(tag: string): Slot {
  const base = tag.split('-')[0].toLowerCase()
  if (base === 'zh' || base === 'ja' || base === 'ko') return 'eastAsia'
  return base === 'ar' || base === 'he' || base === 'fa' || base === 'ur' ? 'bidi' : 'val'
}

export function parseLangAttr(value: unknown): LangSlots | null {
  if (typeof value !== 'string' || !value) return null
  try {
    const parsed: unknown = JSON.parse(value)
    return parsed && typeof parsed === 'object' ? (parsed as LangSlots) : null
  } catch {
    return null
  }
}

/** the selection, or the word a bare caret is in: what Word's Language dialog reads and marks */
function languageRange(editor: Editor): { from: number; to: number } | null {
  const { selection } = editor.state
  return selection.empty ? wordRangeAtCaret(editor) : { from: selection.from, to: selection.to }
}

const EAST_ASIAN = /[\p{sc=Han}\p{sc=Hiragana}\p{sc=Katakana}\p{sc=Hangul}]/u
const RIGHT_TO_LEFT = /[\p{sc=Arabic}\p{sc=Hebrew}]/u

/** the w:lang slot that speaks for a piece of text: the one its first letter is written in */
function slotOfText(text: string): Slot {
  for (const ch of text) {
    if (EAST_ASIAN.test(ch)) return 'eastAsia'
    if (RIGHT_TO_LEFT.test(ch)) return 'bidi'
    if (/\p{L}/u.test(ch)) return 'val'
  }
  return 'val'
}

/** the text formatting the dialog reads: a bare caret reads its word, not the run before it */
function textStyleAttrs(editor: Editor): Record<string, unknown> {
  const word = wordRangeAtCaret(editor)
  if (!word) return editor.getAttributes('docTextStyle')
  const { doc, schema } = editor.state
  return schema.marks.docTextStyle.isInSet(doc.nodeAt(word.from)?.marks ?? [])?.attrs ?? {}
}

/** the language the selection, or the word at the caret, is marked as; '' when unmarked */
export function proofingLanguageOf(editor: Editor): string {
  const slots = parseLangAttr(textStyleAttrs(editor).lang)
  if (!slots) return ''
  const range = languageRange(editor)
  const own = range ? slotOfText(editor.state.doc.textBetween(range.from, range.to, ' ')) : 'val'
  return slots[own] ?? slots.val ?? slots.eastAsia ?? slots.bidi ?? ''
}

/** whether the selection is marked as not to be spell or grammar checked (w:noProof) */
export function noProofOf(editor: Editor): boolean {
  return textStyleAttrs(editor).noProof === true
}

/**
 * Marks the selected text as `tag`, in the slot that language belongs to; a bare caret marks
 * the word it is in, or what is typed next. The pick is the text's one language: the other
 * slots go, so the dialog reads back what was picked. `noProof` left out keeps what each run
 * says about proofing.
 */
export function setProofingLanguage(editor: Editor, tag: string, noProof?: boolean): boolean {
  const { state } = editor
  const type = state.schema.marks.docTextStyle
  const marked = (attrs?: Record<string, unknown>) =>
    type.create({
      ...attrs,
      lang: JSON.stringify({ [languageSlot(tag)]: tag }),
      ...(noProof === undefined ? {} : { noProof }),
    })
  const range = languageRange(editor)
  const tr = state.tr
  if (range) {
    // per text node: each keeps the rest of its formatting
    state.doc.nodesBetween(range.from, range.to, (node, pos) => {
      if (!node.isText) return
      const from = Math.max(pos, range.from)
      const to = Math.min(pos + node.nodeSize, range.to)
      tr.addMark(from, to, marked(type.isInSet(node.marks)?.attrs))
    })
    if (!tr.docChanged) return false
  } else {
    const marks = state.storedMarks ?? state.selection.$from.marks()
    tr.setStoredMarks(marked(type.isInSet(marks)?.attrs).addToSet(marks))
  }
  editor.view.dispatch(tr)
  editor.commands.focus()
  return true
}

import type { Editor } from '@tiptap/core'

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

/** the language the selection is marked as, when it is one of the three slots; '' when unmarked */
export function proofingLanguageOf(editor: Editor): string {
  const slots = parseLangAttr(editor.getAttributes('docTextStyle').lang)
  return slots?.val ?? slots?.eastAsia ?? slots?.bidi ?? ''
}

/** whether the selection is marked as not to be spell or grammar checked (w:noProof) */
export function noProofOf(editor: Editor): boolean {
  return editor.getAttributes('docTextStyle').noProof === true
}

/**
 * Marks the selected text as `tag`, in the slot that language belongs to. The pick is the
 * text's one language: the other slots go, so the dialog reads back what was picked.
 * `noProof` left out keeps what each run says about proofing.
 */
export function setProofingLanguage(editor: Editor, tag: string, noProof?: boolean): boolean {
  if (editor.state.selection.empty) return false
  const { state } = editor
  const type = state.schema.marks.docTextStyle
  const slot = languageSlot(tag)
  const tr = state.tr
  // per text node: each keeps the rest of its formatting
  state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
    if (!node.isText) return
    const from = Math.max(pos, state.selection.from)
    const to = Math.min(pos + node.nodeSize, state.selection.to)
    const mark = type.isInSet(node.marks)
    tr.addMark(
      from,
      to,
      type.create({
        ...mark?.attrs,
        lang: JSON.stringify({ [slot]: tag }),
        ...(noProof === undefined ? {} : { noProof }),
      }),
    )
  })
  if (!tr.docChanged) return false
  editor.view.dispatch(tr)
  editor.commands.focus()
  return true
}

import type { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'

export type CaseMode =
  'upper' | 'lower' | 'title' | 'sentence' | 'toggle' | 'halfWidth' | 'fullWidth'

// Half-width katakana (U+FF61-U+FF9F) and their full-width forms, generated
// from NFKC so voiced pairs like U+FF76 U+FF9E compose to one code point.
const HALF_KANA_START = 0xff61
const HALF_KANA_END = 0xff9f
const HALF_TO_FULL_KANA = new Map<string, string>()
const FULL_TO_HALF_KANA = new Map<string, string>()
for (let code = HALF_KANA_START; code <= HALF_KANA_END; code++) {
  const half = String.fromCharCode(code)
  const full = code === 0xff9e ? '\u309b' : code === 0xff9f ? '\u309c' : half.normalize('NFKC')
  HALF_TO_FULL_KANA.set(half, full)
  FULL_TO_HALF_KANA.set(full, half)
  for (const voiced of ['\uff9e', '\uff9f']) {
    const composed = (half + voiced).normalize('NFKC')
    if (composed.length === 1) {
      HALF_TO_FULL_KANA.set(half + voiced, composed)
      FULL_TO_HALF_KANA.set(composed, half + voiced)
    }
  }
}

function toFullWidth(s: string): string {
  let out = ''
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]!
    const code = ch.charCodeAt(0)
    if (code === 0x20) out += '\u3000'
    else if (code >= 0x21 && code <= 0x7e) out += String.fromCharCode(code + 0xfee0)
    else if (code >= HALF_KANA_START && code <= HALF_KANA_END) {
      const pair = HALF_TO_FULL_KANA.get(ch + (s[i + 1] ?? ''))
      if (pair) {
        out += pair
        i++
      } else out += HALF_TO_FULL_KANA.get(ch) ?? ch
    } else out += ch
  }
  return out
}

function toHalfWidth(s: string): string {
  let out = ''
  for (const ch of s) {
    const code = ch.codePointAt(0)!
    if (code === 0x3000) out += ' '
    else if (code >= 0xff01 && code <= 0xff5e) out += String.fromCharCode(code - 0xfee0)
    else out += FULL_TO_HALF_KANA.get(ch) ?? ch
  }
  return out
}

function toggleCase(s: string): string {
  let out = ''
  for (const ch of s) {
    const lower = ch.toLowerCase()
    out += ch === lower ? ch.toUpperCase() : lower
  }
  return out
}

export function transformCase(s: string, mode: CaseMode): string {
  switch (mode) {
    case 'upper':
      return s.toUpperCase()
    case 'lower':
      return s.toLowerCase()
    case 'title':
      return s.toLowerCase().replace(/(^|\s)(\p{L})/gu, (m) => m.toUpperCase())
    case 'sentence':
      return s.toLowerCase().replace(/(^\s*\p{L})|([.!?。!?]\s*\p{L})/gu, (m) => m.toUpperCase())
    case 'toggle':
      return toggleCase(s)
    case 'halfWidth':
      return toHalfWidth(s)
    case 'fullWidth':
      return toFullWidth(s)
  }
}

/**
 * Word's Shift+F3 cycle: lowercase → UPPERCASE → Capitalize Each Word. Which
 * step comes next is read off the selection, so repeated presses walk the ring
 * (mixed-case text enters it at lowercase, like Word).
 */
export function nextCaseMode(text: string): CaseMode {
  const letters = text.replace(/\P{L}/gu, '')
  if (!letters) return 'lower'
  if (letters === letters.toLowerCase()) return 'upper'
  if (letters === letters.toUpperCase()) return 'title'
  return 'lower'
}

/** rewrite every text run in the selection, keeping its marks */
export function applyCase(editor: Editor, mode: CaseMode): boolean {
  const { from, to } = editor.state.selection
  if (from === to) return false
  return editor
    .chain()
    .focus()
    .command(({ state, tr }) => {
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (!node.isText || !node.text) return
        const start = Math.max(from, pos)
        const end = Math.min(to, pos + node.nodeSize)
        const slice = node.text.slice(start - pos, end - pos)
        const next = transformCase(slice, mode)
        if (next !== slice) {
          tr.replaceWith(
            tr.mapping.map(start),
            tr.mapping.map(end),
            state.schema.text(next, node.marks),
          )
        }
      })
      // Restore the range: without this, replacing the leading run pushes the
      // anchor to its end and the next Shift+F3 only sees the tail of the
      // phrase. Every rewrite happened inside the selection, so the document's
      // growth is the selection's growth (ß → SS and friends lengthen it).
      if (tr.docChanged) {
        const grew = tr.doc.content.size - state.doc.content.size
        tr.setSelection(TextSelection.create(tr.doc, from, to + grew))
      }
      return true
    })
    .run()
}

/** the selection's text, used to decide where the Shift+F3 ring resumes */
export function selectionText(editor: Editor): string {
  const { from, to } = editor.state.selection
  return editor.state.doc.textBetween(from, to, '\n', '\n')
}
